"""
SAP Commerce (xjab2b, P1) -> local snapshot -> diff -> Firebase Realtime Database (+ Excel)

Run on the planning PC (Python 3.9+), e.g. once a day via Windows Task Scheduler:
    py commerce_sync.py                 normal run (uses commerce_sync_config.json next to this file)
    py commerce_sync.py --dry-run       fetch + diff + Excel, but do NOT write Firebase or save the snapshot
    py commerce_sync.py --push-all      push the complete dataset to Firebase (first time / after a reset)
    py commerce_sync.py --full-config   re-fetch the CPQ configuration of every vehicle
    py commerce_sync.py --export-local  also write the Firebase tree as local JSON files (output/site/) for offline viewing

Pip packages: openpyxl (Excel). google-auth + requests only if you use a service-account key for Firebase.

Running on a server without a disk (e.g. a Render cron job)
    Every config key can be set with an environment variable CS_<KEY>, e.g. CS_FIREBASE_ENABLED=true,
    CS_STATE_BACKEND=firebase, CS_EXCEL_ENABLED=false. The service-account key can be given as the whole JSON
    text in FIREBASE_SERVICE_ACCOUNT_JSON. With state_backend=firebase the diff snapshot is stored gzip-compressed
    in Firebase under /<firebase_root>_state (keep that node unreadable in the database rules).

What one vehicle is
    All orders that carry the same chassis number (carFrameNumber) form one vehicle: the dealer Sales_Order,
    the factory Purchase_Order, and any reallocation orders. Orders without a chassis number are grouped by guid.

Firebase tree (under firebase_root, default /commerce)
    meta                 last run time, counts
    index/{key}          one compact row per vehicle (what the list pages read, ~2 MB in total)
    vehicles/{key}       full order history + links of one vehicle (read on demand)
    config/{key}         CPQ configuration of one vehicle (read on demand)
    configStats          option take rates by model (read by the config-analysis page)
    changes/{runId}/{key}  what changed in each run (kept for keep_runs runs)
Only vehicles that changed since the last run are written.
"""
import argparse
import base64
import collections
import datetime as dt
import gzip
import hashlib
import json
import os
import re
import shutil
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

HERE = Path(__file__).resolve().parent
DEFAULT_CONFIG = {
    "sap_base": "https://api.cbxbvu474n-zhejiangd1-p1-public.model-t.ccv2prod.sapcloud.cn",
    "site": "xjab2b",
    "client_id": "trusted_client",
    "client_secret": "secret",
    "admin_user": "admin",
    "firebase_enabled": False,
    "firebase_url": "https://saleslistforplanning-default-rtdb.asia-southeast1.firebasedatabase.app",
    "firebase_root": "commerce",
    "firebase_service_account": "",
    "firebase_db_secret": "",
    "state_dir": "state",
    "output_dir": "output",
    "threads": 12,
    "config_threads": 20,
    "retry_failed_config_per_run": 100,
    "keep_runs": 60,
    "keep_excel": 14,
    "excel_config_sheet": True,
    "excel_enabled": True,
    "state_backend": "local",
}
LIST_FIELDS = ("code,guid,status,placed,total(value,currencyIso),updatedTime,erpSONumber,dealerCode,dealerName,"
               "docType,orderType,quoteReference,customer,goodsReceiverName,goodsReceiverAddress,payerName,"
               "expectedDeliveryDate,description,carFrameNumber,floorPlan,specSheet")
TZ = dt.timezone(dt.timedelta(hours=10))  # display only; Firebase stores ISO UTC strings
STATE_ABBR = {"queensland": "QLD", "victoria": "VIC", "western australia": "WA", "new south wales": "NSW",
              "south australia": "SA", "tasmania": "TAS", "australian capital territory": "ACT",
              "northern territory": "NT"}
NZ_REGIONS = {"northland", "auckland", "canterbury", "southland", "tasman", "waikato", "otago", "wellington",
              "bay of plenty", "hawke's bay", "manawatu-wanganui", "nelson", "marlborough", "taranaki", "gisborne",
              "west coast"}


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


# ----------------------------------------------------------------------------------------------- SAP client
class Sap:
    def __init__(self, cfg):
        self.base = cfg["sap_base"].rstrip("/")
        self.ws = f"{self.base}/xjawebservices/v2/{cfg['site']}"
        self.cred = base64.b64encode(f"{cfg['client_id']}:{cfg['client_secret']}".encode()).decode()
        self._tok, self._t = None, 0
        self._lock = threading.Lock()
        self.calls = 0

    def token(self, force=False):
        with self._lock:
            if self._tok and not force and time.time() - self._t < 200:
                return self._tok
            req = Request(self.base + "/authorizationserver/oauth/token",
                          data=urlencode({"grant_type": "client_credentials", "scope": "extended"}).encode(),
                          headers={"Authorization": "Basic " + self.cred,
                                   "Content-Type": "application/x-www-form-urlencoded"})
            with urlopen(req, timeout=60) as r:
                self._tok = json.load(r)["access_token"]
            self._t = time.time()
            return self._tok

    def get(self, path):
        """Returns (http_status, json). Never raises for HTTP errors; retries network errors."""
        for attempt in range(4):
            self.calls += 1
            req = Request(self.ws + path, headers={"Authorization": "Bearer " + self.token(),
                                                   "Accept": "application/json"})
            try:
                with urlopen(req, timeout=120) as r:
                    return 200, json.load(r)
            except HTTPError as e:
                body = e.read().decode("utf-8", "replace")
                if e.code == 401 and attempt < 3 and ("InvalidTokenError" in body or "expired" in body.lower()):
                    self.token(force=True)
                    continue
                try:
                    return e.code, json.loads(body)
                except ValueError:
                    return e.code, {"errors": [{"type": f"HTTP{e.code}", "message": body[:200]}]}
            except (URLError, TimeoutError, ConnectionError, OSError, ValueError) as e:
                if attempt == 3:
                    return 0, {"errors": [{"type": "NetworkError", "message": str(e)[:200]}]}
                time.sleep(3 * (attempt + 1))

    def users(self, admin):
        out, page = [], 0
        while True:
            c, d = self.get(f"/users/{quote(admin, safe='')}/orgCustomers?fields=FULL&pageSize=100&currentPage={page}")
            if c != 200:
                raise RuntimeError(f"Cannot list B2B users ({c}): {json.dumps(d)[:300]}")
            out += d.get("users", [])
            page += 1
            if page >= d.get("pagination", {}).get("totalPages", 0):
                return out

    def user_orders(self, uid, threads=4):
        """All orders of one account. Pages that SAP cannot render (NullPointerError) are retried in
        smaller pages so that a single broken order does not hide its neighbours."""
        u = quote(uid, safe="")

        def page(size, p):
            return self.get(f"/users/{u}/orders?fields=orders({LIST_FIELDS}),pagination&pageSize={size}&currentPage={p}")

        c, d = page(100, 0)
        if c == 200:
            total, first = d["pagination"]["totalResults"], d.get("orders", [])
        else:
            c1, d1 = self.get(f"/users/{u}/orders?fields=orders(code),pagination&pageSize=1&currentPage=0")
            if c1 != 200:
                return [], ["account unreadable"]
            total, first = d1["pagination"]["totalResults"], None
        if total == 0:
            return [], []

        def block(p):
            if p == 0 and first is not None:
                return first, []
            c, d = page(100, p)
            if c == 200:
                return d.get("orders", []), []
            res, bad = [], []
            for q in range(p * 5, p * 5 + 5):
                if q * 20 >= total:
                    break
                c, d = page(20, q)
                if c == 200:
                    res += d.get("orders", [])
                    continue
                for i in range(q * 20, min(q * 20 + 20, total)):
                    c, d = page(1, i)
                    if c == 200:
                        res += d.get("orders", [])
                    else:
                        bad.append(i)
            return res, bad

        out, bad = [], []
        with ThreadPoolExecutor(threads) as ex:
            for o, b in ex.map(block, range((total + 99) // 100)):
                out += o
                bad += b
        return out, bad

    def config(self, uid, code):
        return self.get(f"/users/{quote(uid, safe='')}/orders/{quote(code, safe='')}/entries/0/configurator/configurationOverview")

    def config_url(self, uid, code):
        return f"{self.ws}/users/{quote(uid, safe='')}/orders/{code}/entries/0/configurator/configurationOverview"


# ----------------------------------------------------------------------------------------------- shaping
def fb_key(s):
    return re.sub(r"[.#$\[\]/\s]", "_", str(s))[:200]


def addr_state(addr):
    m = re.search(r",\s*([A-Za-z' -]+),\s*\d{4}", addr or "")
    if not m:
        return None
    s = m.group(1).strip().lower()
    if s in STATE_ABBR:
        return STATE_ABBR[s]
    if s in NZ_REGIONS:
        return "NZ"
    return m.group(1).strip()


def model_from(o):
    for k in ("floorPlan", "specSheet"):
        u = o.get(k) or ""
        m = re.search(r"/medias/(?:[A-Z0-9]{3,}\d+-)?(20\d\d-[A-Z]{2,4}\d{2}[A-Z0-9.]*)", u)
        if m:
            return m.group(1).rstrip(".")
    return None


def slim_order(uid, o):
    t = o.get("total") or {}
    erp = o.get("erpSONumber")
    exp = o.get("expectedDeliveryDate")
    return {
        "code": o["code"], "type": "SO" if o.get("orderType") == "Sales_Order" else ("PO" if o.get("orderType") == "Purchase_Order" else o.get("orderType")),
        "status": o.get("status"), "placed": o.get("placed"), "updated": o.get("updatedTime"),
        "acct": uid, "dealer": o.get("dealerName"), "dealerCode": o.get("dealerCode"),
        "customer": o.get("customer"), "receiver": o.get("goodsReceiverName"),
        "address": (o.get("goodsReceiverAddress") or "").strip(", ") or None, "payer": o.get("payerName"),
        "total": t.get("value"), "currency": t.get("currencyIso"),
        "erp": erp if erp and erp.strip("0") else None, "docType": o.get("docType"),
        "quote": o.get("quoteReference"), "desc": o.get("description"),
        "expected": exp if exp and not exp.startswith("2050") else None,
        "floorPlan": o.get("floorPlan"), "specSheet": o.get("specSheet"), "guid": o.get("guid"),
    }


def build_vehicles(orders):
    """orders: list of (uid, raw order). Returns {key: vehicle_record}."""
    by_chassis = collections.defaultdict(list)
    no_chassis = []
    for uid, o in orders:
        ch = (o.get("carFrameNumber") or "").strip()
        (by_chassis[ch] if ch else no_chassis).append((uid, o))
    groups = {fb_key(ch): (ch, v) for ch, v in by_chassis.items()}
    g2 = collections.defaultdict(list)
    for uid, o in no_chassis:
        g2[o.get("guid") or o["code"]].append((uid, o))
    for gid, v in g2.items():
        groups[fb_key("G_" + gid)] = (None, v)

    vehicles = {}
    for key, (chassis, v) in groups.items():
        ords = sorted((slim_order(u, o) for u, o in v), key=lambda x: (x["placed"] or "", x["code"]))
        sos = [x for x in ords if x["type"] == "SO"]
        pos = [x for x in ords if x["type"] == "PO"]
        live = lambda xs: [x for x in xs if x["status"] != "CANCELLED"]
        main_so = (live(sos) or sos or [None])[-1]
        main_po = (live(pos) or pos or [None])[-1]
        main = main_so or main_po
        model = next((m for m in (model_from({"floorPlan": x["floorPlan"], "specSheet": x["specSheet"]})
                                  for x in reversed(ords)) if m), None)
        first_so = sos[0] if sos else None
        vehicles[key] = {
            "key": key, "chassis": chassis,
            "so": main_so["code"] if main_so else None, "po": main_po["code"] if main_po else None,
            "soStatus": main_so["status"] if main_so else None, "poStatus": main_po["status"] if main_po else None,
            "active": bool(live(sos) or (not sos and live(pos))),
            "dealer": main["dealer"], "dealerCode": main["dealerCode"], "state": addr_state(main["address"]),
            "customer": main["customer"], "acct": main["acct"], "erp": next((x["erp"] for x in reversed(ords) if x["erp"]), None),
            "docType": main["docType"], "model": model,
            "firstPlaced": (first_so or main)["placed"], "firstDealer": (first_so or main)["dealer"],
            "placed": main["placed"], "updated": max((x["updated"] or "") for x in ords) or None,
            "soTotal": main_so["total"] if main_so else None, "poTotal": main_po["total"] if main_po else None,
            "realloc": sum(1 for x in sos if x["acct"].startswith("realloc")) + (1 if len(sos) > 1 else 0) > 0,
            "floorPlan": next((x["floorPlan"] for x in reversed(ords) if x["floorPlan"]), None),
            "specSheet": next((x["specSheet"] for x in reversed(ords) if x["specSheet"]), None),
            "orders": [{k: x for k, x in o.items() if k not in ("floorPlan", "specSheet")} for o in ords],
        }
    return vehicles


def config_sources(v):
    """Orders to try for the CPQ configuration, best first: live SO, live PO, then the rest (newest first)."""
    ords = list(reversed(v["orders"]))
    rank = lambda x: (x["status"] == "CANCELLED", x["type"] != "SO", x["type"] != "PO")
    return sorted(ords, key=rank)


def compact_config(r):
    groups = collections.OrderedDict()
    for g in r.get("groups", []):
        gd = g.get("groupDescription") or g.get("id") or "Other"
        for it in g.get("characteristicValues", []):
            ch = it.get("characteristic") or it.get("characteristicId")
            groups.setdefault(gd, collections.OrderedDict())
            prev = groups[gd].get(ch)
            val = str(it.get("value"))
            groups[gd][ch] = val if prev is None else f"{prev}; {val}"
    return {"productCode": r.get("productCode"), "groups": groups}


def fb_config(cfg):
    """Firebase keys may not contain . $ # [ ] / (e.g. the option "H.W.S."), so option names are stored as values:
    {"productCode": ..., "groups": [{"g": group, "c": [[option, value], ...]}, ...]}"""
    if not cfg:
        return None
    return {"productCode": cfg.get("productCode"),
            "groups": [{"g": g, "c": [[ch, val] for ch, val in chars.items()]} for g, chars in cfg["groups"].items()]}


BAD_KEY = re.compile(r"[.$#\[\]/\x00-\x1f\x7f]")


def check_keys(obj, path=""):
    """Raise before sending anything Firebase would reject, naming the offending key."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if not isinstance(k, str) or k == "" or BAD_KEY.search(k) or len(k.encode()) > 768:
                raise RuntimeError(f"invalid Firebase key {k!r} at {path or '/'}")
            check_keys(v, f"{path}/{k}")
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            check_keys(v, f"{path}/{i}")


def index_row(v, has_cfg, product_code):
    """Compact list row. Order of fields is fixed and documented in INDEX_FIELDS (also used by the web page)."""
    return [v["chassis"], v["so"], v["po"], v["soStatus"], v["poStatus"], 1 if v["active"] else 0,
            v["dealer"], v["dealerCode"], v["state"], v["customer"], v["acct"], v["erp"], v["docType"],
            v["model"], product_code, v["firstPlaced"], v["firstDealer"], v["placed"], v["updated"],
            v["soTotal"], v["poTotal"], 1 if v["realloc"] else 0, 1 if has_cfg else 0, len(v["orders"]),
            1 if v["floorPlan"] else 0, 1 if v["specSheet"] else 0]


INDEX_FIELDS = ["chassis", "so", "po", "soStatus", "poStatus", "active", "dealer", "dealerCode", "state", "customer",
                "acct", "erp", "docType", "model", "productCode", "firstPlaced", "firstDealer", "placed", "updated",
                "soTotal", "poTotal", "realloc", "hasConfig", "nOrders", "hasFloorPlan", "hasSpecSheet"]


def fingerprint(obj):
    return hashlib.sha1(json.dumps(obj, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()


def config_stats(snapshot):
    """Option take rates over active vehicles that have a configuration."""
    stats = {}
    models = collections.Counter()
    for key, s in snapshot.items():
        if not s.get("cfg") or not s["vehicle"]["active"]:
            continue
        model = s["vehicle"]["model"] or s["cfg"].get("productCode") or "未知"
        models[model] += 1
        for g, chars in s["cfg"]["groups"].items():
            for ch, val in chars.items():
                e = stats.setdefault(ch, {"group": g, "n": 0, "values": {}})
                e["n"] += 1
                vv = e["values"].setdefault(val, {"n": 0, "m": {}})
                vv["n"] += 1
                vv["m"][model] = vv["m"].get(model, 0) + 1
    # Firebase keys cannot contain . $ # [ ] / -> encode characteristic/value/model names into lists
    out = []
    for ch, e in sorted(stats.items(), key=lambda x: (x[1]["group"], -x[1]["n"])):
        vals = sorted(e["values"].items(), key=lambda x: -x[1]["n"])
        out.append({"c": ch, "g": e["group"], "n": e["n"],
                    "v": [{"v": val, "n": d["n"], "m": [[m, c] for m, c in sorted(d["m"].items(), key=lambda x: -x[1])]}
                          for val, d in vals]})
    return {"vehicles": sum(models.values()), "models": [[m, c] for m, c in models.most_common()], "chars": out}


# ----------------------------------------------------------------------------------------------- Firebase
class Firebase:
    def __init__(self, cfg):
        self.url = cfg["firebase_url"].rstrip("/")
        self.root = cfg["firebase_root"].strip("/")
        self.sa = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON") or cfg.get("firebase_service_account")
        self.secret = cfg.get("firebase_db_secret")
        self._creds = None
        self.bytes = 0

    def _auth(self):
        if self.sa:
            if self._creds is None:
                from google.oauth2 import service_account  # pip install google-auth requests
                scopes = ["https://www.googleapis.com/auth/firebase.database",
                          "https://www.googleapis.com/auth/userinfo.email"]
                if self.sa.lstrip().startswith("{"):
                    self._creds = service_account.Credentials.from_service_account_info(json.loads(self.sa), scopes=scopes)
                else:
                    sa_path = self.sa if os.path.isabs(self.sa) else str(HERE / self.sa)
                    self._creds = service_account.Credentials.from_service_account_file(sa_path, scopes=scopes)
            if not self._creds.valid:
                import google.auth.transport.requests
                self._creds.refresh(google.auth.transport.requests.Request())
            return "access_token=" + self._creds.token
        if self.secret:
            return "auth=" + self.secret
        return ""

    def _req(self, method, path, body=None, query="", root=None):
        url = f"{self.url}/{self.root if root is None else root}/{path}.json".replace("/.json", ".json")
        params = [x for x in (self._auth(), query) if x]
        if params:
            url += "?" + "&".join(params)
        data = None
        headers = {"Content-Type": "application/json"}
        if body is not None:
            data = json.dumps(body, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
            self.bytes += len(data)
        for attempt in range(4):
            try:
                with urlopen(Request(url, data=data, method=method, headers=headers), timeout=180) as r:
                    return json.load(r)
            except HTTPError as e:
                msg = e.read().decode("utf-8", "replace")[:300]
                if e.code in (401, 403):
                    raise RuntimeError(f"Firebase refused the {method} ({e.code}): {msg}. Check the service account / "
                                       "database secret in commerce_sync_config.json and the database rules.")
                if attempt == 3:
                    raise RuntimeError(f"Firebase {method} {path} failed ({e.code}): {msg}")
            except (URLError, TimeoutError, OSError) as e:
                if attempt == 3:
                    raise RuntimeError(f"Firebase {method} {path} failed: {e}")
            time.sleep(3 * (attempt + 1))

    STATE_CHUNK = 4_000_000  # Realtime Database strings are limited to 10 MB

    def load_state(self):
        root = self.root + "_state"
        meta = self._req("GET", "meta", root=root)
        if not meta:
            return None
        parts = [self._req("GET", f"chunks/{i}", root=root) or "" for i in range(meta["chunks"])]
        blob = base64.b64decode("".join(parts))
        if hashlib.sha1(blob).hexdigest() != meta.get("sha1"):
            raise RuntimeError("Snapshot in Firebase is incomplete (checksum mismatch) - run with --push-all after deleting "
                               f"/{root} in the console, or restore it.")
        return json.loads(gzip.decompress(blob).decode("utf-8"))

    def save_state(self, obj):
        root = self.root + "_state"
        blob = gzip.compress(json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), 6)
        b64 = base64.b64encode(blob).decode()
        chunks = [b64[i:i + self.STATE_CHUNK] for i in range(0, len(b64), self.STATE_CHUNK)]
        old = self._req("GET", "meta", root=root) or {}
        for i, c in enumerate(chunks):
            self._req("PUT", f"chunks/{i}", c, root=root)
        for i in range(len(chunks), old.get("chunks", 0)):
            self._req("DELETE", f"chunks/{i}", root=root)
        self._req("PUT", "meta", {"chunks": len(chunks), "bytes": len(blob), "sha1": hashlib.sha1(blob).hexdigest(),
                                  "saved": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")}, root=root)
        return len(blob)

    def patch_many(self, updates, chunk_bytes=4_000_000):
        """Multi-path update, split into requests of a few MB."""
        for k, v in updates.items():
            for part in k.split("/"):
                if not part or BAD_KEY.search(part):
                    raise RuntimeError(f"invalid Firebase path {k!r}")
            check_keys(v, k)
        batch, size = {}, 0
        for k, v in updates.items():
            s = len(json.dumps(v, ensure_ascii=False, default=str))
            if batch and size + s > chunk_bytes:
                self._req("PATCH", "", batch)
                batch, size = {}, 0
            batch[k] = v
            size += s
        if batch:
            self._req("PATCH", "", batch)


# ----------------------------------------------------------------------------------------------- Excel
def write_excel(path, snapshot, changes, run_id, with_config):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    F = Font(name="Arial", size=10)
    HF = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    LF = Font(name="Arial", size=10, color="0563C1", underline="single")
    FILL = PatternFill("solid", fgColor="1F3864")

    def to_dt(s):
        if not s:
            return None
        try:
            return dt.datetime.strptime(s, "%Y-%m-%dT%H:%M:%S%z").astimezone(dt.timezone(dt.timedelta(hours=11))).replace(tzinfo=None)
        except ValueError:
            return s

    def head(ws, widths):
        for c in ws[1]:
            c.font, c.fill = HF, FILL
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions

    wb = Workbook(write_only=False)
    ws = wb.active
    ws.title = "车辆"
    H = ["车架号", "销售单", "采购单", "销售单状态", "采购单状态", "在用", "经销商", "州", "客户", "下单账号", "ERP销售单号",
         "单据类型", "车型", "车型编码(CPQ)", "首张销售单时间", "首张销售单经销商", "当前销售单时间", "最后更新", "销售单金额", "采购单金额",
         "经历调拨", "有配置", "订单数", "Floor plan", "Spec sheet", "配置单URL"]
    ws.append(H)
    rows = sorted(snapshot.values(), key=lambda s: s["vehicle"]["placed"] or "", reverse=True)
    for s in rows:
        v = s["vehicle"]
        ws.append([v["chassis"], v["so"], v["po"], v["soStatus"], v["poStatus"], "是" if v["active"] else "否", v["dealer"],
                   v["state"], v["customer"], v["acct"], v["erp"], v["docType"], v["model"],
                   (s.get("cfg") or {}).get("productCode"), to_dt(v["firstPlaced"]), v["firstDealer"], to_dt(v["placed"]),
                   to_dt(v["updated"]), v["soTotal"], v["poTotal"], "是" if v["realloc"] else "", "是" if s.get("cfg") else "",
                   len(v["orders"]), v["floorPlan"], v["specSheet"], s.get("cfgUrl")])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = F
        for i in (14, 16, 17):
            row[i].number_format = "yyyy-mm-dd hh:mm"
        for i in (18, 19):
            row[i].number_format = "$#,##0.00"
        for i, lbl in ((23, "floor plan"), (24, "spec sheet"), (25, "配置单")):
            if row[i].value:
                row[i].hyperlink, row[i].value, row[i].font = row[i].value, lbl, LF
    head(ws, [12, 10, 10, 11, 11, 6, 28, 6, 24, 30, 12, 8, 14, 12, 16, 26, 16, 16, 13, 13, 8, 7, 7, 11, 11, 9])

    wc = wb.create_sheet("本次变化")
    wc.append(["类型", "车架号", "销售单", "经销商", "车型", "变化字段"])
    for k, ch in sorted(changes.items(), key=lambda x: x[1]["t"]):
        wc.append([{"new": "新增", "mod": "修改", "del": "消失"}[ch["t"]], ch.get("c"), ch.get("so"), ch.get("d"), ch.get("m"),
                   ", ".join(ch.get("f", []))])
    for row in wc.iter_rows(min_row=2):
        for c in row:
            c.font = F
    head(wc, [8, 12, 10, 28, 14, 80])

    if with_config:
        chars = collections.OrderedDict()
        for s in rows:
            for g, cs in (s.get("cfg") or {}).get("groups", {}).items():
                for ch in cs:
                    chars.setdefault((g, ch), 0)
                    chars[(g, ch)] += 1
        cols = sorted(chars, key=lambda x: (x[0], -chars[x]))
        wcf = wb.create_sheet("配置")
        wcf.append(["车架号", "销售单", "经销商", "车型", "车型编码"] + [c[1] for c in cols])
        for s in rows:
            if not s.get("cfg"):
                continue
            v, g = s["vehicle"], s["cfg"]["groups"]
            wcf.append([v["chassis"], v["so"], v["dealer"], v["model"], s["cfg"].get("productCode")] +
                       [g.get(c[0], {}).get(c[1]) for c in cols])
        head(wcf, [12, 10, 28, 14, 12] + [22] * len(cols))

    wn = wb.create_sheet("说明")
    for r in [["运行", run_id], ["车辆数", len(snapshot)], ["本次新增", sum(1 for c in changes.values() if c["t"] == "new")],
              ["本次修改", sum(1 for c in changes.values() if c["t"] == "mod")], ["本次消失", sum(1 for c in changes.values() if c["t"] == "del")],
              ["口径", "一台车 = 同一车架号的全部订单（销售单、采购单、调拨单）。当前销售单 = 最新一张未取消的销售单。"]]:
        wn.append(r)
    wn.column_dimensions["A"].width = 12
    wn.column_dimensions["B"].width = 100
    wb.save(path)


# ----------------------------------------------------------------------------------------------- local mirror
def export_local(out_dir, snapshot, meta, stats, runs_changes):
    site = out_dir / "site"
    if site.exists():
        shutil.rmtree(site)
    (site / "vehicles").mkdir(parents=True)
    (site / "config").mkdir()
    w = lambda p, o: p.write_text(json.dumps(o, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    w(site / "meta.json", meta)
    w(site / "index.json", {k: s["idx"] for k, s in snapshot.items()})
    w(site / "configStats.json", stats)
    w(site / "changes.json", runs_changes)
    for k, s in snapshot.items():
        w(site / "vehicles" / f"{k}.json", s["vehicle"])
        if s.get("cfg"):
            w(site / "config" / f"{k}.json", fb_config(s["cfg"]))
    log("local mirror written to", site)


# ----------------------------------------------------------------------------------------------- main
def load_config():
    p = HERE / "commerce_sync_config.json"
    cfg = dict(DEFAULT_CONFIG)
    if p.exists():
        cfg.update(json.loads(p.read_text(encoding="utf-8")))
    elif not any(k.startswith("CS_") for k in os.environ):
        p.write_text(json.dumps(DEFAULT_CONFIG, indent=2, ensure_ascii=False), encoding="utf-8")
        log("created", p.name, "- edit it, then run again if needed")
    for k, default in DEFAULT_CONFIG.items():  # environment variables win: CS_FIREBASE_ENABLED=true etc.
        v = os.environ.get("CS_" + k.upper())
        if v is None:
            continue
        if isinstance(default, bool):
            cfg[k] = v.strip().lower() in ("1", "true", "yes", "on")
        elif isinstance(default, int):
            cfg[k] = int(v)
        else:
            cfg[k] = v
    if cfg["state_backend"] == "firebase" and not cfg["firebase_enabled"]:
        raise RuntimeError("state_backend=firebase needs firebase_enabled=true")
    return cfg


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--push-all", action="store_true")
    ap.add_argument("--full-config", action="store_true")
    ap.add_argument("--export-local", action="store_true")
    ap.add_argument("--no-firebase", action="store_true")
    ap.add_argument("--users", help="comma-separated account ids to limit the run (testing only)")
    args = ap.parse_args()

    cfg = load_config()
    state_dir = HERE / cfg["state_dir"]
    out_dir = HERE / cfg["output_dir"]
    state_dir.mkdir(exist_ok=True)
    out_dir.mkdir(exist_ok=True)
    snap_path = state_dir / "snapshot.json.gz"
    run_at = dt.datetime.now(dt.timezone.utc)
    run_id = run_at.astimezone(dt.timezone(dt.timedelta(hours=11))).strftime("%Y-%m-%d_%H%M")
    t0 = time.time()

    old = {}
    old_meta = {}
    fb_state = Firebase(cfg) if cfg["state_backend"] == "firebase" else None
    saved = None
    if fb_state:
        saved = fb_state.load_state()
    elif snap_path.exists():
        with gzip.open(snap_path, "rt", encoding="utf-8") as f:
            saved = json.load(f)
    if saved:
        old, old_meta = saved["vehicles"], saved.get("meta", {})
    log(f"previous snapshot: {len(old)} vehicles")

    # 1. orders --------------------------------------------------------------------------------------------
    sap = Sap(cfg)
    sap.token()
    users = sap.users(cfg["admin_user"])
    if args.users:
        keep = set(args.users.split(","))
        users = [u for u in users if u["uid"] in keep]
    log(f"{len(users)} accounts")
    orders, unreadable = [], {}
    lock = threading.Lock()
    done = [0]

    def pull(u):
        o, bad = sap.user_orders(u["uid"])
        with lock:
            orders.extend((u["uid"], x) for x in o)
            if bad:
                unreadable[u["uid"]] = len(bad)
            done[0] += 1
            if done[0] % 25 == 0:
                log(f"  accounts {done[0]}/{len(users)}, orders {len(orders)}")

    with ThreadPoolExecutor(max(2, cfg["threads"] // 3)) as ex:
        list(ex.map(pull, users))
    log(f"orders: {len(orders)} (unreadable: {sum(unreadable.values())}), SAP calls so far {sap.calls}")
    if not orders:
        raise RuntimeError("No orders returned - aborting so that the snapshot is not wiped.")
    if old and len(orders) < 0.8 * sum(len(s["vehicle"]["orders"]) for s in old.values()) and not args.users:
        raise RuntimeError("Order count dropped by more than 20% versus the last run - aborting as a safety check.")

    # 2. vehicles ------------------------------------------------------------------------------------------
    vehicles = build_vehicles(orders)
    log(f"vehicles: {len(vehicles)}")

    # 3. configurations: fetch for new vehicles, vehicles whose order set/updatedTime changed, and some failed ones
    need, retry = [], []
    for k, v in vehicles.items():
        prev = old.get(k)
        if args.full_config or not prev:
            need.append(k)
        elif fingerprint(v["orders"]) != prev.get("ordersFp"):
            need.append(k)
        elif not prev.get("cfg"):
            retry.append((prev.get("cfgTried") or "", k))
    retry = [k for _, k in sorted(retry)][: cfg["retry_failed_config_per_run"]]
    todo = need + retry
    log(f"configs to fetch: {len(need)} new/changed + {len(retry)} retries")

    new_cfg = {}
    done[0] = 0

    def fetch(k):
        v = vehicles[k]
        last = None
        for o in config_sources(v)[:3]:
            c, r = sap.config(o["acct"], o["code"])
            if c == 200 and r.get("groups"):
                res = {"cfg": compact_config(r), "cfgFrom": o["code"], "cfgUrl": sap.config_url(o["acct"], o["code"])}
                break
            last = (c, ((r.get("errors") or [{}])[0].get("type")) if isinstance(r, dict) else None)
        else:
            res = {"cfg": None, "cfgErr": last, "cfgUrl": sap.config_url(v["orders"][-1]["acct"], v["orders"][-1]["code"])}
        res["cfgTried"] = run_id
        with lock:
            new_cfg[k] = res
            done[0] += 1
            if done[0] % 250 == 0:
                log(f"  configs {done[0]}/{len(todo)}")

    with ThreadPoolExecutor(cfg["config_threads"]) as ex:
        list(ex.map(fetch, todo))
    log(f"configs fetched: {sum(1 for r in new_cfg.values() if r['cfg'])}/{len(todo)} ok, SAP calls {sap.calls}")

    # 4. new snapshot + diff -------------------------------------------------------------------------------
    snapshot, changes = {}, {}
    for k, v in vehicles.items():
        prev = old.get(k, {})
        if k in new_cfg:
            c = new_cfg[k]
            if not c["cfg"] and prev.get("cfg") and k not in need:
                c = {**c, "cfg": prev["cfg"], "cfgFrom": prev.get("cfgFrom"), "cfgUrl": prev.get("cfgUrl")}
        else:
            c = {x: prev.get(x) for x in ("cfg", "cfgFrom", "cfgUrl", "cfgErr", "cfgTried")}
        snapshot[k] = {"vehicle": v, "ordersFp": fingerprint(v["orders"]), **c}
    # vehicles without floor plan/spec file name: take the model name used by other vehicles with the same CPQ product code
    pc_model = collections.defaultdict(collections.Counter)
    for s in snapshot.values():
        pc = (s.get("cfg") or {}).get("productCode")
        if pc and s["vehicle"]["model"]:
            pc_model[pc][s["vehicle"]["model"]] += 1
    for k, s in snapshot.items():
        v, prev = s["vehicle"], old.get(k, {})
        pc = (s.get("cfg") or {}).get("productCode")
        if not v["model"] and pc in pc_model:
            v["model"] = pc_model[pc].most_common(1)[0][0]
        s["idx"] = index_row(v, bool(s.get("cfg")), pc)
        c = s
        if not prev:
            changes[k] = {"t": "new", "c": v["chassis"], "so": v["so"], "d": v["dealer"], "m": v["model"]}
        else:
            f = [INDEX_FIELDS[i] for i, (a, b) in enumerate(zip(prev.get("idx", []), s["idx"])) if a != b]
            if fingerprint(prev.get("vehicle")) != fingerprint(v) and not f:
                f = ["orders"]
            if fingerprint(prev.get("cfg")) != fingerprint(c.get("cfg")):
                f.append("config")
            if f:
                changes[k] = {"t": "mod", "c": v["chassis"], "so": v["so"], "d": v["dealer"], "m": v["model"], "f": f}
    if not args.users:
        for k, prev in old.items():
            if k not in vehicles:
                pv = prev["vehicle"]
                changes[k] = {"t": "del", "c": pv["chassis"], "so": pv["so"], "d": pv["dealer"], "m": pv["model"]}
    counts = collections.Counter(c["t"] for c in changes.values())
    log(f"diff: {counts.get('new', 0)} new, {counts.get('mod', 0)} changed, {counts.get('del', 0)} gone")

    stats = config_stats(snapshot)
    meta = {"lastRun": run_at.isoformat(timespec="seconds"), "runId": run_id, "vehicles": len(snapshot),
            "active": sum(1 for s in snapshot.values() if s["vehicle"]["active"]),
            "withConfig": sum(1 for s in snapshot.values() if s.get("cfg")), "orders": len(orders),
            "unreadableOrders": sum(unreadable.values()), "sapCalls": sap.calls,
            "changes": {"new": counts.get("new", 0), "mod": counts.get("mod", 0), "del": counts.get("del", 0)},
            "indexFields": INDEX_FIELDS, "seconds": int(time.time() - t0)}

    # 5. Firebase ------------------------------------------------------------------------------------------
    push_firebase = cfg["firebase_enabled"] and not args.dry_run and not args.no_firebase and not args.users
    fb_pushed_all = old_meta.get("firebasePushedAll", False)

    def save_snapshot(pushed_all):
        meta["firebasePushedAll"] = pushed_all
        if fb_state:
            n = fb_state.save_state({"meta": meta, "vehicles": snapshot})
            log(f"snapshot saved to Firebase ({n / 1e6:.1f} MB)")
        else:
            tmp = snap_path.with_suffix(".tmp")
            with gzip.open(tmp, "wt", encoding="utf-8") as f:
                json.dump({"meta": meta, "vehicles": snapshot}, f, ensure_ascii=False)
            os.replace(tmp, snap_path)
            log("snapshot saved")

    can_save = not args.dry_run and not args.users and (push_firebase or not fb_state)
    full_push = push_firebase and (args.push_all or not fb_pushed_all)
    if can_save and full_push:
        # keep the downloaded configurations even if the large first push fails; the next run then only re-pushes
        save_snapshot(False)
    if push_firebase:
        fb = Firebase(cfg)
        full = full_push
        upd = {}
        targets = snapshot.keys() if full else [k for k, c in changes.items() if c["t"] != "del"]
        for k in targets:
            s = snapshot[k]
            upd[f"index/{k}"] = s["idx"]
            upd[f"vehicles/{k}"] = s["vehicle"]
            upd[f"config/{k}"] = fb_config(s.get("cfg"))
        for k, c in changes.items():
            if c["t"] == "del":
                upd[f"index/{k}"] = None
                upd[f"vehicles/{k}"] = None
                upd[f"config/{k}"] = None
        upd["configStats"] = stats
        upd["meta"] = meta
        upd[f"changes/{run_id}"] = changes or {"_none": {"t": "none"}}
        log(f"Firebase: writing {'ALL ' if full else ''}{len(targets)} vehicles ...")
        fb.patch_many(upd)
        try:
            url_keys = fb._req("GET", "changes", query="shallow=true") or {}
            old_runs = sorted(url_keys)[:-cfg["keep_runs"]]
            if old_runs:
                fb.patch_many({f"changes/{r}": None for r in old_runs})
        except RuntimeError as e:
            log("note: could not prune old change logs:", e)
        log(f"Firebase: done, {fb.bytes / 1e6:.1f} MB uploaded")
        fb_pushed_all = True
    else:
        log("Firebase: skipped (" + ("dry run" if args.dry_run else "disabled in config" if not cfg["firebase_enabled"] else "flag") + ")")

    # 6. Excel + local files -------------------------------------------------------------------------------
    if cfg["excel_enabled"]:
        xlsx = out_dir / f"commerce_{run_id}.xlsx"
        write_excel(xlsx, snapshot, changes, run_id, cfg.get("excel_config_sheet", True))
        log("Excel:", xlsx.name)
        for p in sorted(out_dir.glob("commerce_*.xlsx"))[:-cfg["keep_excel"]]:
            p.unlink()
    hist = state_dir / "changes_history.json"
    runs_changes = json.loads(hist.read_text(encoding="utf-8")) if hist.exists() else {}
    if not args.dry_run and not args.users:
        runs_changes[run_id] = changes or {"_none": {"t": "none"}}
        runs_changes = dict(sorted(runs_changes.items())[-cfg["keep_runs"]:])
        hist.write_text(json.dumps(runs_changes, ensure_ascii=False), encoding="utf-8")
    if args.export_local:
        export_local(out_dir, snapshot, meta, stats, runs_changes)

    if can_save:
        save_snapshot(fb_pushed_all)
    elif fb_state and not args.dry_run:
        log("snapshot NOT saved (Firebase writing is off for this run)")
    log(f"finished in {int(time.time() - t0)} s")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # keep a readable message in the Task Scheduler log
        log("FAILED:", exc)
        sys.exit(1)
