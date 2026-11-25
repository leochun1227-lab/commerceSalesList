# Commerce 订单与配置 — Render 部署

部署后有两个 Render 服务：

| 服务 | 类型 | 做什么 | 费用 |
|---|---|---|---|
| `commerce-dashboard` | 静态网站 | 托管网页 `web/index.html`，网页直接读 Firebase | 免费 |
| `commerce-sync` | Cron Job | 每天从 SAP 拉数据，和上次的快照比对，只把变化写进 Firebase | 每天约 7 分钟，按秒计费，大约每月 1 美元（Render 每个 cron job 最低收费 1 美元/月） |

```
render.yaml          Render 蓝图：一次性建好上面两个服务
requirements.txt     Python 依赖
commerce_sync.py     同步脚本（通过环境变量配置，不需要配置文件）
web/index.html       网页
```

## 部署步骤

### 1. 准备 Firebase 服务账号
Firebase 控制台 → 项目设置 → 服务账号 → **生成新的私钥**，下载 JSON 文件。
这个文件等于数据库的管理员密码：**不要提交到 Git**，只粘贴到 Render 的环境变量里。

### 2. 设置 Firebase 数据库规则
Firebase 控制台 → Realtime Database → 规则，和现有规则合并：
```json
{
  "rules": {
    "commerce":       { ".read": true,  ".write": false },
    "commerce_state": { ".read": false, ".write": false }
  }
}
```
- `commerce`：网页读取的数据，所以允许读。
- `commerce_state`：同步脚本用来比对的快照（压缩后约 7 MB），不让外部读。
- 脚本用服务账号写入，不受这些规则限制。

### 3. 建 GitHub 仓库
新建一个 **私有** 仓库，把这个文件夹里的文件推上去（`.gitignore` 已经排除了密钥和快照）。

### 4. 在 Render 用 Blueprint 创建
Render Dashboard → **New → Blueprint** → 选这个仓库 → Render 会读取 `render.yaml`，列出两个服务。
在提示填写 `FIREBASE_SERVICE_ACCOUNT_JSON` 的地方，**把服务账号 JSON 文件的全部内容粘贴进去** → Apply。

### 5. 手动跑第一次
Render → `commerce-sync` → **Trigger Run**。
- 第一次是全量：拉全部订单和每台车的配置，大约 20 分钟，往 Firebase 写入约 100 MB。
- 在 Logs 里看到 `finished in ... s` 就是成功；如果失败，最后一行是 `FAILED: 原因`。
- 之后每天 UTC 18:30 自动运行，也就是悉尼夏令时早上 5:30、冬令时 4:30，每次 6~8 分钟。

### 6. 打开网页
Render → `commerce-dashboard` → 页面上方的 `https://commerce-dashboard-xxxx.onrender.com`。

## 注意

- **只保留一个同步来源**。如果 planning 电脑上也设了计划任务，Render 跑起来以后就把它停掉，避免两边同时写 Firebase。
- **网页是公开地址**：知道网址的人都能打开。数据里有客户姓名和价格。已经加了 `noindex`，搜索引擎不会收录，但这不等于有权限控制。要限制访问，需要给网页加 Firebase 登录（以后可以加）。
- **改同步时间**：修改 `render.yaml` 里的 `schedule`（UTC 时间），推送到 GitHub 后 Render 会自动更新。
- **重新全量**：在 Firebase 控制台删除 `/commerce_state`，然后 Trigger Run。
- **Excel**：Render 上没有磁盘，所以不生成 Excel。需要导出时，在网页的"订单"页点"导出 CSV"。
- **内存**：同步峰值约 600 MB，所以用的是 2 GB 的实例（`1c-2g`）。改成 512 MB 会运行失败。
