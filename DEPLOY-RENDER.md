# 部署到 Render 服务器

## 一、准备代码仓库

1. 在项目目录 `f:\录制` 下初始化 Git（若尚未有仓库）：
   ```bash
   git init
   git add .
   git commit -m "Initial commit for Render"
   ```
2. 在 GitHub 或 GitLab 新建一个仓库，把本仓库推送上去：
   ```bash
   git remote add origin https://github.com/你的用户名/你的仓库名.git
   git branch -M main
   git push -u origin main
   ```
   （若已有仓库，直接 `git push` 即可。）

注意：`.env` 和 `data/` 已在 `.gitignore` 中，不会上传；密码和本地数据需在 Render 上单独配置。

---

## 二、在 Render 创建 Web 服务

1. 打开 [https://render.com](https://render.com)，注册/登录。
2. 点击 **Dashboard** → **New +** → **Web Service**。
3. **Connect a repository**：选择你的 GitHub/GitLab 账号，选中刚推送的仓库。
4. 配置：
   - **Name**：例如 `role-exp-stats`
   - **Region**：选离你近的（如 Singapore）
   - **Branch**：`main`
   - **Runtime**：`Node`
   - **Build Command**：`npm install`
   - **Start Command**：`npm start`
   - **Instance Type**：Free（免费）或选付费机型
5. 点击 **Advanced**，添加环境变量（Environment Variables）：
   - `ADMIN_SECRET` = `你的管理员密码`（必填，否则所有人都能操作分析/合并）
   - 可选：`SYNC_URL`、`SYNC_INTERVAL_MINUTES`（智能表格自动同步）
   - 可选：`WECOM_TOKEN`、`WECOM_ENCODING_AES_KEY` 等（企业微信机器人）
6. 点击 **Create Web Service**，等待构建和部署完成。

---

## 三、访问地址

部署成功后，Render 会给出一个地址，形如：

- **https://role-exp-stats.onrender.com**（名称以你填的 Name 为准）

把该地址发给别人，他们用浏览器打开即可：
- **未登录**：只能「按角色搜索」「各角色总和」「按日期查询」
- **你**：点击「管理员登录」输入 `ADMIN_SECRET` 后，可分析表格、智能表格同步、合并规则

---

## 四、重要说明（Render 免费版）

- **数据持久化**：免费实例重启或休眠后，**磁盘上的 `data/` 会清空**（analyzed.json、role-aliases.json 等）。  
  若需要长期保留分析结果和合并规则，可以：  
  - 使用 Render 的 **Persistent Disk**（付费），或  
  - 每次部署后重新「分析表格」/「从智能表格获取并分析」并重新保存合并规则。
- **休眠**：免费服务约 15 分钟无访问会休眠，别人第一次打开可能稍慢，属于正常现象。

---

## 五、若使用 render.yaml（可选）

若仓库根目录有 `render.yaml`，可在 Render 选择 **Blueprint** 方式创建服务，按文件中的配置一键创建；环境变量仍需在 Dashboard 的 **Environment** 里手动添加（尤其是 `ADMIN_SECRET`）。
