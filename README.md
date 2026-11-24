# KPI 考核

Streamlit 网页应用，用于查看 On Hand INV 和 EO KPI 的 Planner、WHS、实际值、标准值和达标情况，并导出 PDF 或 PowerPoint。

## 网站和数据权限

- GitHub 源仓库 `xiangjianli155-byte/kpi-assessment` 保持私有，网站可设置为公开供所有人查看。
- 网站展示的数据对访客可见；任何访客都可以从网站上传数据。
- 每一批上传都会追加到私有仓库的 `data/weekly/` 历史中，不会覆盖之前批次。
- 多批数据按上传时间合并；同一日期、Planner 和 WHS 重复时，采用较新的记录。
- 不需要每周手动登录 GitHub 上传工作簿。

## 发布网站

1. 在 GitHub 生成细粒度访问令牌：头像 → **Settings** → **Developer settings** → **Personal access tokens** → **Fine-grained tokens** → **Generate new token**。
2. 将 **Repository access** 设为 **Only select repositories**，只选择 `kpi-assessment`；在 **Repository permissions** 中将 **Contents** 设为 **Read and write**，再生成并复制 token。[GitHub 官方创建说明](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens)
3. 在 [Streamlit Community Cloud](https://share.streamlit.io/) 登录并连接 GitHub。如果看不到私有仓库，在 Streamlit 工作区设置中授权访问 private repositories。
4. 点击 **Create app**，选择仓库 `xiangjianli155-byte/kpi-assessment`、分支 `main`、入口 `app.py`。在 **Advanced settings → Secrets** 中粘贴：

   ```toml
   GITHUB_TOKEN = "你的细粒度 GitHub token"
   ```

   这个 token 仅供网站后台保存访客上传的数据，访客不会看到或输入它。不要把真实 token 写进 GitHub 文件或发在聊天里。

5. 部署后，在应用 **Sharing** 设置中设为公开，然后把 Streamlit 网址发给同事。公开应用可能被搜索引擎收录，所有访客都能看到页面展示的 KPI 数据，也都能追加上传批次。

## 每周从网站上传

1. 打开网站左侧的 **共享数据**。
2. 选择一个或多个 `.xlsx`、`.xlsm`、`.xls` 或 `.csv` 文件。
3. 确认文件名后点击 **追加到共享历史**。

上传批次会保存到私有 GitHub 仓库，Streamlit Cloud 同步完成后，访客刷新页面即可看到累计数据。上传的文件需要包含当前 KPI 分析支持的 Planner、WHS、实际值和目标值字段；不同工作簿结构可在页面中进行字段映射。

## 本地运行

使用 Python 3.11 或更新版本，在终端运行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

本地查看已提交的历史批次不需要 GitHub Secrets。要在本地启用共享上传功能，需在 `.streamlit/secrets.toml` 设置 `GITHUB_TOKEN`；该文件已被 `.gitignore` 忽略，切勿提交真实 token。

