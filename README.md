# KPI 考核

Streamlit 网页应用，用于分析 On Hand INV 和 EO KPI 工作簿中的 Planner、WHS、实际值、标准值和达标状态，并导出当前视图为 PowerPoint 或 PDF。

## 网站使用方式

网站默认读取仓库中的共享数据文件 `data/current_kpi.xlsx`。打开网站的人会看到同一份最新数据。左侧临时上传器只会更改当前浏览会话，不会更新其他人的共享数据。

## 发布成可分享的网站

1. 在 GitHub 新建一个**私有仓库**，例如 `kpi-assessment`。
2. 上传本文件夹中的 `app.py`、`requirements.txt`、`README.md`、`.gitignore` 和 `data/README.md`。
3. 将每周使用的 Excel 上传到仓库中的固定路径 `data/current_kpi.xlsx`。不要把仓库改成公开；这样原始工作簿保留在私有仓库中。
4. 登录 Streamlit Community Cloud，连接 GitHub，选择此私有仓库，入口文件填 `app.py` 并部署。
5. 在应用的 Sharing 设置里将应用设为 public，然后把 Streamlit 给出的网址发给同事。

**访问范围：**Streamlit Community Cloud 的公开访问选项是“public and searchable”。这意味着收到网址的人可以访问，网站也可能被其他人发现；页面会显示共享工作簿中的 KPI 数据。请只在公司允许这些 KPI 数据公开给所有访问者时使用此模式。如果要限制为指定同事，需要选择私有访问并邀请他们登录。

## 每周更新共享数据

在私有 GitHub 仓库中替换 `data/current_kpi.xlsx` 并提交更改。Streamlit 会根据仓库更新网站，所有访问者随后会看到新数据。不要将工作簿放进公开仓库。

## 在 VS Code 中本地运行

使用 Python 3.11 或更新版本，在终端运行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

本机运行时若仓库中没有 `data/current_kpi.xlsx`，可在左侧上传 Excel；这种临时上传只对当前会话可见。

