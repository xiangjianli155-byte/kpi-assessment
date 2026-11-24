# KPI 考核网站

这是一个 GitHub Pages 静态网页，直接在浏览器中读取 Excel。网站不使用 Python、Streamlit 或服务器端上传；普通访客只能查看、筛选和导出数据。

## 网站地址

仓库 `xiangjianli155-byte/kpi-assessment` 对应的项目网站地址是：

**https://xiangjianli155-byte.github.io/kpi-assessment/**

## 发布网站

网页入口 `index.html` 已放在仓库根目录。到 GitHub 仓库的 **Settings → Pages**：

1. 在 **Build and deployment** 的 **Source** 中选择 **Deploy from a branch**。
2. Branch 选择 **main**，文件夹选择 **/(root)**，点击 **Save**。
3. 等 Pages 显示已发布后，打开上面的网站地址。

GitHub Pages 的项目网址使用 `用户名.github.io/仓库名/`。Pages 要求入口 `index.html` 位于所选发布目录顶层。[GitHub Pages 发布源说明](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site) · [404 排查说明](https://docs.github.com/en/pages/getting-started-with-github-pages/troubleshooting-404-errors-for-github-pages-sites)

如果仓库目前是私有仓库，GitHub Free 不支持从私有仓库发布 Pages；需要使用支持私有 Pages 的付费方案，或将仓库改为公开。无论仓库是否公开，发布后的 Pages 网站和里面的数据都能被任何访客查看。当前 KPI 工作簿也会通过网站公开提供，请确认工作簿可公开分享后再提交。[GitHub Pages 计划与可见性说明](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site)

## 每周更新 Excel

1. 在仓库打开 `data` 文件夹。
2. 上传最新工作簿，文件名固定为 **`current_kpi.xlsx`**，覆盖旧版本并提交到 **main**。
3. Pages 更新完成后，刷新网站。所有访客会看到同一份最新数据。

网站自动寻找名称包含 `INV` 和 `EO` 的工作表，并识别 Planner、WHS、Actual/实际、Target/目标和日期列。保持这几类字段及日期列结构，替换数据时就不需要改网页代码。数据上传者需要有仓库写入权限；网页访客没有上传按钮。

## 页面功能

- On Hand INV 与 EO 专题切换
- Planner、按周/按月、日期范围和达标方向筛选
- 全部 Planner 实际值与标准值对比；单 Planner 显示期间变化
- Planner × WHS 实际值、标准值、差额和达标情况
- 导出 PDF（浏览器打印窗口中选择“另存为 PDF”）和 PowerPoint

金额单位为 USD。PDF 和 PowerPoint 由访问者的浏览器生成。

## 网页依赖

网页通过 CDN 加载 SheetJS、Chart.js 和 PptxGenJS。浏览器需要能访问这些 CDN；无需安装 Python 或运行 `pip install`。

