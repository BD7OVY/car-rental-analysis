# 租车经营分析平台 · Car Rental Analysis

一个通用的**租车订单分析平台**：导入各租车平台后台导出的原始订单 Excel / JSON → **自动识别平台、统一表头、按城市/平台聚合** → 用图表与城市地图直观展示订单与应收（GMV）。

纯前端单文件应用，**数据只在你的浏览器本地解析计算，不上传任何服务器**。可直接双击 `index.html` 打开，或一键部署到 GitHub Pages 等任意静态托管。

---

## 特性

- **多平台自动识别**：哈啰 / 滴滴 / 悟空 / 携程，按表内「专属列」识别，**与文件名、固定目录无关**。
- **表头自动统一**：任何列名变体都能命中（精确候选 + 分词规则），无需手工整理。
- **跨文件去重**：按订单号全局去重，重叠导出的文件夹/多份文件不会重复计数。
- **丰富看板**：KPI、日/周/月趋势、环比、城市/平台排名、交叉明细表、城市分布地图。
- **离线可用**：地图与 Excel 解析库均本地化，`file://` 双击与静态托管都能跑。
- **少量内置示例**：默认载入 2 平台 / 5 城市 / 2 个月（共 32 单）的演示数据，开箱即看。

---

## 目录结构

```
index.html           主程序（纯前端工作台：导入 / 处理 / 展示）
china-geo.js         本地化中国地图数据（省份边界 + 城市质心，离线可用）
xlsx.full.min.js     本地化 SheetJS（解析 Excel，离线可用）
echarts.min.js       本地化 ECharts（地图与图表）
schema_mapping.json  表头统一「单一映射配置」（前端导入与 Python 管线共用）
data.json            内置示例数据（页面默认载入）
sample-data/         演示原始 Excel（悟空 / 滴滴，表头各不相同，自动识别）
tools/               Python 离线刷新管线（可选，非部署必需）
  generate_dashboard.py  读 Excel → 统一 schema → 算指标 → 生成 data.json
  server.py                本地预览服务（http://localhost:8080）
  config.json              数据源 / 输出目录 / 端口等配置
  make_demo_data.py        生成随仓库附带的演示原始 Excel
.nojekyll               禁用 Jekyll，确保 data.json 与离线库原样发布
启动看板.bat          一键本地运行（Windows）
```

> GitHub Pages 部署只需 `index.html` + `china-geo.js` + `echarts.min.js` + `xlsx.full.min.js` + `schema_mapping.json` + `data.json` 这些文件，其余为本地开发/刷新工具。

---

## 优化后的导入流程（本项目的重点改进）

1. **拖入 / 选择文件 / 选择文件夹**，或一键「使用示例数据」。
2. 平台**自动识别**，导入区逐文件显示解析结果：
   - 绿色 ✓ 成功：显示识别到的平台与订单数；
   - 橙色 ⚠ 跳过：未识别到已知平台订单表（可核对列名）；
   - 红色 ✕ 失败：解析异常，给出原因。
3. 「**查看字段映射**」展开，核对「你的原始列名 → 系统统一字段」是否识别正确。
4. 匹配不上？点「**下载导入模板**」按统一字段填好再导入；或调整 `schema_mapping.json`。
5. 切换上方「**筛选月份**」查看各月分析；选「**全部月份**」看累计口径。

导入区还提供：
- **已支持平台速览**：一眼确认自己的文件会被识别（含各平台关键列）；
- **数据状态条**：显示当前已载入单数、来源，可「重新导入 / 清空」；
- **隐私提示**：明确告知数据仅本地处理、不上传。

---

## 快速开始

**方式一 · 直接打开**
双击 `index.html` 即可（默认载入内置示例）。

**方式二 · 本地服务**（推荐，示例数据稳定加载）
```bash
cd tools
python server.py        # 打开 http://localhost:8080
```

**方式三 · 一键启动（Windows）**
双击 `启动看板.bat`。

---

## 新增平台 / 调整列名

全部收敛在 `schema_mapping.json` 一份文件（前端导入按钮与 Python 管线读同一份）：
- 每个平台的「列名 → 统一字段」候选、城市映射、状态映射、车型归一、节假日；
- 平台识别**不依赖文件名**，而是按「专属标识词 + 列特征」自动判断。

新增渠道或调整列名，**只改这个文件，不用动代码**。

---

## 部署到 GitHub Pages（分支部署，零工作流）

本仓库采用与 STR-Analysis 相同的「从分支发布」方式，**不需要 GitHub Actions 工作流**，配置一次即可：

1. 把本仓库推送到 GitHub 的 `main` 分支；
2. 进入仓库 **Settings → Pages → Build and deployment → Source**，选择
   **Deploy from a branch**，分支选 **main**，目录选 **/ (root)**；
3. 等待约 1 分钟，访问 `https://<你的用户名>.github.io/car-rental-analysis/` 即可。

> 仓库根目录的 `.nojekyll` 已禁用 Jekyll，确保 `data.json` 与离线库等静态资源原样发布。
> 之后每次推送 `main`，GitHub Pages 会自动更新站点，无需任何额外操作。

---

## Python 管线（可选，非部署必需）

```bash
cd tools
python make_demo_data.py        # 重新生成 sample-data 演示 Excel
python generate_dashboard.py    # 读 sample-data → 生成 data.json（刷新内置示例）
python server.py                # 本地预览
```

`config.json` 可配置 `data_source`（原始 Excel 目录）、`output_dir`（产出目录）、`server_port`（端口）。

---

## 隐私说明

本工作台是**纯前端应用**，所有解析与计算都在浏览器本地完成，不会上传任何数据到服务器。你也可以用任意静态托管自行部署，数据始终只留在访问者本地。
