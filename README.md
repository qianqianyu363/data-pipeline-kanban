# 📊 数据采集加工全链路看板

> 🌐 **在线地址**：https://qianqianyu363.github.io/data-pipeline-kanban/
>
> 需求下发 → 采集执行 → 数据加工/审核 → 发布闭环 全链路数据可视化看板，部署在 GitHub Pages，领导打开链接即可查看。

## 📋 看板模块

| 模块 | 内容 |
|------|------|
| 流程 KPI 卡片（8 环节） | 需求下发 / 需求校核 / 数据采集 / 问题闭环 / 质量审核 / 加工审核发布 / 结算与成本，随筛选联动 |
| 📈 月度趋势总览 | 需求总条数 / 采集下发数 / 采集完成数 / 完成率 / 准确率 / 无效数据 / 预算执行率 逐月走势 |
| 📋 采集执行分析 | 4 个 Tab：采集进度 / 采集产能 / 费用执行 / 任务清单 |
| ⚙️ 筛选联动 | 按需求项 / 结算类型 / 月份筛选，KPI 卡片与采集进度、质量抽审联动 |

## 🔄 更新数据（推荐流程）

### 方式1：改模板 + 跑解析器（推荐，持久化）
1. 在 Excel 模板 `看板数据录入模板-调整.xlsx` 中更新数据（6 个 Sheet）
2. 生成 data.json：`python parse_template.py`
3. 提交推送：`git add . && git commit -m "更新看板数据" && git push origin main`
4. 等待 1-2 分钟，GitHub Pages 自动更新，刷新看板即生效

### 方式2：浏览器直接导入模板
1. 打开 https://qianqianyu363.github.io/data-pipeline-kanban/
2. 点击「导入」→ 选择 `看板数据录入模板-调整.xlsx`
3. 自动识别并解析全部 6 个 Sheet，图表即时更新
4. 若需持久化，把更新的 data.json 提交到仓库

### 方式3：在线编辑 / 导出 data.json
1. 打开看板页面 → 点击「✏️ 编辑数据」直接改 JSON
2. 满意后「导出 data.json」替换仓库文件

## 🛠 技术栈

- 纯前端 HTML + CSS + JavaScript
- [ECharts 5](https://echarts.apache.org/) 图表库（柱状图 / 折线图 / 饼图 / 组合图）
- 数据与视图分离（data.json + 模板解析）
- Python 解析器 `parse_template.py`（模板 Excel → data.json，含公式求值 / 对话台账合并）
- 托管于 GitHub Pages（免费 CDN）
