# 📊 数据采集加工全链路看板

企微文档数据可视化看板，部署在 GitHub Pages 上，领导打开链接即可查看。

## 更新数据

1. 在企微文档中修改数据
2. 导出为 CSV 文件
3. 打开看板页面 → 点「导入CSV」
4. 检查数据无误后点「导出 data.json」
5. 将 `data.json` 上传到 GitHub 仓库替换原文件
6. 等待 1-2 分钟，刷新页面即看到最新数据

## 直接编辑

也可以直接修改 `data.json` 文件后提交到 GitHub。

## 技术栈

- 纯前端 HTML + CSS + JavaScript
- ECharts 5 图表库
- PapaParse CSV 解析
- 数据与视图分离（data.json）
