# 需求下发登记全链路看板 — 项目状态与工作约定

> 本文件是会话压缩/重启后恢复上下文的唯一依据，务必保持最新。

## 项目概况

- 目录：`D:\AI实操\看板`，git 仓库，main 分支（2026-08-10：origin 已同步至 e25cbf4；08-10 升级版模板适配改动本次提交）
- 看板已部署到 GitHub Pages（index.html 直接渲染 data.json）
- 业务：数据采集加工全链路（需求下发 → 采集执行 → 数据加工/审核 → 上线发布 → 加工审核发布 → 质量审核），多流程 ECharts 看板

## 数据流水线

```
看板数据录入模板-调整.xlsx ─┐
                           ├─ parse_template.py ─▶ data.json ──index.html──▶ GitHub Pages
需求下发登记.json ───────────┘  (对话登记台账，合并逻辑：
                                去重name+月份、字段归一化、缺失补None、派生自算)
```
- 用户填模板 → 跑 `python parse_template.py` → data.json 更新 → commit+push 部署
- **`需求下发登记.json` 是对话登记的数据源**（Claude 解析自然语言后写入，不碰 Excel），解析器自动合并进 data.json；同「名称+月份」模板优先、台账跳过，模板替换后重跑解析器不会丢台账
- 浏览器导入（index.html 内 `autoImportTemplate`，读 XLSX）是另一条导入路径，字段需与 Python 解析器保持一致

## 关键文件

| 文件 | 作用 |
|---|---|
| `看板数据录入模板-调整.xlsx` | 看板主数据源（**2026-08-10 已同步为用户最新升级版**：约定采集完成时间MMDD/25年对比3列/价格审核/采集产能9列含需求量） |
| `看板数据录入模板-调整_升级版.xlsx` | 升级版源文件（**用户手工维护的最新数据在此**，已 gitignore，08-10 更新后已覆盖主文件） |
| `parse_template.py` | Python 解析器 → data.json（**含对话台账合并 + 质量抽审公式求值 + 08-10 新版模板适配**） |
| `需求下发登记.json` | **对话登记台账**（Claude 对话写入，解析器合并进 data.json） |
| `index.html` | 看板前端（**7项优化 + UI指标去重修复 + dataEmbed 内嵌块自动同步 + 08-10 浏览器导入适配/产能需求量展示/质量抽审按月过滤**） |
| `data.json` | **2026-08-10 已重新生成**（26任务/51结算/4492加工预审，结算金额 sum=902,077.2） |
| `需求下发登记表.xlsx` | 用户手动维护的详细登记台账（"需求下发"sheet 33列含"智能总结"列、"供采试点下发需求""结算明细"等） |
| `_upgrade_template.py` | 模板升级脚本（已验证，可从 HEAD 模板重新生成升级版） |
| `.workbuddy/backup_20260803/` | workbuddy 改动备份 + HEAD 模板备份 |

## 当前工作：回滚 + 7项优化

### 已完成的步骤
1. ✅ **全量回滚到 HEAD**（data.json / index.html / parse_template.py / 模板），清理了 __pycache__、debug_out.txt、embed_summary.txt、qr_fn.txt、_备份/_新版 模板副本
2. ✅ **模板升级** → 生成 `_升级版.xlsx`（已验证）：
   - 各项需求跟踪 第7列插「约定采集完成时间」，其后列右移（下发数量→col13、下发材料→col14、预审问题→col15、需求准确率→col18、采集完成量→col22、采集完成率→col24、无效数据→col25、无效占比→col26、发布数据→col27、需求发布及时率→col12）
   - 预算与结算 第7-11列插去年值5列（去年产品模块/需求类型/需求数量/结算单价/支出费用），「结算一览」副表右移且合并单元格已修复（P2:W2）
   - 加工预审 第9列插「价格审核结果」，「不通过原因」顺延 col10
   - 新建「采集产能」Sheet（7列：月份/计划采集家数/计划采集条数/实际采集家数/实际采集条数/员工总人数/备注）
   - 所有公式引用列号已 +1 修正（如采集完成率 =U/L → =V/M）
3. ✅ **解析器升级完成**（parse_template.py）：列映射适配升级版（timelyRate=col12、dueMonth=col7、去年值5列 col7-11、price_audit=col9、old_enterprise=col7『更新登记表』、采集产能新Sheet）；内置公式求值器自算派生值；新增命令行参数 `python parse_template.py [模板.xlsx] [输出.json]`（默认主文件/data.json）
   - 已用 `_升级版.xlsx` 验证：19任务/25结算/3725加工预审全部正确，与磁盘模板缓存值全量一致
4. ✅ **前端7项优化完成**（index.html）：
   - ① 当月采集完成率（dueMonth==筛选月聚合，renderFlowKPIs 数据采集卡片）
   - ② 加工审核三项（新老供应商占比 + 价格审核通过率，加工审核发布卡片）
   - ③ 指标说明 tooltip（METRIC_TIPS 字典 + ⓘ 悬浮）
   - ④ 月度趋势图例（baseOpt/lines/comboRev 自动提取 series 名）
   - ⑤ 任务清单闭环三态（是/否/跨月跟进，closedState；新增"约定完成"列 + .s-follow CSS）
   - ⑥ 费用虚框（预测费用虚线框）+ 去年费用对比（settlements 去年字段→紫色折线+指标框）
   - ⑦ 采集产能三维度（总产能家数/条数/人均家数/人均条数/总人数）+ 产能人数预测（线性回归）
   - 浏览器导入已同步适配升级版：TPL_COL 0-indexed、parseTasksSheet（dueMonth+自算兜底）、parseSettlementsSheet（去年字段）、parsePreReviewSheet（price_audit/old_enterprise）、新增 parseCollectionCapacitySheet
   - 已用注入新字段的测试数据跑通全部渲染函数（node 冒烟测试无异常）
5. ✅ **主模板已替换为升级版**（2026-08-04）：备份旧模板至 `.workbuddy/backup_20260804/看板数据录入模板-调整_旧版_20260804.xlsx` 后，用 `_升级版.xlsx` 覆盖主文件；旧版磁盘模板经解析为 21任务/26结算（比 HEAD data.json 的 19/25 新，符合"以模板为准"）

### 7项优化点（模板+解析器+前端三者配套，避免死代码）
1. **当月采集完成率** — 前端基于 dueMonth==筛选月聚合
2. **加工审核**：入库通过率 + 新老供应商占比 + 价格审核通过率（price_audit）
3. **指标说明 tooltip**（METRIC_TIPS 字典 + 悬浮提示）
4. **月度趋势图例**（baseOpt 自动提取 series 名）
5. **任务清单闭环三态**：是/否/跨月跟进（区分是否当月任务）
6. **费用未发生虚框 + 去年费用对比**
7. **采集产能三维度（总产能/人均家数/人均条数/总人数）+ 产能人数预测**

### 待办任务（全部完成，2026-08-04）
- ✅ **解析器升级**：公式求值器 + 升级版列映射（详见上文）
- ✅ **前端7项优化**：当月采集完成率/加工审核三项/指标tooltip/图例自动提取/任务闭环三态/费用虚框+去年对比/采集产能三维度
- ✅ **对话式需求下发登记功能**：`需求下发登记.json` 台账 + parse_template.py 合并逻辑（load_ledger/merge_ledger_tasks）。本次再验证：用主模板+测试台账跑通——新增1条(07月测试比价)/模板重复跳过1(北京比价/03月,模板原值保留)/派生值全对(totalPoints=150=50×3,completionRate=80.0=120/150,accuracyRate=50.0=50/100)/07月入 months/monthlyTrend 含07月
- ✅ **前端冒烟**：`node _smoke_test.js [data.json] [index.html]`（自包含：自动从 index.html 提取主脚本，注入 DOM/ECharts/fetch/ResizeObserver stub）。真实 data.json 与含台账合并的 test_ledger_data.json 均 renderAll + 4 tab 5/5 通过无异常
- ✅ **生成 data.json 并全量验证**：`python parse_template.py 看板数据录入模板-调整.xlsx data.json` → 21任务/26结算/3725加工预审/结算金额 sum=902,077.2 与缓存一致；因模板已替换为升级版，无需再替换主文件
- ✅ **提交部署**：05d89e3（模板替换+解析器+台账合并+7项优化+冒烟）→ 已 push；b9cf068（B方案内嵌同步）→ 已 push
- ✅ **前端UI指标去重修复**（4dcbbbb，2026-08-04，已提交待推）：
  - 需求下发卡去「需求总量」（保留需求总条数）；需求校核卡去「反馈问题」「需求问题率」（保留初审问题数/初审问题率）；数据采集卡去「当月承接」（保留采集下发数）；问题闭环卡去「未共识量」「总反馈量」「双方有效共识率」（保留采集反馈问题/共识量/共识率）；加工审核发布卡去「审核人」明细
  - 质量抽审详情重写：按 类目→月份 展开 by_month 嵌套、跳过对象值，修复 [object Object]
- ✅ **质量抽审公式求值**（4dcbbbb）：质量抽审 sheet 的 `=E2-E3` 等公式复用 `cell_val` 求值（eval_formula **实际支持跨行**，CLAUDE.md 记"限同行"不准确）；已验证 初审不通过=96/初审通过率=85.5%/报价偏高=15/高质量报价=73.7%
- ✅ **push 已完成**：`4dcbbbb` 已推送，origin 同步至 `e25cbf4`（定时补推任务随之失效，已无残留）
- ⏳ **指标释义（METRIC_TIPS）内容由用户补充替换**（结构未动，'反馈问题'/'需求问题率'死条目已删）
- ⏳ **模板业务数据待填**：仅剩 加工预审 col9 价格审核结果 全空（前端显示"—"）

## 08-10 升级版模板刷新（已完成）

### 用户模板改动
- 数据更新：各项需求跟踪新增 7/8月 5条任务；加工预审 3725→4492 行；采集产能从空表填上 1-8月
- 结构微调：① 预算与结算数据行 row3→row4（row3 成子表头）、去年值 5列→3列（col7-9，删去年模块/类型）；② 采集产能 7列→9列（新增计划/实际需求量，原列右移）；③ 问题反馈新增公式 `=C2-C3`；④ 约定采集完成时间改填 **MMDD 数值**（730=7月30 等）

### 解析器适配（parse_template.py）
1. settlements：循环 row4 起，去年值改读 col7/8/9，删除 `去年产品模块`/`去年需求类型` 字段
2. 采集产能：映射右移并新增 `计划需求量`(col2)/`实际需求量`(col5)
3. 问题反馈：value 改用 `cell_val` 求值 → `双方未达成共识量=722`（此前存成公式字符串 bug）
4. `fmt_due_month` 支持 MMDD 整数（101-1231 且 商=1-12 判为 `n//100` 月），不误判为日期序列号

### 前端适配（index.html）
- 浏览器导入 TPL_COL：settlements 去年值改 `lastQty:6/lastPrice:7/lastCost:8`、`parseSettlementsSheet` 数据起点 r=2→r=3、删除去年模块/类型；capacity 改 9 列（`month0/planQty1/planHouse2/planItem3/actQty4/actHouse5/actItem6/staff7/note8`）；`fmtDueMonth` 支持 MMDD
- 产能 Tab 展示需求量：新增「计划需求量」「实际需求量」累计指标框 + 「📦 采集需求量（计划 vs 实际）」柱状图（实际为实心、计划为虚线框）
- dataEmbed 内嵌块由解析器自动同步为新数据

### 验证
- `_smoke_test.js` 5/5 通过；结算金额 sum=902,077.2 与历史一致；去年支出 sum=491.3万；dataEmbed 合法 JSON（26/51/8）
- 主模板已备份至 `.workbuddy/backup_20260810/` 后用升级版覆盖

### 数据现状提醒（2026-08-10）
- ✅ **约定采集完成时间(dueMonth) 已填**（MMDD 数值：730=7月30→07月、810=8月10→08月、830=8月30→08月、930=9月30→09月）→ 当月采集完成率 / "约定完成"列有值
- ✅ **采集产能 Sheet 已填**（1-8月；6-8月含计划/实际需求量+条数）→ 产能 Tab 展示需求量
- ✅ **预算结算 25年对比 已填**（col7-9 去年数量/单价/支出）→ 费用"去年对比"折线有值
- ⏳ **加工预审 col9 价格审核结果 仍全空** → 加工审核"价格审核通过率"显示"—"，待用户填数
- ⏳ **采集产能列个别月份不全**（8月 实际需求量/条数空、1-5月无计划量）→ 前端对应显示"—"；`湖北比价`(8月)仅填了基础列，指标全空属正常未完成态

### 重要发现（回滚后）
- **磁盘模板 `看板数据录入模板-调整.xlsx` 的数据比 git HEAD 的 data.json 新**（git 提交里模板与 data.json 本就不同步）。生成 data.json 必须以**模板**为准，HEAD data.json 仅作旧值参考

## 当前 git 工作区（2026-08-10 提交后）

```
origin/main = 本次 08-10 提交（解析器/前端/模板/data.json 适配刷新 + CLAUDE.md 同步）
```
- 历史已推送：`05d89e3`（模板替换+解析器升级+台账合并+7项优化+冒烟）、`b9cf068`（dataEmbed 内嵌同步）、`4dcbbbb`（UI指标去重+质量抽审求值）、`e25cbf4`（CLAUDE.md 更新）
- 本次提交：08-10 升级版模板适配（parse_template.py + index.html TPL_COL/需求量展示 + 主模板同步 + data.json + CLAUDE.md）
- `.workbuddy/`（含 backup_20260810 主模板备份）、`会话.txt`、`指标网-采集检视表.xlsx`、`_升级版.xlsx` 已 gitignore

### 网络方案（重要，可复用）
**github.com 亚太节点（20.205.243.166）443 被阻断时**：在 `C:\Windows\System32\drivers\etc\hosts` 追加 `140.82.112.3 github.com`（GitHub 美国节点，已验证 HTTPS+git 协议可达）→ `ipconfig /flushdns` → push 恢复。该记录当前仍在 hosts，删两行即可撤销。若该节点也 TLS 失败（时段性干扰），可换 `140.82.114.4`/`140.82.113.3`。诊断：`curl -sI https://github.com`（走 hosts）、`curl --resolve github.com:443:<IP> https://github.com/` 测具体节点。
- 下一次提交建议仅当：台账有新登记（`需求下发登记.json` + 重跑解析器 + data.json + CLAUDE.md 状态）或模板/前端有改动

## 验证/测试命令（重要）

- **解析器验证**：`python parse_template.py 看板数据录入模板-调整.xlsx test.json` → 输出与磁盘模板逐字段对比（settlements sum=902,077.2，**26任务/51结算/4492加工预审**；以磁盘模板为准）
- **前端语法**：python 提取主 `<script>` 块 → `node --check`
- **前端冒烟**：`node _smoke_test.js [data.json] [index.html]`（自包含脚本，默认读 data.json + 当前目录 index.html，注入 DOM/ECharts/fetch/ResizeObserver stub 后跑 `renderAll()` + 4 个 tab，输出 5/5 通过）
- **台账合并验证**：临时用测试台账覆盖 `需求下发登记.json`（先备份）→ 跑解析器 → 期望"新增 N 条 + 与模板重复跳过 M 条" → 恢复原台账

## 升级版模板完整列结构（1-indexed，`_升级版.xlsx` 已核验）

### 各项需求跟踪（27列，数据行从 row3 起，max_row=23）
| col | 表头 | | col | 表头 |
|---|---|---|---|---|
| 2 | 月份 | | 15 | 预审问题数据 |
| 3 | 需求项 | | 16 | (无表头, 公式=O*3) |
| 4 | 结算类型 | | 17 | 需求反馈率(=O/I) |
| 5 | 需求名称 | | 18 | **需求准确率**(=N/I) |
| 6 | 需求来源 | | 19 | 需求下发团队时间 |
| 7 | **约定采集完成时间**(dueMonth) | | 20 | 差距原因 |
| 8 | 收到数量(含多家) | | 21 | 4月承诺采集完成率 |
| 9 | **材料数量**(matQty) | | 22 | **实际采集完成量**(collected) |
| 10 | 产品承诺需求发出时间 | | 23 | 实际采集材料 |
| 11 | 实际需求收到时间(日期序列) | | 24 | **采集完成率**(=V/M) |
| 12 | **需求发布及时率**(timelyRate) | | 25 | **无效数据**(invalidData) |
| 13 | **下发数量含多家**(totalPoints) | | 26 | **无效数据占比**(=Y/N) |
| 14 | **下发材料**(itemCount) | | 27 | **发布数据**(published) |

公式列：col13下发数量`=N*3`、col14下发材料`=I-O`、col22采集完成量`=W*3`(山东)/`=I-Y`(广西)、col24完成率`=V/M`、col18准确率`=N/I`、col26无效率`=Y/N`、col8收到数量`=I*3`

### 预算与结算（新版 08-10：row1=标题'结算明细'、row2=分组表头、row3=子表头、**row4 起数据**）
row2: 1=月份 2=产品模块 3=需求类型 **4-6='26年'组**（4=需求数量 5=结算单价 6=支出费用`=D*E`）**7-9='25年'组**（7=去年需求数量 8=去年结算单价 9=去年支出费用`=G*H`）；row3=子表头（需求数量/结算单价/支出费用 各两遍）；col14='结算一览'副表（每月结算明细，解析器不读）
- **结算金额 = 需求数量×结算单价**（col6 公式无缓存，解析器自算）；去年支出费用同理 = 去年数量×去年单价（col9）
- 未来月份（7-12月）26年组常空、25年组有数 → 结算金额=0、去年支出有值（费用"去年对比"折线用）

### 加工预审（28列，max_row=3738 含大量空行/汇总区，按 col3 审核人非空过滤）
1=审核日期 2=月份 3=审核人 4=拓展人 5=供应商名称 6=需求 7=**是否新建企业**(new=是/新增, old=更新登记表) 8=**是否审核通过** 9=**价格审核结果**(price_audit) 10=**不通过原因** 11=是否发布异地 12=报价单文件名称 13=报价单是否有日期 14=备注；col16+ 为汇总区(月份/更新登记表/新增/新拓占比/总计/审核通过/审核通过率)

### 质量抽审（11列）
1=月份 2=抽审任务 3=维度 4=项 5=值；col9-11=问题类型/值/占比

### 采集产能（新版 08-10：9列，**已填 1-8月**）
1=月份 2=计划需求量 3=计划采集家数 4=计划采集条数 5=实际需求量 6=实际采集家数 7=实际采集条数 8=员工总人数 9=备注
- 前端产能 Tab 用：实际采集家数/条数、员工总人数、计划/实际需求量（新）；1-5月仅填实际家数+人数，6-8月才含需求量

### 问题反馈（3列）
读 col2=标签 / col3=值（表头"问题反馈量"在 col2，数值 1000 在 col3）

## 关键坑（务必遵守）

1. **openpyxl 保存会丢公式缓存值**：升级版模板所有公式单元格 data_only=True 读出来是 None。**根因与解法已落地**：parse_template.py 用 `data_only=False` 加载保留公式文本，内置**公式求值器**（`eval_formula`/`cell_val`，支持 `=N5*3`、`=I5-O5`、`=V5/M5` 等算术，列引用限同行递归求值）。公式列自动恢复值，不依赖 Excel 缓存。
   - 模板中的公式列（各项需求跟踪）：下发数量 `=N*3`（=下发材料×3）、下发材料 `=I-O`（=材料数量−预审问题）、实际采集完成量 `=W*3`（=实际采集材料×3，山东）或 `=I-Y`（=材料数量−无效数据，广西）、采集完成率 `=V/M`、需求准确率 `=N/I`、无效占比 `=Y/N`
   - **准确性已用磁盘模板缓存全量验证**：所有数值字段与 Excel 真实计算值一致；settlements 25条 sum=902,077.2 完全匹配
   - **accuracyRate 语义 = 下发材料÷材料数量（模板公式 `=N/I`），不是 (1−预审问题/下发数量)**，此前的记录有误
   - **invalidRate 语义 = 无效数据÷下发材料（模板公式 `=Y/N`），不是 invalid/collected**
   - 解析器自算兜底：completionRate=None→collected/totalPoints；accuracyRate=None→itemCount/matQty；invalidRate=None→invalid/collected（仅当无公式时）
   - 这是 workbuddy 之前数据全空的根因，已根治
2. **workbuddy 的历史 bug**（勿重蹈）：parse_template.py timelyRate 列读错（读11应为12）、缺 old_enterprise 统计、openpyxl 丢缓存导致结算/完成率全空
3. **前端浏览器导入 TPL_COL 为 0-indexed**（=1基col−1，与 Python COL_MAP 一一对应；新版预算结算数据在 aoa 第 3 行起，`parseSettlementsSheet` 循环起点 r=3）：
   - tasks: `{month:1, demandItem:2, settleType:3, name:4, source:5, dueMonth:6, matQty:8, timelyRate:11, totalPoints:12, itemCount:13, reviewIssues:14, accuracyRate:17, collected:21, completionRate:23, invalidData:24, invalidRate:25, publishedData:26}`
   - settlements: `{month:0, module:1, type:2, qty:3, price:4, cost:5, lastQty:6, lastPrice:7, lastCost:8}`（新版无去年模块/类型）
   - preReview: `{reviewer:2, month:1, isNew:6, passed:7, priceAudit:8, reason:9}`
   - capacity: `{month:0, planQty:1, planHouse:2, planItem:3, actQty:4, actHouse:5, actItem:6, staff:7, note:8}`
4. **模板文件正被 Excel 锁定**（`~$` 锁文件），openpyxl 无法覆盖，需先让用户关闭 Excel
5. **浏览器导入与 Python 解析器的字段语义必须一致**——现在两侧均已实现 old_enterprise / price_audit / dueMonth / 去年字段，改任一侧须同步另一侧
6. **对话登记台账**：`需求下发登记.json` 记录 `demandItem` **必须给值**（默认"市场价"），否则前端默认筛选"市场价"下新记录不可见；`timelyRate` 缺省会按"不及时"计入（前端只认 `===1`），对话必须问清及时/延迟
7. **前端零改动原则**：台账新月份 `"MM月"` 严格字符串匹配已核验成立（`MONTH_KEYS_FULL` 仅硬编码 03-06，新月份走 fallback 原文比较）、null 字段显示"—"安全。改台账 schema 或合并逻辑时勿引入前端不认识的字段（`reqQty` 是例外，仅透传前端忽略）
8. **dueMonth MMDD 陷阱**：约定采集完成时间填的是无分隔数值（730=7月30）而非日期序列号。Python `fmt_due_month` 与前端 `fmtDueMonth` 都要先判 MMDD（`101≤n≤1231` 且 `n//100` 在 1-12）再走日期序列号逻辑，否则会误读成 1902 年
9. **问题反馈公式求值**：`问题反馈` sheet 的值可能带公式（如 `双方未达成共识量 =C2-C3`），Python 解析器必须用 `cell_val` 求值，直接读 `.value` 会存成字符串。浏览器导入读 Excel 缓存值通常无此问题，但改解析器时别回退成裸读取

## 对话式「需求下发登记记录」功能（**已实现**，2026-08-03）

### 用户诉求（已定案）
不用手动改表，直接跟 Claude 说自然语言，识别关键字段后**新增登记**（不修改旧的、不重复）。例：现在有1-6月数据，用户登记7月的。

### 已确认的规则
1. **分阶段登记，关联关系跟阶段走**：
   - **需求下发阶段**：登记 月份、需求名称、下发时间（是否延迟→需求发布及时率0/1）、下发材料、要求数量、下发数量
   - **采集阶段**（用户下次输入采集完成情况时）：登记 实际采集完成量、**是否完成/是否采集闭环**（下发阶段未知，不在此阶段登记）
   - **发布阶段**：发布数据
2. **关联关系**：下发数量(含多家) = 下发材料 × 要求数量；**要求数量 = 每个材料要求几家报价**（用户确认）；收到数量 = 材料数量 × 要求数量
3. **台账**：无独立 Excel 台账，指向**看板最终呈现的数据源（data.json）**，对话结果经 `需求下发登记.json` 沉淀后由解析器合并进 data.json
4. **去重**：同「需求名称+月份」已存在则提示，不重复登记

### 对话登记协议（我如何执行，务必遵守）
- **Step0 读上下文**：读 `data.json.tasks` + `需求下发登记.json.records`，建 `(name, month)` 现有键集
- **Step1 解析自然语言**（正则+关键词）：month→`MM月`、name、demandItem(缺省"市场价"标注待确认)、settleType(缺省"价格采集")、timelyRate(按时→1/延迟→0，**未提及必须反问**)、itemCount、reqQty(每家几家)、totalPoints、matQty(可选供准确率)、collected/closed、publishedData
- **Step2 关联核查+派生自算**：totalPoints=itemCount×reqQty；issuedQty=totalPoints；completionRate=collected/totalPoints×100；accuracyRate=itemCount/matQty×100（matQty缺省则null）
- **Step3 去重（对话层）**：`(name,month)` 在 data.json → 提示"模板已存在该需求"，不登记；在台账且 stages 未全完成 → **阶段补录**（按 id=`{month}-{name}` 定位，只覆盖本次字段+置 stages 对应日期）；stages 全完成 → 提示"已登记完整，不重复"
- **Step4 待确认指标清单**：表格展示解析+派生值，缺失字段标"？"，用户确认/修改
- **Step5 写台账**：新记录 append（stages 相应置今天 + createdAt/updatedAt）；补录只覆盖字段并 updatedAt；`json.dump(..., ensure_ascii=False, indent=2)`
- **Step6 更新看板**：`python parse_template.py 看板数据录入模板-调整.xlsx data.json`，确认控制台 `台账合并: 新增 N 条`
- **Step7 部署**：git add（含 `需求下发登记.json`、data.json、parse_template.py、CLAUDE.md）→ commit → push

### 台账 `需求下发登记.json` schema
`{"version":1,"records":[{...}]}`，记录字段 = data.json task 全集 + 新增：
- `id`=`"{month}-{name}"`（阶段补录定位键）
- `reqQty`（要求数量，前端忽略仅透传）
- `stages`=`{"release":"YYYY-MM-DD","collection":...,"publish":...}`（值=日期串已登记，null=未登记）
- `createdAt`/`updatedAt`（ISO 时间戳）

### 用户现有登记表参考
`需求下发登记表.xlsx`：「需求下发」sheet 33列（含"智能总结"自然语言列、需求名称、收到数量、材料数量、下发时间、及时率、下发数量、下发材料、采集闭环时间、是否采集闭环、采集及时率、是否走线上、是否全部闭环、需求月份、初审通过率等）；另有「供采试点下发需求」「结算明细」「工作表*」等。

## 沟通偏好
- 用中文交流
- 用户关心代码整洁、长期可维护，明确反对堆积无用代码（死分支/空数据壳）
- 重大/不可逆操作（回滚、覆盖、删除）先备份再执行
