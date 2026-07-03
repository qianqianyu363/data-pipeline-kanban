#!/usr/bin/env python3
"""
看板数据录入模板-调整.xlsx → data.json 解析脚本
读取新模板4个Sheet + 问题反馈，生成看板可用的 data.json
"""
import openpyxl, json, sys, os
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '看板数据录入模板-调整.xlsx')
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), 'data.json')

wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)

# ============================================================
# 1. 各项需求跟踪 → tasks
# ============================================================
ws1 = wb['各项需求跟踪']
tasks = []
month_names_set = set()

# 列映射 (列号→字段名)
COL_MAP = {
    2: 'month_raw',    # 月份
    3: 'demandItem',   # 需求项
    4: 'settleType',   # 结算类型
    5: 'name',         # 需求名称
    6: 'source',       # 需求来源
    12: 'totalPoints',  # 下发数量(含多家）
    13: 'itemCount',    # 下发材料
    14: 'reviewIssues', # 预审问题数据
    17: 'accuracyRate', # 需求准确率 (比率)
    21: 'collected',    # 实际采集完成量
    23: 'completionRate', # 采集完成率 (比率)
    24: 'invalidData',  # 无效数据
    25: 'invalidRate',  # 无效数据占比
    26: 'published',    # 发布数据
    11: 'timelyRate',   # 需求发布及时率 (0/1)
}

# 月份格式化
def fmt_month(m):
    m = str(m or '').strip()
    if '月' in m:
        num = m.replace('月', '').strip()
        if num.isdigit():
            return f"{int(num):02d}月"
        return m
    if m.isdigit():
        return f"{int(m):02d}月"
    return m

for r in range(3, ws1.max_row + 1):
    name = ws1.cell(row=r, column=5).value
    if not name:
        continue  # 跳过空行

    month_raw = ws1.cell(row=r, column=2).value
    month = fmt_month(month_raw)
    month_names_set.add(month)

    demand_item = ws1.cell(row=r, column=3).value or ''
    settle_type = ws1.cell(row=r, column=4).value or ''
    source = ws1.cell(row=r, column=6).value or ''

    def nv(col):
        v = ws1.cell(row=r, column=col).value
        if v is None:
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            return None

    total_pts = nv(12)
    item_cnt = nv(13)
    review_issues = nv(14)
    collected = nv(21)
    comp_rate = nv(23)
    accuracy = nv(17)
    invalid = nv(24)
    invalid_rate = nv(25)
    published = nv(26)
    timely = nv(11)

    task = {
        "name": str(name).strip(),
        "settleType": settle_type,
        "source": source if source else None,
        "demandItem": demand_item,
        "month": month,
        "totalPoints": int(total_pts) if total_pts is not None else None,
        "itemCount": int(item_cnt) if item_cnt is not None else None,
        "reviewIssues": int(review_issues) if review_issues is not None else None,
        "issuedQty": int(total_pts) if total_pts is not None else None,
        "collected": int(collected) if collected is not None else None,
        "completionRate": round(comp_rate * 100, 1) if comp_rate is not None and comp_rate != '#DIV/0!' else None,
        "accuracyRate": round(accuracy * 100, 1) if accuracy is not None and isinstance(accuracy, (int, float)) and accuracy != '#DIV/0!' else None,
        "invalidData": int(invalid) if invalid is not None else None,
        "invalidRate": round(invalid_rate * 100, 1) if invalid_rate is not None else None,
        "publishedData": int(published) if published is not None else None,
        "timelyRate": int(timely) if timely is not None else None,
        # 缺失字段置空
        "feedbackMat": None,
        "expertVerify": None,
        "supplierRejects": None,
        "closed": None,
        "allClosed": None,
        "notes": None,
    }
    tasks.append(task)

# 按月份排序
month_order = {f"{i:02d}月": i for i in range(1, 13)}
tasks.sort(key=lambda t: (month_order.get(t['month'], 99), t['name'] or ''))

# ============================================================
# 2. 预算与结算 → settlements
# ============================================================
ws2 = wb['预算与结算']
settlements = []
for r in range(3, ws2.max_row + 1):
    month = ws2.cell(row=r, column=1).value
    module = ws2.cell(row=r, column=2).value
    demand_type = ws2.cell(row=r, column=3).value
    qty = ws2.cell(row=r, column=4).value
    price = ws2.cell(row=r, column=5).value
    cost = ws2.cell(row=r, column=6).value

    if not month or not module or not demand_type:
        continue

    def to_num(v):
        if v is None: return 0
        try: return float(v)
        except: return 0

    settlements.append({
        "月份": str(month).strip(),
        "产品模块": str(module).strip(),
        "需求类型": str(demand_type).strip(),
        "需求数量": int(to_num(qty)),
        "结算单价": to_num(price),
        "结算金额": to_num(cost),
    })

# ============================================================
# 3. 质量抽审 → qualityReview
# ============================================================
ws3 = wb['质量抽审']
quality_review = {
    "公章抽审": {},
    "报价抽审": {},
}

current_task = None
for r in range(1, ws3.max_row + 1):
    a = ws3.cell(row=r, column=1).value  # 月份
    b = ws3.cell(row=r, column=2).value  # 抽审任务
    c = ws3.cell(row=r, column=3).value  # 维度
    d = ws3.cell(row=r, column=4).value  # 项
    e = ws3.cell(row=r, column=5).value  # 值

    if a and b:
        current_task = str(b).strip()
        if current_task not in quality_review:
            quality_review[current_task] = {"月份": str(a).strip()}

    if current_task and d and e is not None:
        quality_review[current_task][str(d).strip()] = e

# ============================================================
# 4. 加工预审 → preReview (统计摘要)
# ============================================================
ws4 = wb['加工预审']
preview_stats = {
    "total": 0,
    "passed": 0,
    "failed": 0,
    "by_reviewer": defaultdict(lambda: {"total": 0, "passed": 0}),
    "by_month": defaultdict(lambda: {"total": 0, "passed": 0}),
    "fail_reasons": defaultdict(int),
    "new_enterprise": 0,
}

has_data = False
for r in range(3, ws4.max_row + 1):
    reviewer = ws4.cell(row=r, column=3).value
    passed = ws4.cell(row=r, column=8).value
    month = ws4.cell(row=r, column=2).value
    reason = ws4.cell(row=r, column=9).value
    is_new = ws4.cell(row=r, column=7).value

    if not reviewer:
        continue
    has_data = True

    preview_stats["total"] += 1
    passed_flag = str(passed or '').strip() == '是'
    if passed_flag:
        preview_stats["passed"] += 1
    else:
        preview_stats["failed"] += 1

    reviewer_name = str(reviewer).strip()
    preview_stats["by_reviewer"][reviewer_name]["total"] += 1
    if passed_flag:
        preview_stats["by_reviewer"][reviewer_name]["passed"] += 1

    m = str(month or '').strip()
    if m:
        preview_stats["by_month"][m]["total"] += 1
        if passed_flag:
            preview_stats["by_month"][m]["passed"] += 1

    if reason and not passed_flag:
        rsn = str(reason).strip()
        if rsn:
            preview_stats["fail_reasons"][rsn] += 1

    if is_new and str(is_new).strip() in ('是', '新增'):
        preview_stats["new_enterprise"] += 1

# 转换defaultdict为普通dict
if has_data:
    preview_stats["by_reviewer"] = dict(preview_stats["by_reviewer"])
    preview_stats["by_month"] = dict(preview_stats["by_month"])
    preview_stats["fail_reasons"] = dict(preview_stats["fail_reasons"])
    # 计算通过率
    total = preview_stats["total"]
    preview_stats["passRate"] = round(preview_stats["passed"] / total * 100, 1) if total else 0

    for rv in preview_stats["by_reviewer"].values():
        rv["passRate"] = round(rv["passed"] / rv["total"] * 100, 1) if rv["total"] else 0
    for bm in preview_stats["by_month"].values():
        bm["passRate"] = round(bm["passed"] / bm["total"] * 100, 1) if bm["total"] else 0
else:
    preview_stats = None

# ============================================================
# 5. 问题反馈 → issueFeedback
# ============================================================
ws5 = wb['问题反馈']
issue_feedback = {}
current_section = "main"
issue_labels = {}

for r in range(1, ws5.max_row + 1):
    label = ws5.cell(row=r, column=2).value
    val = ws5.cell(row=r, column=3).value

    if not label:
        continue
    label = str(label).strip()

    if label == "已共识问题标签":
        current_section = "labels"
        continue

    if current_section == "main":
        try:
            issue_feedback[label] = int(float(val)) if val is not None else None
        except (ValueError, TypeError):
            issue_feedback[label] = val
    else:
        try:
            issue_labels[label] = int(float(val)) if val is not None else None
        except (ValueError, TypeError):
            issue_labels[label] = val

if issue_labels:
    issue_feedback["已共识问题标签"] = issue_labels

# ============================================================
# 6. 计算月度趋势 monthlyTrend
# ============================================================
monthly_trend = defaultdict(list)
months_in_data = sorted(set(
    f"{int(m.split('月')[0]):02d}月" for m in month_names_set
    if m and '月' in m and m.split('月')[0].isdigit()
), key=lambda x: int(x.split('月')[0]))

for m in months_in_data:
    month_tasks = [t for t in tasks if t['month'] == m]
    total_pts = sum(t['totalPoints'] or 0 for t in month_tasks)
    total_issued = sum(t['issuedQty'] or 0 for t in month_tasks)
    total_collected = sum(t['collected'] or 0 for t in month_tasks)
    total_review = sum(t['reviewIssues'] or 0 for t in month_tasks)
    total_invalid = sum(t['invalidData'] or 0 for t in month_tasks)

    # 采集完成率
    comp_rate = round(total_collected / total_issued * 100, 1) if total_issued else 0
    # 需求准确率
    acc_rate = round((1 - total_review / total_pts) * 100, 1) if total_pts else 0
    # 无效数据率
    inv_rate = round(total_invalid / total_collected * 100, 1) if total_collected else 0

    monthly_trend["需求总条数"].append(total_pts)
    monthly_trend["采集下发数"].append(total_issued)
    monthly_trend["采集完成数"].append(total_collected)
    monthly_trend["采集完成率"].append(comp_rate)
    monthly_trend["需求准确率"].append(acc_rate)
    monthly_trend["无效数据数"].append(total_invalid)
    monthly_trend["无效数据率"].append(inv_rate)

# ============================================================
# 7. 构建完整 data.json
# ============================================================
# 收集唯一值
demand_items = sorted(set(t['demandItem'] for t in tasks if t['demandItem']))
settle_types = sorted(set(t['settleType'] for t in tasks if t['settleType']))
sources = sorted(set(t['source'] for t in tasks if t['source']))

month_keys = [m.replace('月', '月') for m in months_in_data]
month_names = {m.replace('月', '').zfill(2): m for m in months_in_data}

# 构建monthlyTrend的key（按月）
monthly_trend_dict = dict(monthly_trend)

# 额外：从结算数据算预算执行率和成本执行率
total_budget = 3900000
total_cost = sum(s['结算金额'] for s in settlements)
budget_rate = round(total_cost / total_budget * 100, 1) if total_budget else 0

# 为monthlyTrend补充预算执行率和成本执行率
# 按月份计算预算执行率
for m_idx, m_name in enumerate(months_in_data):
    m_num = m_name.replace('月', '')
    month_settlements = [s for s in settlements if s['月份'].replace('月', '') == m_num or s['月份'] == m_num + '月']
    month_cost = sum(s['结算金额'] for s in month_settlements)
    month_budget_rate = round(month_cost / (total_budget / 6) * 100, 1) if (total_budget / 6) else 0

    if '预算执行率' not in monthly_trend_dict:
        monthly_trend_dict['预算执行率'] = []
    if '成本执行率' not in monthly_trend_dict:
        monthly_trend_dict['成本执行率'] = []

    # 累计执行率
    cumulative_cost = sum(s['结算金额'] for s in settlements
                         if int(s['月份'].replace('月', '')) <= int(m_num))
    cum_budget_rate = round(cumulative_cost / total_budget * 100, 1) if total_budget else 0

    if len(monthly_trend_dict['预算执行率']) < len(months_in_data):
        monthly_trend_dict['预算执行率'].append(cum_budget_rate)
    if len(monthly_trend_dict['成本执行率']) < len(months_in_data):
        monthly_trend_dict['成本执行率'].append(cum_budget_rate)

data = {
    "title": "需求下发登记全链路看板",
    "updateDate": "2026-07-02",
    "months": month_keys + ["07月", "08月"],
    "monthNames": month_names,
    "filterOptions": {
        "demandItems": ["全部"] + demand_items,
        "settleTypes": ["全部"] + settle_types,
        "sources": ["全部"] + sources,
        "months": ["全部"] + month_keys + ["07月", "08月"],
    },
    "tasks": tasks,
    "settlements": settlements,
    "totalBudget": total_budget,
    "stepConfig": {
        "需求下发": {
            "color": "#E74C3C",
            "icon": "📋",
            "metrics": [
                {"key": "totalPoints", "label": "需求总条数", "unit": "条"},
                {"key": "itemCount", "label": "需求项数", "unit": "项"},
                {"key": "reviewIssues", "label": "初审问题数", "unit": "条"},
                {"key": "reviewIssueRate", "label": "初审问题率", "unit": "%"},
                {"key": "reviewPassRate", "label": "初审通过率", "unit": "%"},
                {"key": "timelyRate", "label": "发布及时率", "unit": "%"},
            ]
        },
        "采集执行": {
            "color": "#F39C12",
            "icon": "🔧",
            "metrics": [
                {"key": "issuedQty", "label": "采集下发数", "unit": "条"},
                {"key": "collected", "label": "采集完成数", "unit": "条"},
                {"key": "completionRate", "label": "采集完成率", "unit": "%"},
                {"key": "invalidData", "label": "无效数据", "unit": "条"},
                {"key": "publishedData", "label": "发布数据", "unit": "条"},
            ]
        },
        "数据加工/审核": {
            "color": "#F1C40F",
            "icon": "🔍",
            "metrics": [
                {"key": "accuracyRate", "label": "需求准确率", "unit": "%"},
                {"key": "feedbackMat", "label": "材料反馈数", "unit": "条"},
                {"key": "expertVerify", "label": "专家核实通过数", "unit": "条"},
            ]
        },
        "上线发布": {
            "color": "#2ECC71",
            "icon": "🚀",
            "metrics": [
                {"key": "closedRate", "label": "闭环率", "unit": "%"},
                {"key": "allClosedRate", "label": "全部闭环率", "unit": "%"},
            ]
        },
    },
    "specialProjects": {
        "采购价更新": {
            "color": "#9B59B6",
            "icon": "🟣",
            "desc": "月度大批量更新需求专项",
            "metrics": [
                {"key": "updateDemand", "label": "更新需求量", "unit": "条"},
                {"key": "updateCompletion", "label": "更新完成率", "unit": "%"},
            ]
        },
        "结算与成本": {
            "color": "#1ABC9C",
            "icon": "💰",
            "desc": "外包、专家人力成本预算管控",
            "metrics": [
                {"key": "budgetRate", "label": "预算执行率", "unit": "%"},
                {"key": "totalCost", "label": "累计成本", "unit": "万"},
            ]
        },
    },
    "alerts": [
        {
            "condition": "completionRate<60",
            "priority": "P0",
            "problem": "{name} 采集完成率 {value}",
            "action": "专项攻坚，增配采集资源",
        },
        {
            "condition": "reviewIssues>0",
            "priority": "P1",
            "problem": "{name} 初审问题数 {value}",
            "action": "追溯源头，优化需求质量",
        },
        {
            "condition": "invalidData>0",
            "priority": "P1",
            "problem": "{name} 无效数据 {value} 条",
            "action": "排查无效数据原因，提升数据质量",
        },
    ],
    "monthlyTrend": monthly_trend_dict,
    # 新模板扩展数据
    "qualityReview": quality_review,
    "preReview": preview_stats,
    "issueFeedback": issue_feedback if issue_feedback else None,
}

# 写入文件
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f"✅ 解析完成！共 {len(tasks)} 条任务, {len(settlements)} 条结算记录")
print(f"   月份: {months_in_data}")
print(f"   需求项: {demand_items}")
print(f"   结算类型: {settle_types}")
print(f"   问题反馈: {issue_feedback.get('问题反馈量', 'N/A') if issue_feedback else '无'}")
if preview_stats:
    print(f"   加工预审: {preview_stats['total']} 条, 通过率 {preview_stats.get('passRate', 0)}%")
print(f"   已写入: {OUTPUT_PATH}")
