#!/usr/bin/env python3
"""
看板数据录入模板-调整.xlsx → data.json 解析脚本

架构：**Excel 模板是唯一数据源**。日常数据（任务/结算/产能/质量抽审）通过两种方式进 Excel：
  1. 手动编辑模板（Excel 或 Claude 直接改）
  2. 对话登记：Claude 解析自然语言 → 写入"待追加文件" → `--append` 追加进 Excel
然后本脚本把 Excel 解析成 data.json（含公式自算、内嵌同步、GitHub Pages 部署）。

用法：
  python parse_template.py                       # 默认模板 → data.json
  python parse_template.py [模板.xlsx] [输出.json]
  python parse_template.py [模板.xlsx] [输出.json] --append 对话登记_待追加.json
      # --append: 把待追加文件里的记录先追加进 Excel（只补不覆盖，幂等），再解析
"""
import openpyxl, json, sys, os, re, datetime, subprocess
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')


def sync_embed_data(index_html_path, data):
    """把数据【加密后】同步到 index.html 内嵌块，并生成 data.enc.json。

    ⚠️ 安全约束：内嵌块与 data.enc.json 只允许存放密文。
       明文 data.json 仅保留在本地（已加入 .gitignore），切勿提交。
       加密由 encrypt_data.js 完成（AES-256-GCM / PBKDF2-SHA256），
       与浏览器 WebCrypto 完全兼容，前端零依赖解密。
    """
    base = os.path.dirname(os.path.abspath(__file__))
    script = os.path.join(base, 'encrypt_data.js')
    if not os.path.exists(script):
        print("   ⚠️ 未找到 encrypt_data.js，跳过加密同步（内嵌块保持原样）")
        return False
    try:
        proc = subprocess.run(
            ['node', script],
            cwd=base, capture_output=True, text=True,
            encoding='utf-8', errors='replace', timeout=180,
        )
    except FileNotFoundError:
        print("   ⚠️ 未找到 node，跳过加密同步（请安装 Node.js 后重跑本脚本）")
        return False
    except subprocess.TimeoutExpired:
        print("   ⚠️ 加密超时，跳过内嵌同步")
        return False

    for line in (proc.stdout or '').strip().splitlines():
        print("  " + line.strip())
    if proc.returncode != 0:
        print(f"   ❌ 加密失败: {(proc.stderr or '').strip()[:300]}")
        return False
    return True


# 支持命令行：python parse_template.py [模板.xlsx] [输出.json] [--append 待追加.json]
#   --append <file>  对话登记待追加文件（Claude 解析自然语言后写入；先追加到 Excel 再解析）。
#                    追加遵循"只补不覆盖"：缺失键追加新行、已有行只填空单元格；幂等可重跑。
_default_tpl = '看板数据录入模板-调整.xlsx'
_default_out = 'data.json'
_append_file = None
_pos_args = []
_args = sys.argv[1:]
_i = 0
while _i < len(_args):
    if _args[_i] == '--append':
        if _i + 1 < len(_args):
            _append_file = _args[_i + 1]
            _i += 2
            continue
    else:
        _pos_args.append(_args[_i])
    _i += 1
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), _pos_args[0] if len(_pos_args) > 0 else _default_tpl)
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), _pos_args[1] if len(_pos_args) > 1 else _default_out)
APPEND_PATH = os.path.join(os.path.dirname(__file__), _append_file) if _append_file else None

# 注意：用 data_only=False 保留公式文本，配合下方公式求值器自算派生值。
# openpyxl 保存会丢公式缓存（升级版模板 data_only=True 读出来全是 None），故不依赖缓存。
wb = openpyxl.load_workbook(TEMPLATE_PATH)


def colnum(letter):
    """列字母 → 数字（A=1, B=2 …）"""
    n = 0
    for ch in str(letter).upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def eval_formula(formula, row, get_val, seen=None):
    """求值模板公式（=A{r}*K、=A-B、=A/B 等算术，列引用限同行）。
    非公式值原样返回；公式求值失败返回 None。"""
    if seen is None:
        seen = set()
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    expr = formula[1:]
    def repl(m):
        col, r = m.group(1).replace('$', ''), int(m.group(2))
        key = (col.upper(), r)
        if key in seen:  # 防循环引用
            return '0'
        seen.add(key)
        v = get_val(col.upper(), r)
        if isinstance(v, str) and v.startswith('='):
            v = eval_formula(v, row, get_val, seen)
        try:
            return str(float(v))
        except (TypeError, ValueError):
            return '0'
    expr = re.sub(r'(\$?[A-Za-z]{1,3}\$?)(\d+)', repl, expr)
    try:
        return float(eval(expr, {'__builtins__': {}}, {}))
    except Exception:
        return None


def cell_val(ws, col, r):
    """取单元格值：公式单元格递归求值，非公式返回原值。"""
    if isinstance(col, int):
        from openpyxl.utils import get_column_letter
        col = get_column_letter(col)
    v = ws.cell(row=r, column=colnum(col)).value
    if isinstance(v, str) and v.startswith('='):
        return eval_formula(v, r, lambda c, rr: cell_val(ws, c, rr))
    return v


# ============================================================
# 月份/日期 格式化工具
# ============================================================
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


def fmt_settle_month(m):
    """结算/产能/质量 的月份 → 'N月'（无前导零，对齐模板 settlements/collectionCapacity/qualityReview）。
    兼容 '08月'/'8月'/'08'/'8'；非 1-12 月或无法解析时原样返回。"""
    m = str(m or '').strip()
    digits = ''.join(c for c in m if c.isdigit())
    if not digits:
        return m
    n = int(digits)
    if 1 <= n <= 12:
        return f"{n}月"
    return m


def fmt_due_month(v):
    """约定采集完成时间 → 'MM月'。兼容 datetime/日期序列号/文本('7月'、'2026-07-31')/MMDD数值(730→07月)"""
    if v is None or v == '':
        return None
    if isinstance(v, datetime.datetime):
        return f"{v.month:02d}月"
    if isinstance(v, datetime.date):
        return f"{v.month:02d}月"
    if isinstance(v, (int, float)):
        # MMDD 无分隔数值（如 730=7月30、1030=10月30）；真实日期序列号远大于此范围
        n = int(v)
        if 101 <= n <= 1231:
            month = n // 100
            if 1 <= month <= 12:
                return f"{month:02d}月"
        try:  # Excel 日期序列号（1900-01-01 基准，实际用 1899-12-30）
            dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=float(v))
            return f"{dt.month:02d}月"
        except Exception:
            return None
    s = str(v).strip()
    if '月' in s:  # '7月' / '2026年7月'
        digits = ''.join(c for c in s.replace('年', '') if c.isdigit())
        return f"{int(digits):02d}月" if digits else None
    for sep in ('-', '.', '/'):  # '2026-07-31'
        if sep in s:
            parts = s.split(sep)
            if len(parts) >= 2 and parts[1].strip().isdigit():
                return f"{int(parts[1].strip()):02d}月"
    return None


def fmt_collect_time(v):
    """采集完成时间 → 'YYYY-MM-DD'（完整日期序列号）或 'MM月'（MMDD 数值）。"""
    if v is None or v == '':
        return None
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, (int, float)):
        n = int(v)
        if 101 <= n <= 1231 and 1 <= n // 100 <= 12:  # MMDD 无分隔（如 730=7月30）
            return f"{n // 100:02d}月"
        try:  # Excel 日期序列号
            dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=float(v))
            return dt.isoformat()
        except Exception:
            return None
    s = str(v).strip()
    for sep in ('-', '.', '/'):  # '2026-08-01'
        parts = s.split(sep)
        if len(parts) >= 3 and parts[0].strip().isdigit() and len(parts[0].strip()) == 4:
            try:
                return datetime.date(int(parts[0]), int(parts[1]), int(parts[2])).isoformat()
            except Exception:
                pass
    return s or None


# ============================================================
# 1.5 对话登记追加（--append）：把待追加记录写进 Excel，Excel 是唯一数据源
# 原则：只补不覆盖。缺失键追加新行、已有行只填空单元格（None/空串），绝不动已有值/公式。
# 幂等：已追加过的记录重跑会被跳过（键已存在）；有变更才备份模板并保存。
# ============================================================
def _norm_num(v):
    if v is None or v == '':
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _to_int(v):
    return int(v) if v is not None else None


def _to_pct(v):
    return round(v, 1) if v is not None else None


def _coerce_item(v):
    """质量抽审 item 值：数字转 int/float，文本保留。"""
    n = _norm_num(v)
    if n is not None:
        return int(n) if float(n).is_integer() else n
    return v


def _fill_row_cells(ws, r, col_map, rec, zero_as_empty_cols=()):
    """按 col_map{列号: 字段名} 把 rec 中非空、对应单元格为空(None/空串)的字段写入该行。
    zero_as_empty_cols 内的列号，单元格值 0 也视为空（如结算未来月 qty=0）。
    返回是否写入了任何单元格。"""
    changed = False
    for c, field in col_map.items():
        v = rec.get(field)
        if v is None:
            continue
        cur = ws.cell(row=r, column=c).value
        if cur is None or (isinstance(cur, str) and not cur.strip()):
            ws.cell(row=r, column=c).value = v
            changed = True
        elif c in zero_as_empty_cols and _norm_num(cur) == 0 and _norm_num(v) is not None:
            ws.cell(row=r, column=c).value = v
            changed = True
    return changed


def _append_row_values(ws, r, col_map, rec):
    """把 rec 按 col_map 写入新行 r（仅写非空字段），返回写入的字段数。"""
    n = 0
    for c, field in col_map.items():
        v = rec.get(field)
        if v is None:
            continue
        ws.cell(row=r, column=c).value = v
        n += 1
    return n


def _backup_template():
    import shutil
    bdir = os.path.join(os.path.dirname(__file__), '.workbuddy', 'backup_' + datetime.date.today().strftime('%Y%m%d'))
    os.makedirs(bdir, exist_ok=True)
    dst = os.path.join(bdir, os.path.basename(TEMPLATE_PATH))
    if not os.path.exists(dst):
        shutil.copy2(TEMPLATE_PATH, dst)
        print(f"   🗂️ 模板已备份至: {dst}")
    return dst


def apply_dialogue_to_template(reg):
    """把待追加 reg 的 4 个 section（records/settlements/capacity/quality）写进 Excel 对应 sheet。
    只补不覆盖；有变更才备份+保存；Excel 被占用时警告（data.json 仍可由内存生成）。返回变更行数。"""
    if not reg or not any(reg.get(k) for k in ('records', 'settlements', 'capacity', 'quality')):
        return 0
    changes = 0

    # ---- 1. 各项需求跟踪（任务）：键=(月份, 名称) ----
    ws = wb['各项需求跟踪']
    TASK_APPEND = {2: 'month', 3: 'demandItem', 4: 'settleType', 5: 'name', 6: 'source', 12: 'dueMonth',
                   8: 'matQty', 14: 'timelyRate', 15: 'totalPoints', 16: 'itemCount', 17: 'reviewIssues',
                   20: 'accuracyRate', 22: 'collectedTime', 24: 'collected', 28: 'completionRate',
                   29: 'invalidData', 30: 'invalidRate', 31: 'publishedData', 35: 'closed'}
    TASK_FILL = {8: 'matQty', 14: 'timelyRate', 15: 'totalPoints', 16: 'itemCount', 17: 'reviewIssues',
                 20: 'accuracyRate', 22: 'collectedTime', 24: 'collected', 28: 'completionRate',
                 29: 'invalidData', 30: 'invalidRate', 31: 'publishedData', 35: 'closed'}
    tpl_rows = {}
    for r in range(3, ws.max_row + 1):
        nm = str(ws.cell(row=r, column=5).value or '').strip()
        if nm:
            tpl_rows[(fmt_month(ws.cell(row=r, column=2).value), nm)] = r
    next_r = ws.max_row + 1
    for rec in (reg.get('records') or []):
        if not isinstance(rec, dict):
            continue
        name = str(rec.get('name') or '').strip()
        month = fmt_month(rec.get('month'))
        if not name or not month or '月' not in month:
            continue
        if (month, name) in tpl_rows:
            if _fill_row_cells(ws, tpl_rows[(month, name)], TASK_FILL, rec):
                changes += 1
                print(f"   🔄 追加填空: 任务 {name}/{month}")
        else:
            _append_row_values(ws, next_r, TASK_APPEND, dict(rec, month=month))
            tpl_rows[(month, name)] = next_r
            next_r += 1
            changes += 1
            print(f"   ➕ 追加新任务行: {name}/{month}")

    # ---- 2. 预算与结算（结算）：键=(月份, 模块, 类型) ----
    ws2 = wb['预算与结算']
    st_rows = {}
    for r in range(4, ws2.max_row + 1):
        mod = str(ws2.cell(row=r, column=2).value or '').strip()
        typ = str(ws2.cell(row=r, column=3).value or '').strip()
        mon = fmt_settle_month(ws2.cell(row=r, column=1).value)
        if mod and typ and mon:
            st_rows[(mon, mod, typ)] = r
    next_r = ws2.max_row + 1
    for rec in (reg.get('settlements') or []):
        if not isinstance(rec, dict):
            continue
        month = fmt_settle_month(rec.get('月份') or rec.get('month'))
        module = str(rec.get('产品模块') or rec.get('module') or '市场价').strip()
        dtype = str(rec.get('需求类型') or rec.get('type') or '价格采集').strip()
        qty = _norm_num(rec.get('需求数量') or rec.get('qty'))
        price = _norm_num(rec.get('结算单价') or rec.get('price'))
        if not month or qty is None or price is None:
            continue
        if (month, module, dtype) in st_rows:
            r = st_rows[(month, module, dtype)]
            # 未来月占位行 qty=0 视为空，只填空并重算金额
            if _fill_row_cells(ws2, r, {4: '需求数量', 5: '结算单价'}, rec, zero_as_empty_cols=(4,)):
                ws2.cell(row=r, column=6).value = round(qty * price, 2)
                changes += 1
                print(f"   🔄 追加填空: 结算 {module}/{dtype}/{month}")
        else:
            ws2.cell(row=next_r, column=1).value = month
            ws2.cell(row=next_r, column=2).value = module
            ws2.cell(row=next_r, column=3).value = dtype
            ws2.cell(row=next_r, column=4).value = int(qty)
            ws2.cell(row=next_r, column=5).value = round(price, 2)
            ws2.cell(row=next_r, column=6).value = round(qty * price, 2)
            st_rows[(month, module, dtype)] = next_r
            next_r += 1
            changes += 1
            print(f"   ➕ 追加新结算行: {module}/{dtype}/{month} qty={qty} price={price}")

    # ---- 3. 采集产能：键=月份 ----
    ws3 = wb['采集产能']
    CAP_APPEND = {2: '计划需求量', 3: '计划采集家数', 4: '计划采集条数', 5: '实际需求量',
                  6: '实际采集家数', 7: '实际采集条数', 8: '员工总人数', 9: '备注'}
    cap_rows = {}
    for r in range(2, ws3.max_row + 1):
        mon = fmt_settle_month(ws3.cell(row=r, column=1).value)
        if mon:
            cap_rows[mon] = r
    next_r = ws3.max_row + 1
    for rec in (reg.get('capacity') or []):
        if not isinstance(rec, dict):
            continue
        month = fmt_settle_month(rec.get('月份') or rec.get('month'))
        if not month:
            continue
        if month in cap_rows:
            if _fill_row_cells(ws3, cap_rows[month], CAP_APPEND, rec):
                changes += 1
                print(f"   🔄 追加填空: 产能 {month}")
        else:
            _append_row_values(ws3, next_r, CAP_APPEND, rec)
            ws3.cell(row=next_r, column=1).value = month
            cap_rows[month] = next_r
            next_r += 1
            changes += 1
            print(f"   ➕ 追加新产能行: {month}")

    # ---- 4. 质量抽审：键=(任务, 月份)，新块逐 item 追加行 ----
    ws4 = wb['质量抽审']
    qz_blocks = {}  # (任务, 月份) -> 起始行
    for r in range(2, ws4.max_row + 1):
        task = str(ws4.cell(row=r, column=2).value or '').strip()
        mon = fmt_settle_month(ws4.cell(row=r, column=1).value)
        if task and mon:
            qz_blocks[(task, mon)] = r
    next_r = ws4.max_row + 1
    for rec in (reg.get('quality') or []):
        if not isinstance(rec, dict):
            continue
        task = str(rec.get('任务') or rec.get('task') or '').strip()
        month = fmt_settle_month(rec.get('月份') or rec.get('month'))
        items = rec.get('items') or {}
        if not task or not month or not isinstance(items, dict):
            continue
        if (task, month) in qz_blocks:
            # 已有块：只填空（逐项查同块内该 项 行）
            blk_start = qz_blocks[(task, month)]
            r = blk_start
            while r <= ws4.max_row:
                nm = str(ws4.cell(row=r, column=4).value or '').strip()
                if ws4.cell(row=r, column=2).value and r > blk_start:
                    break  # 进入下一块
                if nm in items and ws4.cell(row=r, column=5).value is None:
                    ws4.cell(row=r, column=5).value = items[nm]
                    changes += 1
                r += 1
            if changes:
                print(f"   🔄 追加填空: 质量 {task}/{month}")
        else:
            first = True
            for k, v in items.items():
                if first:
                    ws4.cell(row=next_r, column=1).value = month
                    ws4.cell(row=next_r, column=2).value = task
                    first = False
                ws4.cell(row=next_r, column=4).value = k
                ws4.cell(row=next_r, column=5).value = v
                next_r += 1
            qz_blocks[(task, month)] = next_r
            changes += 1
            print(f"   ➕ 追加新质量块: {task}/{month} ({len(items)} 项)")

    if changes:
        _backup_template()
        try:
            wb.save(TEMPLATE_PATH)
            print(f"   ✅ 对话登记已追加进模板（{changes} 处变更）: {TEMPLATE_PATH}")
        except PermissionError:
            print("   ⚠️ 模板被 Excel 占用，跳过保存（data.json 仍按内存生成；关闭 Excel 后重跑 --append 即可补上）")
        except Exception as e:
            print(f"   ⚠️ 模板追加保存失败: {e}")
    else:
        print("   ℹ️ 待追加记录已在模板中（或为空），无新增")
    return changes


# ---- 对话登记追加：--append 提供 Claude 解析好的待追加记录，先写进 Excel 再解析 ----
if APPEND_PATH:
    if not os.path.exists(APPEND_PATH):
        print(f"   ⚠️ 待追加文件不存在，跳过追加: {APPEND_PATH}")
    else:
        try:
            with open(APPEND_PATH, 'r', encoding='utf-8') as f:
                _reg = json.load(f)
            apply_dialogue_to_template(_reg)
        except Exception as e:
            print(f"   ⚠️ 待追加文件处理失败: {e}")


# ============================================================
# 1. 各项需求跟踪 → tasks
# ============================================================
ws1 = wb['各项需求跟踪']
tasks = []
month_names_set = set()

# 列映射 (列号→字段名) — 适配 08-28 升级版模板（各项需求跟踪 39 列重排：条数/项数细分）
COL_MAP = {
    2: 'month_raw',    # 月份
    3: 'demandItem',   # 需求项
    4: 'settleType',   # 结算类型
    5: 'name',         # 需求名称
    6: 'source',       # 需求来源
    8: 'matQty',       # 需求项数（需求准确率分母）
    12: 'dueMonth',    # 采集承诺完成时间（MMDD）
    14: 'timelyRate',  # 需求发布及时率 (0/1)
    15: 'totalPoints', # 下发条数（含多家）
    16: 'itemCount',   # 下发项数
    17: 'reviewIssues',# 预审问题项数
    20: 'accuracyRate',# 需求准确率（自算 = itemCount/matQty）
    22: 'collectedTime',  # 实际采集完成时间
    24: 'collected',   # 实际采集条数
    28: 'completionRate', # 采集完成率（自算 = collected/totalPoints）
    29: 'invalidData', # 无效数据
    30: 'invalidRate', # 无效数据占比（自算 = invalidData/itemCount）
    31: 'published',   # 发布数据
    35: 'closed',      # 闭环状态
}

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
        v = cell_val(ws1, col, r)   # 公式单元格自动求值
        if v is None:
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            return None

    total_pts = nv(15)
    item_cnt = nv(16)
    mat_qty = nv(8)        # 需求项数（需求准确率分母）
    review_issues = nv(17)
    collected = nv(24)
    comp_rate = nv(28)     # 模板缓存（升级版 openpyxl 保存后无缓存，通常 None）
    accuracy = nv(20)      # 模板缓存
    invalid = nv(29)
    invalid_rate = nv(30)  # 模板缓存
    published = nv(31)
    timely = nv(14)
    due_month_raw = ws1.cell(row=r, column=12).value   # 采集承诺完成时间（MMDD）
    collect_time_raw = ws1.cell(row=r, column=22).value  # 实际采集完成时间
    closed_raw = ws1.cell(row=r, column=35).value     # 闭环状态

    # ---- 派生值自算（openpyxl 保存丢公式缓存，必须自算，遵循模板公式语义）----
    # 下发条数兜底 = 下发项数×3（模板公式 =P*3；下发条数空但下发项数有值时）
    if total_pts is None and item_cnt is not None:
        total_pts = item_cnt * 3
    # 采集完成率 = 实际采集条数 / 下发条数
    if comp_rate is None and total_pts:
        comp_rate = collected / total_pts if collected is not None else None
    # 需求准确率 = 下发项数 / 需求项数（模板公式 =P/H 语义，非 1-预审问题占比）
    if accuracy is None and mat_qty:
        accuracy = item_cnt / mat_qty if item_cnt is not None else None
    # 无效数据占比 = 无效数据 / 下发项数（模板公式 =AC/P 语义）
    if invalid_rate is None and item_cnt:
        invalid_rate = invalid / item_cnt if invalid is not None else None
    due_month = fmt_due_month(due_month_raw)

    task = {
        "name": str(name).strip(),
        "settleType": settle_type,
        "source": source if source else None,
        "demandItem": demand_item,
        "month": month,
        "dueMonth": due_month,   # 约定采集完成时间（新，用于当月采集完成率）
        "collectedTime": fmt_collect_time(collect_time_raw),  # 采集完成时间（新，08-17）
        "totalPoints": int(total_pts) if total_pts is not None else None,
        "itemCount": int(item_cnt) if item_cnt is not None else None,
        "matQty": int(mat_qty) if mat_qty is not None else None,   # 需求项数（准确率分母，前端聚合用）
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
        "closed": str(closed_raw).strip() if closed_raw is not None else None,  # 闭环状态（新，08-17）
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
# 新版模板：row1=标题'结算明细'、row2=分组表头(26年 col4-6 / 25年 col7-9)、row3=子表头(需求数量/结算单价/支出费用)、row4 起数据
for r in range(4, ws2.max_row + 1):
    month = ws2.cell(row=r, column=1).value
    module = ws2.cell(row=r, column=2).value
    demand_type = ws2.cell(row=r, column=3).value
    qty = cell_val(ws2, 4, r)
    price = cell_val(ws2, 5, r)
    # 25年对比 3 列（col7-9：去年需求数量/结算单价/支出费用；模板已删去年模块/需求类型）
    last_qty = cell_val(ws2, 7, r)
    last_price = cell_val(ws2, 8, r)

    if not month or not module or not demand_type:
        continue

    def to_num(v):
        if v is None: return 0
        try: return float(v)
        except: return 0

    q = to_num(qty); p = to_num(price)
    lq = to_num(last_qty); lp = to_num(last_price)
    has_last = last_qty is not None or last_price is not None

    settlements.append({
        "月份": str(month).strip(),
        "产品模块": str(module).strip(),
        "需求类型": str(demand_type).strip(),
        "需求数量": int(q),
        "结算单价": p,
        # 结算金额 = 需求数量 × 结算单价（openpyxl 丢缓存，自算）
        "结算金额": round(q * p, 2),
        # 去年对比（新，供前端"费用执行去年对比"）
        "去年需求数量": int(lq) if has_last else None,
        "去年结算单价": round(lp, 2) if last_price is not None else None,
        # 去年支出费用 = 去年需求数量 × 去年结算单价
        "去年支出费用": round(lq * lp, 2) if last_price is not None else None,
    })

# ============================================================
# 3. 质量抽审 → qualityReview (by_month 月度格式)
# ============================================================
ws3 = wb['质量抽审']
quality_review = {}

current_task = None
current_month = None
for r in range(1, ws3.max_row + 1):
    a = ws3.cell(row=r, column=1).value  # 月份
    b = ws3.cell(row=r, column=2).value  # 抽审任务
    c = ws3.cell(row=r, column=3).value  # 维度
    d = ws3.cell(row=r, column=4).value  # 项
    e = ws3.cell(row=r, column=5).value  # 值

    if a and b:
        current_task = str(b).strip()
        current_month = str(a).strip()
        if current_task not in quality_review:
            quality_review[current_task] = {"by_month": {}}
        if current_month not in quality_review[current_task]["by_month"]:
            quality_review[current_task]["by_month"][current_month] = {}

    if current_task and current_month and d is not None and e is not None:
        # 公式值（如 =E2-E3、=(E16+E14)/E13）经求值器还原为数值，避免前端显示公式文本
        if isinstance(e, str) and e.startswith('='):
            e = cell_val(ws3, 5, r)
        quality_review[current_task]["by_month"][current_month][str(d).strip()] = e

# ============================================================
# 4. 加工预审 → preReview (统计摘要)
# ============================================================
ws4 = wb['加工预审']
preview_stats = {
    "total": 0,
    "passed": 0,
    "failed": 0,
    "by_reviewer": defaultdict(lambda: {"total": 0, "passed": 0}),
    "by_month": defaultdict(lambda: {"total": 0, "passed": 0, "new_enterprise": 0, "old_enterprise": 0, "price_audit": {"total": 0, "passed": 0}}),
    "fail_reasons": defaultdict(int),
    "new_enterprise": 0,
    "old_enterprise": 0,                       # 老供应商（更新登记表）
    "price_audit": {"total": 0, "passed": 0},  # 价格审核（新，col9）
}

has_data = False
for r in range(3, ws4.max_row + 1):
    reviewer = ws4.cell(row=r, column=3).value
    passed = ws4.cell(row=r, column=8).value
    month = ws4.cell(row=r, column=2).value
    reason = ws4.cell(row=r, column=10).value      # 升级后"不通过原因"顺延 col10
    is_new = ws4.cell(row=r, column=7).value
    price_audit = ws4.cell(row=r, column=9).value  # 新增"价格审核结果"

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
        bm = preview_stats["by_month"][m]
        bm["total"] += 1
        if passed_flag:
            bm["passed"] += 1
        _nf = str(is_new or '').strip()
        if _nf in ('是', '新增'):
            bm["new_enterprise"] += 1
        elif _nf in ('更新登记表', '否'):
            bm["old_enterprise"] += 1
        _pa = str(price_audit or '').strip()
        if _pa in ('是', '通过'):
            bm["price_audit"]["total"] += 1
            bm["price_audit"]["passed"] += 1
        elif _pa in ('否', '不通过'):
            bm["price_audit"]["total"] += 1

    if reason and not passed_flag:
        rsn = str(reason).strip()
        if rsn:
            preview_stats["fail_reasons"][rsn] += 1

    # 新老供应商占比
    new_flag = str(is_new or '').strip()
    if new_flag in ('是', '新增'):
        preview_stats["new_enterprise"] += 1
    elif new_flag in ('更新登记表', '否'):
        preview_stats["old_enterprise"] += 1
    # 价格审核通过统计（col9）
    pa = str(price_audit or '').strip()
    if pa in ('是', '通过'):
        preview_stats["price_audit"]["total"] += 1
        preview_stats["price_audit"]["passed"] += 1
    elif pa in ('否', '不通过'):
        preview_stats["price_audit"]["total"] += 1

# 转换defaultdict为普通dict
if has_data:
    preview_stats["by_reviewer"] = dict(preview_stats["by_reviewer"])
    preview_stats["by_month"] = dict(preview_stats["by_month"])
    preview_stats["fail_reasons"] = dict(preview_stats["fail_reasons"])
    # 计算通过率
    total = preview_stats["total"]
    preview_stats["passRate"] = round(preview_stats["passed"] / total * 100, 1) if total else 0
    pa_total = preview_stats["price_audit"]["total"]
    preview_stats["price_audit"]["passRate"] = round(preview_stats["price_audit"]["passed"] / pa_total * 100, 1) if pa_total else 0

    for rv in preview_stats["by_reviewer"].values():
        rv["passRate"] = round(rv["passed"] / rv["total"] * 100, 1) if rv["total"] else 0
    for bm in preview_stats["by_month"].values():
        bm["passRate"] = round(bm["passed"] / bm["total"] * 100, 1) if bm["total"] else 0
        _pa_t = bm["price_audit"]["total"]
        bm["price_audit"]["passRate"] = round(bm["price_audit"]["passed"] / _pa_t * 100, 1) if _pa_t else 0
        _sup = bm["new_enterprise"] + bm["old_enterprise"]
        bm["newRate"] = round(bm["new_enterprise"] / _sup * 100, 1) if _sup else 0
        bm["oldRate"] = round(bm["old_enterprise"] / _sup * 100, 1) if _sup else 0
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
    val = cell_val(ws5, 3, r)   # 公式列（如 =C2-C3 双方未达成共识量）求值

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
# 6. 产品验收 → productAcceptance
# ============================================================
product_acceptance = {"by_month": {}}
if '产品验收' in wb.sheetnames:
    ws6 = wb['产品验收']
    for r in range(2, ws6.max_row + 1):
        month = ws6.cell(row=r, column=1).value
        label = ws6.cell(row=r, column=2).value
        val = ws6.cell(row=r, column=3).value
        if not month or not label:
            continue
        m = str(month).strip()
        if m == '月份':
            continue
        if m not in product_acceptance["by_month"]:
            product_acceptance["by_month"][m] = {}
        try:
            product_acceptance["by_month"][m][str(label).strip()] = float(val) if val else val
        except (ValueError, TypeError):
            product_acceptance["by_month"][m][str(label).strip()] = val

# ============================================================
# 7. 采集产能 → collectionCapacity（新 Sheet）
# ============================================================
collection_capacity = []
if '采集产能' in wb.sheetnames:
    ws7 = wb['采集产能']
    for r in range(2, ws7.max_row + 1):
        month = ws7.cell(row=r, column=1).value
        if not month:
            continue

        def cap_num(v):
            if v is None: return None
            try: return float(v)
            except: return None

        collection_capacity.append({
            "月份": str(month).strip(),
            "计划需求量": cap_num(ws7.cell(row=r, column=2).value),
            "计划采集家数": cap_num(ws7.cell(row=r, column=3).value),
            "计划采集条数": cap_num(ws7.cell(row=r, column=4).value),
            "实际需求量": cap_num(ws7.cell(row=r, column=5).value),
            "实际采集家数": cap_num(ws7.cell(row=r, column=6).value),
            "实际采集条数": cap_num(ws7.cell(row=r, column=7).value),
            "员工总人数": cap_num(ws7.cell(row=r, column=8).value),
            "备注": ws7.cell(row=r, column=9).value,
        })

# ============================================================
# 8. 计算月度趋势 monthlyTrend
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
# 9. 构建完整 data.json
# ============================================================
# 收集唯一值
demand_items = sorted(set(t['demandItem'] for t in tasks if t['demandItem']))
settle_types = sorted(set(t['settleType'] for t in tasks if t['settleType']))
sources = sorted(set(t['source'] for t in tasks if t['source']))

# 去重月份
month_keys = list(dict.fromkeys(list(months_in_data) + ["07月", "08月"]))
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
    "updateDate": datetime.date.today().isoformat(),
    "months": month_keys,
    "monthNames": month_names,
    "filterOptions": {
        "demandItems": ["全部"] + demand_items,
        "settleTypes": ["全部"] + settle_types,
        "sources": ["全部"] + sources,
        "months": ["全部"] + month_keys,
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
    "productAcceptance": product_acceptance if product_acceptance.get("by_month") else None,
    "collectionCapacity": collection_capacity,
}

# 写入文件
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

# 同步内嵌数据到 index.html 的 dataEmbed 块（离线 fallback 保持新鲜，避免显示过时快照）
# 仅在输出为默认 data.json 时同步；跑 test.json 等自定义输出不碰 index.html
if os.path.basename(OUTPUT_PATH) == _default_out:
    sync_embed_data(os.path.join(os.path.dirname(__file__), 'index.html'), data)

print(f"✅ 解析完成！共 {len(tasks)} 条任务, {len(settlements)} 条结算记录")
print(f"   月份: {months_in_data}")
print(f"   需求项: {demand_items}")
print(f"   结算类型: {settle_types}")
print(f"   问题反馈: {issue_feedback.get('问题反馈量', 'N/A') if issue_feedback else '无'}")
if preview_stats:
    print(f"   加工预审: {preview_stats['total']} 条, 通过率 {preview_stats.get('passRate', 0)}%, 新供应商 {preview_stats['new_enterprise']}, 老供应商 {preview_stats['old_enterprise']}, 价格审核 {preview_stats['price_audit']}")
print(f"   采集产能: {len(collection_capacity)} 条")
print(f"   已写入: {OUTPUT_PATH}")
