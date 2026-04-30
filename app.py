"""
NailVesta 发货系统
- 上传 Lark 水单 CSV
- (可选) 上传产品图册（图片）
- 选择洛杉矶日期（默认最晚日期）
- 生成 3 个文件：
  1. 拣货表（按库位排序）
  2. Packing Slip（客户发货单）
  3. 买 Label 的 Excel（物流水单）
"""

import io
import re
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 常量：发件人信息（NailVesta 固定信息）
# ============================================================
SENDER_INFO = {
    "name": "NailVesta",
    "company": None,
    "phone": "5105089943",
    "country": "US",
    "state": "CA",
    "city": "Los Angeles",
    "zip": "90071",
    "address": "515 S Flower St, Floor 18 & 19, STE 1901",
}

# 包裹规格默认值
PACKAGE_DEFAULTS = {
    "weight": 0.3,
    "length": 20,
    "width": 15,
    "height": 2,
    "cn_name": "穿戴甲",
    "en_name": "Press-On Nails",
    "qty": 1,
    "declare_price": 5,
    "net_weight": 0.3,
}

# 水单 Excel 列头（与现有模板完全一致）
SHUIDAN_HEADERS = [
    "客户订单号", "物流产品(产品编号)", "重量", "长", "宽", "高",
    "发件人姓名", "发件人公司", "发件人电话", "发件人国家", "发件人省/州",
    "发件人城市", "发件人邮编", "发件人地址",
    "收件人姓名", "收件人公司", "收件人电话", "收件人国家", "收件人省/州",
    "收件人城市", "收件人地址一", "收件人地址二", "收件人邮编",
    "中文品名1", "英文品名1", "SKU1", "数量1", "配货备注1",
    "申报单价1", "单位净重(kg)1",
]


# ============================================================
# 工具函数
# ============================================================
_US_STATES_FULL = {
    "ALABAMA","ALASKA","ARIZONA","ARKANSAS","CALIFORNIA","COLORADO","CONNECTICUT",
    "DELAWARE","FLORIDA","GEORGIA","HAWAII","IDAHO","ILLINOIS","INDIANA","IOWA",
    "KANSAS","KENTUCKY","LOUISIANA","MAINE","MARYLAND","MASSACHUSETTS","MICHIGAN",
    "MINNESOTA","MISSISSIPPI","MISSOURI","MONTANA","NEBRASKA","NEVADA",
    "NEW HAMPSHIRE","NEW JERSEY","NEW MEXICO","NEW YORK","NORTH CAROLINA",
    "NORTH DAKOTA","OHIO","OKLAHOMA","OREGON","PENNSYLVANIA","RHODE ISLAND",
    "SOUTH CAROLINA","SOUTH DAKOTA","TENNESSEE","TEXAS","UTAH","VERMONT",
    "VIRGINIA","WASHINGTON","WEST VIRGINIA","WISCONSIN","WYOMING",
    "DISTRICT OF COLUMBIA","PUERTO RICO",
}
_STATE_ABBR = {
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA",
    "KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
    "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
    "VA","WA","WV","WI","WY","DC","PR",
}


def _is_zip(s: str) -> bool:
    return bool(re.fullmatch(r"\d{5}(-\d{4})?", s.strip()))


def _is_phone_line(s: str) -> bool:
    if re.match(
        r"^[\sa]*\(?\+?1?\)?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4}",
        s.strip(),
    ):
        return True
    return bool(re.match(r"^(Tel|Phone|WhatsApp)\s*[:\uff1a]", s.strip(), re.I))


def _normalize_state(tok: str) -> Optional[str]:
    t = tok.strip().upper()
    if t in _US_STATES_FULL:
        return t
    if t in _STATE_ABBR:
        return t
    return None


def _clean_phone(s: str) -> str:
    digits = re.sub(r"\D", "", s)
    if digits.startswith("1") and len(digits) == 11:
        digits = digits[1:]
    return digits


def _clean_street(s: str) -> str:
    # "(Apt 5)" / "(Leave Inside Porch)" -> 用空格替代两端括号
    s = re.sub(r"\(([^)]*)\)", r" \1", s)
    s = re.sub(r"#\s+", "#", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_shipping_info(text: str) -> dict:
    """
    鲁棒解析多种格式的 Shipping Info：
      - 标准 5 行：姓名 / 电话 / 街道 / 城市,州,国家 / 邮编
      - 紧凑 4 行：街道+城市挤一行
      - 6 行（含 Address Line 2 / 顺序变化）
      - 7 行带 "Name:/Tel:/Address:" 前缀
      - 开头中文备注（自动跳过）
    """
    if not isinstance(text, str):
        return {}
    lines = [ln.strip() for ln in text.replace("\r", "").split("\n") if ln.strip()]
    while lines and re.search(r"[\u4e00-\u9fff]", lines[0]) and len(lines[0]) > 15:
        lines.pop(0)
    if not lines:
        return {}

    info = {
        "name": "", "phone": "", "street": "", "street2": "",
        "city": "", "state": "", "country": "United States", "zip": "",
    }

    # ZIP（最后一个独立 5 位邮编）
    zip_idx = None
    for i in range(len(lines) - 1, -1, -1):
        if _is_zip(lines[i]):
            zip_idx = i
            info["zip"] = lines[i].strip()
            break

    # Phone
    phone_idx = None
    for i, ln in enumerate(lines):
        if i == zip_idx:
            continue
        if _is_phone_line(ln):
            info["phone"] = _clean_phone(ln)
            phone_idx = i
            break

    # City, State, Country
    csc_idx = None
    csc_inline_street = None
    for i, ln in enumerate(lines):
        if i in (zip_idx, phone_idx):
            continue
        if "," in ln:
            parts = [p.strip() for p in ln.split(",")]
            if (
                any("UNITED STATES" in p.upper() or "USA" in p.upper() for p in parts)
                and len(parts) >= 3
            ):
                st = _normalize_state(parts[-2])
                if st:
                    info["state"] = st
                    info["country"] = parts[-1]
                    head = parts[0]
                    csc_idx = i
                    has_digit = bool(re.search(r"\d", head))
                    m = re.match(
                        r"^(.*?)\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s*$", head
                    )
                    if has_digit and m and re.search(r"\d", m.group(1)):
                        info["street"] = m.group(1).strip()
                        info["city"] = m.group(2).strip()
                        csc_inline_street = info["street"]
                    else:
                        info["city"] = head
                    break

    used = {zip_idx, phone_idx, csc_idx}
    others = [(i, lines[i]) for i in range(len(lines)) if i not in used]

    # 姓名：第一个不含数字、逗号、冒号的行
    name_pos = None
    for k, (i, ln) in enumerate(others):
        clean = re.sub(r"^Name\s*[:\uff1a]\s*", "", ln, flags=re.I)
        if not re.search(r"\d", clean) and "," not in clean and ":" not in clean:
            info["name"] = clean
            name_pos = k
            break

    street_lines = [
        re.sub(
            r"^(Address|Street|Apt|Suite)\s*[:\uff1a]\s*", "", ln, flags=re.I
        )
        for k, (i, ln) in enumerate(others) if k != name_pos
    ]

    if not info["street"]:
        if street_lines:
            info["street"] = street_lines[0]
            if len(street_lines) > 1:
                info["street2"] = " ".join(street_lines[1:])
    else:
        if street_lines:
            info["street2"] = " ".join(street_lines)

    info["street"] = _clean_street(info["street"])
    info["street2"] = _clean_street(info["street2"]) if info["street2"] else ""

    return info


def state_to_abbr(state: str) -> str:
    """全名 → 缩写（数据多用全名，水单需要缩写）"""
    mapping = {
        "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR",
        "CALIFORNIA": "CA", "COLORADO": "CO", "CONNECTICUT": "CT",
        "DELAWARE": "DE", "FLORIDA": "FL", "GEORGIA": "GA", "HAWAII": "HI",
        "IDAHO": "ID", "ILLINOIS": "IL", "INDIANA": "IN", "IOWA": "IA",
        "KANSAS": "KS", "KENTUCKY": "KY", "LOUISIANA": "LA", "MAINE": "ME",
        "MARYLAND": "MD", "MASSACHUSETTS": "MA", "MICHIGAN": "MI",
        "MINNESOTA": "MN", "MISSISSIPPI": "MS", "MISSOURI": "MO",
        "MONTANA": "MT", "NEBRASKA": "NE", "NEVADA": "NV",
        "NEW HAMPSHIRE": "NH", "NEW JERSEY": "NJ", "NEW MEXICO": "NM",
        "NEW YORK": "NY", "NORTH CAROLINA": "NC", "NORTH DAKOTA": "ND",
        "OHIO": "OH", "OKLAHOMA": "OK", "OREGON": "OR", "PENNSYLVANIA": "PA",
        "RHODE ISLAND": "RI", "SOUTH CAROLINA": "SC", "SOUTH DAKOTA": "SD",
        "TENNESSEE": "TN", "TEXAS": "TX", "UTAH": "UT", "VERMONT": "VT",
        "VIRGINIA": "VA", "WASHINGTON": "WA", "WEST VIRGINIA": "WV",
        "WISCONSIN": "WI", "WYOMING": "WY", "DISTRICT OF COLUMBIA": "DC",
        "PUERTO RICO": "PR",
    }
    if not state:
        return ""
    s = state.strip().upper()
    return mapping.get(s, state if len(state) == 2 else state)


@st.cache_data(show_spinner=False)
def load_lark_data(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes))
    df["日期"] = df["日期"].astype(str).str.strip()
    return df


def filter_orders_for_date(df: pd.DataFrame, target_date: str) -> pd.DataFrame:
    """筛选指定日期、有 Order ID 和 Shipping Info 的订单"""
    out = df[
        (df["日期"] == target_date)
        & df["Order ID"].notna()
        & df["Shipping Info"].notna()
    ].copy()
    out["Order ID"] = out["Order ID"].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x)
    )
    return out


# ============================================================
# 文件 1：拣货表（按库位排序）
# ============================================================
def build_picking_list(orders: pd.DataFrame, date_str: str) -> bytes:
    rows = []
    for _, r in orders.iterrows():
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        rows.append({
            "库位": r.get("库位") if pd.notna(r.get("库位")) else "",
            "Full SKU": r.get("Full SKU") if pd.notna(r.get("Full SKU")) else "",
            "款式": r.get("款式") if pd.notna(r.get("款式")) else "",
            "Size": r.get("Size") if pd.notna(r.get("Size")) else "",
            "数量": 1,
            "Order ID": r["Order ID"],
            "收件人": ship.get("name", ""),
        })
    df = pd.DataFrame(rows)
    df = df.sort_values(by=["库位", "Full SKU"], na_position="last").reset_index(drop=True)
    df.insert(0, "拣货顺序", range(1, len(df) + 1))
    df["√"] = ""

    wb = Workbook()
    ws = wb.active
    ws.title = "拣货表"

    title_font = Font(name="Arial", size=14, bold=True)
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="305496")
    border = Border(*[Side(style="thin", color="BFBFBF")] * 4)

    ws["A1"] = f"NailVesta 拣货表 - {date_str}"
    ws["A1"].font = title_font
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"] = f"共 {len(df)} 单"
    ws["A2"].font = Font(name="Arial", size=10, italic=True, color="595959")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(df.columns))

    cols = list(df.columns)
    for c, h in enumerate(cols, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for r_idx, row in enumerate(df.itertuples(index=False), 5):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(
                horizontal="center" if cols[c_idx - 1] in ("拣货顺序", "数量", "√", "Size") else "left",
                vertical="center",
            )
            cell.border = border
            if cols[c_idx - 1] == "库位":
                cell.font = Font(name="Arial", size=10, bold=True, color="C00000")

    widths = {
        "拣货顺序": 8, "库位": 12, "Full SKU": 14, "款式": 22, "Size": 8,
        "数量": 7, "Order ID": 22, "收件人": 22, "√": 5,
    }
    for c, h in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(h, 12)
    ws.row_dimensions[4].height = 22

    ws.freeze_panes = "A5"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# 文件 2：Packing Slip
# ============================================================
def build_packing_slip(orders: pd.DataFrame, date_str: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing Slips"

    title_font = Font(name="Arial", size=16, bold=True, color="305496")
    label_font = Font(name="Arial", size=9, bold=True, color="595959")
    val_font = Font(name="Arial", size=11)
    name_font = Font(name="Arial", size=12, bold=True)
    thank_font = Font(name="Arial", size=10, italic=True, color="595959")

    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    cur = 1
    for _, r in orders.iterrows():
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        order_id = r["Order ID"]
        sku = r.get("Full SKU") if pd.notna(r.get("Full SKU")) else ""
        style = r.get("款式") if pd.notna(r.get("款式")) else ""
        size = r.get("Size") if pd.notna(r.get("Size")) else ""

        # 标题
        ws.cell(row=cur, column=1, value="NailVesta").font = title_font
        ws.cell(row=cur, column=4, value=f"Date: {date_str}").font = Font(name="Arial", size=10)
        ws.cell(row=cur, column=4).alignment = Alignment(horizontal="right")
        cur += 1

        ws.cell(row=cur, column=1, value="PACKING SLIP").font = Font(name="Arial", size=11, bold=True, color="305496")
        cur += 2

        # 订单信息
        ws.cell(row=cur, column=1, value="ORDER #").font = label_font
        ws.cell(row=cur, column=2, value=order_id).font = val_font
        cur += 1

        # 收件人区块
        ws.cell(row=cur, column=1, value="SHIP TO").font = label_font
        cur += 1
        ws.cell(row=cur, column=1, value=ship.get("name", "").title()).font = name_font
        cur += 1
        ws.cell(row=cur, column=1, value=ship.get("street", "")).font = val_font
        cur += 1
        ws.cell(
            row=cur, column=1,
            value=f"{ship.get('city','')}, {state_to_abbr(ship.get('state',''))} {ship.get('zip','')}",
        ).font = val_font
        cur += 1
        ws.cell(row=cur, column=1, value=ship.get("country", "United States")).font = val_font
        cur += 2

        # 商品表头
        for c, h in enumerate(["SKU", "Style", "Size", "Qty"], 1):
            cell = ws.cell(row=cur, column=c, value=h)
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", start_color="305496")
            cell.alignment = Alignment(horizontal="center")
            cell.border = box
        cur += 1

        # 商品行
        items = [(sku, style, size, 1)]
        for sku_v, style_v, size_v, qty_v in items:
            for c, val in enumerate([sku_v, style_v, size_v, qty_v], 1):
                cell = ws.cell(row=cur, column=c, value=val)
                cell.font = val_font
                cell.alignment = Alignment(
                    horizontal="center" if c in (3, 4) else "left",
                    vertical="center",
                )
                cell.border = box
            cur += 1

        cur += 1
        ws.cell(
            row=cur, column=1,
            value="Thank you for shopping with NailVesta! 💅 We hope you love your nails!",
        ).font = thank_font
        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=4)
        cur += 1
        ws.cell(
            row=cur, column=1,
            value="Tag us @nailvesta on Instagram & TikTok to be featured ✨",
        ).font = thank_font
        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=4)
        cur += 2

        # 分隔
        for c in range(1, 5):
            ws.cell(row=cur, column=c).border = Border(top=Side(style="medium", color="305496"))
        cur += 2

    widths = [22, 28, 12, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 每张 Slip 设置打印分页（每订单约 17 行）
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# 文件 3：买 Label 的水单 Excel（与现有模板完全一致）
# ============================================================
def build_shuidan(orders: pd.DataFrame, date_str: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"

    # 表头
    for c, h in enumerate(SHUIDAN_HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name="宋体", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 数据行
    for r_idx, (_, r) in enumerate(orders.iterrows(), 2):
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        order_id = r["Order ID"]

        row_vals = [
            order_id,                                # 客户订单号
            None,                                    # 物流产品(产品编号)
            PACKAGE_DEFAULTS["weight"],              # 重量
            PACKAGE_DEFAULTS["length"],              # 长
            PACKAGE_DEFAULTS["width"],               # 宽
            PACKAGE_DEFAULTS["height"],              # 高
            SENDER_INFO["name"],                     # 发件人姓名
            SENDER_INFO["company"],                  # 发件人公司
            SENDER_INFO["phone"],                    # 发件人电话
            SENDER_INFO["country"],                  # 发件人国家
            SENDER_INFO["state"],                    # 发件人省/州
            SENDER_INFO["city"],                     # 发件人城市
            SENDER_INFO["zip"],                      # 发件人邮编
            SENDER_INFO["address"],                  # 发件人地址
            ship.get("name", ""),                    # 收件人姓名
            None,                                    # 收件人公司
            ship.get("phone", ""),                   # 收件人电话
            "US",                                    # 收件人国家
            state_to_abbr(ship.get("state", "")),    # 收件人省/州
            ship.get("city", ""),                    # 收件人城市
            ship.get("street", ""),                  # 收件人地址一
            ship.get("street2", "") or None,         # 收件人地址二
            ship.get("zip", ""),                     # 收件人邮编
            PACKAGE_DEFAULTS["cn_name"],             # 中文品名1
            PACKAGE_DEFAULTS["en_name"],             # 英文品名1
            None,                                    # SKU1
            PACKAGE_DEFAULTS["qty"],                 # 数量1
            None,                                    # 配货备注1
            PACKAGE_DEFAULTS["declare_price"],       # 申报单价1
            PACKAGE_DEFAULTS["net_weight"],          # 单位净重(kg)1
        ]
        for c, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c, value=v)
            cell.font = Font(name="宋体", size=10)

    # 列宽
    for c in range(1, len(SHUIDAN_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(
    page_title="NailVesta 发货系统",
    page_icon="💅",
    layout="wide",
)

st.title("💅 NailVesta 发货系统")
st.caption("上传 Lark 水单 → 选日期 → 一键生成拣货表 / Packing Slip / 买 Label 水单")

# 侧栏：上传
with st.sidebar:
    st.header("📤 文件上传")
    lark_file = st.file_uploader(
        "Lark 水单 CSV",
        type=["csv"],
        help="从 Lark 多维表格导出的 CSV",
    )
    catalog_files = st.file_uploader(
        "产品图册（可选）",
        type=["png", "jpg", "jpeg", "pdf"],
        accept_multiple_files=True,
        help="可上传产品图册供拣货时参考（不强制）",
    )

    st.divider()
    st.markdown("**📦 发件人信息**")
    st.caption(f"{SENDER_INFO['name']}\n{SENDER_INFO['address']}\n{SENDER_INFO['city']}, {SENDER_INFO['state']} {SENDER_INFO['zip']}")

if lark_file is None:
    st.info("👈 请先在左侧上传 Lark 水单 CSV 文件")
    st.stop()

# 加载数据
df = load_lark_data(lark_file.read())
all_dates = sorted(
    [d for d in df["日期"].dropna().unique() if re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", str(d))],
    key=lambda x: datetime.strptime(x, "%Y/%m/%d"),
    reverse=True,
)

if not all_dates:
    st.error("CSV 中找不到合法日期，请检查 Lark 导出格式")
    st.stop()

# 日期选择
col1, col2 = st.columns([1, 2])
with col1:
    st.subheader("📅 选择日期")
    st.caption("洛杉矶时间，默认最晚日期")
    selected_date = st.selectbox(
        "发货日期",
        options=all_dates,
        index=0,
        format_func=lambda x: f"{x} {'⬅️ 最新' if x == all_dates[0] else ''}",
    )

# 筛选订单
orders = filter_orders_for_date(df, selected_date)

with col2:
    st.subheader("📊 当日订单概况")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("订单数", len(orders))
    with m2:
        unique_states = orders.apply(
            lambda r: parse_shipping_info(r.get("Shipping Info", "")).get("state", ""),
            axis=1,
        ).nunique()
        st.metric("覆盖州", unique_states)
    with m3:
        unique_skus = orders["Full SKU"].dropna().nunique()
        st.metric("SKU 种类", unique_skus)
    with m4:
        location_count = orders["库位"].dropna().nunique()
        st.metric("拣货库位", location_count)

if len(orders) == 0:
    st.warning(f"⚠️ {selected_date} 没有可发货订单（需同时有 Order ID 和 Shipping Info）")
    st.stop()

# 订单预览
with st.expander("👀 订单明细预览", expanded=False):
    preview_cols = ["Order ID", "款式", "Size", "Full SKU", "库位", "Shipping Info"]
    preview = orders[[c for c in preview_cols if c in orders.columns]].copy()
    st.dataframe(preview, use_container_width=True, height=300)

st.divider()

# 生成按钮
st.subheader("🚀 生成发货文件")
if st.button("✨ 一键生成 3 个文件", type="primary", use_container_width=True):
    with st.spinner("生成中..."):
        date_compact = selected_date.replace("/", "")
        picking_bytes = build_picking_list(orders, selected_date)
        slip_bytes = build_packing_slip(orders, selected_date)
        shuidan_bytes = build_shuidan(orders, selected_date)

    st.success(f"✅ 已生成 {selected_date} 的 {len(orders)} 单发货文件！")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "📋 1. 拣货表（按库位排序）",
            data=picking_bytes,
            file_name=f"拣货表_{date_compact}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "📦 2. Packing Slips",
            data=slip_bytes,
            file_name=f"PackingSlip_{date_compact}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c3:
        st.download_button(
            "🚚 3. 水单（买 Label）",
            data=shuidan_bytes,
            file_name=f"{date_compact}客人水单.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# 产品图册预览
if catalog_files:
    st.divider()
    st.subheader("📷 产品图册参考")
    cols = st.columns(min(4, len(catalog_files)))
    for i, f in enumerate(catalog_files):
        with cols[i % len(cols)]:
            if f.type.startswith("image/"):
                st.image(f, caption=f.name, use_container_width=True)
            else:
                st.markdown(f"📄 {f.name}")
