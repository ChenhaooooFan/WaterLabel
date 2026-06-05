"""
NailVesta 发货系统（两阶段工作流）
========================================================
阶段 1（生成水单）:
    Lark CSV → 选日期 → 生成水单 → 发给打 Label 的人
    分 3 类输出:
      - 深度达人水单 (NailVesta 发出, 寄给达人)
      - 客户水单     (NailVesta 发出, 寄给客户)
      - Return Label 包裹 (顾客寄回, 顾客=发件人, NailVesta=收件人)

    重要逻辑：
      如果 Return Label 包裹列打勾，这一行仍然需要生成一张「正常发货水单」；
      同时额外生成一张「Return Label 水单」。也就是同一行会产出 2 张 label 需求。

阶段 2（拣货 + 发货 + 核对）:
    Lark CSV + 图册 CSV + Label PDF → OCR 核对 + 生成拣货单/发货单
    Return Label 勾选订单仍参与正常发货拣货；顾客寄回给我们的 Return Label 本身不参与拣货。
"""

import io
import os
import re
import subprocess
import tempfile
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

# ============================================================
# 常量
# ============================================================
SENDER_INFO = {
    "name": "NailVesta", "company": None, "phone": "5105089943",
    "country": "US", "state": "CA", "city": "Los Angeles",
    "zip": "90071", "address": "515 S Flower St, Floor 18 & 19, STE 1901",
}
PACKAGE_DEFAULTS = {
    "weight": 0.3, "length": 20, "width": 15, "height": 2,
    "cn_name": "穿戴甲", "en_name": "Press-On Nails",
    "qty": 1, "declare_price": 5, "net_weight": 0.3,
}
SHUIDAN_HEADERS = [
    "客户订单号","物流产品(产品编号)","重量","长","宽","高",
    "发件人姓名","发件人公司","发件人电话","发件人国家","发件人省/州",
    "发件人城市","发件人邮编","发件人地址",
    "收件人姓名","收件人公司","收件人电话","收件人国家","收件人省/州",
    "收件人城市","收件人地址一","收件人地址二","收件人邮编",
    "中文品名1","英文品名1","SKU1","数量1","配货备注1",
    "申报单价1","单位净重(kg)1",
]
SIZE_COL = "Size'"
RETURN_LABEL_COL = "Return Label 包裹"

# ============================================================
# 地址解析
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
_STATE_ABBR = {"AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL",
"IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH",
"NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
"VA","WA","WV","WI","WY","DC","PR"}
_STATE_FULL_TO_ABBR = {
    "ALABAMA":"AL","ALASKA":"AK","ARIZONA":"AZ","ARKANSAS":"AR","CALIFORNIA":"CA",
    "COLORADO":"CO","CONNECTICUT":"CT","DELAWARE":"DE","FLORIDA":"FL","GEORGIA":"GA",
    "HAWAII":"HI","IDAHO":"ID","ILLINOIS":"IL","INDIANA":"IN","IOWA":"IA","KANSAS":"KS",
    "KENTUCKY":"KY","LOUISIANA":"LA","MAINE":"ME","MARYLAND":"MD","MASSACHUSETTS":"MA",
    "MICHIGAN":"MI","MINNESOTA":"MN","MISSISSIPPI":"MS","MISSOURI":"MO","MONTANA":"MT",
    "NEBRASKA":"NE","NEVADA":"NV","NEW HAMPSHIRE":"NH","NEW JERSEY":"NJ","NEW MEXICO":"NM",
    "NEW YORK":"NY","NORTH CAROLINA":"NC","NORTH DAKOTA":"ND","OHIO":"OH","OKLAHOMA":"OK",
    "OREGON":"OR","PENNSYLVANIA":"PA","RHODE ISLAND":"RI","SOUTH CAROLINA":"SC",
    "SOUTH DAKOTA":"SD","TENNESSEE":"TN","TEXAS":"TX","UTAH":"UT","VERMONT":"VT",
    "VIRGINIA":"VA","WASHINGTON":"WA","WEST VIRGINIA":"WV","WISCONSIN":"WI","WYOMING":"WY",
    "DISTRICT OF COLUMBIA":"DC","PUERTO RICO":"PR",
}


def state_to_abbr(state: str) -> str:
    if not state: return ""
    s = state.strip().upper()
    if s in _STATE_FULL_TO_ABBR: return _STATE_FULL_TO_ABBR[s]
    if s in _STATE_ABBR: return s
    return state if len(state) == 2 else state


def _is_zip(s): return bool(re.fullmatch(r"\d{5}(-\d{4})?", s.strip()))


def _is_phone_line(s):
    if re.match(r"^[\sa]*\(?\+?1?\)?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4}", s.strip()):
        return True
    return bool(re.match(r"^(Tel|Phone|WhatsApp)\s*[:：]", s.strip(), re.I))


def _normalize_state(tok):
    t = tok.strip().upper()
    if t in _US_STATES_FULL: return t
    if t in _STATE_ABBR: return t
    return None


def _clean_phone(s):
    digits = re.sub(r"\D", "", s)
    if digits.startswith("1") and len(digits) == 11: digits = digits[1:]
    return digits


def _clean_street(s):
    s = re.sub(r"\(([^)]*)\)", r" \1", s)
    s = re.sub(r"#\s+", "#", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_return_label(val) -> bool:
    if val is None: return False
    if isinstance(val, bool): return val
    if isinstance(val, (int, float)):
        try:
            if pd.isna(val): return False
        except Exception:
            pass
        return float(val) >= 1.0
    if pd.isna(val): return False
    s = str(val).strip().lower()
    if not s or s == "nan" or s == "none" or s == "0" or s == "0.0":
        return False
    truthy = {"true", "1", "1.0", "yes", "y", "是", "✓", "✔", "✅",
              "勾选", "checked", "x", "t", "已选"}
    return s in truthy


def parse_free_address(text: str) -> dict:
    if not isinstance(text, str): return {}
    s = text.strip().rstrip(",").strip()
    s = s.replace("，", ",").replace("　", " ")
    s = re.sub(r"\s*\n\s*", " ", s)
    s = re.sub(r"\s+", " ", s)

    zip_matches = list(re.finditer(r"\b(\d{5})(-\d{4})?\b", s))
    zip_code = ""
    zip_match = None
    if zip_matches:
        zip_match = zip_matches[-1]
        zip_code = zip_match.group(1)

    state = ""
    state_pos = None
    if zip_match:
        zs = zip_match.start()
        ze = zip_match.end()
        ws = max(0, zs - 30)
        we = min(len(s), ze + 10)
        window = s[ws:we]
        window_upper = window.upper()
        for full, abbr in sorted(_STATE_FULL_TO_ABBR.items(), key=lambda x: -len(x[0])):
            m = re.search(r"\b" + re.escape(full) + r"\b", window_upper)
            if m:
                state = abbr
                state_pos = ws + m.start()
                break
        if not state:
            for m in re.finditer(r"\b([A-Za-z]{2})\b\.?", window):
                tok = m.group(1).upper()
                if tok in _STATE_ABBR:
                    state = tok
                    state_pos = ws + m.start()
    if not state and not zip_match:
        s_upper = s.upper()
        for full, abbr in sorted(_STATE_FULL_TO_ABBR.items(), key=lambda x: -len(x[0])):
            m = re.search(r"\b" + re.escape(full) + r"\b", s_upper)
            if m:
                state = abbr
                state_pos = m.start()
                break

    if state_pos is not None and zip_match is not None:
        cut = min(state_pos, zip_match.start())
    elif zip_match is not None:
        cut = zip_match.start()
    elif state_pos is not None:
        cut = state_pos
    else:
        cut = len(s)
    before = s[:cut].rstrip(" ,.\t\n")

    if "," in before:
        parts = [p.strip() for p in before.split(",") if p.strip()]
        if len(parts) >= 2:
            city = parts[-1]
            street = ", ".join(parts[:-1])
        else:
            street = parts[0]; city = ""
    else:
        words = before.split()
        if len(words) >= 4:
            street_kw = {"st","street","dr","drive","rd","road","ave","avenue","blvd",
                         "ln","lane","ct","court","way","cir","circle","pkwy","parkway",
                         "pl","place","hwy","highway","xing","crossing","ter","terrace",
                         "sq","square","trl","trail","plz","plaza"}
            apt_kw = {"apt","suite","ste","unit","#","no.","no"}
            last_idx = -1
            for i in range(len(words) - 1, -1, -1):
                wc = words[i].lower().rstrip(".,#").lstrip("#")
                if wc in street_kw:
                    last_idx = i; break
                if wc in apt_kw:
                    j = i + 1
                    while j < len(words) and re.match(r"^[\d\w\.\-#]+$", words[j]) and not (words[j] and words[j][0].isupper()):
                        j += 1
                    last_idx = j - 1 if j > i + 1 else i
                    break
            if 0 <= last_idx < len(words) - 1:
                street = " ".join(words[:last_idx + 1])
                city = " ".join(words[last_idx + 1:])
            elif len(words) >= 2:
                city = words[-1]; street = " ".join(words[:-1])
            else:
                street = before; city = ""
        else:
            street = before; city = ""

    return {
        "street": street.strip(" ,."),
        "city": city.strip(" ,.").title() if city else "",
        "state": state,
        "zip": zip_code,
    }


def _state_from_segment(seg: str):
    original = str(seg or "").strip(" ,")
    if not original:
        return None
    upper = original.upper().replace(".", "")

    if upper in _STATE_ABBR:
        return "", upper
    if upper in _STATE_FULL_TO_ABBR:
        return "", _STATE_FULL_TO_ABBR[upper]

    for full, abbr in sorted(_STATE_FULL_TO_ABBR.items(), key=lambda x: -len(x[0])):
        m = re.search(r"(?:^|\s)" + re.escape(full) + r"$", upper)
        if m:
            before = original[:m.start()].strip(" ,")
            return before, abbr

    m = re.search(r"(?:^|\s)([A-Z]{2})$", upper)
    if m and m.group(1) in _STATE_ABBR:
        before = original[:m.start()].strip(" ,")
        return before, m.group(1)

    return None


def _split_address_components(addr_lines: list, zip_code: str = "") -> dict:
    info = {
        "street": "", "street2": "", "city": "", "state": "",
        "country": "United States", "zip": zip_code or "",
    }

    cleaned = []
    for ln in addr_lines:
        if not ln:
            continue
        t = str(ln).strip()
        if re.fullmatch(r"(?i)(United States|USA|US)", t.strip(" ,")):
            continue
        if _is_zip(t):
            continue
        cleaned.append(t)

    combined = ", ".join(cleaned)
    combined = re.sub(r"\b(United States|USA|US)\b", "", combined, flags=re.I)

    zips = list(re.finditer(r"\b(\d{5})(?:-\d{4})?\b", combined))
    if zips:
        zip_match = None
        if info["zip"]:
            for z in reversed(zips):
                if z.group(1) == info["zip"]:
                    zip_match = z
                    break
        if zip_match is None and not info["zip"]:
            zip_match = zips[-1]
            info["zip"] = zip_match.group(1)
        if zip_match is not None:
            combined = (combined[:zip_match.start()] + combined[zip_match.end():]).strip(" ,")

    parts = [p.strip() for p in combined.split(",") if p.strip()]

    state_part_idx = None
    city_in_state_part = ""
    state_abbr = ""
    for idx in range(len(parts) - 1, -1, -1):
        cand = _state_from_segment(parts[idx])
        if cand:
            city_in_state_part, state_abbr = cand
            state_part_idx = idx
            break

    if state_part_idx is not None:
        info["state"] = state_abbr
        trailing_parts = parts[state_part_idx + 1:]

        if city_in_state_part:
            info["city"] = city_in_state_part
            street_parts = parts[:state_part_idx]
        else:
            if state_part_idx - 1 >= 0:
                info["city"] = parts[state_part_idx - 1]
                street_parts = parts[:state_part_idx - 1]
            else:
                street_parts = []

        if street_parts:
            info["street"] = _clean_street(", ".join(street_parts))
        if trailing_parts:
            info["street2"] = _clean_street(" ".join(trailing_parts))
        return info

    if cleaned:
        info["street"] = _clean_street(cleaned[0])
        if len(cleaned) > 1:
            info["street2"] = _clean_street(" ".join(cleaned[1:]))
    return info


def parse_shipping_info(text):
    if not isinstance(text, str):
        return {}

    text = text.replace("，", ",").replace("　", " ")

    lines = [ln.strip() for ln in text.replace("\r", "").split("\n") if ln.strip()]
    while lines and re.search(r"[一-鿿]", lines[0]) and len(lines[0]) > 15:
        lines.pop(0)
    if not lines:
        return {}

    info = {
        "name": "", "phone": "", "street": "", "street2": "",
        "city": "", "state": "", "country": "United States", "zip": "",
    }

    zip_idx = None
    for i in range(len(lines) - 1, -1, -1):
        if _is_zip(lines[i]):
            zip_idx = i
            info["zip"] = lines[i].strip()[:5]
            break
    if not info["zip"]:
        zip_matches = list(re.finditer(r"\b(\d{5})(?:-\d{4})?\b", text))
        if zip_matches:
            info["zip"] = zip_matches[-1].group(1)

    _PHONE_RE = re.compile(
        r'\(?\+?1?\)?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4}'
    )
    phone_idx = None
    phone_suffix = ""
    for i, ln in enumerate(lines):
        if i == zip_idx:
            continue
        m = _PHONE_RE.search(ln)
        if m:
            info["phone"] = _clean_phone(m.group(0))
            phone_idx = i
            after = ln[m.end():].strip(" ,\t")
            if after:
                phone_suffix = after
            break

    if phone_idx is not None:
        m = _PHONE_RE.search(lines[phone_idx])
        before_phone = lines[phone_idx][:m.start()].strip(" ,") if m else ""
        if before_phone and not re.fullmatch(r"(?i)(United States|USA|US)", before_phone):
            info["name"] = before_phone
        else:
            for i in range(phone_idx):
                clean = re.sub(r"^Name\s*[:：]\s*", "", lines[i], flags=re.I).strip(" ,")
                if (
                    clean
                    and not _is_phone_line(clean)
                    and not _is_zip(clean)
                    and not re.fullmatch(r"(?i)(United States|USA|US)", clean)
                ):
                    info["name"] = clean
                    break

    if not info["name"]:
        for i, ln in enumerate(lines):
            clean = str(ln).strip(" ,")
            if i in ([phone_idx] if phone_idx is not None else []) or i == zip_idx:
                continue
            if (
                clean
                and not _is_phone_line(clean)
                and not _is_zip(clean)
                and not re.fullmatch(r"(?i)(United States|USA|US)", clean)
            ):
                info["name"] = clean
                break

    start = (phone_idx + 1) if phone_idx is not None else 1
    addr_lines = []
    if phone_suffix:
        addr_lines.append(phone_suffix)
    for i in range(start, len(lines)):
        if i == zip_idx:
            continue
        ln = re.sub(r"^(Address|Street|Apt|Suite)\s*[:：]\s*", "", lines[i], flags=re.I).strip()
        if ln:
            addr_lines.append(ln)

    addr = _split_address_components(addr_lines, info["zip"])
    info.update(addr)
    info["state"] = state_to_abbr(info.get("state", ""))
    info["street"] = _clean_street(info.get("street", ""))
    info["street2"] = _clean_street(info.get("street2", "")) if info.get("street2") else ""
    return info


# ============================================================
# 数据加载
# ============================================================
@st.cache_data(show_spinner=False)
def load_lark_data(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes))
    df["日期"] = df["日期"].astype(str).str.strip()
    return df


@st.cache_data(show_spinner=False)
def load_catalog(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes))
    keep = ["SKU","中文名称","款式英文名称","甲型","图片","库位","所属系列"]
    keep = [c for c in keep if c in df.columns]
    df = df[keep].copy()
    df["SKU"] = df["SKU"].astype(str).str.strip()
    if "款式英文名称" in df.columns:
        df["款式英文名称"] = df["款式英文名称"].astype(str).str.strip()
    return df


def filter_orders_for_date(df, target_date):
    """
    筛选当日可发货订单。包括两种类型：
    A) 普通订单：必须有 Order ID + Shipping Info
    B) 深度达人订单：客诉类型为"深度达人"，有地址信息。
       兼容两种 Lark View 导出格式：
         格式 1（达人专属视图）：有"地址"+"达人Name"列
         格式 2（普通水单视图）：有"地址"+"Handle"列，无"达人Name"列
    """
    is_date = df["日期"] == target_date
    has_order_id = df["Order ID"].notna() if "Order ID" in df.columns else pd.Series([False] * len(df), index=df.index)
    has_shipping = df["Shipping Info"].notna() if "Shipping Info" in df.columns else pd.Series([False] * len(df), index=df.index)
    type_a = is_date & has_order_id & has_shipping

    is_deep_kol = (df["客诉类型"] == "深度达人") if "客诉类型" in df.columns else pd.Series([False] * len(df), index=df.index)
    has_addr = df["地址"].notna() if "地址" in df.columns else pd.Series([False] * len(df), index=df.index)
    # 达人名字：优先"达人Name"列，fallback 到"Handle"列
    has_kol_name = (
        df["达人Name"].notna() if "达人Name" in df.columns
        else df["Handle"].notna() if "Handle" in df.columns
        else pd.Series([False] * len(df), index=df.index)
    )
    type_b = is_date & is_deep_kol & has_addr & has_kol_name

    out = df[type_a | type_b].copy()

    def make_id(r):
        oid = r["Order ID"]
        if pd.notna(oid):
            return str(int(oid)) if isinstance(oid, float) and float(oid).is_integer() else str(oid)
        handle = str(r.get("Handle", "") or "").strip()
        name = str(r.get("达人Name", "") or "").strip()
        return f"KOL-{handle}" if handle and handle.lower() != "nan" else f"KOL-{name}"

    out["Order ID"] = out.apply(make_id, axis=1)
    return out


def split_orders_by_type(orders: pd.DataFrame) -> dict:
    if RETURN_LABEL_COL in orders.columns:
        is_ret = orders[RETURN_LABEL_COL].apply(is_return_label)
    else:
        is_ret = pd.Series([False] * len(orders), index=orders.index)

    is_kol = orders["客诉类型"] == "深度达人"

    return {
        "kol":               orders[ is_kol].copy(),
        "customer":          orders[~is_kol].copy(),
        "return_label_kol":  orders[ is_ret &  is_kol].copy(),
        "return_label_cust": orders[ is_ret & ~is_kol].copy(),
    }


def is_kol_order(row) -> bool:
    return str(row.get("客诉类型", "") or "").strip() == "深度达人"


def _get_kol_name(row) -> str:
    """获取达人名字：优先"达人Name"列，fallback 到"Handle"列"""
    name = str(row.get("达人Name", "") or "").strip()
    if name and name.lower() != "nan":
        return name
    handle = str(row.get("Handle", "") or "").strip()
    if handle and handle.lower() != "nan":
        return handle
    return ""


def get_recipient_info(row) -> dict:
    if is_kol_order(row):
        addr_raw = str(row.get("地址", "") or "").strip()
        parsed = parse_free_address(addr_raw)
        kol_name = _get_kol_name(row)
        return {
            "name": kol_name, "phone": "",
            "street": parsed.get("street", ""), "street2": "",
            "city": parsed.get("city", ""), "state": parsed.get("state", ""),
            "country": "United States", "zip": parsed.get("zip", ""),
        }
    else:
        ship = parse_shipping_info(row.get("Shipping Info", ""))
        return {
            "name": ship.get("name", ""), "phone": ship.get("phone", ""),
            "street": ship.get("street", ""), "street2": ship.get("street2", ""),
            "city": ship.get("city", ""), "state": ship.get("state", ""),
            "country": ship.get("country", "United States"), "zip": ship.get("zip", ""),
        }


# ============================================================
# 地址完整性校验
# ============================================================
def _is_blank_value(v) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except Exception:
        pass
    s = str(v).strip()
    return s == "" or s.lower() in {"nan", "none", "null"}


def _valid_us_zip(v) -> bool:
    if _is_blank_value(v):
        return False
    return bool(re.fullmatch(r"\d{5}(-\d{4})?", str(v).strip()))


def _valid_us_state(v) -> bool:
    if _is_blank_value(v):
        return False
    return state_to_abbr(str(v).strip()).upper() in _STATE_ABBR


def _format_original_address(row) -> str:
    if is_kol_order(row):
        return str(row.get("地址", "") or "").strip()
    return str(row.get("Shipping Info", "") or "").strip()


def validate_address_info(row, label_kind: str = "正常发货") -> list:
    info = get_recipient_info(row)
    role = "收件人" if label_kind == "正常发货" else "发件人"
    issues = []

    def add_issue(field, message, level="严重"):
        issues.append({
            "影响Label": label_kind,
            "对象": role,
            "订单号": str(row.get("Order ID", "") or "").strip(),
            "客诉类型": str(row.get("客诉类型", "") or "").strip(),
            "问题等级": level,
            "问题字段": field,
            "问题说明": message,
            "解析姓名": str(info.get("name", "") or "").strip(),
            "解析电话": str(info.get("phone", "") or "").strip(),
            "解析地址一": str(info.get("street", "") or "").strip(),
            "解析地址二": str(info.get("street2", "") or "").strip(),
            "解析城市": str(info.get("city", "") or "").strip(),
            "解析州": state_to_abbr(str(info.get("state", "") or "").strip()),
            "解析邮编": str(info.get("zip", "") or "").strip(),
            "原始地址信息": _format_original_address(row),
        })

    raw_addr = _format_original_address(row)
    if _is_blank_value(raw_addr):
        add_issue("原始地址", "原始地址信息为空，无法生成可用 Label")
        return issues

    if _is_blank_value(info.get("name")):
        add_issue(f"{role}姓名", f"{role}姓名缺失或没有被程序识别出来")

    phone = str(info.get("phone", "") or "").strip()
    if _is_blank_value(phone):
        if is_kol_order(row):
            add_issue(f"{role}电话", f"{role}电话缺失；深度达人单可人工确认是否允许为空", level="提醒")
        else:
            add_issue(f"{role}电话", f"{role}电话缺失")
    elif not re.fullmatch(r"\d{10}", phone):
        add_issue(f"{role}电话", f"{role}电话格式异常：{phone}，建议为 10 位美国电话", level="提醒")

    if _is_blank_value(info.get("street")):
        add_issue(f"{role}地址一", f"{role}街道地址 Address 1 缺失")
    if _is_blank_value(info.get("city")):
        add_issue(f"{role}城市", f"{role}城市 City 缺失或没有被程序识别出来")

    state = str(info.get("state", "") or "").strip()
    if _is_blank_value(state):
        add_issue(f"{role}州", f"{role}州 State 缺失或没有被程序识别出来")
    elif not _valid_us_state(state):
        add_issue(f"{role}州", f"{role}州 State 格式异常：{state}，应为美国州缩写或州全称")

    zip_code = str(info.get("zip", "") or "").strip()
    if _is_blank_value(zip_code):
        add_issue(f"{role}邮编", f"{role}邮编 Zip Code 缺失或没有被程序识别出来")
    elif not _valid_us_zip(zip_code):
        add_issue(f"{role}邮编", f"{role}邮编 Zip Code 格式异常：{zip_code}，应为 5 位或 ZIP+4")

    if str(info.get("name", "") or "").strip().lower() in {"united states", "usa", "us"}:
        add_issue(f"{role}姓名", f"{role}姓名被误识别为国家，请检查原始 Shipping Info")

    return issues


def build_address_issue_report(orders: pd.DataFrame, label_kind: str = "正常发货") -> pd.DataFrame:
    all_issues = []
    if orders is None or len(orders) == 0:
        return pd.DataFrame(columns=[
            "影响Label", "对象", "订单号", "客诉类型", "问题等级", "问题字段",
            "问题说明", "解析姓名", "解析电话", "解析地址一", "解析地址二",
            "解析城市", "解析州", "解析邮编", "原始地址信息",
        ])
    for _, row in orders.iterrows():
        all_issues.extend(validate_address_info(row, label_kind=label_kind))
    return pd.DataFrame(all_issues)


def address_report_csv(report: pd.DataFrame) -> bytes:
    return report.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# 组合装拆单 + 图册富集
# ============================================================
def _parse_size(size_raw):
    s = str(size_raw).strip() if size_raw else ""
    if ";" in s:
        mapping = {"_default": s}
        for part in s.split(";"):
            part = part.strip()
            m = re.match(r"^(.+?)\s+(S|M|L|XL|XS|\d+\s*个)$", part, re.I)
            if m: mapping[m.group(1).strip()] = m.group(2).strip()
        return mapping
    return {"_default": s}


def explode_orders(orders, catalog):
    cat_by_name = cat_by_sku = None
    if catalog is not None and not catalog.empty:
        if "款式英文名称" in catalog.columns:
            tmp = catalog.dropna(subset=["款式英文名称"]).copy()
            tmp["_key"] = tmp["款式英文名称"].astype(str).str.strip().str.lower()
            cat_by_name = tmp.set_index("_key")
        cat_by_sku = catalog.set_index("SKU")

    def lookup_by_name(name):
        if cat_by_name is None or not name: return None
        key = name.strip().lower()
        if key in cat_by_name.index:
            row = cat_by_name.loc[key]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            return row
        return None

    def lookup_by_sku(sku):
        if cat_by_sku is None or not sku: return None
        if sku in cat_by_sku.index:
            row = cat_by_sku.loc[sku]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            return row
        return None

    rows = []
    for _, r in orders.iterrows():
        order_id = r["Order ID"]
        ship_info = r.get("Shipping Info", "")
        lark_loc = str(r.get("库位", "") or "").strip()
        if lark_loc.lower() == "nan": lark_loc = ""

        is_kol = is_kol_order(r)

        if is_kol:
            pn_raw = str(r.get("款式", "") or "").strip()
            sku_raw = str(r.get("SKU", "") or "").strip()
            size_raw = str(r.get("Size", "") or "").strip()
        else:
            pn_raw = str(r.get("Product Name", "") or "").strip()
            sku_raw = str(r.get("SKU", "") or "").strip()
            size_raw = str(r.get(SIZE_COL, "") or "").strip()
        if pn_raw.lower() == "nan": pn_raw = ""
        if sku_raw.lower() == "nan": sku_raw = ""
        if size_raw.lower() == "nan": size_raw = ""

        pn_list = [p.strip() for p in pn_raw.split(",") if p.strip()] if pn_raw else []
        sku_list = [s.strip() for s in sku_raw.split(",") if s.strip()] if sku_raw else []
        size_map = _parse_size(size_raw)
        n = max(len(pn_list), len(sku_list), 1)

        for i in range(n):
            name_i = pn_list[i] if i < len(pn_list) else ""
            sku_i = sku_list[i] if i < len(sku_list) else ""
            cat_row = lookup_by_name(name_i) if name_i else None
            if cat_row is None:
                cat_row = lookup_by_sku(sku_i) if sku_i else None
            cn_name = cat_loc = ""
            cat_en_name = name_i
            if cat_row is not None:
                cn_v = cat_row.get("中文名称", "")
                cn_name = "" if pd.isna(cn_v) else str(cn_v).strip()
                loc_v = cat_row.get("库位", "")
                cat_loc = "" if pd.isna(loc_v) else str(loc_v).strip()
                if not name_i:
                    en_v = cat_row.get("款式英文名称", "")
                    cat_en_name = "" if pd.isna(en_v) else str(en_v).strip()
                if not sku_i and cat_row.name:
                    sku_i = str(cat_row.name)
            size_i = size_map.get(name_i, size_map.get("_default", size_raw))
            final_loc = cat_loc or lark_loc
            rows.append({
                "Order ID": order_id, "英文款式": cat_en_name, "SKU": sku_i,
                "中文名": cn_name, "库位": final_loc, "Size": size_i,
                "Shipping Info": ship_info,
                "Full SKU": str(r.get("Full SKU", "") or "").strip(),
                "客诉类型": str(r.get("客诉类型", "") or "").strip(),
                "is_kol": is_kol,
                "达人Name": _get_kol_name(r),
                "Handle": str(r.get("Handle", "") or "").strip(),
                "地址_raw": str(r.get("地址", "") or "").strip(),
            })
    return pd.DataFrame(rows)


# ============================================================
# 文件 1：拣货表（按库位汇总）
# ============================================================
def build_picking_summary_csv(exploded: pd.DataFrame) -> bytes:
    if len(exploded) == 0:
        return pd.DataFrame(columns=["库位","Product Name","S","M","L","其他","Total"]).to_csv(index=False).encode("utf-8-sig")

    df = exploded.copy()
    df["Size"] = df["Size"].fillna("").astype(str).str.strip().str.upper()

    def size_bucket(s):
        if s in ("S","M","L"): return s
        if s in ("XS","XL"): return s
        return "其他"
    df["SizeBucket"] = df["Size"].apply(size_bucket)

    grouped = df.groupby(["库位", "英文款式", "SizeBucket"]).size().unstack(fill_value=0)
    for c in ["S","M","L"]:
        if c not in grouped.columns: grouped[c] = 0
    other_cols = [c for c in grouped.columns if c not in ("S","M","L")]
    grouped["其他"] = grouped[other_cols].sum(axis=1) if other_cols else 0
    grouped["Total"] = grouped[["S","M","L","其他"]].sum(axis=1)
    grouped = grouped[["S","M","L","其他","Total"]].reset_index()
    grouped = grouped.rename(columns={"英文款式": "Product Name"})

    grouped["_loc_sort"] = grouped["库位"].apply(lambda x: (x == "", x))
    grouped = grouped.sort_values(by=["_loc_sort", "Product Name"]).drop(columns=["_loc_sort"]).reset_index(drop=True)

    if (grouped["其他"] == 0).all():
        grouped = grouped.drop(columns=["其他"])

    return grouped.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# 文件 2：Packing Slip CSV
# ============================================================
def build_packing_slip_csv(exploded: pd.DataFrame, date_str: str) -> bytes:
    rows = []
    for order_id, grp in exploded.groupby("Order ID", sort=False):
        first = grp.iloc[0]
        if first.get("is_kol", False):
            recip_addr = parse_free_address(first.get("地址_raw", ""))
            recip = {
                "name": first.get("达人Name", ""), "phone": "",
                "street": recip_addr.get("street", ""), "street2": "",
                "city": recip_addr.get("city", ""), "state": recip_addr.get("state", ""),
                "zip": recip_addr.get("zip", ""), "country": "United States",
            }
        else:
            ship = parse_shipping_info(first.get("Shipping Info", ""))
            recip = {
                "name": ship.get("name", ""), "phone": ship.get("phone", ""),
                "street": ship.get("street", ""), "street2": ship.get("street2", ""),
                "city": ship.get("city", ""), "state": ship.get("state", ""),
                "zip": ship.get("zip", ""), "country": ship.get("country", "United States"),
            }
        addr_parts = [p for p in [
            recip["street"], recip["street2"],
            f"{recip['city']}, {state_to_abbr(recip['state'])} {recip['zip']}".strip(", "),
            recip["country"],
        ] if p.strip(" ,")]
        full_address = " | ".join(addr_parts)
        for _, r in grp.iterrows():
            rows.append({
                "Order ID": order_id, "Date": date_str,
                "Recipient": recip["name"].title() if recip["name"] else "",
                "Phone": recip["phone"], "Address": full_address,
                "客诉类型": r.get("客诉类型", ""),
                "SKU": r["SKU"] or "", "Style": r["英文款式"] or "",
                "Chinese Name": r["中文名"] or "", "Size": r["Size"] or "", "Qty": 1,
            })
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")


# ============================================================
# 文件 3：水单 Excel（正常发货：NailVesta → 客户）
# ============================================================
def _write_shuidan_workbook(rows_data: list) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"

    for c, h in enumerate(SHUIDAN_HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name="宋体", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row_vals in enumerate(rows_data, 2):
        for c, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c, value=v)
            cell.font = Font(name="宋体", size=10)

    text_cols = [1, 9, 13, 17, 23]
    for col_idx in text_cols:
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                cell.number_format = "@"

    for c in range(1, len(SHUIDAN_HEADERS) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def build_shuidan_xlsx(orders: pd.DataFrame) -> bytes:
    rows_data = []
    for _, r in orders.iterrows():
        recip = get_recipient_info(r)
        rows_data.append([
            r["Order ID"], None,
            PACKAGE_DEFAULTS["weight"], PACKAGE_DEFAULTS["length"],
            PACKAGE_DEFAULTS["width"], PACKAGE_DEFAULTS["height"],
            SENDER_INFO["name"], None, SENDER_INFO["phone"],
            SENDER_INFO["country"], SENDER_INFO["state"],
            SENDER_INFO["city"], SENDER_INFO["zip"], SENDER_INFO["address"],
            recip["name"], None, recip["phone"], "US",
            state_to_abbr(recip["state"]), recip["city"],
            recip["street"], recip["street2"] or None, recip["zip"],
            PACKAGE_DEFAULTS["cn_name"], PACKAGE_DEFAULTS["en_name"],
            None, PACKAGE_DEFAULTS["qty"], None,
            PACKAGE_DEFAULTS["declare_price"], PACKAGE_DEFAULTS["net_weight"],
        ])
    return _write_shuidan_workbook(rows_data)


def build_return_label_xlsx(orders: pd.DataFrame) -> bytes:
    rows_data = []
    for _, r in orders.iterrows():
        sender = get_recipient_info(r)
        rows_data.append([
            r["Order ID"], None,
            PACKAGE_DEFAULTS["weight"], PACKAGE_DEFAULTS["length"],
            PACKAGE_DEFAULTS["width"], PACKAGE_DEFAULTS["height"],
            sender["name"], None, sender["phone"],
            "US", state_to_abbr(sender["state"]),
            sender["city"], sender["zip"], sender["street"],
            SENDER_INFO["name"], None, SENDER_INFO["phone"], "US",
            SENDER_INFO["state"], SENDER_INFO["city"],
            SENDER_INFO["address"], None, SENDER_INFO["zip"],
            PACKAGE_DEFAULTS["cn_name"], PACKAGE_DEFAULTS["en_name"],
            None, PACKAGE_DEFAULTS["qty"], None,
            PACKAGE_DEFAULTS["declare_price"], PACKAGE_DEFAULTS["net_weight"],
        ])
    return _write_shuidan_workbook(rows_data)


# ============================================================
# Label PDF OCR & 核对
# ============================================================
def check_ocr_available() -> tuple:
    has_tess = subprocess.run(["which","tesseract"], capture_output=True).returncode == 0
    has_popp = subprocess.run(["which","pdftoppm"], capture_output=True).returncode == 0
    return has_tess, has_popp


def parse_label_pdf(pdf_bytes: bytes) -> list:
    has_tess, has_popp = check_ocr_available()
    if not (has_tess and has_popp):
        raise RuntimeError("OCR 工具未安装：需要 tesseract + poppler-utils")

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, "labels.pdf")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)
        subprocess.run(
            ["pdftoppm", "-jpeg", "-r", "200", pdf_path, os.path.join(tmp, "lbl")],
            capture_output=True, timeout=120,
        )
        jpgs = sorted([f for f in os.listdir(tmp) if f.startswith("lbl") and f.endswith(".jpg")])

        results = []
        for jpg in jpgs:
            jpg_path = os.path.join(tmp, jpg)
            page_num = int(re.search(r"lbl-(\d+)", jpg).group(1))
            text = subprocess.run(
                ["tesseract", jpg_path, "-"],
                capture_output=True, text=True, timeout=30,
            ).stdout
            results.append(_parse_label_text(text, page_num))
        results.sort(key=lambda x: x["page"])
        return results


def _parse_label_text(text: str, page: int) -> dict:
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    info = {"page": page, "order_id": "", "tracking": "",
            "name": "", "address_lines": [], "raw_text": text}

    for ln in reversed(lines):
        digits = re.sub(r"\s", "", ln)
        m = re.match(r"^(\d{17,20})(-\d+)?$", digits)
        if m:
            info["order_id"] = digits
            break

    for ln in lines:
        if re.search(r"9\d{3}\s+\d{4}\s+\d{4}\s+\d{4}", ln):
            info["tracking"] = re.sub(r"\s+", "", ln)
            break

    in_addr = False
    for ln in lines:
        if "LOS ANGELES" in ln.upper() and "CA" in ln.upper():
            in_addr = True; continue
        if in_addr:
            if "TRACKING" in ln.upper(): break
            cleaned = re.sub(r"^([a-zA-Z]+:?\s*[=>:;©]+\s*|[=>:;©]+\s*|[a-z]\s+|[a-zA-Z0-9]{1,4}[=:;©]\s*)", "", ln)
            cleaned = re.sub(r"^[a-z]:\s*", "", cleaned)
            cleaned = re.sub(r"^[=>:;]+\s*", "", cleaned).strip()
            if cleaned and re.search(r"[A-Za-z0-9]", cleaned):
                info["address_lines"].append(cleaned)

    if info["address_lines"]:
        info["name"] = info["address_lines"][0].upper().strip()
    return info


def reconcile_labels(orders: pd.DataFrame, labels: list) -> pd.DataFrame:
    label_by_oid = {}
    for lbl in labels:
        oid = lbl["order_id"]
        oid_base = re.sub(r"-\d+$", "", oid)
        if oid_base:
            label_by_oid[oid_base] = lbl

    rows = []
    for _, r in orders.iterrows():
        oid = str(r["Order ID"]).strip()
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        lark_name = ship.get("name", "")
        lark_zip = ship.get("zip", "")
        lark_state = state_to_abbr(ship.get("state", ""))

        lbl = label_by_oid.get(oid)
        if not lbl:
            rows.append({
                "Order ID": oid, "状态": "❌ 缺面单",
                "Lark 收件人": lark_name, "Label 收件人": "—",
                "Lark 邮编": lark_zip, "Label 末行": "—",
                "Tracking": "—", "备注": "PDF 中找不到此订单的面单",
            })
            continue

        def norm(s): return re.sub(r"[^A-Z0-9]", "", s.upper()) if s else ""
        lark_n = norm(lark_name); lbl_n = norm(lbl["name"])
        last_line = lbl["address_lines"][-1] if lbl["address_lines"] else ""
        zip_in_label = lark_zip and lark_zip in last_line

        name_match = lark_n and lbl_n and (
            lark_n in lbl_n or lbl_n in lark_n
            or _ratio(lark_n, lbl_n) >= 0.7
        )

        if name_match and zip_in_label:
            status = "✅ 一致"; note = ""
        elif name_match:
            status = "⚠️ 名字一致但邮编对不上"
            note = f"Lark={lark_zip} | Label={last_line}"
        elif zip_in_label:
            status = "⚠️ 邮编一致但名字 OCR 模糊"
            note = f"Lark={lark_name} | Label={lbl['name']}"
        else:
            status = "❌ 不匹配"
            note = f"Lark={lark_name} {lark_zip} | Label={lbl['name']} {last_line}"

        rows.append({
            "Order ID": oid, "状态": status,
            "Lark 收件人": lark_name, "Label 收件人": lbl["name"],
            "Lark 邮编": lark_zip, "Label 末行": last_line,
            "Tracking": lbl["tracking"], "备注": note,
        })

    lark_oids = set(str(o).strip() for o in orders["Order ID"])
    for oid_base, lbl in label_by_oid.items():
        if oid_base not in lark_oids:
            rows.append({
                "Order ID": lbl["order_id"], "状态": "❌ 多余面单",
                "Lark 收件人": "—", "Label 收件人": lbl["name"],
                "Lark 邮编": "—",
                "Label 末行": lbl["address_lines"][-1] if lbl["address_lines"] else "",
                "Tracking": lbl["tracking"], "备注": "Label PDF 含此订单但 Lark 当日无此单",
            })

    return pd.DataFrame(rows)


def _ratio(a, b):
    if not a or not b: return 0
    if a == b: return 1
    m, n = len(a), len(b)
    dp = [[0]*(n+1) for _ in range(m+1)]
    for i in range(1, m+1):
        for j in range(1, n+1):
            if a[i-1] == b[j-1]: dp[i][j] = dp[i-1][j-1] + 1
            else: dp[i][j] = max(dp[i-1][j], dp[i][j-1])
    return dp[m][n] / max(m, n)


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(page_title="NailVesta 发货系统", page_icon="💅", layout="wide")
st.title("💅 NailVesta 发货系统")

tab1, tab2 = st.tabs([
    "1️⃣ 第一步：生成水单（发给打 Label 的人）",
    "2️⃣ 第二步：拿到 Label PDF 后，核对 + 生成拣货单/发货单",
])

# ============================================================
# Tab 1: 生成水单
# ============================================================
with tab1:
    st.subheader("第 1 步 · 生成水单 Excel")
    st.caption(
        "上传 Lark CSV → 选日期 → 下载水单（正常发货 + 额外 Return Label）→ 发给打 Label 的同事"
    )

    lark_file_1 = st.file_uploader("Lark 水单 CSV", type=["csv"], key="lark1")

    if lark_file_1 is None:
        st.info("👆 请上传 Lark CSV")
    else:
        df1 = load_lark_data(lark_file_1.read())
        all_dates_1 = sorted(
            [d for d in df1["日期"].dropna().unique()
             if re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", str(d))],
            key=lambda x: datetime.strptime(x, "%Y/%m/%d"), reverse=True,
        )
        if not all_dates_1:
            st.error("CSV 中找不到合法日期")
        else:
            sel_date_1 = st.selectbox(
                "发货日期（洛杉矶时间，默认最晚）",
                options=all_dates_1, index=0,
                format_func=lambda x: f"{x} {'⬅️ 最新' if x == all_dates_1[0] else ''}",
                key="d1",
            )
            orders_1 = filter_orders_for_date(df1, sel_date_1)
            split_1 = split_orders_by_type(orders_1)
            kol_orders = split_1["kol"]
            customer_orders = split_1["customer"]
            return_kol_orders = split_1["return_label_kol"]
            return_cust_orders = split_1["return_label_cust"]
            return_total = len(return_kol_orders) + len(return_cust_orders)

            mc1, mc2, mc3, mc4, mc5 = st.columns(5)
            mc1.metric("当日总订单", len(orders_1))
            mc2.metric("深度达人单", len(kol_orders))
            mc3.metric("客户单", len(customer_orders))
            mc4.metric("↩️ 退货达人", len(return_kol_orders))
            mc5.metric("↩️ 退货客户", len(return_cust_orders))

            if len(orders_1) == 0:
                st.warning(f"⚠️ {sel_date_1} 没有可发货订单")
            else:
                with st.expander("📊 客诉类型分布", expanded=False):
                    type_counts = orders_1["客诉类型"].fillna("(空)").value_counts()
                    st.dataframe(
                        type_counts.rename_axis("客诉类型").reset_index(name="单数"),
                        use_container_width=True, hide_index=True,
                    )

                with st.expander("👀 订单预览", expanded=False):
                    preview_cols = ["Order ID", "客诉类型", RETURN_LABEL_COL,
                                    "Product Name", "款式",
                                    "Full SKU", SIZE_COL, "Size",
                                    "达人Name", "Handle", "地址", "Shipping Info"]
                    preview = orders_1[[c for c in preview_cols if c in orders_1.columns]]
                    st.dataframe(preview, use_container_width=True, height=300)

                date_compact = sel_date_1.replace("/", "")

                # —— 地址完整性校验 ——
                normal_orders_for_check = pd.concat([kol_orders, customer_orders], ignore_index=False)
                normal_address_issues = build_address_issue_report(normal_orders_for_check, label_kind="正常发货")
                return_address_issues = pd.concat([
                    build_address_issue_report(return_kol_orders, label_kind="Return Label"),
                    build_address_issue_report(return_cust_orders, label_kind="Return Label"),
                ], ignore_index=True)
                address_issues_all = pd.concat(
                    [normal_address_issues, return_address_issues],
                    ignore_index=True,
                )

                if len(address_issues_all) > 0:
                    serious_cnt = (address_issues_all["问题等级"] == "严重").sum()
                    remind_cnt = (address_issues_all["问题等级"] == "提醒").sum()
                    st.warning(
                        f"⚠️ 地址信息检查发现 {len(address_issues_all)} 条问题："
                        f"{serious_cnt} 条严重 / {remind_cnt} 条提醒。"
                        "建议先修正 Lark 水单表后再下载或上传水单。"
                    )
                    with st.expander("🚨 查看地址问题明细", expanded=True):
                        show_cols = [
                            "影响Label", "订单号", "客诉类型", "问题等级", "问题字段", "问题说明",
                            "解析姓名", "解析电话", "解析地址一", "解析地址二",
                            "解析城市", "解析州", "解析邮编", "原始地址信息",
                        ]
                        st.dataframe(
                            address_issues_all[[c for c in show_cols if c in address_issues_all.columns]],
                            use_container_width=True, height=320, hide_index=True,
                        )
                        st.download_button(
                            f"📥 下载地址问题报告（{len(address_issues_all)} 条）",
                            data=address_report_csv(address_issues_all),
                            file_name=f"{date_compact}地址问题报告.csv",
                            mime="text/csv",
                            use_container_width=True,
                            key="dl_address_issues_tab1",
                        )
                else:
                    st.success("✅ 地址信息检查通过：姓名 / 电话 / 地址 / 城市 / 州 / 邮编未发现明显缺失。")

                st.divider()
                st.markdown("### 📥 下载水单文件（正常发货）")

                col_kol, col_other = st.columns(2)

                with col_kol:
                    st.markdown("#### 🌟 深度达人水单")
                    if len(kol_orders) == 0:
                        st.info("今日无深度达人单")
                    else:
                        kol_xlsx = build_shuidan_xlsx(kol_orders)
                        st.download_button(
                            f"🌟 下载深度达人水单（{len(kol_orders)} 单）",
                            data=kol_xlsx,
                            file_name=f"{date_compact}深度达人水单.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary", use_container_width=True, key="dl_kol",
                        )

                with col_other:
                    st.markdown("#### 📦 客户水单")
                    if len(customer_orders) == 0:
                        st.info("今日无客户单")
                    else:
                        other_xlsx = build_shuidan_xlsx(customer_orders)
                        st.download_button(
                            f"📦 下载客户水单（{len(customer_orders)} 单）",
                            data=other_xlsx,
                            file_name=f"{date_compact}客人水单.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary", use_container_width=True, key="dl_other",
                        )

                # —— Return Label 单独区域 ——
                st.divider()
                st.markdown("### ↩️ Return Label 包裹（顾客/达人寄回）")
                st.caption(
                    "**这些是额外生成的寄回包裹 label**：发件人 = 顾客/达人，收件人 = NailVesta。"
                    "Return Label 打勾的订单仍然已经包含在上面的正常发货水单里；这里是额外再生成一张寄回 label ⚠️"
                )

                if return_total == 0:
                    st.info("今日无 Return Label 包裹 ✅")
                else:
                    col_ret_kol, col_ret_cust = st.columns(2)

                    with col_ret_kol:
                        st.markdown("#### 🌟 退货 · 深度达人")
                        if len(return_kol_orders) == 0:
                            st.info("今日无深度达人退货")
                        else:
                            with st.expander(f"👀 查看 {len(return_kol_orders)} 单详情", expanded=False):
                                ret_kol_cols = ["Order ID", "客诉类型", "款式", "Handle", "达人Name", "地址"]
                                ret_kol_preview = return_kol_orders[
                                    [c for c in ret_kol_cols if c in return_kol_orders.columns]
                                ]
                                st.dataframe(ret_kol_preview, use_container_width=True, height=200)
                            ret_kol_xlsx = build_return_label_xlsx(return_kol_orders)
                            st.download_button(
                                f"↩️ 下载深度达人退货水单（{len(return_kol_orders)} 单）",
                                data=ret_kol_xlsx,
                                file_name=f"{date_compact}Return_Label_深度达人水单.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary", use_container_width=True, key="dl_return_kol",
                            )

                    with col_ret_cust:
                        st.markdown("#### 📦 退货 · 客户")
                        if len(return_cust_orders) == 0:
                            st.info("今日无客户退货")
                        else:
                            with st.expander(f"👀 查看 {len(return_cust_orders)} 单详情", expanded=False):
                                ret_cust_cols = ["Order ID", "客诉类型", "Product Name", "Shipping Info"]
                                ret_cust_preview = return_cust_orders[
                                    [c for c in ret_cust_cols if c in return_cust_orders.columns]
                                ]
                                st.dataframe(ret_cust_preview, use_container_width=True, height=200)
                            ret_cust_xlsx = build_return_label_xlsx(return_cust_orders)
                            st.download_button(
                                f"↩️ 下载客户退货水单（{len(return_cust_orders)} 单）",
                                data=ret_cust_xlsx,
                                file_name=f"{date_compact}Return_Label_客人水单.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary", use_container_width=True, key="dl_return_cust",
                            )

                st.caption(
                    "💡 已强制邮编/电话/Order ID 为文本格式，"
                    "Excel 打开不会丢失前导零或 Order ID 转科学计数法"
                )

# ============================================================
# Tab 2: 核对 + 拣货单 + 发货单
# ============================================================
with tab2:
    st.subheader("第 2 步 · 核对面单 + 生成拣货单/发货单")
    st.caption(
        "上传 Lark CSV + 图册 + Label PDF → 自动核对 + 生成拣货单/发货单。"
        "Return Label 勾选订单仍参与正常发货拣货；寄回给我们的 Return Label 本身不需要拣货。"
        "面单核对建议只上传正常发货 Label PDF，不要把寄回 Return Label PDF 混在一起。"
    )

    c_left, c_right = st.columns(2)
    with c_left:
        lark_file_2 = st.file_uploader("Lark 水单 CSV", type=["csv"], key="lark2")
        catalog_file_2 = st.file_uploader("产品图册 CSV", type=["csv"], key="cat2")
    with c_right:
        label_pdf = st.file_uploader("Label PDF（打 Label 后拿到的面单）", type=["pdf"], key="pdf2")

    if lark_file_2 is None or catalog_file_2 is None:
        st.info("👆 请上传 Lark CSV 和产品图册 CSV")
        st.stop()

    df2 = load_lark_data(lark_file_2.read())
    catalog2 = load_catalog(catalog_file_2.read())

    all_dates_2 = sorted(
        [d for d in df2["日期"].dropna().unique()
         if re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", str(d))],
        key=lambda x: datetime.strptime(x, "%Y/%m/%d"), reverse=True,
    )
    if not all_dates_2:
        st.error("CSV 中找不到合法日期")
        st.stop()

    sel_date_2 = st.selectbox(
        "发货日期", options=all_dates_2, index=0,
        format_func=lambda x: f"{x} {'⬅️ 最新' if x == all_dates_2[0] else ''}",
        key="d2",
    )
    orders_2_all = filter_orders_for_date(df2, sel_date_2)

    split_2 = split_orders_by_type(orders_2_all)
    return_count_2 = len(split_2["return_label_kol"]) + len(split_2["return_label_cust"])
    orders_2 = pd.concat([split_2["kol"], split_2["customer"]], ignore_index=False)

    if return_count_2 > 0:
        st.info(
            f"ℹ️ 已保留 {return_count_2} 单 Return Label 勾选订单参与正常发货拣货"
            f"（{len(split_2['return_label_kol'])} 深度达人 + "
            f"{len(split_2['return_label_cust'])} 客户）。"
            "寄回给我们的 Return Label 本身不需要拣货；面单核对请上传正常发货 Label PDF。"
        )

    exploded_2 = explode_orders(orders_2, catalog2)

    m1, m2, m3 = st.columns(3)
    m1.metric("待发货订单", len(orders_2))
    m2.metric("拣货行数", len(exploded_2),
              delta=f"+{len(exploded_2)-len(orders_2)} 组合装"
                    if len(exploded_2) > len(orders_2) else None)
    m3.metric("拣货库位",
              exploded_2["库位"].replace("", pd.NA).dropna().nunique() if len(exploded_2) else 0)

    if len(orders_2) == 0:
        st.warning(f"⚠️ {sel_date_2} 没有需要拣货的订单")
        st.stop()

    # ---- 面单核对 ----
    if label_pdf is not None:
        st.divider()
        st.subheader("🔍 面单核对结果")
        has_tess, has_popp = check_ocr_available()
        if not (has_tess and has_popp):
            st.error(
                "OCR 未启用：服务器缺少 tesseract / poppler-utils。"
                "本地：`brew install tesseract poppler` 或 `apt install tesseract-ocr poppler-utils`。"
                "Streamlit Cloud：在仓库根目录的 `packages.txt` 加上 `tesseract-ocr` 和 `poppler-utils`。"
            )
        else:
            with st.spinner("OCR 解析 PDF 中（每页约 1-2 秒）..."):
                try:
                    labels = parse_label_pdf(label_pdf.read())
                    rec = reconcile_labels(orders_2, labels)
                except Exception as e:
                    st.error(f"PDF 解析失败：{e}")
                    rec = None

            if rec is not None:
                ok = (rec["状态"] == "✅ 一致").sum()
                warn = rec["状态"].str.startswith("⚠️").sum()
                bad = rec["状态"].str.startswith("❌").sum()
                cc1, cc2, cc3 = st.columns(3)
                cc1.metric("✅ 一致", ok)
                cc2.metric("⚠️ 警告", warn)
                cc3.metric("❌ 不一致", bad)

                def color_status(val):
                    if val.startswith("✅"): return "background-color: #d4edda"
                    if val.startswith("⚠️"): return "background-color: #fff3cd"
                    if val.startswith("❌"): return "background-color: #f8d7da"
                    return ""
                styled = rec.style.map(color_status, subset=["状态"])
                st.dataframe(styled, use_container_width=True, height=400)

                if bad == 0 and warn == 0:
                    st.success(f"🎉 所有 {len(rec)} 单面单核对通过，可以发货！")
                elif bad > 0:
                    st.error(f"⚠️ {bad} 单严重不一致，请人工核对再发货")

    # ---- 生成拣货单 + 发货单 ----
    st.divider()
    st.subheader("📦 生成拣货单 + 发货单")

    if catalog2 is not None:
        missing = exploded_2[(exploded_2["中文名"] == "") & (exploded_2["英文款式"] != "")]
        if len(missing) > 0:
            unmatched = missing["英文款式"].unique().tolist()
            st.warning(
                f"⚠️ {len(missing)} 行未在图册匹配到："
                + ", ".join(unmatched[:10]) + ("..." if len(unmatched) > 10 else "")
            )

    with st.expander("👀 拣货明细预览", expanded=False):
        st.dataframe(
            exploded_2[["Order ID","中文名","英文款式","SKU","Size","库位"]],
            use_container_width=True, height=300,
        )

    if st.button("✨ 生成拣货单 + 发货单", type="primary", use_container_width=True):
        with st.spinner("生成中..."):
            picking_csv = build_picking_summary_csv(exploded_2)
            slip_csv = build_packing_slip_csv(exploded_2, sel_date_2)
        st.success(f"✅ 已生成 {sel_date_2} 的发货文件")

        date_compact = sel_date_2.replace("/", "")
        cc1, cc2 = st.columns(2)
        with cc1:
            st.download_button(
                "📋 拣货单（按库位汇总）",
                data=picking_csv,
                file_name=f"拣货单_{date_compact}.csv",
                mime="text/csv", use_container_width=True,
            )
        with cc2:
            st.download_button(
                "📦 发货单 / Packing Slip",
                data=slip_csv,
                file_name=f"发货单_{date_compact}.csv",
                mime="text/csv", use_container_width=True,
            )
