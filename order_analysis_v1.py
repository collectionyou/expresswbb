from __future__ import annotations

import argparse
import csv
import hashlib
import html
import io
import json
import pathlib
import re
import sqlite3
import urllib.parse
import zipfile
import cgi
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


SCRIPT_PATH = pathlib.Path(__file__).resolve()
WORKSPACE_ROOT = SCRIPT_PATH.parent.parent if SCRIPT_PATH.parent.name == "outputs" else SCRIPT_PATH.parent
DEFAULT_STATE_DIR = WORKSPACE_ROOT / "work" / "order_analysis_v1"
DEFAULT_DB_PATH = DEFAULT_STATE_DIR / "order_analysis_v1.sqlite3"
DEFAULT_UPLOAD_DIR = DEFAULT_STATE_DIR / "imports"
DEFAULT_SUMMARY_PATH = SCRIPT_PATH.with_name("order_analysis_v1.latest.json")

XLSX_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
TIME_FIELD_MAP = {
    "imported_at": "imported_at",
    "order_time": "order_time",
    "pay_time": "pay_time",
    "print_time": "print_time",
    "ship_time": "ship_time",
    "last_print_time": "last_print_time",
}
SIZE_RE = re.compile(r"(?:XXXL|XXL|XL|XS|L|M|S|2XL|3XL|4XL|5XL|6XL|\u5747\u7801|\b4\b)", re.IGNORECASE)
DATE_TIME_RE = re.compile(r"^\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2}:\d{2})?$")


def now_text() -> str:
    from datetime import datetime

    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", text).strip()


def csv_escape(value: object) -> str:
    return clean_text(value)


def sha1_bytes(data: bytes) -> str:
    return hashlib.sha1(data).hexdigest()


def parse_money(value: object) -> Optional[float]:
    text = clean_text(value)
    if not text:
        return None
    text = text.replace("￥", "").replace(",", "").replace("元", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def parse_int(value: object, default: int = 0) -> int:
    text = clean_text(value)
    if not text:
        return default
    match = re.search(r"-?\d+", text)
    if not match:
        return default
    try:
        return int(match.group(0))
    except ValueError:
        return default


def parse_float(value: object) -> Optional[float]:
    text = clean_text(value)
    if not text:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def normalize_header(value: object) -> str:
    return clean_text(value).replace(" ", "")


def split_multi_value(value: object) -> List[str]:
    text = clean_text(value).strip(";；")
    if not text:
        return []
    parts = [clean_text(item) for item in re.split(r"[;；]+", text) if clean_text(item)]
    return parts or ([text] if text else [])


def detect_refund_status(*texts: object) -> str:
    joined = " ".join(clean_text(text) for text in texts if clean_text(text))
    if not joined:
        return "正常"
    for keyword in ("退款中", "退款成功", "已退款", "退货中", "退货成功", "退单", "已取消", "已关闭"):
        if keyword in joined:
            return keyword
    if "退款" in joined or "退货" in joined or "取消" in joined:
        return "异常"
    return "正常"


def derive_shipped_status(ship_time: object, order_status: object, waybill: object) -> str:
    if clean_text(ship_time):
        return "已发货"
    text = " ".join(clean_text(item) for item in (order_status, waybill) if clean_text(item))
    if any(keyword in text for keyword in ("待收货", "已发货", "已签收", "运输中")):
        return "已发货"
    return "未发货"


def derive_print_status(print_time: object, waybill: object, raw_status: object) -> str:
    if clean_text(raw_status):
        return clean_text(raw_status)
    if clean_text(print_time) or clean_text(waybill):
        return "已打印"
    return "未知"


def derive_product_category(name: object, spec: object) -> str:
    text = f"{clean_text(name)} {clean_text(spec)}"
    if not text.strip():
        return "未分类"

    shape = "其他"
    if "睡袍" in text or "袍子" in text:
        shape = "睡袍"
    elif any(keyword in text for keyword in ("家居服", "套装", "两件套")):
        shape = "家居服套装"

    tags: List[str] = [shape]
    for keyword in ("珊瑚绒", "法兰绒", "半边绒", "拉链", "连帽", "立领", "豹纹", "方格", "小浣熊", "男女同款"):
        if keyword in text:
            tags.append(keyword)
    return " / ".join(dict.fromkeys(tags))


def derive_style_key(
    product_name: object,
    spec_name: object,
    product_code: object,
    sku_code: object,
    product_short_name: object,
    product_id: object,
) -> str:
    for candidate in (
        clean_text(sku_code),
        clean_text(product_code),
        clean_text(product_short_name),
        clean_text(product_id),
    ):
        if candidate and candidate not in {"未设置", "0"}:
            return candidate
    title = clean_text(product_name)
    if not title:
        return "未命名款式"
    title = re.sub(r"202\d新款|网红风|很划算|外穿|加厚|加绒|保暖", "", title)
    title = clean_text(title)
    spec = clean_text(spec_name)
    if spec:
        spec_head = re.split(r"[,，]", spec, maxsplit=1)[0]
        if spec_head and spec_head not in title:
            return clean_text(f"{title} / {spec_head}")[:48]
    return title[:48]


def derive_variant_tag(spec_name: object) -> str:
    spec = clean_text(spec_name)
    if not spec:
        return ""
    return clean_text(re.split(r"[,，]", spec, maxsplit=1)[0])


def derive_size_tag(spec_name: object) -> str:
    spec = clean_text(spec_name)
    if not spec:
        return ""
    match = SIZE_RE.search(spec)
    return clean_text(match.group(0)) if match else ""


def join_address(region: object, detail: object) -> str:
    region_text = clean_text(region)
    detail_text = clean_text(detail)
    if region_text and detail_text:
        return clean_text(f"{region_text} {detail_text}")
    return region_text or detail_text


def excel_col_to_index(value: str) -> int:
    index = 0
    for char in value:
        if char.isalpha():
            index = index * 26 + ord(char.upper()) - 64
    return index


def read_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    result: List[str] = []
    for si in root.findall("a:si", XLSX_NS):
        parts = [item.text or "" for item in si.iterfind(".//a:t", XLSX_NS)]
        result.append("".join(parts))
    return result


def read_xlsx_rows(data: bytes) -> List[List[str]]:
    rows: List[List[str]] = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        shared = read_shared_strings(zf)
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        first_sheet = workbook.find("a:sheets/a:sheet", XLSX_NS)
        if first_sheet is None:
            return []
        rel_id = first_sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")

        rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        target = ""
        for rel in rel_root:
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib.get("Target", "")
                break
        if not target:
            target = "worksheets/sheet1.xml"
        worksheet_path = f"xl/{target}".replace("\\", "/")

        root = ET.fromstring(zf.read(worksheet_path))
        for row in root.findall(".//a:sheetData/a:row", XLSX_NS):
            values: Dict[int, str] = {}
            for cell in row.findall("a:c", XLSX_NS):
                ref = cell.attrib.get("r", "")
                match = re.match(r"([A-Z]+)(\d+)", ref)
                if not match:
                    continue
                column_index = excel_col_to_index(match.group(1))
                cell_type = cell.attrib.get("t", "")
                value = ""
                if cell_type == "inlineStr":
                    parts = [item.text or "" for item in cell.iterfind(".//a:t", XLSX_NS)]
                    value = "".join(parts)
                else:
                    value_node = cell.find("a:v", XLSX_NS)
                    if value_node is not None and value_node.text is not None:
                        raw = value_node.text
                        if cell_type == "s":
                            try:
                                value = shared[int(raw)]
                            except (ValueError, IndexError):
                                value = raw
                        else:
                            value = raw
                values[column_index] = clean_text(value)
            if values:
                max_column = max(values)
                rows.append([values.get(index, "") for index in range(1, max_column + 1)])
    return rows


def decode_csv_bytes(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "ignore")


def read_csv_rows(data: bytes) -> List[List[str]]:
    text = decode_csv_bytes(data)
    buffer = io.StringIO(text)
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample)
    except csv.Error:
        dialect = csv.excel
    reader = csv.reader(buffer, dialect)
    return [[clean_text(item) for item in row] for row in reader]


def rows_to_dicts(rows: Sequence[Sequence[str]]) -> Tuple[List[str], List[Dict[str, str]]]:
    if not rows:
        return [], []
    headers = [clean_text(item) for item in rows[0]]
    records: List[Dict[str, str]] = []
    for raw_row in rows[1:]:
        if not any(clean_text(item) for item in raw_row):
            continue
        item = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            item[header] = clean_text(raw_row[index]) if index < len(raw_row) else ""
        records.append(item)
    return headers, records


@dataclass
class TemplateSpec:
    key: str
    label: str
    platform: str
    required_headers: Tuple[str, ...]


TEMPLATES = [
    TemplateSpec(
        key="pdd_ship_export",
        label="拼多多-发货单导出",
        platform="拼多多",
        required_headers=("订单号", "店铺名称", "发货时间", "商品名称", "规格名称", "数量", "实收"),
    ),
    TemplateSpec(
        key="pdd_print_sheet",
        label="拼多多-打印单",
        platform="拼多多",
        required_headers=("订单编号", "收件人", "收件人电话", "商品信息", "订单状态"),
    ),
    TemplateSpec(
        key="pdd_order_export",
        label="拼多多-订单导出",
        platform="拼多多",
        required_headers=("订单编号", "下单时间", "付款时间", "快递单打印时间", "发货时间", "商品简称", "销售规格", "货号"),
    ),
]


def detect_template(headers: Sequence[str]) -> Optional[TemplateSpec]:
    normalized = {normalize_header(item) for item in headers if clean_text(item)}
    for template in TEMPLATES:
        required = {normalize_header(item) for item in template.required_headers}
        if required.issubset(normalized):
            return template
    return None


def build_record(
    *,
    source_file: str,
    source_path: str,
    source_template: str,
    platform: str,
    shop_name: str,
    order_id: str,
    order_status: str = "",
    refund_status: str = "正常",
    shipped_status: str = "",
    print_status: str = "",
    presale_status: str = "",
    buyer_nick: str = "",
    receiver_name: str = "",
    receiver_phone: str = "",
    region: str = "",
    detail_address: str = "",
    order_time: str = "",
    pay_time: str = "",
    print_time: str = "",
    ship_time: str = "",
    last_print_time: str = "",
    product_name: str = "",
    product_short_name: str = "",
    product_code: str = "",
    product_id: str = "",
    spec_name: str = "",
    spec_short_name: str = "",
    sku_code: str = "",
    quantity: int = 0,
    weight_kg: Optional[float] = None,
    total_price: Optional[float] = None,
    paid_amount: Optional[float] = None,
    courier: str = "",
    waybill: str = "",
    remark: str = "",
    buyer_message: str = "",
    seller_remark: str = "",
    gift_info: str = "",
    raw: Optional[Dict[str, str]] = None,
) -> Dict[str, object]:
    refund_status = refund_status or "正常"
    shipped_status = shipped_status or derive_shipped_status(ship_time, order_status, waybill)
    print_status = print_status or derive_print_status(print_time or last_print_time, waybill, print_status)
    full_address = join_address(region, detail_address)
    category = derive_product_category(product_name or product_short_name, spec_name)
    style_key = derive_style_key(product_name, spec_name, product_code, sku_code, product_short_name, product_id)
    variant_tag = derive_variant_tag(spec_name)
    size_tag = derive_size_tag(spec_name)
    return {
        "source_file": clean_text(source_file),
        "source_path": clean_text(source_path),
        "source_template": clean_text(source_template),
        "platform": clean_text(platform),
        "shop_name": clean_text(shop_name),
        "order_id": clean_text(order_id),
        "order_status": clean_text(order_status),
        "refund_status": clean_text(refund_status),
        "shipped_status": clean_text(shipped_status),
        "print_status": clean_text(print_status),
        "presale_status": clean_text(presale_status),
        "buyer_nick": clean_text(buyer_nick),
        "receiver_name": clean_text(receiver_name),
        "receiver_phone": clean_text(receiver_phone),
        "region": clean_text(region),
        "detail_address": clean_text(detail_address),
        "full_address": clean_text(full_address),
        "order_time": clean_text(order_time),
        "pay_time": clean_text(pay_time),
        "print_time": clean_text(print_time),
        "ship_time": clean_text(ship_time),
        "last_print_time": clean_text(last_print_time),
        "product_category": clean_text(category),
        "style_key": clean_text(style_key),
        "variant_tag": clean_text(variant_tag),
        "size_tag": clean_text(size_tag),
        "product_name": clean_text(product_name),
        "product_short_name": clean_text(product_short_name),
        "product_code": clean_text(product_code),
        "product_id": clean_text(product_id),
        "spec_name": clean_text(spec_name),
        "spec_short_name": clean_text(spec_short_name),
        "sku_code": clean_text(sku_code),
        "quantity": quantity or 0,
        "weight_kg": weight_kg,
        "total_price": total_price,
        "paid_amount": paid_amount,
        "courier": clean_text(courier),
        "waybill": clean_text(waybill),
        "remark": clean_text(remark),
        "buyer_message": clean_text(buyer_message),
        "seller_remark": clean_text(seller_remark),
        "gift_info": clean_text(gift_info),
        "raw_json": json.dumps(raw or {}, ensure_ascii=False),
    }


def parse_pdd_ship_export(
    rows: Sequence[Dict[str, str]],
    source_file: str,
    source_path: str,
    platform_override: str,
    shop_override: str,
) -> List[Dict[str, object]]:
    result: List[Dict[str, object]] = []
    for row in rows:
        result.append(
            build_record(
                source_file=source_file,
                source_path=source_path,
                source_template="拼多多-发货单导出",
                platform=platform_override or "拼多多",
                shop_name=shop_override or row.get("店铺名称", ""),
                order_id=row.get("订单号", ""),
                shipped_status=derive_shipped_status(row.get("发货时间", ""), "", row.get("运单号", "")),
                print_status=row.get("打印状态", ""),
                presale_status=row.get("预售", ""),
                receiver_name=row.get("收件人", ""),
                receiver_phone=row.get("联系方式", ""),
                region=row.get("省/市/区", ""),
                detail_address=row.get("详细地址", ""),
                ship_time=row.get("发货时间", ""),
                last_print_time=row.get("最后打印时间", ""),
                product_name=row.get("商品名称", ""),
                product_short_name=row.get("商品简称", ""),
                product_code=row.get("商品编码", ""),
                product_id=row.get("商品ID", ""),
                spec_name=row.get("规格名称", ""),
                spec_short_name=row.get("规格简称", ""),
                sku_code=row.get("商品编码", "") or row.get("商品ID", ""),
                quantity=parse_int(row.get("数量", ""), 1),
                weight_kg=parse_float(row.get("重量（kg）", "")),
                total_price=parse_money(row.get("总价", "")),
                paid_amount=parse_money(row.get("实收", "")),
                courier=row.get("快递", ""),
                waybill=row.get("运单号", ""),
                remark=row.get("备注", ""),
                buyer_message=row.get("买家留言", ""),
                gift_info=row.get("赠品信息", ""),
                refund_status=detect_refund_status(row.get("备注", ""), row.get("买家留言", "")),
                raw=row,
            )
        )
    return result


def parse_print_goods_entry(text: str) -> List[str]:
    parts = [clean_text(item) for item in text.split(",")]
    while parts and not parts[-1]:
        parts.pop()
    return parts


def parse_pdd_print_sheet(
    rows: Sequence[Dict[str, str]],
    source_file: str,
    source_path: str,
    platform_override: str,
    shop_override: str,
) -> List[Dict[str, object]]:
    result: List[Dict[str, object]] = []
    for row in rows:
        items = split_multi_value(row.get("商品信息", ""))
        if not items:
            items = [""]
        for item_text in items:
            fields = parse_print_goods_entry(item_text)
            product_name = fields[0] if len(fields) > 0 else ""
            sku_code = fields[1] if len(fields) > 1 else ""
            spec_name = fields[2] if len(fields) > 2 else ""
            quantity = parse_int(fields[3] if len(fields) > 3 else "", 1)
            trailing = " ".join(fields[4:]) if len(fields) > 4 else ""
            refund_status = detect_refund_status(trailing, row.get("订单状态", ""))
            result.append(
                build_record(
                    source_file=source_file,
                    source_path=source_path,
                    source_template="拼多多-打印单",
                    platform=platform_override or "拼多多",
                    shop_name=shop_override,
                    order_id=row.get("订单编号", ""),
                    order_status=row.get("订单状态", ""),
                    refund_status=refund_status,
                    shipped_status=derive_shipped_status("", row.get("订单状态", ""), row.get("快递单号", "")),
                    print_status=derive_print_status("", row.get("快递单号", ""), "已打印" if clean_text(row.get("快递单号", "")) else ""),
                    receiver_name=row.get("收件人", ""),
                    receiver_phone=row.get("收件人电话", ""),
                    detail_address=row.get("收件人地址", ""),
                    product_name=product_name,
                    sku_code=sku_code,
                    spec_name=spec_name,
                    quantity=quantity,
                    courier=row.get("快递公司", ""),
                    waybill=row.get("快递单号", ""),
                    remark=trailing,
                    raw={**row, "商品信息明细": item_text},
                )
            )
    return result


def split_aligned_cells(row: Dict[str, str], *keys: str) -> Dict[str, List[str]]:
    split_map = {key: split_multi_value(row.get(key, "")) for key in keys}
    max_len = max((len(values) for values in split_map.values()), default=0)
    if max_len <= 1:
        return split_map
    for key, values in split_map.items():
        while len(values) < max_len:
            values.append("")
    return split_map


def allocate_paid_amounts(total_paid: Optional[float], unit_prices: List[Optional[float]], quantities: List[int]) -> List[Optional[float]]:
    line_totals = [(unit_prices[index] or 0.0) * max(quantities[index], 0) for index in range(len(quantities))]
    total_line_amount = sum(line_totals)
    if total_paid is None:
        return [line_totals[index] if line_totals[index] else None for index in range(len(line_totals))]
    if total_line_amount <= 0:
        return [total_paid if index == 0 else 0.0 for index in range(len(line_totals))]
    allocations: List[Optional[float]] = []
    running = 0.0
    for index, line_total in enumerate(line_totals):
        if index == len(line_totals) - 1:
            value = round(total_paid - running, 2)
        else:
            value = round(total_paid * (line_total / total_line_amount), 2)
            running += value
        allocations.append(value)
    return allocations


def parse_pdd_order_export(
    rows: Sequence[Dict[str, str]],
    source_file: str,
    source_path: str,
    platform_override: str,
    shop_override: str,
) -> List[Dict[str, object]]:
    result: List[Dict[str, object]] = []
    for row in rows:
        aligned = split_aligned_cells(row, "商品简称", "销售规格", "货号", "商品数量", "商品单价")
        item_count = max((len(values) for values in aligned.values()), default=1)
        product_names = aligned.get("商品简称", [""])
        spec_names = aligned.get("销售规格", [""])
        sku_codes = aligned.get("货号", [""])
        quantities = [parse_int(item, 1) for item in aligned.get("商品数量", ["1"])]
        unit_prices = [parse_money(item) for item in aligned.get("商品单价", [""])]
        paid_allocations = allocate_paid_amounts(parse_money(row.get("支付金额", "")), unit_prices, quantities)

        for index in range(item_count):
            result.append(
                build_record(
                    source_file=source_file,
                    source_path=source_path,
                    source_template="拼多多-订单导出",
                    platform=platform_override or "拼多多",
                    shop_name=shop_override or row.get("店铺", ""),
                    order_id=row.get("订单编号", ""),
                    shipped_status=derive_shipped_status(row.get("发货时间", ""), "", row.get("快递单号", "")),
                    print_status=derive_print_status(row.get("快递单打印时间", ""), row.get("快递单号", ""), ""),
                    buyer_nick=row.get("买家昵称", ""),
                    receiver_name=row.get("收件人名称", ""),
                    receiver_phone=row.get("收件人手机号码", ""),
                    detail_address=row.get("收件人省市区详细地址", ""),
                    order_time=row.get("下单时间", ""),
                    pay_time=row.get("付款时间", ""),
                    print_time=row.get("快递单打印时间", ""),
                    ship_time=row.get("发货时间", ""),
                    product_name=product_names[index] if index < len(product_names) else "",
                    product_short_name=product_names[index] if index < len(product_names) else "",
                    spec_name=spec_names[index] if index < len(spec_names) else "",
                    sku_code=sku_codes[index] if index < len(sku_codes) else "",
                    quantity=quantities[index] if index < len(quantities) else 0,
                    total_price=(unit_prices[index] or 0.0) * (quantities[index] if index < len(quantities) else 0) if index < len(unit_prices) and unit_prices[index] is not None else None,
                    paid_amount=paid_allocations[index] if index < len(paid_allocations) else None,
                    courier=row.get("快递公司", ""),
                    waybill=row.get("快递单号", ""),
                    buyer_message=row.get("买家留言", ""),
                    seller_remark=row.get("卖家备注", ""),
                    refund_status=detect_refund_status(row.get("买家留言", ""), row.get("卖家备注", "")),
                    raw={**row, "line_index": str(index + 1)},
                )
            )
    return result


def parse_by_template(
    template: TemplateSpec,
    rows: Sequence[Dict[str, str]],
    source_file: str,
    source_path: str,
    platform_override: str,
    shop_override: str,
) -> List[Dict[str, object]]:
    if template.key == "pdd_ship_export":
        return parse_pdd_ship_export(rows, source_file, source_path, platform_override, shop_override)
    if template.key == "pdd_print_sheet":
        return parse_pdd_print_sheet(rows, source_file, source_path, platform_override, shop_override)
    if template.key == "pdd_order_export":
        return parse_pdd_order_export(rows, source_file, source_path, platform_override, shop_override)
    raise ValueError(f"unsupported template: {template.key}")


class OrderAnalysisService:
    def __init__(self, args: argparse.Namespace) -> None:
        self.args = args
        self.state_dir = pathlib.Path(args.state_dir)
        self.db_path = pathlib.Path(args.db_path)
        self.upload_dir = pathlib.Path(args.upload_dir)
        self.summary_path = pathlib.Path(args.summary_path)
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.upload_dir.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_db()

    def close(self) -> None:
        self.conn.close()

    def _init_db(self) -> None:
        self.conn.executescript(
            """
            PRAGMA journal_mode=WAL;

            CREATE TABLE IF NOT EXISTS imports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                imported_at TEXT NOT NULL,
                file_name TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                source_path TEXT NOT NULL,
                sha1 TEXT NOT NULL UNIQUE,
                template_key TEXT NOT NULL,
                template_label TEXT NOT NULL,
                platform TEXT NOT NULL,
                shop_label TEXT NOT NULL,
                row_count INTEGER NOT NULL,
                line_count INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                import_id INTEGER NOT NULL,
                imported_at TEXT NOT NULL,
                source_file TEXT NOT NULL,
                source_path TEXT NOT NULL,
                source_template TEXT NOT NULL,
                platform TEXT NOT NULL,
                shop_name TEXT NOT NULL,
                order_id TEXT NOT NULL,
                order_status TEXT NOT NULL,
                refund_status TEXT NOT NULL,
                shipped_status TEXT NOT NULL,
                print_status TEXT NOT NULL,
                presale_status TEXT NOT NULL,
                buyer_nick TEXT NOT NULL,
                receiver_name TEXT NOT NULL,
                receiver_phone TEXT NOT NULL,
                region TEXT NOT NULL,
                detail_address TEXT NOT NULL,
                full_address TEXT NOT NULL,
                order_time TEXT NOT NULL,
                pay_time TEXT NOT NULL,
                print_time TEXT NOT NULL,
                ship_time TEXT NOT NULL,
                last_print_time TEXT NOT NULL,
                product_category TEXT NOT NULL,
                style_key TEXT NOT NULL,
                variant_tag TEXT NOT NULL,
                size_tag TEXT NOT NULL,
                product_name TEXT NOT NULL,
                product_short_name TEXT NOT NULL,
                product_code TEXT NOT NULL,
                product_id TEXT NOT NULL,
                spec_name TEXT NOT NULL,
                spec_short_name TEXT NOT NULL,
                sku_code TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                weight_kg REAL,
                total_price REAL,
                paid_amount REAL,
                courier TEXT NOT NULL,
                waybill TEXT NOT NULL,
                remark TEXT NOT NULL,
                buyer_message TEXT NOT NULL,
                seller_remark TEXT NOT NULL,
                gift_info TEXT NOT NULL,
                raw_json TEXT NOT NULL,
                FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_records_order_time ON records(order_time);
            CREATE INDEX IF NOT EXISTS idx_records_pay_time ON records(pay_time);
            CREATE INDEX IF NOT EXISTS idx_records_print_time ON records(print_time);
            CREATE INDEX IF NOT EXISTS idx_records_ship_time ON records(ship_time);
            CREATE INDEX IF NOT EXISTS idx_records_platform ON records(platform);
            CREATE INDEX IF NOT EXISTS idx_records_shop_name ON records(shop_name);
            CREATE INDEX IF NOT EXISTS idx_records_style_key ON records(style_key);
            CREATE INDEX IF NOT EXISTS idx_records_product_category ON records(product_category);
            """
        )
        self.conn.commit()

    def clear_all(self) -> None:
        self.conn.execute("DELETE FROM records")
        self.conn.execute("DELETE FROM imports")
        self.conn.commit()
        if self.upload_dir.exists():
            for child in self.upload_dir.iterdir():
                if child.is_file():
                    child.unlink(missing_ok=True)

    def import_path(self, path_text: str, platform_label: str = "", shop_label: str = "") -> Dict[str, object]:
        path = pathlib.Path(path_text).expanduser()
        data = path.read_bytes()
        return self.import_bytes(path.name, data, str(path), platform_label, shop_label)

    def import_bytes(
        self,
        file_name: str,
        data: bytes,
        source_path: str = "",
        platform_label: str = "",
        shop_label: str = "",
    ) -> Dict[str, object]:
        file_hash = sha1_bytes(data)
        existing = self.conn.execute("SELECT * FROM imports WHERE sha1 = ?", (file_hash,)).fetchone()
        if existing is not None:
            return {
                "status": "duplicate",
                "message": f"{file_name} already imported",
                "import_id": existing["id"],
                "template_label": existing["template_label"],
                "line_count": existing["line_count"],
            }

        suffix = pathlib.Path(file_name).suffix.lower()
        if suffix == ".xlsx":
            raw_rows = read_xlsx_rows(data)
        elif suffix == ".csv":
            raw_rows = read_csv_rows(data)
        else:
            raise ValueError(f"unsupported file type: {file_name}")

        headers, row_dicts = rows_to_dicts(raw_rows)
        template = detect_template(headers)
        if template is None:
            raise ValueError(f"unrecognized template: {file_name}")

        records = parse_by_template(template, row_dicts, file_name, source_path, platform_label, shop_label)
        if not records:
            raise ValueError(f"no data rows found: {file_name}")

        imported_at = now_text()
        stored_name = f"{imported_at.replace(':', '').replace(' ', '_')}_{pathlib.Path(file_name).name}"
        stored_path = self.upload_dir / stored_name
        stored_path.write_bytes(data)

        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO imports (
                imported_at, file_name, stored_path, source_path, sha1,
                template_key, template_label, platform, shop_label, row_count, line_count
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                imported_at,
                file_name,
                str(stored_path),
                source_path,
                file_hash,
                template.key,
                template.label,
                platform_label or template.platform,
                shop_label,
                len(row_dicts),
                len(records),
            ),
        )
        import_id = cursor.lastrowid

        insert_sql = """
            INSERT INTO records (
                import_id, imported_at, source_file, source_path, source_template, platform, shop_name,
                order_id, order_status, refund_status, shipped_status, print_status, presale_status,
                buyer_nick, receiver_name, receiver_phone, region, detail_address, full_address,
                order_time, pay_time, print_time, ship_time, last_print_time,
                product_category, style_key, variant_tag, size_tag,
                product_name, product_short_name, product_code, product_id, spec_name, spec_short_name, sku_code,
                quantity, weight_kg, total_price, paid_amount,
                courier, waybill, remark, buyer_message, seller_remark, gift_info, raw_json
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?
            )
        """
        for record in records:
            cursor.execute(
                insert_sql,
                (
                    import_id,
                    imported_at,
                    record["source_file"],
                    record["source_path"],
                    record["source_template"],
                    record["platform"],
                    record["shop_name"],
                    record["order_id"],
                    record["order_status"],
                    record["refund_status"],
                    record["shipped_status"],
                    record["print_status"],
                    record["presale_status"],
                    record["buyer_nick"],
                    record["receiver_name"],
                    record["receiver_phone"],
                    record["region"],
                    record["detail_address"],
                    record["full_address"],
                    record["order_time"],
                    record["pay_time"],
                    record["print_time"],
                    record["ship_time"],
                    record["last_print_time"],
                    record["product_category"],
                    record["style_key"],
                    record["variant_tag"],
                    record["size_tag"],
                    record["product_name"],
                    record["product_short_name"],
                    record["product_code"],
                    record["product_id"],
                    record["spec_name"],
                    record["spec_short_name"],
                    record["sku_code"],
                    int(record["quantity"]),
                    record["weight_kg"],
                    record["total_price"],
                    record["paid_amount"],
                    record["courier"],
                    record["waybill"],
                    record["remark"],
                    record["buyer_message"],
                    record["seller_remark"],
                    record["gift_info"],
                    record["raw_json"],
                ),
            )
        self.conn.commit()
        self.write_summary()
        return {
            "status": "ok",
            "message": f"imported {file_name}",
            "import_id": import_id,
            "template_label": template.label,
            "line_count": len(records),
        }

    def write_summary(self) -> None:
        stats = self.query_stats({})
        imports = [dict(row) for row in self.conn.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 20")]
        payload = {
            "generated_at": now_text(),
            "stats": stats,
            "imports": imports,
        }
        self.summary_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def parse_filters(self, query: Dict[str, List[str]]) -> Dict[str, str]:
        def take(name: str, default: str = "") -> str:
            values = query.get(name, [])
            return clean_text(values[0]) if values else default

        return {
            "time_field": take("time_field", "ship_time"),
            "date_from": take("date_from"),
            "date_to": take("date_to"),
            "platform": take("platform"),
            "shop_name": take("shop_name"),
            "template": take("template"),
            "product_category": take("product_category"),
            "shipped_status": take("shipped_status"),
            "refund_mode": take("refund_mode"),
            "print_status": take("print_status"),
            "keyword": take("keyword"),
            "limit": take("limit", "200"),
            "message": take("message"),
        }

    def _build_where(self, filters: Dict[str, str]) -> Tuple[str, List[object]]:
        clauses: List[str] = []
        params: List[object] = []

        time_column = TIME_FIELD_MAP.get(filters.get("time_field", ""), "ship_time")
        if filters.get("date_from"):
            clauses.append(f"{time_column} >= ?")
            params.append(f"{filters['date_from']} 00:00:00")
        if filters.get("date_to"):
            clauses.append(f"{time_column} <= ?")
            params.append(f"{filters['date_to']} 23:59:59")

        for field in ("platform", "shop_name", "product_category"):
            if filters.get(field):
                clauses.append(f"{field} = ?")
                params.append(filters[field])
        if filters.get("template"):
            clauses.append("source_template = ?")
            params.append(filters["template"])
        if filters.get("shipped_status"):
            clauses.append("shipped_status = ?")
            params.append(filters["shipped_status"])
        if filters.get("print_status"):
            clauses.append("print_status = ?")
            params.append(filters["print_status"])

        refund_mode = filters.get("refund_mode", "")
        if refund_mode == "normal":
            clauses.append("refund_status = '正常'")
        elif refund_mode == "refund":
            clauses.append("refund_status <> '正常'")

        keyword = filters.get("keyword", "")
        if keyword:
            like_value = f"%{keyword}%"
            clauses.append(
                "("
                "order_id LIKE ? OR product_name LIKE ? OR spec_name LIKE ? OR sku_code LIKE ? OR "
                "style_key LIKE ? OR receiver_name LIKE ? OR waybill LIKE ? OR full_address LIKE ?"
                ")"
            )
            params.extend([like_value] * 8)

        if not clauses:
            return "", params
        return " WHERE " + " AND ".join(clauses), params

    def query_stats(self, filters: Dict[str, str]) -> Dict[str, object]:
        where_sql, params = self._build_where(filters)
        row = self.conn.execute(
            f"""
            SELECT
                COUNT(*) AS line_count,
                COUNT(DISTINCT order_id) AS order_count,
                COALESCE(SUM(quantity), 0) AS quantity_sum,
                COALESCE(SUM(paid_amount), 0) AS paid_sum,
                COALESCE(SUM(CASE WHEN shipped_status = '已发货' THEN quantity ELSE 0 END), 0) AS shipped_quantity,
                COALESCE(SUM(CASE WHEN refund_status <> '正常' THEN quantity ELSE 0 END), 0) AS refund_quantity,
                COALESCE(SUM(CASE WHEN print_status LIKE '%打印%' OR print_status = '已打印' THEN quantity ELSE 0 END), 0) AS printed_quantity
            FROM records
            {where_sql}
            """,
            params,
        ).fetchone()
        import_count = self.conn.execute("SELECT COUNT(*) FROM imports").fetchone()[0]
        return {
            "import_count": import_count,
            "line_count": int(row["line_count"]),
            "order_count": int(row["order_count"]),
            "quantity_sum": int(row["quantity_sum"]),
            "paid_sum": round(float(row["paid_sum"] or 0), 2),
            "shipped_quantity": int(row["shipped_quantity"]),
            "refund_quantity": int(row["refund_quantity"]),
            "printed_quantity": int(row["printed_quantity"]),
        }

    def query_summary_rows(self, filters: Dict[str, str]) -> List[sqlite3.Row]:
        where_sql, params = self._build_where(filters)
        return list(
            self.conn.execute(
                f"""
                SELECT
                    platform,
                    shop_name,
                    product_category,
                    style_key,
                    MIN(product_name) AS sample_product_name,
                    MIN(spec_name) AS sample_spec_name,
                    COALESCE(SUM(quantity), 0) AS quantity_sum,
                    COUNT(DISTINCT order_id) AS order_count,
                    ROUND(COALESCE(SUM(paid_amount), 0), 2) AS paid_sum,
                    COALESCE(SUM(CASE WHEN refund_status <> '正常' THEN quantity ELSE 0 END), 0) AS refund_quantity
                FROM records
                {where_sql}
                GROUP BY platform, shop_name, product_category, style_key
                ORDER BY quantity_sum DESC, paid_sum DESC, style_key ASC
                LIMIT 300
                """,
                params,
            )
        )

    def query_detail_rows(self, filters: Dict[str, str]) -> List[sqlite3.Row]:
        where_sql, params = self._build_where(filters)
        try:
            limit = max(50, min(2000, int(filters.get("limit", "200"))))
        except ValueError:
            limit = 200
        return list(
            self.conn.execute(
                f"""
                SELECT *
                FROM records
                {where_sql}
                ORDER BY
                    COALESCE(NULLIF(ship_time, ''), NULLIF(pay_time, ''), NULLIF(order_time, ''), imported_at) DESC,
                    id DESC
                LIMIT ?
                """,
                params + [limit],
            )
        )

    def query_import_rows(self) -> List[sqlite3.Row]:
        return list(self.conn.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 50"))

    def query_filter_options(self) -> Dict[str, List[str]]:
        def values(column: str, table: str = "records") -> List[str]:
            rows = self.conn.execute(
                f"SELECT DISTINCT {column} AS value FROM {table} WHERE {column} <> '' ORDER BY {column} ASC LIMIT 200"
            ).fetchall()
            return [clean_text(row["value"]) for row in rows if clean_text(row["value"])]

        return {
            "platforms": values("platform"),
            "shops": values("shop_name"),
            "templates": values("source_template"),
            "categories": values("product_category"),
            "print_statuses": values("print_status"),
        }

    def export_details_csv(self, filters: Dict[str, str]) -> str:
        rows = self.query_detail_rows(filters)
        buffer = io.StringIO()
        writer = csv.writer(buffer)
        writer.writerow(
            [
                "平台",
                "店铺",
                "模板",
                "订单号",
                "下单时间",
                "付款时间",
                "打印时间",
                "发货时间",
                "发货状态",
                "退款状态",
                "订单状态",
                "打印状态",
                "收件人",
                "手机号",
                "地址",
                "商品种类",
                "款式键",
                "商品名称",
                "规格",
                "货号/编码",
                "数量",
                "实收金额",
                "总价",
                "快递",
                "运单号",
                "买家留言",
                "卖家备注",
                "来源文件",
            ]
        )
        for row in rows:
            writer.writerow(
                [
                    row["platform"],
                    row["shop_name"],
                    row["source_template"],
                    row["order_id"],
                    row["order_time"],
                    row["pay_time"],
                    row["print_time"] or row["last_print_time"],
                    row["ship_time"],
                    row["shipped_status"],
                    row["refund_status"],
                    row["order_status"],
                    row["print_status"],
                    row["receiver_name"],
                    row["receiver_phone"],
                    row["full_address"],
                    row["product_category"],
                    row["style_key"],
                    row["product_name"],
                    row["spec_name"],
                    row["sku_code"] or row["product_code"] or row["product_id"],
                    row["quantity"],
                    row["paid_amount"],
                    row["total_price"],
                    row["courier"],
                    row["waybill"],
                    row["buyer_message"],
                    row["seller_remark"] or row["remark"],
                    row["source_file"],
                ]
            )
        return buffer.getvalue()

    def export_summary_csv(self, filters: Dict[str, str]) -> str:
        rows = self.query_summary_rows(filters)
        buffer = io.StringIO()
        writer = csv.writer(buffer)
        writer.writerow(["平台", "店铺", "商品种类", "款式键", "示例商品", "示例规格", "销量件数", "订单数", "金额", "退款件数"])
        for row in rows:
            writer.writerow(
                [
                    row["platform"],
                    row["shop_name"],
                    row["product_category"],
                    row["style_key"],
                    row["sample_product_name"],
                    row["sample_spec_name"],
                    row["quantity_sum"],
                    row["order_count"],
                    row["paid_sum"],
                    row["refund_quantity"],
                ]
            )
        return buffer.getvalue()

    def json_summary(self, filters: Dict[str, str]) -> Dict[str, object]:
        return {
            "generated_at": now_text(),
            "filters": filters,
            "stats": self.query_stats(filters),
            "summary_rows": [dict(row) for row in self.query_summary_rows(filters)[:100]],
            "import_rows": [dict(row) for row in self.query_import_rows()[:20]],
        }

    def build_http_server(self) -> ThreadingHTTPServer:
        service = self

        class Handler(BaseHTTPRequestHandler):
            server_version = "OrderAnalysisV1/1.0"

            def log_message(self, format: str, *args: object) -> None:
                return

            def _redirect(self, location: str) -> None:
                self.send_response(302)
                self.send_header("Location", location)
                self.end_headers()

            def _html(self, text: str, status: int = 200) -> None:
                data = text.encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def _json(self, payload: Dict[str, object], status: int = 200) -> None:
                data = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def _csv(self, name: str, text: str) -> None:
                data = text.encode("utf-8-sig")
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{urllib.parse.quote(name)}")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def do_GET(self) -> None:
                parsed = urllib.parse.urlparse(self.path)
                query = urllib.parse.parse_qs(parsed.query)
                filters = service.parse_filters(query)

                if parsed.path == "/":
                    self._html(service.dashboard_html(filters))
                    return
                if parsed.path == "/api/summary":
                    self._json(service.json_summary(filters))
                    return
                if parsed.path == "/export/details.csv":
                    self._csv("明细导出.csv", service.export_details_csv(filters))
                    return
                if parsed.path == "/export/summary.csv":
                    self._csv("汇总导出.csv", service.export_summary_csv(filters))
                    return
                if parsed.path == "/healthz":
                    self._json({"ok": True, "generated_at": now_text()})
                    return
                self._html("<h1>Not Found</h1>", status=404)

            def do_POST(self) -> None:
                parsed = urllib.parse.urlparse(self.path)
                if parsed.path == "/reset":
                    service.clear_all()
                    service.write_summary()
                    self._redirect("/?message=" + urllib.parse.quote("已清空数据"))
                    return
                if parsed.path == "/import-paths":
                    length = int(self.headers.get("Content-Length", "0") or "0")
                    body = self.rfile.read(length).decode("utf-8", "ignore")
                    form = urllib.parse.parse_qs(body)
                    platform = clean_text(form.get("platform_label", [""])[0])
                    shop = clean_text(form.get("shop_label", [""])[0])
                    paths_text = form.get("paths", [""])[0]
                    messages: List[str] = []
                    for raw_line in paths_text.splitlines():
                        path_text = clean_text(raw_line)
                        if not path_text:
                            continue
                        try:
                            result = service.import_path(path_text, platform, shop)
                            messages.append(f"{pathlib.Path(path_text).name}: {result['status']}")
                        except Exception as exc:  # noqa: BLE001
                            messages.append(f"{pathlib.Path(path_text).name}: {exc}")
                    service.write_summary()
                    self._redirect("/?message=" + urllib.parse.quote(" | ".join(messages) if messages else "未提供路径"))
                    return
                if parsed.path == "/import-files":
                    environ = {
                        "REQUEST_METHOD": "POST",
                        "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                        "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                    }
                    form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=environ)
                    platform = clean_text(form.getfirst("platform_label", ""))
                    shop = clean_text(form.getfirst("shop_label", ""))
                    file_items = form["files"] if "files" in form else []
                    if not isinstance(file_items, list):
                        file_items = [file_items]
                    messages: List[str] = []
                    for item in file_items:
                        if not getattr(item, "filename", ""):
                            continue
                        file_name = pathlib.Path(item.filename).name
                        data = item.file.read()
                        try:
                            result = service.import_bytes(file_name, data, "", platform, shop)
                            messages.append(f"{file_name}: {result['status']}")
                        except Exception as exc:  # noqa: BLE001
                            messages.append(f"{file_name}: {exc}")
                    service.write_summary()
                    self._redirect("/?message=" + urllib.parse.quote(" | ".join(messages) if messages else "未选择文件"))
                    return
                self._html("<h1>Not Found</h1>", status=404)

        return ThreadingHTTPServer((self.args.host, self.args.port), Handler)

    def dashboard_html(self, filters: Dict[str, str]) -> str:
        stats = self.query_stats(filters)
        summary_rows = self.query_summary_rows(filters)
        detail_rows = self.query_detail_rows(filters)
        imports = self.query_import_rows()
        options = self.query_filter_options()
        query_string = urllib.parse.urlencode(
            {
                key: value
                for key, value in filters.items()
                if value and key != "message"
            }
        )

        def option_html(values: Iterable[str], selected: str) -> str:
            parts = ['<option value="">全部</option>']
            for value in values:
                escaped = html.escape(value)
                current = ' selected' if value == selected else ''
                parts.append(f'<option value="{escaped}"{current}>{escaped}</option>')
            return "".join(parts)

        summary_html = "".join(
            f"""
            <tr>
                <td>{html.escape(row['platform'])}</td>
                <td>{html.escape(row['shop_name'])}</td>
                <td>{html.escape(row['product_category'])}</td>
                <td>{html.escape(row['style_key'])}</td>
                <td title="{html.escape(row['sample_product_name'] or '')}">{html.escape(row['sample_product_name'] or '')}</td>
                <td title="{html.escape(row['sample_spec_name'] or '')}">{html.escape(row['sample_spec_name'] or '')}</td>
                <td class="num">{row['quantity_sum']}</td>
                <td class="num">{row['order_count']}</td>
                <td class="num">{row['paid_sum']}</td>
                <td class="num">{row['refund_quantity']}</td>
            </tr>
            """
            for row in summary_rows
        )
        detail_html = "".join(
            f"""
            <tr>
                <td>{html.escape(row['platform'])}</td>
                <td>{html.escape(row['shop_name'])}</td>
                <td>{html.escape(row['source_template'])}</td>
                <td class="mono">{html.escape(row['order_id'])}</td>
                <td>{html.escape(row['order_time'])}</td>
                <td>{html.escape(row['pay_time'])}</td>
                <td>{html.escape(row['print_time'] or row['last_print_time'])}</td>
                <td>{html.escape(row['ship_time'])}</td>
                <td>{html.escape(row['shipped_status'])}</td>
                <td>{html.escape(row['refund_status'])}</td>
                <td>{html.escape(row['order_status'])}</td>
                <td>{html.escape(row['print_status'])}</td>
                <td>{html.escape(row['receiver_name'])}</td>
                <td class="mono">{html.escape(row['receiver_phone'])}</td>
                <td title="{html.escape(row['full_address'])}">{html.escape(row['full_address'])}</td>
                <td>{html.escape(row['product_category'])}</td>
                <td>{html.escape(row['style_key'])}</td>
                <td title="{html.escape(row['product_name'])}">{html.escape(row['product_name'])}</td>
                <td title="{html.escape(row['spec_name'])}">{html.escape(row['spec_name'])}</td>
                <td>{html.escape(row['variant_tag'])}</td>
                <td>{html.escape(row['size_tag'])}</td>
                <td>{html.escape(row['sku_code'] or row['product_code'] or row['product_id'])}</td>
                <td class="num">{row['quantity']}</td>
                <td class="num">{'' if row['paid_amount'] is None else row['paid_amount']}</td>
                <td class="num">{'' if row['total_price'] is None else row['total_price']}</td>
                <td>{html.escape(row['courier'])}</td>
                <td class="mono">{html.escape(row['waybill'])}</td>
                <td title="{html.escape(row['buyer_message'])}">{html.escape(row['buyer_message'])}</td>
                <td title="{html.escape(row['seller_remark'] or row['remark'])}">{html.escape(row['seller_remark'] or row['remark'])}</td>
                <td title="{html.escape(row['source_file'])}">{html.escape(row['source_file'])}</td>
            </tr>
            """
            for row in detail_rows
        )
        import_html = "".join(
            f"""
            <tr>
                <td>{row['id']}</td>
                <td>{html.escape(row['imported_at'])}</td>
                <td>{html.escape(row['file_name'])}</td>
                <td>{html.escape(row['template_label'])}</td>
                <td>{html.escape(row['platform'])}</td>
                <td>{html.escape(row['shop_label'])}</td>
                <td class="num">{row['row_count']}</td>
                <td class="num">{row['line_count']}</td>
            </tr>
            """
            for row in imports
        )

        return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>多平台订单导入分析 V1</title>
  <style>
    :root {{
      --bg: #f6f7fb;
      --panel: #ffffff;
      --line: #d9dde7;
      --text: #1f2937;
      --sub: #6b7280;
      --blue: #2563eb;
      --green: #15803d;
      --red: #dc2626;
      --amber: #b45309;
    }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; background: var(--bg); color: var(--text); font: 14px/1.45 "Microsoft YaHei", "PingFang SC", sans-serif; }}
    .page {{ max-width: 1680px; margin: 0 auto; padding: 20px; }}
    .band {{ background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 16px; margin-bottom: 16px; }}
    h1, h2, h3 {{ margin: 0 0 12px; font-weight: 700; }}
    .muted {{ color: var(--sub); }}
    .mono {{ font-family: Consolas, "SFMono-Regular", monospace; }}
    .grid {{ display: grid; gap: 12px; }}
    .hero {{ display: grid; grid-template-columns: 1.6fr 1fr; gap: 16px; }}
    .stats {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }}
    .stat {{ background: #f9fafb; border: 1px solid var(--line); border-radius: 8px; padding: 12px; min-height: 92px; }}
    .stat .label {{ color: var(--sub); font-size: 12px; }}
    .stat .value {{ font-size: 26px; font-weight: 700; margin-top: 8px; }}
    form.inline {{ display: inline; }}
    .form-grid {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 10px; }}
    label {{ display: grid; gap: 4px; font-size: 12px; color: var(--sub); }}
    input, select, textarea, button {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      color: var(--text);
      padding: 8px 10px;
      font: inherit;
    }}
    textarea {{ min-height: 110px; resize: vertical; }}
    button {{ cursor: pointer; font-weight: 600; }}
    .btn-primary {{ background: var(--blue); color: #fff; border-color: var(--blue); }}
    .btn-danger {{ background: #fff; color: var(--red); border-color: #f3b3b3; }}
    .actions {{ display: flex; gap: 10px; flex-wrap: wrap; }}
    .link-btn {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 38px;
      padding: 8px 12px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      color: var(--text);
      text-decoration: none;
      font-weight: 600;
    }}
    .link-btn.primary {{ background: var(--blue); color: #fff; border-color: var(--blue); }}
    .table-wrap {{ overflow: auto; border: 1px solid var(--line); border-radius: 8px; }}
    table {{ width: 100%; border-collapse: collapse; min-width: 1080px; }}
    th, td {{ border-bottom: 1px solid var(--line); padding: 8px 10px; text-align: left; vertical-align: top; }}
    th {{ position: sticky; top: 0; background: #f8fafc; z-index: 1; white-space: nowrap; }}
    td {{ max-width: 280px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .pill {{ display: inline-block; padding: 2px 8px; border-radius: 999px; background: #eef2ff; color: #4338ca; font-size: 12px; }}
    .message {{ padding: 10px 12px; border-radius: 8px; background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe; margin-bottom: 16px; }}
    .small {{ font-size: 12px; }}
    .template-list li {{ margin: 6px 0; }}
    @media (max-width: 1100px) {{
      .hero {{ grid-template-columns: 1fr; }}
      .stats {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
      .form-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
  </style>
</head>
<body>
  <div class="page">
    {f'<div class="message">{html.escape(filters["message"])}</div>' if filters.get("message") else ''}

    <section class="band hero">
      <div>
        <h1>多平台订单导入分析 V1</h1>
        <p class="muted">先把不同平台导出的表扔进来，系统按模板自动识别、拆单、汇总。当前内置了你给我的 3 套拼多多模板，后面继续加淘宝、抖店、京东时，只要补模板解析器就行。</p>
        <div class="actions small">
          <span class="pill">支持 .xlsx / .csv</span>
          <span class="pill">区分下单 / 付款 / 打印 / 发货时间</span>
          <span class="pill">区分发货状态 / 退款状态</span>
          <span class="pill">按商品种类 / 款式 / 店铺筛选</span>
        </div>
      </div>
      <div class="grid">
        <div class="small muted">服务地址: <span class="mono">http://{html.escape(self.args.host)}:{self.args.port}</span></div>
        <div class="small muted">数据库: <span class="mono">{html.escape(str(self.db_path))}</span></div>
        <div class="small muted">导入目录: <span class="mono">{html.escape(str(self.upload_dir))}</span></div>
        <div class="small muted">摘要文件: <span class="mono">{html.escape(str(self.summary_path))}</span></div>
      </div>
    </section>

    <section class="band">
      <h2>导入数据</h2>
      <div class="grid" style="grid-template-columns: 1fr 1fr; gap: 16px;">
        <form method="post" action="/import-files" enctype="multipart/form-data" class="grid">
          <div class="form-grid">
            <label>平台标记
              <input type="text" name="platform_label" placeholder="留空则按模板默认，例如 拼多多">
            </label>
            <label>店铺标记
              <input type="text" name="shop_label" placeholder="可选，覆盖表内店铺">
            </label>
            <label style="grid-column: span 4;">选择文件
              <input type="file" name="files" multiple accept=".xlsx,.csv">
            </label>
          </div>
          <div class="actions">
            <button class="btn-primary" type="submit">上传并导入</button>
          </div>
        </form>

        <form method="post" action="/import-paths" class="grid">
          <div class="form-grid">
            <label>平台标记
              <input type="text" name="platform_label" placeholder="可选">
            </label>
            <label>店铺标记
              <input type="text" name="shop_label" placeholder="可选">
            </label>
            <label style="grid-column: span 4;">本地文件路径（每行一个）
              <textarea name="paths" placeholder="例如：C:\\Users\\18461\\Desktop\\...\\订单导出列表.xlsx"></textarea>
            </label>
          </div>
          <div class="actions">
            <button class="btn-primary" type="submit">按路径导入</button>
            <button class="btn-danger" type="submit" formaction="/reset" formmethod="post">清空全部数据</button>
          </div>
        </form>
      </div>
      <ul class="template-list small muted">
        <li><strong>拼多多-发货单导出</strong>：订单号 / 店铺名称 / 发货时间 / 商品名称 / 规格名称 / 数量 / 实收</li>
        <li><strong>拼多多-打印单</strong>：订单编号 / 收件人 / 收件人电话 / 商品信息 / 订单状态</li>
        <li><strong>拼多多-订单导出</strong>：订单编号 / 下单时间 / 付款时间 / 快递单打印时间 / 发货时间 / 商品简称 / 销售规格 / 货号</li>
      </ul>
    </section>

    <section class="band">
      <h2>筛选与导出</h2>
      <form method="get" action="/" class="grid">
        <div class="form-grid">
          <label>时间字段
            <select name="time_field">
              <option value="ship_time"{' selected' if filters['time_field'] == 'ship_time' else ''}>发货时间</option>
              <option value="order_time"{' selected' if filters['time_field'] == 'order_time' else ''}>下单时间</option>
              <option value="pay_time"{' selected' if filters['time_field'] == 'pay_time' else ''}>付款时间</option>
              <option value="print_time"{' selected' if filters['time_field'] == 'print_time' else ''}>打印时间</option>
              <option value="last_print_time"{' selected' if filters['time_field'] == 'last_print_time' else ''}>最后打印时间</option>
              <option value="imported_at"{' selected' if filters['time_field'] == 'imported_at' else ''}>导入时间</option>
            </select>
          </label>
          <label>开始日期
            <input type="date" name="date_from" value="{html.escape(filters['date_from'])}">
          </label>
          <label>结束日期
            <input type="date" name="date_to" value="{html.escape(filters['date_to'])}">
          </label>
          <label>平台
            <select name="platform">{option_html(options['platforms'], filters['platform'])}</select>
          </label>
          <label>店铺
            <select name="shop_name">{option_html(options['shops'], filters['shop_name'])}</select>
          </label>
          <label>模板
            <select name="template">{option_html(options['templates'], filters['template'])}</select>
          </label>
          <label>商品种类
            <select name="product_category">{option_html(options['categories'], filters['product_category'])}</select>
          </label>
          <label>发货状态
            <select name="shipped_status">
              <option value=""{' selected' if not filters['shipped_status'] else ''}>全部</option>
              <option value="已发货"{' selected' if filters['shipped_status'] == '已发货' else ''}>已发货</option>
              <option value="未发货"{' selected' if filters['shipped_status'] == '未发货' else ''}>未发货</option>
            </select>
          </label>
          <label>退款状态
            <select name="refund_mode">
              <option value=""{' selected' if not filters['refund_mode'] else ''}>全部</option>
              <option value="normal"{' selected' if filters['refund_mode'] == 'normal' else ''}>正常</option>
              <option value="refund"{' selected' if filters['refund_mode'] == 'refund' else ''}>退款/退货/异常</option>
            </select>
          </label>
          <label>打印状态
            <select name="print_status">{option_html(options['print_statuses'], filters['print_status'])}</select>
          </label>
          <label>关键字
            <input type="text" name="keyword" value="{html.escape(filters['keyword'])}" placeholder="订单号 / 商品 / 收件人 / 运单号">
          </label>
          <label>明细条数
            <select name="limit">
              <option value="100"{' selected' if filters['limit'] == '100' else ''}>100</option>
              <option value="200"{' selected' if filters['limit'] == '200' else ''}>200</option>
              <option value="500"{' selected' if filters['limit'] == '500' else ''}>500</option>
              <option value="1000"{' selected' if filters['limit'] == '1000' else ''}>1000</option>
            </select>
          </label>
        </div>
        <div class="actions">
          <button class="btn-primary" type="submit">应用筛选</button>
          <a class="link-btn" href="/">清空筛选</a>
          <a class="link-btn" href="/export/details.csv?{html.escape(query_string)}">导出明细 CSV</a>
          <a class="link-btn" href="/export/summary.csv?{html.escape(query_string)}">导出汇总 CSV</a>
          <a class="link-btn" href="/api/summary?{html.escape(query_string)}">查看 JSON 摘要</a>
        </div>
      </form>
    </section>

    <section class="band">
      <h2>总览</h2>
      <div class="stats">
        <div class="stat"><div class="label">已导入文件</div><div class="value">{stats['import_count']}</div></div>
        <div class="stat"><div class="label">明细行数</div><div class="value">{stats['line_count']}</div></div>
        <div class="stat"><div class="label">订单数</div><div class="value">{stats['order_count']}</div></div>
        <div class="stat"><div class="label">销量件数</div><div class="value">{stats['quantity_sum']}</div></div>
        <div class="stat"><div class="label">已发货件数</div><div class="value">{stats['shipped_quantity']}</div></div>
        <div class="stat"><div class="label">退款/异常件数</div><div class="value">{stats['refund_quantity']}</div></div>
        <div class="stat"><div class="label">已打印件数</div><div class="value">{stats['printed_quantity']}</div></div>
        <div class="stat"><div class="label">实收金额</div><div class="value">{stats['paid_sum']}</div></div>
      </div>
    </section>

    <section class="band">
      <h2>按款式汇总</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>平台</th>
              <th>店铺</th>
              <th>商品种类</th>
              <th>款式键</th>
              <th>示例商品</th>
              <th>示例规格</th>
              <th>销量件数</th>
              <th>订单数</th>
              <th>金额</th>
              <th>退款件数</th>
            </tr>
          </thead>
          <tbody>{summary_html or '<tr><td colspan="10" class="muted">暂无数据</td></tr>'}</tbody>
        </table>
      </div>
    </section>

    <section class="band">
      <h2>订单明细</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>平台</th>
              <th>店铺</th>
              <th>模板</th>
              <th>订单号</th>
              <th>下单</th>
              <th>付款</th>
              <th>打印</th>
              <th>发货</th>
              <th>发货状态</th>
              <th>退款状态</th>
              <th>订单状态</th>
              <th>打印状态</th>
              <th>收件人</th>
              <th>手机</th>
              <th>地址</th>
              <th>商品种类</th>
              <th>款式键</th>
              <th>商品名称</th>
              <th>规格</th>
              <th>变体</th>
              <th>尺码</th>
              <th>货号/编码</th>
              <th>数量</th>
              <th>实收</th>
              <th>总价</th>
              <th>快递</th>
              <th>运单号</th>
              <th>买家留言</th>
              <th>备注</th>
              <th>来源文件</th>
            </tr>
          </thead>
          <tbody>{detail_html or '<tr><td colspan="30" class="muted">暂无数据</td></tr>'}</tbody>
        </table>
      </div>
    </section>

    <section class="band">
      <h2>导入记录</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>导入时间</th>
              <th>文件名</th>
              <th>模板</th>
              <th>平台</th>
              <th>店铺标记</th>
              <th>原始行数</th>
              <th>拆分后明细行</th>
            </tr>
          </thead>
          <tbody>{import_html or '<tr><td colspan="8" class="muted">暂无导入记录</td></tr>'}</tbody>
        </table>
      </div>
    </section>
  </div>
</body>
</html>"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Multi-template order analysis dashboard")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8791)
    parser.add_argument("--state-dir", default=str(DEFAULT_STATE_DIR))
    parser.add_argument("--db-path", default=str(DEFAULT_DB_PATH))
    parser.add_argument("--upload-dir", default=str(DEFAULT_UPLOAD_DIR))
    parser.add_argument("--summary-path", default=str(DEFAULT_SUMMARY_PATH))
    parser.add_argument("--import-path", action="append", default=[], help="Preload a local file path")
    parser.add_argument("--platform-label", default="", help="Optional platform label for preloaded imports")
    parser.add_argument("--shop-label", default="", help="Optional shop label for preloaded imports")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    service = OrderAnalysisService(args)
    try:
        for path_text in args.import_path:
            try:
                result = service.import_path(path_text, args.platform_label, args.shop_label)
                print(f"[import] {path_text} -> {result['status']} ({result.get('template_label', '-')}, {result.get('line_count', 0)} lines)")
            except Exception as exc:  # noqa: BLE001
                print(f"[import-failed] {path_text}: {exc}")
        service.write_summary()
        server = service.build_http_server()
        print(f"[watching] http://{args.host}:{args.port}")
        print(f"[db] {service.db_path}")
        print(f"[imports] {service.upload_dir}")
        print(f"[summary] {service.summary_path}")
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            print("[stopped]")
        finally:
            server.server_close()
    finally:
        service.close()


if __name__ == "__main__":
    main()
