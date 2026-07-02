"""
XTX 竞品规格书解析与对标推荐工具
=================================
运行方式：
    pip install streamlit pandas requests beautifulsoup4 pymupdf rapidfuzz openpyxl
    streamlit run xtx_competitor_mapping_tool.py

维护文件：
    推荐上传 xtx_competitor_maintenance_template.xlsx，至少包含以下工作表：
    1) Company_Master
    2) XTX_Product_Library
    3) Match_Weights
    4) History_Log，可选

公司区分方式：
    Company_Master.company_role = XTX         表示我司：芯天下/XTX
    Company_Master.company_role = Competitor  表示友商/竞品公司

注意：
    1) 官网自动查找 PDF 受竞品官网结构影响，建议第一版优先使用“手动上传 PDF”或“输入 PDF 链接”。
    2) 封装字段经常会列出多个封装，最终建议人工确认具体后缀和 Pin to Pin 兼容性。
    3) 扫描版 PDF 暂不支持 OCR。
"""

from __future__ import annotations

import io
import os
import re
from datetime import datetime
from urllib.parse import urljoin, urlparse

import fitz  # PyMuPDF
import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from rapidfuzz import fuzz


# =========================================================
# Streamlit 页面配置
# =========================================================

st.set_page_config(
    page_title="竞品规格书解析与 XTX 对标推荐工具",
    page_icon="🔎",
    layout="wide",
)

REQUEST_TIMEOUT = 15
MAX_CRAWL_PAGES = 25
MAX_PDF_CANDIDATES = 20
LOCAL_HISTORY_PATH = "competitor_analysis_history.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}


# =========================================================
# 默认数据：没有上传 XLSX 时也能运行
# =========================================================

DEFAULT_COMPANY_MASTER = pd.DataFrame(
    [
        ["芯天下", "XTX", "xtxtech.com", "https://www.xtxtech.com", "", 1, "我司官网，只展示，不参与竞品 PDF 检索"],
        ["Winbond", "Competitor", "winbond.com", "https://www.winbond.com", "", 1, "友商官网"],
        ["Macronix", "Competitor", "macronix.com", "https://www.macronix.com", "", 1, "友商官网"],
        ["GigaDevice", "Competitor", "gigadevice.com", "https://www.gigadevice.com", "", 1, "友商官网"],
        ["ISSI", "Competitor", "issi.com", "https://www.issi.com", "", 1, "友商官网"],
        ["Puya", "Competitor", "puyasemi.com", "https://www.puyasemi.com", "", 1, "友商官网"],
        ["Boya", "Competitor", "boyamicro.com", "https://www.boyamicro.com", "", 1, "友商官网"],
        ["Zbit", "Competitor", "zbitsemi.com", "https://www.zbitsemi.com", "", 1, "友商官网"],
    ],
    columns=["company_name", "company_role", "domain", "base_url", "search_url_template", "enabled", "note"],
)

DEFAULT_XTX_PRODUCT_LIBRARY = pd.DataFrame(
    [
        ["XT25Q256FWSIGT", "SPI NOR", 256, "1.8V", 1.65, 1.95, "WSON8", "8x6mm", -40, 85, 133, "MP", "256Mb 1.8V SPI NOR"],
        ["XT25Q512FWSIGA", "SPI NOR", 512, "1.8V", 1.65, 1.95, "WSON8", "8x6mm", -40, 85, 133, "MP", "512Mb 1.8V SPI NOR"],
        ["XT25F256BWSIGT", "SPI NOR", 256, "3.3V", 2.70, 3.60, "WSON8", "8x6mm", -40, 85, 133, "MP", "256Mb 3.3V SPI NOR"],
        ["XT25F512BWSIGA", "SPI NOR", 512, "3.3V", 2.70, 3.60, "WSON8", "8x6mm", -40, 85, 133, "MP", "512Mb 3.3V SPI NOR"],
        ["XT26G01F", "SPI NAND", 1024, "3.3V", 2.70, 3.60, "WSON8", "8x6mm", -40, 85, 104, "MP", "1Gb 3.3V SPI NAND"],
        ["XT26G02E", "SPI NAND", 2048, "3.3V", 2.70, 3.60, "WSON8", "8x6mm", -40, 85, 104, "MP", "2Gb 3.3V SPI NAND"],
    ],
    columns=[
        "xtx_model", "product_type", "density_mb", "voltage_type", "vcc_min", "vcc_max",
        "package", "package_size", "temp_min", "temp_max", "frequency_mhz", "status", "note"
    ],
)

DEFAULT_MATCH_WEIGHTS = pd.DataFrame(
    [
        ["product_type", 10, 1, "产品类型一致性，如 SPI NOR / SPI NAND"],
        ["density", 35, 1, "容量一致性"],
        ["voltage", 30, 1, "电压范围一致或覆盖"],
        ["package", 15, 1, "封装形式和尺寸兼容性"],
        ["temperature", 20, 1, "温度等级覆盖"],
    ],
    columns=["criterion", "weight", "enabled", "description"],
)

DEFAULT_HISTORY_LOG = pd.DataFrame(
    columns=[
        "run_time", "competitor_company", "competitor_model", "product_type", "datasheet_source",
        "capacity", "density_mb", "voltage_range", "voltage_type", "package", "package_size",
        "temperature", "temp_grade", "capacity_confirm", "voltage_confirm", "package_confirm",
        "temperature_confirm", "recommended_xtx_model", "match_score", "match_percent",
        "match_level", "risk_points", "operator_note",
    ]
)


# =========================================================
# 数据加载与导出
# =========================================================

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_bool_series(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower().isin(["1", "true", "yes", "y", "启用", "是"])


@st.cache_data(show_spinner=False)
def read_maintenance_xlsx(file_bytes: bytes | None) -> dict[str, pd.DataFrame]:
    """读取维护 XLSX，缺失的 sheet 用默认数据补齐。"""
    if not file_bytes:
        return {
            "Company_Master": DEFAULT_COMPANY_MASTER.copy(),
            "XTX_Product_Library": DEFAULT_XTX_PRODUCT_LIBRARY.copy(),
            "Match_Weights": DEFAULT_MATCH_WEIGHTS.copy(),
            "History_Log": DEFAULT_HISTORY_LOG.copy(),
        }

    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    output = {}
    output["Company_Master"] = clean_columns(sheets.get("Company_Master", DEFAULT_COMPANY_MASTER.copy()))
    output["XTX_Product_Library"] = clean_columns(sheets.get("XTX_Product_Library", DEFAULT_XTX_PRODUCT_LIBRARY.copy()))
    output["Match_Weights"] = clean_columns(sheets.get("Match_Weights", DEFAULT_MATCH_WEIGHTS.copy()))
    output["History_Log"] = clean_columns(sheets.get("History_Log", DEFAULT_HISTORY_LOG.copy()))
    return output


def read_history_xlsx(history_upload) -> pd.DataFrame:
    """优先读取上传历史；否则如果本地存在历史文件，则读取本地；否则为空。"""
    if history_upload is not None:
        try:
            sheets = pd.read_excel(history_upload, sheet_name=None)
            return clean_columns(sheets.get("History_Log", next(iter(sheets.values()))))
        except Exception:
            return DEFAULT_HISTORY_LOG.copy()

    if os.path.exists(LOCAL_HISTORY_PATH):
        try:
            sheets = pd.read_excel(LOCAL_HISTORY_PATH, sheet_name=None)
            return clean_columns(sheets.get("History_Log", next(iter(sheets.values()))))
        except Exception:
            return DEFAULT_HISTORY_LOG.copy()

    return DEFAULT_HISTORY_LOG.copy()


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


def save_history_local(history_df: pd.DataFrame) -> None:
    with pd.ExcelWriter(LOCAL_HISTORY_PATH, engine="openpyxl") as writer:
        history_df.to_excel(writer, index=False, sheet_name="History_Log")


# =========================================================
# 通用工具函数
# =========================================================

def normalize_model(model: str) -> str:
    if not model:
        return ""
    return re.sub(r"[\s\-_./]", "", str(model).upper())


def safe_float(value):
    try:
        if pd.isna(value):
            return None
        return float(value)
    except Exception:
        return None


def safe_int(value):
    try:
        if pd.isna(value):
            return None
        return int(float(value))
    except Exception:
        return None


def valid_url(url: str) -> bool:
    try:
        parsed = urlparse(str(url))
        return parsed.scheme in ["http", "https"] and bool(parsed.netloc)
    except Exception:
        return False


def get_domain(url: str) -> str:
    try:
        return urlparse(str(url)).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def same_domain_or_subdomain(url: str, allowed_domain: str) -> bool:
    current = get_domain(url)
    allowed = str(allowed_domain).lower().replace("www.", "").strip()
    return current == allowed or current.endswith("." + allowed)


def request_url(url: str):
    return requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)


# =========================================================
# 官网 PDF 查找
# =========================================================

def score_pdf_candidate(url: str, link_text: str, model: str) -> int:
    model_norm = normalize_model(model)
    url_norm = normalize_model(url)
    text_norm = normalize_model(link_text)
    combined = f"{url_norm} {text_norm}"

    score = 0
    if model_norm and model_norm in url_norm:
        score += 90
    if model_norm and model_norm in text_norm:
        score += 90
    if "DATASHEET" in combined or "DATASHEET" in combined.replace(" ", "") or "DATA SHEET" in combined:
        score += 40
    if "SPECIFICATION" in combined or "SPEC" in combined:
        score += 25
    if "PDF" in combined:
        score += 10
    if model_norm:
        score += int(fuzz.partial_ratio(model_norm, combined) * 0.4)

    bad_words = [
        "ANNUAL", "REPORT", "ESG", "FINANCIAL", "PRESENTATION", "BROCHURE",
        "CATALOG", "QUALITY", "RELIABILITYREPORT", "PACKAGEINFOONLY",
    ]
    if any(word in combined for word in bad_words):
        score -= 40

    return score


def extract_links_from_html(html: str, current_url: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = str(a.get("href", "")).strip()
        text = a.get_text(" ", strip=True)
        if href:
            links.append({"url": urljoin(current_url, href), "text": text})
    return links


def crawl_company_for_pdf(company_row: pd.Series, model: str) -> pd.DataFrame:
    company = str(company_row.get("company_name", "")).strip()
    domain = str(company_row.get("domain", "")).strip()
    base_url = str(company_row.get("base_url", "")).strip()
    search_template = str(company_row.get("search_url_template", "")).strip()

    if not domain or not base_url:
        return pd.DataFrame()

    seed_urls = []
    if search_template and search_template.lower() not in ["nan", "none", ""]:
        seed_urls.append(search_template.replace("{model}", model))
    seed_urls.append(base_url)

    visited = set()
    queue = list(dict.fromkeys(seed_urls))
    candidates = []
    model_norm = normalize_model(model)

    while queue and len(visited) < MAX_CRAWL_PAGES:
        current_url = queue.pop(0)
        if current_url in visited:
            continue
        visited.add(current_url)

        try:
            resp = request_url(current_url)
            resp.raise_for_status()
        except Exception:
            continue

        content_type = resp.headers.get("Content-Type", "").lower()
        if "pdf" in content_type or current_url.lower().endswith(".pdf"):
            candidates.append({
                "company_name": company,
                "url": current_url,
                "link_text": current_url,
                "score": score_pdf_candidate(current_url, current_url, model),
            })
            continue

        if "html" not in content_type and "text" not in content_type:
            continue

        links = extract_links_from_html(resp.text, current_url)
        for item in links:
            link_url = item["url"]
            link_text = item["text"]
            if not same_domain_or_subdomain(link_url, domain):
                continue

            link_url_lower = link_url.lower()
            link_url_norm = normalize_model(link_url)
            link_text_norm = normalize_model(link_text)
            is_pdf = ".pdf" in link_url_lower or "pdf" in link_text.lower()

            if is_pdf:
                candidates.append({
                    "company_name": company,
                    "url": link_url,
                    "link_text": link_text,
                    "score": score_pdf_candidate(link_url, link_text, model),
                })
                continue

            relevant_keywords = [
                "product", "products", "memory", "flash", "nor", "nand", "serial", "spi",
                "datasheet", "download", "document", "spec",
            ]
            looks_relevant = (
                model_norm in link_url_norm
                or model_norm in link_text_norm
                or any(k in link_url_lower for k in relevant_keywords)
            )
            if looks_relevant and link_url not in visited and len(queue) < MAX_CRAWL_PAGES:
                queue.append(link_url)

    if not candidates:
        return pd.DataFrame()
    df = pd.DataFrame(candidates).sort_values("score", ascending=False).drop_duplicates(subset=["url"])
    return df.head(MAX_PDF_CANDIDATES)


def download_pdf(pdf_url: str) -> bytes:
    resp = request_url(pdf_url)
    resp.raise_for_status()
    return resp.content


# =========================================================
# PDF 文本解析
# =========================================================

def extract_text_from_pdf_bytes(pdf_bytes: bytes, max_pages: int = 30) -> str:
    text_parts = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        total_pages = min(len(doc), max_pages)
        for idx in range(total_pages):
            page = doc[idx]
            text = page.get_text("text", sort=True)
            if text:
                text_parts.append(f"\n\n--- Page {idx + 1} ---\n{text}")
    return "\n".join(text_parts)


# =========================================================
# 字段抽取：容量、电压、封装、温度
# =========================================================

def extract_density(text: str) -> dict:
    t = text.upper()
    candidates = []
    patterns = [
        r"(\d+)\s*G\s*[- ]?\s*BIT",
        r"(\d+)\s*GBIT",
        r"(\d+)\s*M\s*[- ]?\s*BIT",
        r"(\d+)\s*MBIT",
        r"(\d+)\s*M\s*BYTE",
        r"(\d+)\s*MBYTE",
    ]
    for pattern in patterns:
        for m in re.finditer(pattern, t):
            value = int(m.group(1))
            raw = m.group(0).upper()
            if "GBIT" in raw or ("G" in raw and "BIT" in raw):
                density_mb = value * 1024
                display = f"{value}Gb"
            elif "MBIT" in raw or ("M" in raw and "BIT" in raw):
                density_mb = value
                display = f"{value}Mb"
            elif "MBYTE" in raw or "M BYTE" in raw:
                density_mb = value * 8
                display = f"{density_mb}Mb"
            else:
                continue
            if 1 <= density_mb <= 131072:
                candidates.append({"density_mb": density_mb, "display": display, "raw": raw})

    if not candidates:
        return {"density_mb": None, "display": "未识别", "confidence": "低"}

    df = pd.DataFrame(candidates)
    grouped = (
        df.groupby(["density_mb", "display"])
        .size()
        .reset_index(name="count")
        .sort_values(["count", "density_mb"], ascending=[False, False])
    )
    best = grouped.iloc[0]
    return {
        "density_mb": int(best["density_mb"]),
        "display": str(best["display"]),
        "confidence": "高" if int(best["count"]) >= 2 else "中",
    }


def parse_density_from_text(value: str):
    if not value:
        return None, "未识别"
    s = str(value).upper().replace(" ", "")
    m = re.search(r"(\d+(?:\.\d+)?)(GB|G|GBIT|G-BIT)", s)
    if m:
        density = int(float(m.group(1)) * 1024)
        return density, f"{int(float(m.group(1)))}Gb"
    m = re.search(r"(\d+(?:\.\d+)?)(MB|M|MBIT|M-BIT)", s)
    if m:
        density = int(float(m.group(1)))
        return density, f"{density}Mb"
    m = re.search(r"(\d+)", s)
    if m:
        density = int(m.group(1))
        return density, f"{density}Mb"
    return None, "未识别"


def extract_voltage(text: str) -> dict:
    t = text.replace("–", "-").replace("—", "-").replace("～", "-")
    t = re.sub(r"\s+", " ", t)
    patterns = [
        r"VCC\s*[:=]?\s*([0-9]\.[0-9]{1,2})\s*V?\s*(?:TO|-|~)\s*([0-9]\.[0-9]{1,2})\s*V",
        r"OPERATING VOLTAGE\s*[:=]?\s*([0-9]\.[0-9]{1,2})\s*V?\s*(?:TO|-|~)\s*([0-9]\.[0-9]{1,2})\s*V",
        r"([0-9]\.[0-9]{1,2})\s*V?\s*(?:TO|-|~)\s*([0-9]\.[0-9]{1,2})\s*V",
    ]
    candidates = []
    for pattern in patterns:
        for m in re.finditer(pattern, t, flags=re.IGNORECASE):
            v1 = safe_float(m.group(1))
            v2 = safe_float(m.group(2))
            if v1 is None or v2 is None:
                continue
            vmin = min(v1, v2)
            vmax = max(v1, v2)
            if 1.0 <= vmin <= 5.5 and 1.0 <= vmax <= 5.5:
                candidates.append({"vcc_min": vmin, "vcc_max": vmax, "raw": m.group(0)})

    if not candidates:
        return {"vcc_min": None, "vcc_max": None, "display": "未识别", "voltage_type": "未识别", "confidence": "低"}

    def priority(item):
        vmin = item["vcc_min"]
        vmax = item["vcc_max"]
        if 1.60 <= vmin <= 1.75 and 1.85 <= vmax <= 2.05:
            return 100
        if 2.60 <= vmin <= 2.85 and 3.45 <= vmax <= 3.70:
            return 95
        if 1.60 <= vmin <= 1.80 and 3.45 <= vmax <= 3.70:
            return 90
        if 2.20 <= vmin <= 2.50 and 3.45 <= vmax <= 3.70:
            return 80
        return 50

    best = sorted(candidates, key=priority, reverse=True)[0]
    return build_voltage_dict(best["vcc_min"], best["vcc_max"], "中")


def build_voltage_dict(vmin, vmax, confidence="中") -> dict:
    if vmin is None or vmax is None:
        return {"vcc_min": None, "vcc_max": None, "display": "未识别", "voltage_type": "未识别", "confidence": "低"}
    if 1.60 <= vmin <= 1.75 and 1.85 <= vmax <= 2.05:
        voltage_type = "1.8V"
    elif 2.60 <= vmin <= 2.85 and 3.45 <= vmax <= 3.70:
        voltage_type = "3.3V"
    elif 1.60 <= vmin <= 1.80 and 3.45 <= vmax <= 3.70:
        voltage_type = "宽压"
    elif 2.20 <= vmin <= 2.50 and 3.45 <= vmax <= 3.70:
        voltage_type = "宽压/3V"
    else:
        voltage_type = "其他"
    return {"vcc_min": vmin, "vcc_max": vmax, "display": f"{vmin:g}V–{vmax:g}V", "voltage_type": voltage_type, "confidence": confidence}


def parse_voltage_from_text(value: str):
    if not value:
        return None, None
    s = str(value).replace("–", "-").replace("—", "-").replace("~", "-").replace("～", "-")
    nums = re.findall(r"\d+\.\d+|\d+", s)
    nums = [float(x) for x in nums]
    if len(nums) >= 2:
        return min(nums[0], nums[1]), max(nums[0], nums[1])
    return None, None


def extract_package(text: str) -> dict:
    t = text.upper().replace("×", "x").replace("X", "x")
    package_words = [
        "WSON8", "WSON 8", "USON8", "USON 8", "SOP8", "SOP 8", "SOIC8", "SOIC 8",
        "SOP16", "SOP 16", "SOIC16", "SOIC 16", "VSOP8", "VSOP 8", "TSSOP8", "TSSOP 8",
        "DFN8", "DFN 8", "BGA24", "BGA 24", "TFBGA", "FBGA", "WLCSP", "CSP",
    ]
    packages = []
    for word in package_words:
        if word in t:
            packages.append(word.replace(" ", ""))
    packages = list(dict.fromkeys(packages))

    size_patterns = [
        r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)\s*MM",
        r"(\d+(?:\.\d+)?)\s*MM\s*x\s*(\d+(?:\.\d+)?)\s*MM",
    ]
    sizes = []
    for pattern in size_patterns:
        for m in re.finditer(pattern, t):
            sizes.append(f"{m.group(1)}x{m.group(2)}mm")
    sizes = list(dict.fromkeys(sizes))

    if not packages and not sizes:
        return {"package": "未识别", "package_size": "未识别", "confidence": "低"}
    return {
        "package": " / ".join(packages[:10]) if packages else "未识别",
        "package_size": " / ".join(sizes[:10]) if sizes else "未识别",
        "confidence": "中",
    }


def extract_temperature(text: str) -> dict:
    t = text.replace("℃", "°C").replace("–", "-").replace("—", "-").replace("～", "-")
    t = re.sub(r"\s+", " ", t)
    patterns = [
        r"(-?\d{1,3})\s*°?\s*C\s*(?:TO|-|~)\s*(-?\d{1,3})\s*°?\s*C",
        r"(-?\d{1,3})\s*(?:TO|-|~)\s*(-?\d{1,3})\s*°?\s*C",
    ]
    candidates = []
    for pattern in patterns:
        for m in re.finditer(pattern, t, flags=re.IGNORECASE):
            t1 = int(m.group(1))
            t2 = int(m.group(2))
            temp_min = min(t1, t2)
            temp_max = max(t1, t2)
            if -65 <= temp_min <= 25 and 70 <= temp_max <= 150:
                candidates.append({"temp_min": temp_min, "temp_max": temp_max, "raw": m.group(0)})

    if not candidates:
        return {"temp_min": None, "temp_max": None, "display": "未识别", "temp_grade": "未识别", "confidence": "低"}

    df = pd.DataFrame(candidates)
    grouped = (
        df.groupby(["temp_min", "temp_max"])
        .size()
        .reset_index(name="count")
        .sort_values(["count", "temp_max"], ascending=[False, False])
    )
    best = grouped.iloc[0]
    return build_temperature_dict(int(best["temp_min"]), int(best["temp_max"]), "高" if int(best["count"]) >= 2 else "中")


def build_temperature_dict(temp_min, temp_max, confidence="中") -> dict:
    if temp_min is None or temp_max is None:
        return {"temp_min": None, "temp_max": None, "display": "未识别", "temp_grade": "未识别", "confidence": "低"}
    if temp_min >= 0 and temp_max <= 70:
        grade = "Commercial"
    elif temp_min <= -40 and temp_max <= 85:
        grade = "Industrial"
    elif temp_min <= -40 and 100 <= temp_max < 125:
        grade = "Industrial Plus / Automotive Grade 2"
    elif temp_min <= -40 and temp_max >= 125:
        grade = "Automotive Grade 1 / Extended"
    elif temp_min <= -55 and temp_max >= 125:
        grade = "Military / Extended"
    else:
        grade = "其他"
    return {"temp_min": temp_min, "temp_max": temp_max, "display": f"{temp_min}°C~{temp_max}°C", "temp_grade": grade, "confidence": confidence}


def parse_temperature_from_text(value: str):
    if not value:
        return None, None
    s = str(value).replace("–", "-").replace("—", "-").replace("~", "-").replace("～", "-")
    nums = re.findall(r"-?\d+", s)
    nums = [int(x) for x in nums]
    if len(nums) >= 2:
        return min(nums[0], nums[1]), max(nums[0], nums[1])
    return None, None


def analyze_spec_text(text: str) -> dict:
    density = extract_density(text)
    voltage = extract_voltage(text)
    package = extract_package(text)
    temp = extract_temperature(text)
    return {
        "capacity": density["display"],
        "density_mb": density["density_mb"],
        "capacity_confidence": density["confidence"],
        "voltage_range": voltage["display"],
        "voltage_type": voltage["voltage_type"],
        "vcc_min": voltage["vcc_min"],
        "vcc_max": voltage["vcc_max"],
        "voltage_confidence": voltage["confidence"],
        "package": package["package"],
        "package_size": package["package_size"],
        "package_confidence": package["confidence"],
        "temperature": temp["display"],
        "temp_grade": temp["temp_grade"],
        "temp_min": temp["temp_min"],
        "temp_max": temp["temp_max"],
        "temperature_confidence": temp["confidence"],
    }


# =========================================================
# 人工确认编辑
# =========================================================

def spec_to_review_df(spec: dict) -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["capacity", "容量", spec.get("capacity", ""), "", spec.get("capacity_confidence", ""), "未确认"],
            ["voltage", "电压范围", spec.get("voltage_range", ""), "", spec.get("voltage_confidence", ""), "未确认"],
            ["package", "封装形式", spec.get("package", ""), "", spec.get("package_confidence", ""), "未确认"],
            ["package_size", "封装尺寸", spec.get("package_size", ""), "", spec.get("package_confidence", ""), "未确认"],
            ["temperature", "温度范围", spec.get("temperature", ""), "", spec.get("temperature_confidence", ""), "未确认"],
        ],
        columns=["field_key", "field_name", "extracted_value", "manual_value", "confidence", "confirm_status"],
    )


def build_confirmed_spec(original_spec: dict, review_df: pd.DataFrame) -> tuple[dict, dict]:
    spec = original_spec.copy()
    confirm_map = {}
    for _, row in review_df.iterrows():
        key = str(row.get("field_key", ""))
        manual = str(row.get("manual_value", "")).strip()
        status = str(row.get("confirm_status", "未确认")).strip()
        confirm_map[key] = status

        if not manual:
            continue

        if key == "capacity":
            density_mb, display = parse_density_from_text(manual)
            spec["density_mb"] = density_mb
            spec["capacity"] = display
        elif key == "voltage":
            vmin, vmax = parse_voltage_from_text(manual)
            vd = build_voltage_dict(vmin, vmax, "人工确认")
            spec.update({
                "vcc_min": vd["vcc_min"], "vcc_max": vd["vcc_max"], "voltage_range": vd["display"],
                "voltage_type": vd["voltage_type"], "voltage_confidence": vd["confidence"],
            })
        elif key == "package":
            spec["package"] = manual
            spec["package_confidence"] = "人工确认"
        elif key == "package_size":
            spec["package_size"] = manual
            spec["package_confidence"] = "人工确认"
        elif key == "temperature":
            tmin, tmax = parse_temperature_from_text(manual)
            td = build_temperature_dict(tmin, tmax, "人工确认")
            spec.update({
                "temp_min": td["temp_min"], "temp_max": td["temp_max"], "temperature": td["display"],
                "temp_grade": td["temp_grade"], "temperature_confidence": td["confidence"],
            })
    return spec, confirm_map


# =========================================================
# 匹配权重与 XTX 推荐逻辑
# =========================================================

def get_enabled_weights(weights_df: pd.DataFrame) -> dict[str, float]:
    if weights_df.empty:
        weights_df = DEFAULT_MATCH_WEIGHTS.copy()
    df = clean_columns(weights_df)
    if "enabled" in df.columns:
        df = df[normalize_bool_series(df["enabled"])]
    weights = {}
    for _, row in df.iterrows():
        criterion = str(row.get("criterion", "")).strip()
        weight = safe_float(row.get("weight"))
        if criterion and weight is not None:
            weights[criterion] = weight
    return weights or {"product_type": 10, "density": 35, "voltage": 30, "package": 15, "temperature": 20}


def calc_match_score(spec: dict, xtx_row: pd.Series, selected_product_type: str, weights: dict[str, float]) -> tuple[float, float, str, list[str]]:
    score = 0.0
    max_score = sum(weights.values()) if weights else 100.0
    risks = []

    competitor_density = spec.get("density_mb")
    xtx_density = safe_int(xtx_row.get("density_mb"))

    competitor_vmin = spec.get("vcc_min")
    competitor_vmax = spec.get("vcc_max")
    xtx_vmin = safe_float(xtx_row.get("vcc_min"))
    xtx_vmax = safe_float(xtx_row.get("vcc_max"))

    competitor_pkg = str(spec.get("package", "")).upper()
    competitor_pkg_size = str(spec.get("package_size", "")).upper()
    xtx_pkg = str(xtx_row.get("package", "")).upper()
    xtx_pkg_size = str(xtx_row.get("package_size", "")).upper()

    competitor_tmin = spec.get("temp_min")
    competitor_tmax = spec.get("temp_max")
    xtx_tmin = safe_float(xtx_row.get("temp_min"))
    xtx_tmax = safe_float(xtx_row.get("temp_max"))

    xtx_product_type = str(xtx_row.get("product_type", ""))

    # 产品类型
    if weights.get("product_type", 0) > 0:
        if selected_product_type and selected_product_type != "未指定":
            if selected_product_type.upper() == xtx_product_type.upper():
                score += weights["product_type"]
            else:
                risks.append("产品类型不一致")
        else:
            max_score -= weights["product_type"]

    # 容量
    if weights.get("density", 0) > 0:
        if competitor_density and xtx_density:
            if int(competitor_density) == int(xtx_density):
                score += weights["density"]
            else:
                risks.append("容量不一致")
        else:
            risks.append("容量未识别，需人工确认")

    # 电压
    if weights.get("voltage", 0) > 0:
        if competitor_vmin and competitor_vmax and xtx_vmin and xtx_vmax:
            if abs(competitor_vmin - xtx_vmin) <= 0.15 and abs(competitor_vmax - xtx_vmax) <= 0.15:
                score += weights["voltage"]
            elif xtx_vmin <= competitor_vmin and xtx_vmax >= competitor_vmax:
                score += weights["voltage"] * 0.75
                risks.append("XTX电压范围覆盖竞品，但系统电压仍需确认")
            else:
                risks.append("电压范围不完全匹配")
        else:
            risks.append("电压未识别，需人工确认")

    # 封装
    if weights.get("package", 0) > 0:
        if competitor_pkg and competitor_pkg != "未识别" and xtx_pkg:
            pkg_same = xtx_pkg in competitor_pkg or competitor_pkg in xtx_pkg
            size_same = xtx_pkg_size and competitor_pkg_size and (xtx_pkg_size in competitor_pkg_size or competitor_pkg_size in xtx_pkg_size)
            if pkg_same and size_same:
                score += weights["package"]
            elif pkg_same:
                score += weights["package"] * 0.75
                risks.append("封装名称一致，具体尺寸需确认")
            else:
                same_family = False
                for key in ["WSON", "SOP", "SOIC", "BGA", "WLCSP", "USON", "DFN"]:
                    if key in competitor_pkg and key in xtx_pkg:
                        same_family = True
                        break
                if same_family:
                    score += weights["package"] * 0.5
                    risks.append("封装大类接近，具体尺寸和 Pin 脚需确认")
                else:
                    risks.append("封装可能不兼容")
        else:
            risks.append("封装未识别，需人工确认")

    # 温度
    if weights.get("temperature", 0) > 0:
        if competitor_tmin is not None and competitor_tmax is not None and xtx_tmin is not None and xtx_tmax is not None:
            if xtx_tmin <= competitor_tmin and xtx_tmax >= competitor_tmax:
                score += weights["temperature"]
            elif xtx_tmin <= competitor_tmin and xtx_tmax + 20 >= competitor_tmax:
                score += weights["temperature"] * 0.5
                risks.append("温度范围接近，但高温等级可能不足")
            else:
                risks.append("温度等级不满足")
        else:
            risks.append("温度未识别，需人工确认")

    match_percent = score / max_score if max_score else 0
    if match_percent >= 0.85:
        level = "高"
    elif match_percent >= 0.60:
        level = "中"
    elif match_percent >= 0.40:
        level = "低"
    else:
        level = "不推荐"

    return round(score, 2), round(match_percent * 100, 1), level, risks


def recommend_xtx(spec: dict, xtx_df: pd.DataFrame, selected_product_type: str, weights_df: pd.DataFrame, top_n: int = 8) -> pd.DataFrame:
    if xtx_df.empty:
        return pd.DataFrame()
    weights = get_enabled_weights(weights_df)
    rows = []
    for _, row in xtx_df.iterrows():
        score, pct, level, risks = calc_match_score(spec, row, selected_product_type, weights)
        rows.append({
            "推荐XTX型号": row.get("xtx_model", ""),
            "产品类型": row.get("product_type", ""),
            "容量": f"{row.get('density_mb', '')}Mb",
            "电压类型": row.get("voltage_type", ""),
            "电压范围": f"{row.get('vcc_min', '')}V–{row.get('vcc_max', '')}V",
            "封装": f"{row.get('package', '')} {row.get('package_size', '')}",
            "温度": f"{row.get('temp_min', '')}°C~{row.get('temp_max', '')}°C",
            "频率MHz": row.get("frequency_mhz", ""),
            "状态": row.get("status", ""),
            "匹配分": score,
            "匹配率": f"{pct}%",
            "匹配度": level,
            "风险点": "；".join(risks) if risks else "无明显风险",
            "备注": row.get("note", ""),
        })
    result = pd.DataFrame(rows).sort_values("匹配分", ascending=False)
    return result.head(top_n)


# =========================================================
# 页面 UI
# =========================================================

st.title("🔎 竞品规格书解析与 XTX 对标推荐工具")
st.caption("输入竞品型号，获取/上传规格书，自动解析容量、电压、封装、温度，并推荐芯天下对标型号。")

with st.sidebar:
    st.header("维护数据上传")
    maintenance_file = st.file_uploader(
        "上传维护数据库 XLSX",
        type=["xlsx"],
        help="建议使用模板 xtx_competitor_maintenance_template.xlsx。包含 Company_Master、XTX_Product_Library、Match_Weights、History_Log。",
    )
    history_file = st.file_uploader(
        "上传历史记录 XLSX（可选）",
        type=["xlsx"],
        help="如果不上传，程序会尝试读取本地 competitor_analysis_history.xlsx。",
    )
    max_pages = st.slider("PDF解析页数", min_value=5, max_value=100, value=30, step=5)
    save_local_history = st.checkbox("保存历史到本地 xlsx", value=True, help="适合在公司内网/本地部署时使用。")

    st.divider()
    st.markdown("**公司区分规则**")
    st.caption("Company_Master.company_role = XTX 表示芯天下；Competitor 表示友商。自动检索只使用 Competitor。")

maintenance_bytes = maintenance_file.getvalue() if maintenance_file is not None else None
maintenance = read_maintenance_xlsx(maintenance_bytes)
company_df = clean_columns(maintenance["Company_Master"])
xtx_df = clean_columns(maintenance["XTX_Product_Library"])
weights_df = clean_columns(maintenance["Match_Weights"])
history_df = read_history_xlsx(history_file)

if "enabled" in company_df.columns:
    company_df = company_df[normalize_bool_series(company_df["enabled"])]
if "enabled" in weights_df.columns:
    weights_df = weights_df[normalize_bool_series(weights_df["enabled"])]

competitor_company_df = company_df[company_df["company_role"].astype(str).str.upper().eq("COMPETITOR")].copy()
xtx_company_df = company_df[company_df["company_role"].astype(str).str.upper().eq("XTX")].copy()

# 顶部下载模板
with st.expander("下载/查看维护文件格式", expanded=False):
    st.markdown(
        """
维护数据库 XLSX 建议包含四个 Sheet：

| Sheet | 用途 | 是否必须 |
|---|---|---|
| Company_Master | 公司官网白名单，区分芯天下和友商 | 必须 |
| XTX_Product_Library | 我司可对标产品库 | 必须 |
| Match_Weights | 匹配规则权重配置 | 建议 |
| History_Log | 历史分析记录 | 可选 |

`Company_Master.company_role` 用于区分公司身份：`XTX` = 芯天下，`Competitor` = 友商。自动检索 PDF 只会检索 `Competitor` 行。
"""
    )
    template_bytes = to_excel_bytes(
        {
            "Company_Master": DEFAULT_COMPANY_MASTER,
            "XTX_Product_Library": DEFAULT_XTX_PRODUCT_LIBRARY,
            "Match_Weights": DEFAULT_MATCH_WEIGHTS,
            "History_Log": DEFAULT_HISTORY_LOG,
        }
    )
    st.download_button(
        "下载维护数据库 XLSX 模板",
        data=template_bytes,
        file_name="xtx_competitor_maintenance_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.subheader("1. 输入竞品信息")
col1, col2, col3 = st.columns(3)
with col1:
    competitor_options = ["自动识别/全部友商"] + sorted(competitor_company_df["company_name"].dropna().astype(str).unique().tolist())
    selected_company = st.selectbox("竞品厂商", competitor_options)
with col2:
    competitor_model = st.text_input("竞品型号", placeholder="例如：W25Q512JV / MX25L51245G / GD25Q256E")
with col3:
    product_type = st.selectbox("产品类型", ["未指定", "SPI NOR", "SPI NAND", "Parallel NOR", "eMMC", "DDR"], index=0)

st.subheader("2. 选择规格书来源")
source_mode = st.radio(
    "规格书来源",
    ["自动从友商官网白名单查找", "输入 PDF 链接", "手动上传 PDF"],
    horizontal=True,
)

pdf_url_input = ""
uploaded_pdf = None
if source_mode == "输入 PDF 链接":
    pdf_url_input = st.text_input("PDF 链接", placeholder="https://...pdf")
elif source_mode == "手动上传 PDF":
    uploaded_pdf = st.file_uploader("上传竞品规格书 PDF", type=["pdf"])

run = st.button("开始解析规格书", type="primary")

if run:
    if not competitor_model:
        st.error("请先输入竞品型号。")
        st.stop()

    pdf_bytes = None
    selected_pdf_url = ""
    candidate_df = pd.DataFrame()

    with st.status("正在处理...", expanded=True) as status:
        if source_mode == "手动上传 PDF":
            if uploaded_pdf is None:
                st.error("请上传 PDF。")
                st.stop()
            st.write("读取上传的 PDF...")
            pdf_bytes = uploaded_pdf.read()
            selected_pdf_url = "用户手动上传 PDF"

        elif source_mode == "输入 PDF 链接":
            if not valid_url(pdf_url_input):
                st.error("请输入有效的 PDF 链接。")
                st.stop()
            st.write("下载 PDF...")
            try:
                pdf_bytes = download_pdf(pdf_url_input)
                selected_pdf_url = pdf_url_input
            except Exception as e:
                st.error(f"PDF 下载失败：{e}")
                st.stop()

        else:
            if competitor_company_df.empty:
                st.error("友商官网白名单为空，请检查 Company_Master 中 company_role=Competitor 的记录。")
                st.stop()

            if selected_company != "自动识别/全部友商":
                search_df = competitor_company_df[competitor_company_df["company_name"].astype(str) == selected_company]
            else:
                search_df = competitor_company_df

            st.write("根据友商官网白名单查找 PDF...")
            all_candidates = []
            progress = st.progress(0)
            for idx, (_, row) in enumerate(search_df.iterrows()):
                st.write(f"检索：{row.get('company_name', '')}")
                try:
                    df = crawl_company_for_pdf(row, competitor_model)
                    if not df.empty:
                        all_candidates.append(df)
                except Exception as e:
                    st.write(f"检索失败：{e}")
                progress.progress((idx + 1) / max(len(search_df), 1))

            if not all_candidates:
                st.warning("未自动找到规格书 PDF。建议切换到“输入 PDF 链接”或“手动上传 PDF”。")
                st.stop()

            candidate_df = pd.concat(all_candidates, ignore_index=True)
            candidate_df = candidate_df.sort_values("score", ascending=False).drop_duplicates(subset=["url"]).head(MAX_PDF_CANDIDATES)
            best = candidate_df.iloc[0]
            selected_pdf_url = best["url"]
            st.write("默认选择匹配分最高的 PDF：")
            st.write(selected_pdf_url)

            try:
                pdf_bytes = download_pdf(selected_pdf_url)
            except Exception as e:
                st.error(f"PDF 下载失败：{e}")
                st.stop()

        st.write("解析 PDF 文本...")
        try:
            pdf_text = extract_text_from_pdf_bytes(pdf_bytes, max_pages=max_pages)
        except Exception as e:
            st.error(f"PDF 解析失败：{e}")
            st.stop()

        if not pdf_text.strip():
            st.error("PDF 未提取到有效文本，可能是扫描版 PDF，需要后续增加 OCR。")
            st.stop()

        st.write("抽取容量、电压、封装、温度...")
        spec = analyze_spec_text(pdf_text)

        st.session_state["last_result"] = {
            "spec": spec,
            "candidate_df": candidate_df,
            "selected_pdf_url": selected_pdf_url,
            "competitor_model": competitor_model,
            "selected_company": selected_company,
            "product_type": product_type,
            "review_df": spec_to_review_df(spec),
        }
        status.update(label="规格书解析完成", state="complete", expanded=False)

# 结果展示与人工确认
if "last_result" in st.session_state:
    result = st.session_state["last_result"]
    spec = result["spec"]
    selected_pdf_url = result["selected_pdf_url"]
    candidate_df = result["candidate_df"]
    competitor_model_result = result["competitor_model"]
    selected_company_result = result["selected_company"]
    product_type_result = result["product_type"]

    st.divider()
    st.subheader("3. 自动解析结果")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("容量", spec["capacity"])
        st.caption(f"置信度：{spec['capacity_confidence']}")
    with m2:
        st.metric("电压范围", spec["voltage_range"])
        st.caption(f"{spec['voltage_type']}，置信度：{spec['voltage_confidence']}")
    with m3:
        st.metric("封装形式", spec["package"])
        st.caption(f"尺寸：{spec['package_size']}，置信度：{spec['package_confidence']}")
    with m4:
        st.metric("温度范围", spec["temperature"])
        st.caption(f"{spec['temp_grade']}，置信度：{spec['temperature_confidence']}")

    st.markdown("**规格书来源：**")
    if str(selected_pdf_url).startswith("http"):
        st.markdown(f"[打开规格书 PDF]({selected_pdf_url})")
    else:
        st.write(selected_pdf_url)

    if not candidate_df.empty:
        with st.expander("查看自动检索到的 PDF 候选链接"):
            st.dataframe(candidate_df, use_container_width=True)

    st.divider()
    st.subheader("4. 人工确认 / 修正")
    st.caption("如自动解析不准确，请在 manual_value 中填写人工确认值；例如容量填 512Mb，电压填 2.7V-3.6V，温度填 -40~85。")

    review_df = st.data_editor(
        result["review_df"],
        use_container_width=True,
        hide_index=True,
        column_config={
            "confirm_status": st.column_config.SelectboxColumn(
                "confirm_status",
                options=["未确认", "已确认", "需FAE确认", "不适用"],
                required=True,
            )
        },
        disabled=["field_key", "field_name", "extracted_value", "confidence"],
        key="review_editor",
    )

    confirmed_spec, confirm_map = build_confirmed_spec(spec, review_df)

    st.divider()
    st.subheader("5. XTX 对标型号推荐")
    recommend_df = recommend_xtx(confirmed_spec, xtx_df, product_type_result, weights_df, top_n=10)
    if recommend_df.empty:
        st.warning("XTX 产品库为空或无法完成推荐。")
    else:
        st.dataframe(recommend_df, use_container_width=True)
        best = recommend_df.iloc[0]
        st.success(
            f"优先推荐对标型号：{best['推荐XTX型号']}；"
            f"匹配度：{best['匹配度']}；匹配分：{best['匹配分']}；匹配率：{best['匹配率']}。"
        )

    operator_note = st.text_area("本次分析备注", placeholder="例如：封装后缀需确认；客户要求 QE=1；需确认是否车规等级。")

    # 生成本次结果
    if not recommend_df.empty:
        best_dict = recommend_df.iloc[0].to_dict()
        new_row = {
            "run_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "competitor_company": selected_company_result,
            "competitor_model": competitor_model_result,
            "product_type": product_type_result,
            "datasheet_source": selected_pdf_url,
            "capacity": confirmed_spec.get("capacity"),
            "density_mb": confirmed_spec.get("density_mb"),
            "voltage_range": confirmed_spec.get("voltage_range"),
            "voltage_type": confirmed_spec.get("voltage_type"),
            "package": confirmed_spec.get("package"),
            "package_size": confirmed_spec.get("package_size"),
            "temperature": confirmed_spec.get("temperature"),
            "temp_grade": confirmed_spec.get("temp_grade"),
            "capacity_confirm": confirm_map.get("capacity", "未确认"),
            "voltage_confirm": confirm_map.get("voltage", "未确认"),
            "package_confirm": confirm_map.get("package", "未确认"),
            "temperature_confirm": confirm_map.get("temperature", "未确认"),
            "recommended_xtx_model": best_dict.get("推荐XTX型号"),
            "match_score": best_dict.get("匹配分"),
            "match_percent": best_dict.get("匹配率"),
            "match_level": best_dict.get("匹配度"),
            "risk_points": best_dict.get("风险点"),
            "operator_note": operator_note,
        }
        current_result_df = pd.DataFrame([new_row])

        st.divider()
        st.subheader("6. 导出 / 保存")
        ec1, ec2, ec3 = st.columns(3)
        with ec1:
            st.download_button(
                "下载本次分析结果 XLSX",
                data=to_excel_bytes({"Current_Result": current_result_df, "Recommendation": recommend_df}),
                file_name=f"{competitor_model_result}_xtx_mapping_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with ec2:
            if st.button("保存本次结果到历史记录"):
                history_df_new = pd.concat([history_df, current_result_df], ignore_index=True)
                st.session_state["history_df_saved"] = history_df_new
                if save_local_history:
                    save_history_local(history_df_new)
                st.success("已保存到历史记录。")
        with ec3:
            hist_download_df = st.session_state.get("history_df_saved", history_df)
            st.download_button(
                "下载历史记录 XLSX",
                data=to_excel_bytes({"History_Log": hist_download_df}),
                file_name="competitor_analysis_history.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# 页面底部：维护数据展示
st.divider()
st.subheader("7. 当前维护数据")

tab1, tab2, tab3, tab4 = st.tabs(["公司白名单", "XTX产品库", "匹配权重", "历史记录"])
with tab1:
    st.markdown("**芯天下/XTX 公司记录**")
    st.dataframe(xtx_company_df, use_container_width=True)
    st.markdown("**友商/竞品公司记录**")
    st.dataframe(competitor_company_df, use_container_width=True)
with tab2:
    st.dataframe(xtx_df, use_container_width=True)
with tab3:
    st.dataframe(weights_df, use_container_width=True)
with tab4:
    hist_display_df = st.session_state.get("history_df_saved", history_df)
    st.dataframe(hist_display_df, use_container_width=True)

st.download_button(
    "下载当前完整维护数据库 XLSX",
    data=to_excel_bytes(
        {
            "Company_Master": company_df,
            "XTX_Product_Library": xtx_df,
            "Match_Weights": weights_df,
            "History_Log": st.session_state.get("history_df_saved", history_df),
        }
    ),
    file_name="xtx_competitor_maintenance_current.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("使用说明 / 限制"):
    st.markdown(
        """
### 维护文件格式说明

1. **Company_Master**：公司官网白名单。
   - `company_role = XTX`：我司芯天下，只展示，不参与竞品 PDF 检索。
   - `company_role = Competitor`：友商，参与自动检索。
   - `search_url_template` 可为空；如果官网搜索地址明确，可写成 `https://xxx.com/search?keyword={model}`。

2. **XTX_Product_Library**：我司产品库。
   - 推荐型号只从该 Sheet 中匹配，避免 AI 临时猜型号。
   - 后续新增产品、封装、温度等级，只需维护该 Sheet。

3. **Match_Weights**：匹配权重。
   - 可以调整容量、电压、封装、温度的权重。
   - `enabled=0` 可临时关闭某一项匹配。

4. **History_Log**：历史记录。
   - 可上传历史记录，也可在本地部署时保存到 `competitor_analysis_history.xlsx`。

### 当前限制

- 自动官网下载依赖竞品官网结构，不保证 100% 成功。
- 扫描版 PDF 暂不支持 OCR。
- 封装可能会识别到多个规格，最终仍建议人工确认具体订购后缀。
- 第一版重点覆盖容量、电压、封装、温度；后续可以继续加入 QE、SFDP、ECC、DTR、RPMC、AEC-Q100 等字段。
"""
    )
