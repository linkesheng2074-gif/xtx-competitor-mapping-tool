"""
XTX 竞品规格书解析与对标推荐工具 V4
=================================
运行方式：
    pip install -r requirements.txt
    streamlit run xtx_competitor_mapping_tool_v2.py

V2 优化点：
1. 自动检索提速：根据型号前缀优先锁定友商；并发检索；减少默认爬取页数；PDF 下载后会校验是否包含输入型号。
2. 避免错抓其他公司规格书：候选 PDF 必须来自白名单域名，并且解析前会验证 PDF 文本中包含目标型号。
3. 增加产品类型解析与显示：SPI NOR / SPI NAND / PPI NOR / PPI NAND / eMMC / DDR3 / DDR4 等。
4. 未识别封装、温度时显示为空，不再显示大量无关封装/温度。
5. 人工确认后自动重新匹配；推荐型号必须满足：产品类型 + 容量 + 电压范围 三项核心条件一致。
6. 维护数据库继续使用 XLSX，兼容旧模板；可选新增 Company_Master.model_prefixes 字段提升自动识别速度。
"""

from __future__ import annotations

import io
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import quote_plus, urljoin, urlparse

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
    page_title="竞品规格书解析与 XTX 对标推荐工具 V4",
    page_icon="🔎",
    layout="wide",
)

REQUEST_TIMEOUT = 5
DEFAULT_MAX_CRAWL_PAGES = 3
MAX_PDF_CANDIDATES = 8
LOCAL_HISTORY_PATH = "competitor_analysis_history.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    )
}

PRODUCT_TYPE_OPTIONS = [
    "未指定",
    "SPI NOR",
    "SPI NAND",
    "PPI NOR",
    "PPI NAND",
    "Parallel NOR",
    "Parallel NAND",
    "eMMC",
    "DDR2",
    "DDR3",
    "DDR4",
    "DDR5",
    "LPDDR4",
    "LPDDR5",
    "EEPROM",
]

# 内置型号前缀规则：可在 Company_Master 中新增 model_prefixes 覆盖或补充
DEFAULT_PREFIX_MAP = {
    "Boya": "BY,BY25,BY26,BY27",
    "Winbond": "W,W25,W25Q,W25N,W29,W39",
    "Macronix": "MX,MX25,MX30,MX35,MX66,MX29",
    "GigaDevice": "GD,GD25,GD5F,GD55,GD32",
    "ISSI": "IS,IS25,IS26,IS29",
    "Puya": "P,P25,PY,PY25,PY26",
    "Zbit": "ZB,ZB25,ZB26",
    "XMC": "XM,XM25,XM26,XM29",
    "ESMT": "F,F25,F59",
    "DS": "DS,DS25,DS26,FM25",
    "Dosilicon": "DS,DS25,DS26,FM25",
    "Giantec": "GT,GT25,GT26",
    "giantec-semi": "GT,GT25,GT26",
    "Tsingteng": "GT,GT25,GT26",
    "tsingtengms": "GT,GT25,GT26",
    "Micron": "MT,MT25,N25,MT29",
    "Longsys": "FORESEE,FS,F25,LGS",
    "Infineon": "S25,S26,S29,SL",
    "Kioxia": "TC58,TH58",
    "SK hynix": "H25,H26,H27",
    "Samsung": "K9,KLM,KLMBG,KLMAG",
}

# =========================================================
# 默认数据：没有上传 XLSX 时也能运行
# =========================================================

DEFAULT_COMPANY_MASTER = pd.DataFrame(
    [
        ["芯天下", "XTX", "xtxtech.com", "https://www.xtxtech.com", "", "XT,XT25,XT26,XT27,XT28", 1, "我司官网，只展示，不参与竞品 PDF 检索"],
        ["Winbond", "Competitor", "winbond.com", "https://www.winbond.com", "", "W,W25,W25Q,W25N,W29,W39", 1, "友商官网"],
        ["Macronix", "Competitor", "macronix.com", "https://www.macronix.com", "", "MX,MX25,MX30,MX35,MX66,MX29", 1, "友商官网"],
        ["GigaDevice", "Competitor", "gigadevice.com", "https://www.gigadevice.com", "", "GD,GD25,GD5F,GD55,GD32", 1, "友商官网"],
        ["ISSI", "Competitor", "issi.com", "https://www.issi.com", "", "IS,IS25,IS26,IS29", 1, "友商官网"],
        ["Puya", "Competitor", "puyasemi.com", "https://www.puyasemi.com", "", "P25,PY,PY25,PY26", 1, "友商官网"],
        ["Boya", "Competitor", "boyamicro.com", "https://www.boyamicro.com", "https://www.boyamicro.com/?zh/products/2", "BY,BY25,BY26,BY27", 1, "友商官网；SPI NOR 产品页作为检索种子"],
        ["Zbit", "Competitor", "zbitsemi.com", "https://www.zbitsemi.com", "", "ZB,ZB25,ZB26", 1, "友商官网"],
        ["XMC", "Competitor", "xmcwh.com", "https://www.xmcwh.com", "", "XM,XM25,XM26,XM29", 1, "友商官网"],
        ["tsingtengms", "Competitor", "tsingtengms.com", "https://www.tsingtengms.com", "", "GT,GT25,GT26", 1, "友商官网"],
        ["DS", "Competitor", "dosilicon.com", "https://www.dosilicon.com", "", "DS,DS25,DS26,FM25", 1, "友商官网"],
        ["ESMT", "Competitor", "esmt.com.tw", "https://www.esmt.com.tw", "", "F,F25,F59", 1, "友商官网"],
        ["giantec-semi", "Competitor", "giantec-semi.com", "https://www.giantec-semi.com", "", "GT,GT25,GT26", 1, "友商官网"],
        ["Micron", "Competitor", "micron.cn", "https://www.micron.cn", "", "MT,MT25,N25,MT29", 1, "友商官网"],
        ["longsys", "Competitor", "longsys.com", "https://cn.longsys.com", "", "FORESEE,FS,F25,LGS", 1, "友商官网"],
        ["infineon", "Competitor", "infineon.com", "https://www.infineon.com", "", "S25,S26,S29,SL", 1, "友商官网"],
        ["Kioxia", "Competitor", "kioxia.com", "https://www.kioxia.com", "", "TC58,TH58", 1, "友商官网"],
        ["Samsung", "Competitor", "semiconductor.samsung.com", "https://semiconductor.samsung.com", "", "K9,KLM,KLMBG,KLMAG", 1, "友商官网"],
        ["SK hynix", "Competitor", "skhynix.com", "https://www.skhynix.com", "", "H25,H26,H27", 1, "友商官网"],
    ],
    columns=["company_name", "company_role", "domain", "base_url", "search_url_template", "model_prefixes", "enabled", "note"],
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
        "package", "package_size", "temp_min", "temp_max", "frequency_mhz", "status", "note",
    ],
)

DEFAULT_MATCH_WEIGHTS = pd.DataFrame(
    [
        ["product_type", 20, 1, "核心条件：产品类型一致，V3 中必须匹配"],
        ["density", 35, 1, "核心条件：容量一致，V3 中必须匹配"],
        ["voltage", 30, 1, "核心条件：电压范围一致，V3 中必须匹配"],
        ["package", 10, 1, "封装形式和尺寸兼容性"],
        ["temperature", 10, 1, "温度等级覆盖"],
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


def ensure_company_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    for col in DEFAULT_COMPANY_MASTER.columns:
        if col not in df.columns:
            df[col] = ""
    # 兼容旧模板：自动补 model_prefixes
    for idx, row in df.iterrows():
        if str(row.get("model_prefixes", "")).strip() in ["", "nan", "None"]:
            company = str(row.get("company_name", "")).strip()
            for key, prefixes in DEFAULT_PREFIX_MAP.items():
                if key.lower() in company.lower() or company.lower() in key.lower():
                    df.at[idx, "model_prefixes"] = prefixes
                    break
    return df


def ensure_xtx_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    for col in DEFAULT_XTX_PRODUCT_LIBRARY.columns:
        if col not in df.columns:
            df[col] = ""
    return df


@st.cache_data(show_spinner=False)
def read_maintenance_xlsx(file_bytes: bytes | None) -> dict[str, pd.DataFrame]:
    if not file_bytes:
        return {
            "Company_Master": DEFAULT_COMPANY_MASTER.copy(),
            "XTX_Product_Library": DEFAULT_XTX_PRODUCT_LIBRARY.copy(),
            "Match_Weights": DEFAULT_MATCH_WEIGHTS.copy(),
            "History_Log": DEFAULT_HISTORY_LOG.copy(),
        }
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    return {
        "Company_Master": ensure_company_columns(sheets.get("Company_Master", DEFAULT_COMPANY_MASTER.copy())),
        "XTX_Product_Library": ensure_xtx_columns(sheets.get("XTX_Product_Library", DEFAULT_XTX_PRODUCT_LIBRARY.copy())),
        "Match_Weights": clean_columns(sheets.get("Match_Weights", DEFAULT_MATCH_WEIGHTS.copy())),
        "History_Log": clean_columns(sheets.get("History_Log", DEFAULT_HISTORY_LOG.copy())),
    }


def read_history_xlsx(history_upload) -> pd.DataFrame:
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
            df.to_excel(writer, index=False, sheet_name=str(sheet_name)[:31])
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


def normalize_product_type(pt: str) -> str:
    s = str(pt or "").strip().upper().replace("PARALLEL", "PPI")
    s = re.sub(r"\s+", " ", s)
    if s in ["", "NAN", "NONE", "未指定"]:
        return ""
    if "SPI" in s and "NOR" in s:
        return "SPI NOR"
    if "SPI" in s and "NAND" in s:
        return "SPI NAND"
    if ("PPI" in s or "PARALLEL" in s) and "NOR" in s:
        return "PPI NOR"
    if ("PPI" in s or "PARALLEL" in s) and "NAND" in s:
        return "PPI NAND"
    if "EMMC" in s:
        return "eMMC"
    if "LPDDR5" in s:
        return "LPDDR5"
    if "LPDDR4" in s:
        return "LPDDR4"
    if "DDR5" in s:
        return "DDR5"
    if "DDR4" in s:
        return "DDR4"
    if "DDR3" in s:
        return "DDR3"
    if "DDR2" in s:
        return "DDR2"
    if "EEPROM" in s:
        return "EEPROM"
    return str(pt or "").strip()


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
    return bool(allowed) and (current == allowed or current.endswith("." + allowed))


class CachedResponse:
    def __init__(self, data: dict):
        self.url = data.get("url", "")
        self.status_code = int(data.get("status_code", 0))
        self.headers = data.get("headers", {}) or {}
        self.content = data.get("content", b"") or b""

    @property
    def text(self) -> str:
        encoding = "utf-8"
        ctype = self.headers.get("Content-Type", "") or self.headers.get("content-type", "")
        m = re.search(r"charset=([\w\-]+)", ctype, flags=re.IGNORECASE)
        if m:
            encoding = m.group(1)
        try:
            return self.content.decode(encoding, errors="replace")
        except Exception:
            return self.content.decode("utf-8", errors="replace")

    def raise_for_status(self):
        if self.status_code >= 400 or self.status_code == 0:
            raise requests.HTTPError(f"HTTP {self.status_code} for {self.url}")


@st.cache_data(ttl=1800, show_spinner=False)
def cached_get_url(url: str) -> dict:
    """缓存官网页面/PDF下载结果，避免 Streamlit rerun 时重复请求同一链接。"""
    resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
    return {
        "url": resp.url,
        "status_code": resp.status_code,
        "headers": dict(resp.headers),
        "content": resp.content,
    }


def request_url(url: str):
    return CachedResponse(cached_get_url(url))


def values_match(a, b, tol=0.15) -> bool:
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) <= tol


# =========================================================
# 友商自动识别与 PDF 查找
# =========================================================

def parse_prefixes(value: str) -> list[str]:
    parts = re.split(r"[,，;；/\s]+", str(value or ""))
    return [normalize_model(p) for p in parts if normalize_model(p)]


def infer_company_df_by_model(competitor_df: pd.DataFrame, model: str, selected_company: str) -> tuple[pd.DataFrame, str]:
    if selected_company and selected_company != "自动识别/全部友商":
        return competitor_df[competitor_df["company_name"].astype(str) == selected_company].copy(), "手动选择厂商"

    model_norm = normalize_model(model)
    if not model_norm:
        return competitor_df.copy(), "未输入型号，使用全部友商"

    matched_rows = []
    for _, row in competitor_df.iterrows():
        prefixes = parse_prefixes(row.get("model_prefixes", ""))
        # 长前缀优先，避免 B / BY 这种过短前缀造成误判
        prefixes = sorted(prefixes, key=len, reverse=True)
        for p in prefixes:
            if len(p) >= 2 and model_norm.startswith(p):
                matched_rows.append(row)
                break

    if matched_rows:
        df = pd.DataFrame(matched_rows)
        names = ", ".join(df["company_name"].astype(str).tolist())
        return df, f"根据型号前缀自动锁定：{names}"

    return competitor_df.copy(), "未匹配到型号前缀，使用全部友商"


def score_pdf_candidate(url: str, link_text: str, model: str) -> int:
    model_norm = normalize_model(model)
    url_norm = normalize_model(url)
    text_norm = normalize_model(link_text)
    combined = f"{url_norm} {text_norm}"

    score = 0
    if model_norm and model_norm in url_norm:
        score += 120
    if model_norm and model_norm in text_norm:
        score += 120

    # 型号不完整时，给系列号弱加分；避免仅凭 25Q/25D 误匹配到其他公司
    if model_norm and len(model_norm) >= 6:
        family = model_norm[:6]
        if family in combined:
            score += 30

    if "DATASHEET" in combined or "DATASHEET" in combined.replace(" ", "") or "DATA SHEET" in combined:
        score += 45
    if "SPECIFICATION" in combined or "SPEC" in combined:
        score += 25
    if "PDF" in combined:
        score += 10

    if model_norm and len(model_norm) >= 5:
        score += int(fuzz.partial_ratio(model_norm, combined) * 0.15)

    bad_words = [
        "ANNUAL", "REPORT", "ESG", "FINANCIAL", "PRESENTATION", "BROCHURE",
        "CATALOG", "QUALITY", "RELIABILITY", "PACKAGEINFOONLY", "CERTIFICATE",
        "ROHS", "REACH", "CONFLICTMINERALS",
    ]
    if any(word in combined for word in bad_words):
        score -= 60
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


def page_text_contains_model(html: str, model: str) -> bool:
    model_norm = normalize_model(model)
    if not model_norm:
        return False
    text = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)
    return model_norm in normalize_model(text)



def normalize_header_name(name: str) -> str:
    """把产品页表头映射成统一字段。"""
    n = re.sub(r"[\s\n\r\t:：.。/()（）_-]+", "", str(name or "")).lower()
    if n in ["density", "capacity", "memory", "容量"] or "density" in n or "capacity" in n:
        return "density"
    if n in ["partno", "partnumber", "model", "型号", "料号"] or "part" in n or "model" in n:
        return "part_no"
    if n in ["vcc", "voltage", "电压"] or "vcc" in n or "voltage" in n:
        return "vcc"
    if n in ["frequency", "freq", "speed", "频率"] or "frequency" in n or "freq" in n:
        return "frequency"
    if n in ["feature", "features", "功能", "特性"] or "feature" in n:
        return "feature"
    if n in ["package", "pkg", "封装"] or "package" in n:
        return "package"
    if n in ["tem", "temp", "temperature", "operatingtemperature", "温度"] or "temp" in n or n == "tem":
        return "temperature"
    if n in ["status", "状态"] or "status" in n:
        return "status"
    return n or "col"


def table_rows_containing_model(html: str, model: str, current_url: str) -> list[dict]:
    """从官网产品页 HTML 表格中抽取包含目标型号的行。"""
    soup = BeautifulSoup(html, "html.parser")
    model_norm = normalize_model(model)
    rows_out = []
    if not model_norm:
        return rows_out

    for table_idx, table in enumerate(soup.find_all("table")):
        trs = table.find_all("tr")
        if not trs:
            continue

        headers = []
        header_tr_index = 0
        for i, tr in enumerate(trs[:3]):
            cells = tr.find_all(["th", "td"])
            texts = [c.get_text(" ", strip=True) for c in cells]
            if any(normalize_header_name(t) in ["density", "part_no", "vcc", "package", "temperature"] for t in texts):
                headers = [normalize_header_name(t) for t in texts]
                header_tr_index = i
                break

        if not headers:
            first_cells = trs[0].find_all(["th", "td"])
            headers = [normalize_header_name(c.get_text(" ", strip=True)) for c in first_cells]
            header_tr_index = 0

        for tr in trs[header_tr_index + 1:]:
            cells = tr.find_all(["td", "th"])
            values = [c.get_text(" ", strip=True) for c in cells]
            if not values:
                continue
            row_text = " ".join(values)
            if model_norm not in normalize_model(row_text):
                continue

            row = {"source_url": current_url, "table_index": table_idx, "row_text": row_text}
            for idx, value in enumerate(values):
                key = headers[idx] if idx < len(headers) and headers[idx] else f"col_{idx}"
                # 避免重复表头覆盖，把重复列拼接起来
                if key in row and str(row[key]).strip():
                    row[key] = f"{row[key]} | {value}"
                else:
                    row[key] = value
            rows_out.append(row)
    return rows_out


def extract_pdf_links_from_page(html: str, current_url: str, domain: str, model: str) -> list[dict]:
    links = extract_links_from_html(html, current_url)
    pdfs = []
    for item in links:
        link_url = item["url"]
        link_text = item["text"]
        if not same_domain_or_subdomain(link_url, domain):
            continue
        combined = (link_url + " " + link_text).lower()
        if ".pdf" in combined or "datasheet" in combined or "download" in combined or "规格书" in combined:
            pdfs.append({
                "url": link_url,
                "link_text": link_text,
                "score": score_pdf_candidate(link_url, link_text, model),
            })
    return sorted(pdfs, key=lambda x: x["score"], reverse=True)


def crawl_company_for_product_page(company_row: pd.Series, model: str, max_pages: int = 3) -> pd.DataFrame:
    """产品页优先：如果官网产品表格已经有目标型号，直接返回该行数据，避免慢速 PDF 下载/扫描。"""
    company = str(company_row.get("company_name", "")).strip()
    domain = str(company_row.get("domain", "")).strip()
    base_url = str(company_row.get("base_url", "")).strip()
    search_template = str(company_row.get("search_url_template", "")).strip()
    if not domain or not base_url:
        return pd.DataFrame()

    visited = set()
    queue = build_seed_urls(company, base_url, search_template, model)
    results = []
    model_norm = normalize_model(model)

    while queue and len(visited) < max_pages:
        current_url = queue.pop(0)
        if current_url in visited:
            continue
        visited.add(current_url)
        try:
            resp = request_url(current_url)
            resp.raise_for_status()
        except Exception:
            continue
        final_url = resp.url
        if not same_domain_or_subdomain(final_url, domain):
            continue
        ctype = resp.headers.get("Content-Type", "").lower()
        if "html" not in ctype and "text" not in ctype:
            continue

        html = resp.text
        page_text = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)
        page_has_model = model_norm in normalize_model(page_text)
        pdfs = extract_pdf_links_from_page(html, final_url, domain, model)
        best_pdf = pdfs[0]["url"] if pdfs else ""

        # 1) 优先使用表格行，最准确、最快
        for row in table_rows_containing_model(html, model, final_url):
            row.update({
                "company_name": company,
                "page_url": final_url,
                "datasheet_url": best_pdf,
                "source_kind": "官网产品页表格",
                "page_score": 500 + score_pdf_candidate(final_url, row.get("row_text", ""), model),
            })
            results.append(row)

        # 2) 没有表格行，但页面文本包含型号，也作为兜底
        if page_has_model and not results:
            results.append({
                "company_name": company,
                "page_url": final_url,
                "source_url": final_url,
                "datasheet_url": best_pdf,
                "source_kind": "官网产品页文本",
                "row_text": page_text[:5000],
                "page_score": 250 + score_pdf_candidate(final_url, page_text[:1000], model),
            })

        # 3) 页面没命中时，只继续少量高相关链接，不做全站漫游
        if not page_has_model:
            links = extract_links_from_html(html, final_url)
            for item in links:
                link_url = item["url"]
                link_text = item["text"]
                if not same_domain_or_subdomain(link_url, domain):
                    continue
                combined = normalize_model(link_url + " " + link_text)
                if model_norm in combined or any(k in link_url.lower() for k in ["product", "products", "flash", "nor", "nand"]):
                    if link_url not in visited and len(queue) < max_pages * 2:
                        queue.append(link_url)

    if not results:
        return pd.DataFrame()
    df = pd.DataFrame(results).sort_values("page_score", ascending=False).drop_duplicates(subset=["page_url", "row_text"])
    return df


def crawl_product_pages_parallel(search_df: pd.DataFrame, model: str, max_pages: int) -> pd.DataFrame:
    if search_df.empty:
        return pd.DataFrame()
    dfs = []
    max_workers = min(4, max(1, len(search_df)))
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(crawl_company_for_product_page, row, model, max_pages): row.get("company_name", "") for _, row in search_df.iterrows()}
        for future in as_completed(futures):
            try:
                df = future.result()
                if not df.empty:
                    dfs.append(df)
            except Exception:
                pass
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True).sort_values("page_score", ascending=False).drop_duplicates(subset=["page_url", "row_text"])


def spec_from_product_page_row(row: pd.Series, selected_type: str, model: str) -> dict:
    """把官网产品页表格行转换成统一 spec。"""
    row_text = str(row.get("row_text", ""))
    density_text = str(row.get("density", "")) or row_text
    vcc_text = str(row.get("vcc", "")) or row_text
    package_text = str(row.get("package", "")) or row_text
    temp_text = str(row.get("temperature", "")) or row_text

    product_type = extract_product_type(row_text, selected_type=selected_type, model=model)
    density_mb, density_display = parse_density_from_text(density_text)
    if density_mb is None:
        d = extract_density(row_text)
        density_mb, density_display = d.get("density_mb"), d.get("display", "")
    vmin, vmax = parse_voltage_from_text(vcc_text)
    voltage = build_voltage_dict(vmin, vmax, "产品页")
    package = extract_package(package_text)
    temp = extract_temperature(temp_text)

    return {
        "product_type": product_type.get("product_type", ""),
        "product_type_confidence": product_type.get("confidence", "产品页"),
        "capacity": density_display or "",
        "density_mb": density_mb,
        "capacity_confidence": "产品页" if density_display else "低",
        "voltage_range": voltage.get("display", ""),
        "voltage_type": voltage.get("voltage_type", ""),
        "vcc_min": voltage.get("vcc_min"),
        "vcc_max": voltage.get("vcc_max"),
        "voltage_confidence": voltage.get("confidence", "产品页"),
        "package": package.get("package", ""),
        "package_size": package.get("package_size", ""),
        "package_confidence": "产品页" if package.get("package") or package.get("package_size") else "低",
        "temperature": temp.get("display", ""),
        "temp_grade": temp.get("temp_grade", ""),
        "temp_min": temp.get("temp_min"),
        "temp_max": temp.get("temp_max"),
        "temperature_confidence": "产品页" if temp.get("display") else "低",
    }


def build_seed_urls(company: str, base_url: str, search_template: str, model: str) -> list[str]:
    """构造少量高价值入口，避免全站漫游拖慢速度。search_url_template 可以是含 {model} 的搜索地址，也可以是固定产品页。"""
    seeds = []
    model_q = quote_plus(model)
    if search_template and search_template.lower() not in ["nan", "none", ""]:
        if "{model}" in search_template:
            seeds.append(search_template.replace("{model}", model_q))
        else:
            seeds.append(search_template)

    seeds.append(base_url)
    base = base_url.rstrip("/")
    company_l = str(company).lower()

    # 少量通用搜索入口，控制数量，避免拖慢
    seeds.extend([
        f"{base}/?s={model_q}",
        f"{base}/search?keyword={model_q}",
        f"{base}/search?keywords={model_q}",
        f"{base}/search?q={model_q}",
    ])

    # 友商常用产品页作为额外种子；这些是官网页面，不是第三方来源。
    if "boya" in company_l:
        seeds.extend([
            "https://www.boyamicro.com/?zh/products/2",
            "https://www.boyamicro.com/zh/products/2",
        ])
    elif "gigadevice" in company_l:
        seeds.extend([
            "https://www.gigadevice.com/product/memory/",
            "https://www.gigadevice.com.cn/product/memory/",
        ])
    elif "macronix" in company_l:
        seeds.append("https://www.macronix.com/en-us/products/NOR-Flash/Pages/default.aspx")
    elif "winbond" in company_l:
        seeds.append("https://www.winbond.com/hq/product/code-storage-flash-memory/serial-nor-flash/")
    elif "issi" in company_l:
        seeds.append("https://www.issi.com/US/product-flash.shtml")
    elif "zbit" in company_l:
        seeds.append("https://www.zbitsemi.com/product")
    elif "puy" in company_l:
        seeds.append("https://www.puyasemi.com/product.html")

    return list(dict.fromkeys([u for u in seeds if valid_url(u)]))


def crawl_company_for_pdf(company_row: pd.Series, model: str, max_pages: int = DEFAULT_MAX_CRAWL_PAGES) -> pd.DataFrame:
    company = str(company_row.get("company_name", "")).strip()
    domain = str(company_row.get("domain", "")).strip()
    base_url = str(company_row.get("base_url", "")).strip()
    search_template = str(company_row.get("search_url_template", "")).strip()

    if not domain or not base_url:
        return pd.DataFrame()

    visited = set()
    queue = build_seed_urls(company, base_url, search_template, model)
    candidates = []
    model_norm = normalize_model(model)

    while queue and len(visited) < max_pages:
        current_url = queue.pop(0)
        if current_url in visited:
            continue
        visited.add(current_url)

        try:
            resp = request_url(current_url)
            resp.raise_for_status()
        except Exception:
            continue

        final_url = resp.url
        if not same_domain_or_subdomain(final_url, domain):
            continue

        content_type = resp.headers.get("Content-Type", "").lower()
        if "pdf" in content_type or final_url.lower().endswith(".pdf"):
            candidates.append({
                "company_name": company,
                "url": final_url,
                "link_text": final_url,
                "score": score_pdf_candidate(final_url, final_url, model),
            })
            continue

        if "html" not in content_type and "text" not in content_type:
            continue

        html = resp.text
        links = extract_links_from_html(html, final_url)
        current_page_has_model = page_text_contains_model(html, model)

        for item in links:
            link_url = item["url"]
            link_text = item["text"]
            if not same_domain_or_subdomain(link_url, domain):
                continue

            link_url_lower = link_url.lower()
            link_url_norm = normalize_model(link_url)
            link_text_norm = normalize_model(link_text)
            is_pdf = ".pdf" in link_url_lower or "pdf" in link_text.lower() or "download" in link_url_lower

            if is_pdf:
                # 如果当前页或链接文本/URL包含型号，则优先；否则仍收集但分数会较低，后续 PDF 会校验型号
                candidates.append({
                    "company_name": company,
                    "url": link_url,
                    "link_text": link_text,
                    "score": score_pdf_candidate(link_url, link_text, model) + (40 if current_page_has_model else 0),
                })
                continue

            relevant_keywords = [
                "product", "products", "memory", "flash", "nor", "nand", "serial", "spi",
                "datasheet", "download", "document", "spec", "support", "resource", "file",
            ]
            looks_relevant = (
                model_norm in link_url_norm
                or model_norm in link_text_norm
                or (current_page_has_model and any(k in link_url_lower for k in ["download", "pdf", "file", "datasheet"]))
                or any(k in link_url_lower for k in relevant_keywords)
            )
            if looks_relevant and link_url not in visited and len(queue) < max_pages * 2:
                queue.append(link_url)

    if not candidates:
        return pd.DataFrame()
    df = pd.DataFrame(candidates)
    df = df.sort_values("score", ascending=False).drop_duplicates(subset=["url"])
    # 分数过低的候选基本是无关 PDF，先过滤一层，后续还会做 PDF 文本校验
    return df[df["score"] >= 20].head(MAX_PDF_CANDIDATES)


def crawl_companies_parallel(search_df: pd.DataFrame, model: str, max_pages: int) -> pd.DataFrame:
    dfs = []
    if search_df.empty:
        return pd.DataFrame()
    max_workers = min(5, max(1, len(search_df)))
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(crawl_company_for_pdf, row, model, max_pages): row.get("company_name", "") for _, row in search_df.iterrows()}
        for future in as_completed(futures):
            try:
                df = future.result()
                if not df.empty:
                    dfs.append(df)
            except Exception:
                pass
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True).sort_values("score", ascending=False).drop_duplicates(subset=["url"]).head(MAX_PDF_CANDIDATES)


def download_pdf(pdf_url: str) -> bytes:
    resp = request_url(pdf_url)
    resp.raise_for_status()
    content = resp.content
    # 有些错误页是 HTML，不要当 PDF 解析
    if not content.startswith(b"%PDF") and "pdf" not in resp.headers.get("Content-Type", "").lower():
        raise ValueError("下载内容不是 PDF，可能是错误页或需要权限访问")
    return content


def extract_text_from_pdf_bytes(pdf_bytes: bytes, max_pages: int = 30) -> str:
    """详细解析 PDF 文本，用于字段抽取。"""
    text_parts = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        total_pages = min(len(doc), max_pages) if max_pages else len(doc)
        for idx in range(total_pages):
            page = doc[idx]
            text = page.get_text("text", sort=True)
            if text:
                text_parts.append(f"\n\n--- Page {idx + 1} ---\n{text}")
    return "\n".join(text_parts)


def pdf_text_mentions_model(pdf_text: str, model: str) -> bool:
    model_norm = normalize_model(model)
    text_norm = normalize_model(pdf_text)
    if not model_norm:
        return False
    if model_norm in text_norm:
        return True
    # 对带封装/温度后缀的型号，允许主体型号命中，但主体长度不能太短，避免误匹配。
    for n in [10, 9, 8, 7]:
        if len(model_norm) >= n + 2 and model_norm[:n] in text_norm:
            return True
    return False


def scan_pdf_for_model_fast(pdf_bytes: bytes, model: str, max_pages: int | None = None) -> tuple[bool, int | None, str]:
    """
    快速扫描 PDF 全文定位型号，找到即停止。
    只做型号校验，不把全文都拼接进内存；比“详细解析全部 PDF”更快。
    返回：(是否命中, 命中页码, 命中方式)
    """
    model_norm = normalize_model(model)
    if not model_norm:
        return False, None, "empty_model"

    prefixes = []
    for n in [10, 9, 8, 7]:
        if len(model_norm) >= n + 2:
            prefixes.append(model_norm[:n])

    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        page_count = len(doc)
        scan_pages = min(page_count, max_pages) if max_pages else page_count
        for idx in range(scan_pages):
            try:
                text = doc[idx].get_text("text", sort=False) or ""
            except Exception:
                continue
            norm = normalize_model(text)
            if model_norm in norm:
                return True, idx + 1, "exact_model"
            if any(p in norm for p in prefixes):
                return True, idx + 1, "model_prefix"
    return False, None, f"not_found_in_{scan_pages}_pages"


def get_pdf_response_metadata(pdf_url: str) -> dict:
    meta = {"final_url": pdf_url, "last_modified": "", "content_length": ""}
    try:
        resp = requests.head(pdf_url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        meta["final_url"] = resp.url
        meta["last_modified"] = resp.headers.get("Last-Modified", "")
        meta["content_length"] = resp.headers.get("Content-Length", "")
    except Exception:
        pass
    return meta


def choose_valid_pdf(candidate_df: pd.DataFrame, model: str, max_scan_pages: int | None = None, max_candidates: int = 3) -> tuple[bytes | None, str, pd.DataFrame, str]:
    """
    V3：对候选 PDF 逐个下载后“快速扫描全文”校验目标型号。
    不再只看前几页，避免型号在 PDF 后部时误判；同时记录 final_url/Last-Modified，优先选择官网可访问、型号命中的候选。
    """
    if candidate_df.empty:
        return None, "", candidate_df, "未找到候选 PDF"

    checked_rows = []
    valid_items = []

    for _, row in candidate_df.sort_values("score", ascending=False).head(max_candidates).iterrows():
        url = str(row.get("url", "")).strip()
        check = row.to_dict()
        check.update(get_pdf_response_metadata(url))
        try:
            pdf_bytes = download_pdf(url)
            found, page_no, method = scan_pdf_for_model_fast(pdf_bytes, model, max_pages=max_scan_pages)
            check["model_found"] = "是" if found else "否"
            check["model_page"] = page_no if page_no else ""
            check["validation_method"] = method

            if found:
                # 稳定性评分：来源官网 + URL/文本含型号 + datasheet关键词 + PDF可访问 + 型号命中页
                validation_score = float(check.get("score", 0)) + 500
                combined = normalize_model(str(check.get("url", "")) + " " + str(check.get("link_text", "")))
                model_norm = normalize_model(model)
                if model_norm in combined:
                    validation_score += 120
                if "DATASHEET" in combined or "SPEC" in combined:
                    validation_score += 50
                # 页码越靠前，通常越可能是标题页/订购页
                if page_no:
                    validation_score += max(0, 50 - min(page_no, 50))
                check["validation"] = "通过：PDF全文扫描包含目标型号"
                check["validation_score"] = round(validation_score, 2)
                valid_items.append((validation_score, pdf_bytes, check))
            else:
                check["validation"] = "未通过：PDF全文扫描未找到目标型号"
                check["validation_score"] = float(check.get("score", 0))
        except Exception as e:
            check["model_found"] = "否"
            check["model_page"] = ""
            check["validation_method"] = "download_or_parse_error"
            check["validation"] = f"未通过：{e}"
            check["validation_score"] = float(check.get("score", 0))
        checked_rows.append(check)

    checked_df = pd.DataFrame(checked_rows)
    if valid_items:
        valid_items.sort(key=lambda x: x[0], reverse=True)
        _, best_pdf_bytes, best_check = valid_items[0]
        best_url = str(best_check.get("final_url") or best_check.get("url") or "")
        checked_df = checked_df.sort_values("validation_score", ascending=False)
        return best_pdf_bytes, best_url, checked_df, ""

    return None, "", checked_df, "没有找到通过型号全文校验的 PDF。建议手动输入官网 PDF 链接或上传规格书 PDF。"


# =========================================================
# 字段抽取：产品类型、容量、电压、封装、温度
# =========================================================

def blank_field(extra: dict | None = None) -> dict:
    data = {"display": "", "confidence": "低"}
    if extra:
        data.update(extra)
    return data


def extract_product_type(text: str, selected_type: str = "未指定", model: str = "") -> dict:
    selected_norm = normalize_product_type(selected_type)
    if selected_norm:
        return {"product_type": selected_norm, "display": selected_norm, "confidence": "人工选择"}

    t = text.upper()
    checks = [
        ("SPI NAND", ["SPI NAND", "SERIAL NAND"]),
        ("SPI NOR", ["SPI NOR", "SERIAL NOR", "SERIAL FLASH MEMORY", "SERIAL FLASH"]),
        ("PPI NAND", ["PARALLEL NAND", "PPI NAND"]),
        ("PPI NOR", ["PARALLEL NOR", "PPI NOR"]),
        ("eMMC", ["EMMC", "EMBEDDED MULTI MEDIA CARD"]),
        ("DDR5", ["DDR5"]),
        ("DDR4", ["DDR4"]),
        ("DDR3", ["DDR3"]),
        ("DDR2", ["DDR2"]),
        ("LPDDR5", ["LPDDR5"]),
        ("LPDDR4", ["LPDDR4"]),
        ("EEPROM", ["EEPROM"]),
    ]
    for pt, keywords in checks:
        if any(k in t for k in keywords):
            return {"product_type": pt, "display": pt, "confidence": "中"}

    model_norm = normalize_model(model)
    if model_norm.startswith(("BY25", "W25", "MX25", "GD25", "P25", "ZB25", "IS25", "XM25", "GT25")):
        return {"product_type": "SPI NOR", "display": "SPI NOR", "confidence": "低：根据型号前缀推断"}
    if model_norm.startswith(("BY26", "W25N", "GD5F", "ZB26", "IS26", "XM26")):
        return {"product_type": "SPI NAND", "display": "SPI NAND", "confidence": "低：根据型号前缀推断"}

    return {"product_type": "", "display": "", "confidence": "低"}


def extract_density(text: str) -> dict:
    t = text.upper()
    candidates = []
    patterns = [
        r"(\d+(?:\.\d+)?)\s*G\s*[- ]?\s*BIT",
        r"(\d+(?:\.\d+)?)\s*GBIT",
        r"(\d+(?:\.\d+)?)\s*M\s*[- ]?\s*BIT",
        r"(\d+(?:\.\d+)?)\s*MBIT",
        r"(\d+(?:\.\d+)?)\s*K\s*[- ]?\s*BIT",
        r"(\d+(?:\.\d+)?)\s*KBIT",
        r"(\d+(?:\.\d+)?)\s*M\s*BYTE",
        r"(\d+(?:\.\d+)?)\s*MBYTE",
    ]
    for pattern in patterns:
        for m in re.finditer(pattern, t):
            value = float(m.group(1))
            raw = m.group(0).upper()
            if "GBIT" in raw or ("G" in raw and "BIT" in raw):
                density_mb = int(value * 1024)
                display = f"{value:g}Gb"
            elif "MBIT" in raw or ("M" in raw and "BIT" in raw):
                density_mb = int(value)
                display = f"{density_mb}Mb"
            elif "KBIT" in raw or ("K" in raw and "BIT" in raw):
                density_mb = value / 1024
                display = f"{value:g}Kb"
            elif "MBYTE" in raw or "M BYTE" in raw:
                density_mb = int(value * 8)
                display = f"{density_mb}Mb"
            else:
                continue
            if 0.001 <= density_mb <= 131072:
                candidates.append({"density_mb": density_mb, "display": display, "raw": raw})

    if not candidates:
        return {"density_mb": None, "display": "", "confidence": "低"}

    df = pd.DataFrame(candidates)
    grouped = (
        df.groupby(["density_mb", "display"])
        .size()
        .reset_index(name="count")
        .sort_values(["count", "density_mb"], ascending=[False, False])
    )
    best = grouped.iloc[0]
    density = best["density_mb"]
    density_value = int(density) if float(density).is_integer() else float(density)
    return {
        "density_mb": density_value,
        "display": str(best["display"]),
        "confidence": "高" if int(best["count"]) >= 2 else "中",
    }


def parse_density_from_text(value: str):
    if not value:
        return None, ""
    s = str(value).upper().replace(" ", "")
    m = re.search(r"(\d+(?:\.\d+)?)(GBIT|G-BIT|GB|G)", s)
    if m:
        density = int(float(m.group(1)) * 1024)
        return density, f"{float(m.group(1)):g}Gb"
    m = re.search(r"(\d+(?:\.\d+)?)(MBIT|M-BIT|MB|M)", s)
    if m:
        density = int(float(m.group(1)))
        return density, f"{density}Mb"
    m = re.search(r"(\d+(?:\.\d+)?)(KBIT|K-BIT|KB|K)", s)
    if m:
        kb = float(m.group(1))
        return kb / 1024, f"{kb:g}Kb"
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if m:
        density = int(float(m.group(1)))
        return density, f"{density}Mb"
    return None, ""


def build_voltage_dict(vmin, vmax, confidence="中") -> dict:
    if vmin is None or vmax is None:
        return {"vcc_min": None, "vcc_max": None, "display": "", "voltage_type": "", "confidence": "低"}
    vmin = float(vmin)
    vmax = float(vmax)
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
            vmin, vmax = min(v1, v2), max(v1, v2)
            if 1.0 <= vmin <= 5.5 and 1.0 <= vmax <= 5.5:
                candidates.append({"vcc_min": vmin, "vcc_max": vmax, "raw": m.group(0)})
    if not candidates:
        return build_voltage_dict(None, None, "低")

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
        "BGA", "LGA", "WFBGA",
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
        return {"package": "", "package_size": "", "confidence": "低"}
    return {
        "package": " / ".join(packages[:10]) if packages else "",
        "package_size": " / ".join(sizes[:10]) if sizes else "",
        "confidence": "中",
    }


def build_temperature_dict(temp_min, temp_max, confidence="中") -> dict:
    if temp_min is None or temp_max is None:
        return {"temp_min": None, "temp_max": None, "display": "", "temp_grade": "", "confidence": "低"}
    temp_min = int(temp_min)
    temp_max = int(temp_max)
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
            t1, t2 = int(m.group(1)), int(m.group(2))
            temp_min, temp_max = min(t1, t2), max(t1, t2)
            if -65 <= temp_min <= 25 and 70 <= temp_max <= 150:
                candidates.append({"temp_min": temp_min, "temp_max": temp_max, "raw": m.group(0)})
    if not candidates:
        return build_temperature_dict(None, None, "低")

    df = pd.DataFrame(candidates)
    grouped = (
        df.groupby(["temp_min", "temp_max"])
        .size()
        .reset_index(name="count")
        .sort_values(["count", "temp_max"], ascending=[False, False])
    )
    best = grouped.iloc[0]
    return build_temperature_dict(int(best["temp_min"]), int(best["temp_max"]), "高" if int(best["count"]) >= 2 else "中")


def parse_temperature_from_text(value: str):
    if not value:
        return None, None
    s = str(value).replace("–", "-").replace("—", "-").replace("~", "-").replace("～", "-")
    nums = re.findall(r"-?\d+", s)
    nums = [int(x) for x in nums]
    if len(nums) >= 2:
        return min(nums[0], nums[1]), max(nums[0], nums[1])
    return None, None


def build_focus_text_around_model(text: str, model: str, window: int = 1800, max_hits: int = 8) -> str:
    """提取目标型号附近文本，优先用于容量/封装/温度等与订购型号强相关字段，减少被其它型号表格干扰。"""
    model_norm = normalize_model(model)
    if not model_norm or not text:
        return ""

    # 在标准化文本里定位不方便映射原始下标；这里用宽松正则在原文中找型号字符序列。
    pattern = r"\\s*[-_/\\.]*".join([re.escape(ch) for ch in model])
    hits = []
    try:
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            hits.append((m.start(), m.end()))
            if len(hits) >= max_hits:
                break
    except Exception:
        hits = []

    # 如果完整型号找不到，尝试主体前缀
    if not hits and len(model_norm) >= 8:
        prefix = model_norm[:8]
        for m in re.finditer(re.escape(prefix), normalize_model(text), flags=re.IGNORECASE):
            # 无法精确映射原文位置，直接返回前几页文本兜底
            break

    chunks = []
    for s, e in hits:
        chunks.append(text[max(0, s - window): min(len(text), e + window)])
    return "\\n".join(chunks)


def build_analysis_texts(text: str, model: str) -> tuple[str, str]:
    """返回：focus_text, full_text。focus_text 供强相关字段优先识别；full_text 供电压等通用字段兜底。"""
    focus = build_focus_text_around_model(text, model)
    # 加上第一页/前若干字符，标题页常有产品类型/容量/电压
    first_part = text[:12000]
    if focus:
        return focus + "\\n" + first_part, text
    return first_part, text


def analyze_spec_text(text: str, selected_type: str, model: str) -> dict:
    focus_text, full_text = build_analysis_texts(text, model)

    # 产品类型/容量/封装/温度优先看目标型号附近文本，避免同一 PDF 多个型号表格干扰。
    product_type = extract_product_type(focus_text or full_text, selected_type=selected_type, model=model)
    density = extract_density(focus_text or full_text)
    package = extract_package(focus_text)
    temp = extract_temperature(focus_text)

    # 电压通常在 Recommended Operating Conditions，可能不在目标型号附近，使用全文。
    voltage = extract_voltage(full_text)

    # 如果 focus_text 未识别温度，则不要强行从全文抓其它产品/可靠性温度；保持空，交给人工确认。
    if not temp.get("display"):
        temp = build_temperature_dict(None, None, "低")

    return {
        "product_type": product_type["product_type"],
        "product_type_confidence": product_type["confidence"],
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
            ["product_type", "产品类型", spec.get("product_type", ""), "", spec.get("product_type_confidence", ""), "未确认"],
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
        if key == "product_type":
            spec["product_type"] = normalize_product_type(manual)
            spec["product_type_confidence"] = "人工确认"
        elif key == "capacity":
            density_mb, display = parse_density_from_text(manual)
            spec["density_mb"] = density_mb
            spec["capacity"] = display
            spec["capacity_confidence"] = "人工确认"
        elif key == "voltage":
            vmin, vmax = parse_voltage_from_text(manual)
            vd = build_voltage_dict(vmin, vmax, "人工确认")
            spec.update({
                "vcc_min": vd["vcc_min"], "vcc_max": vd["vcc_max"],
                "voltage_range": vd["display"], "voltage_type": vd["voltage_type"],
                "voltage_confidence": vd["confidence"],
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
                "temp_min": td["temp_min"], "temp_max": td["temp_max"],
                "temperature": td["display"], "temp_grade": td["temp_grade"],
                "temperature_confidence": td["confidence"],
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
    return weights or {"product_type": 20, "density": 35, "voltage": 30, "package": 10, "temperature": 10}


def core_match_filter(spec: dict, xtx_df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """V2 核心规则：产品类型 + 容量 + 电压必须一致，才允许推荐。"""
    warnings = []
    if xtx_df.empty:
        return pd.DataFrame(), ["XTX 产品库为空。"]

    spec_pt = normalize_product_type(spec.get("product_type", ""))
    spec_density = spec.get("density_mb")
    spec_vmin = spec.get("vcc_min")
    spec_vmax = spec.get("vcc_max")

    missing = []
    if not spec_pt:
        missing.append("产品类型")
    if spec_density is None or spec_density == "":
        missing.append("容量")
    if spec_vmin is None or spec_vmax is None:
        missing.append("电压范围")
    if missing:
        return pd.DataFrame(), [f"核心字段缺失：{'、'.join(missing)}。请在人工确认区补充后再推荐。"]

    df = xtx_df.copy()
    df["_norm_product_type"] = df["product_type"].apply(normalize_product_type)
    df["_density_mb"] = df["density_mb"].apply(safe_float)
    df["_vcc_min"] = df["vcc_min"].apply(safe_float)
    df["_vcc_max"] = df["vcc_max"].apply(safe_float)

    mask = (
        df["_norm_product_type"].eq(spec_pt)
        & df["_density_mb"].apply(lambda x: x is not None and abs(float(x) - float(spec_density)) <= 0.001)
        & df["_vcc_min"].apply(lambda x: values_match(x, spec_vmin))
        & df["_vcc_max"].apply(lambda x: values_match(x, spec_vmax))
    )
    matched = df[mask].drop(columns=["_norm_product_type", "_density_mb", "_vcc_min", "_vcc_max"], errors="ignore")
    if matched.empty:
        warnings.append(
            f"没有找到满足核心条件的 XTX 型号：产品类型={spec_pt}，容量={spec_density}Mb，电压={spec_vmin:g}V–{spec_vmax:g}V。"
        )
    return matched, warnings


def calc_secondary_score(spec: dict, xtx_row: pd.Series, weights: dict[str, float]) -> tuple[float, float, str, list[str]]:
    # 核心三项已过滤，这里只做完整度评分和风险提示
    core_score = weights.get("product_type", 20) + weights.get("density", 35) + weights.get("voltage", 30)
    score = float(core_score)
    max_score = float(sum(weights.values())) if weights else 100.0
    risks = []

    competitor_pkg = str(spec.get("package", "")).upper()
    competitor_pkg_size = str(spec.get("package_size", "")).upper()
    xtx_pkg = str(xtx_row.get("package", "")).upper()
    xtx_pkg_size = str(xtx_row.get("package_size", "")).upper()

    competitor_tmin = spec.get("temp_min")
    competitor_tmax = spec.get("temp_max")
    xtx_tmin = safe_float(xtx_row.get("temp_min"))
    xtx_tmax = safe_float(xtx_row.get("temp_max"))

    if weights.get("package", 0) > 0:
        if competitor_pkg and xtx_pkg:
            pkg_same = xtx_pkg in competitor_pkg or competitor_pkg in xtx_pkg
            size_same = bool(xtx_pkg_size and competitor_pkg_size and (xtx_pkg_size in competitor_pkg_size or competitor_pkg_size in xtx_pkg_size))
            if pkg_same and size_same:
                score += weights["package"]
            elif pkg_same:
                score += weights["package"] * 0.75
                risks.append("封装名称一致，具体尺寸需确认")
            else:
                same_family = any(key in competitor_pkg and key in xtx_pkg for key in ["WSON", "SOP", "SOIC", "BGA", "WLCSP", "USON", "DFN"])
                if same_family:
                    score += weights["package"] * 0.5
                    risks.append("封装大类接近，具体尺寸和 Pin 脚需确认")
                else:
                    risks.append("封装可能不兼容")
        else:
            risks.append("封装未识别，需人工确认")

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
    else:
        level = "低"
    return round(score, 2), round(match_percent * 100, 1), level, risks


def recommend_xtx(spec: dict, xtx_df: pd.DataFrame, weights_df: pd.DataFrame, top_n: int = 8) -> tuple[pd.DataFrame, list[str]]:
    if xtx_df.empty:
        return pd.DataFrame(), ["XTX 产品库为空或未上传。"]
    weights = get_enabled_weights(weights_df)
    filtered_df, warnings = core_match_filter(spec, xtx_df)
    if filtered_df.empty:
        return pd.DataFrame(), warnings

    rows = []
    for _, row in filtered_df.iterrows():
        score, pct, level, risks = calc_secondary_score(spec, row, weights)
        capacity_display = row.get("capacity_display", "")
        if not capacity_display:
            capacity_display = f"{row.get('density_mb', '')}Mb"
        voltage_display = row.get("voltage_range", "")
        if not voltage_display:
            voltage_display = f"{row.get('vcc_min', '')}V–{row.get('vcc_max', '')}V"
        temp_display = row.get("temperature", "")
        if not temp_display:
            temp_display = f"{row.get('temp_min', '')}°C~{row.get('temp_max', '')}°C"

        row_out = {
            "推荐XTX型号": row.get("xtx_model", ""),
            "产品类型": row.get("product_type", ""),
            "容量": capacity_display,
            "电压类型": row.get("voltage_type", ""),
            "电压范围": voltage_display,
            "封装": row.get("package", ""),
            "封装尺寸": row.get("package_size", ""),
            "温度": temp_display,
            "频率MHz": row.get("frequency_mhz", ""),
            "状态": row.get("status", ""),
            "包装": row.get("packing", ""),
            "最小包装量": row.get("min_package_qty", ""),
            "最小订货量": row.get("moq", ""),
            "箱数量": row.get("carton_qty", ""),
            "内盒包装量": row.get("inner_box_qty", ""),
            "匹配分": score,
            "匹配率": f"{pct}%",
            "匹配度": level,
            "风险点": "；".join(risks) if risks else "核心字段完全匹配，未发现明显风险",
            "备注": row.get("note", ""),
        }
        # 兼容未来产品库新增字段，常用字段已在上面展示，避免丢失重要信息。
        for extra_col in ["original_product_category", "original_capacity", "original_voltage", "original_temperature"]:
            if extra_col in row.index and extra_col not in row_out:
                row_out[extra_col] = row.get(extra_col, "")
        rows.append(row_out)
    result = pd.DataFrame(rows).sort_values(["匹配分", "推荐XTX型号"], ascending=[False, True])
    return result.head(top_n), warnings


# =========================================================
# 页面 UI
# =========================================================

st.title("🔎 竞品规格书解析与 XTX 对标推荐工具 V4")
st.caption("V4：产品页优先匹配；无 PDF 时直接解析官网产品表格；PDF 只校验少量高分候选，显著提升检索速度。")

with st.sidebar:
    st.header("维护数据上传")
    maintenance_file = st.file_uploader(
        "上传维护数据库 XLSX",
        type=["xlsx"],
        help="建议使用模板，包含 Company_Master、XTX_Product_Library、Match_Weights、History_Log。V2 可选新增 Company_Master.model_prefixes。",
    )
    history_file = st.file_uploader(
        "上传历史记录 XLSX（可选）",
        type=["xlsx"],
        help="如果不上传，程序会尝试读取本地 competitor_analysis_history.xlsx。",
    )
    max_pages = st.slider("每家官网最大检索页数", min_value=1, max_value=10, value=DEFAULT_MAX_CRAWL_PAGES, step=1, help="建议 2~4。数值越大越慢。")
    pdf_validate_limit = st.slider("最多校验 PDF 候选数", min_value=1, max_value=8, value=3, step=1, help="自动检索时最多下载并全文扫描几个高分 PDF。建议 2~3。")
    pdf_scan_pages_input = st.slider("PDF型号快速扫描页数", min_value=0, max_value=300, value=0, step=20, help="0 表示扫描全文；非 0 表示只扫描前 N 页。型号可能靠后时用 0，但会慢一些。")
    pdf_parse_pages = st.slider("PDF详细解析页数", min_value=5, max_value=80, value=20, step=5, help="只影响选中 PDF 后的详细字段抽取。值越大越慢，一般 20~30 页足够。")
    product_page_first = st.checkbox("产品页优先，命中后跳过 PDF 解析", value=True, help="推荐打开。官网产品表格已包含容量/电压/封装/温度时，直接用产品页数据。")
    save_local_history = st.checkbox("保存历史到本地 xlsx", value=True, help="适合公司内网/本地部署；Streamlit Cloud 重启后本地文件可能丢失。")

    st.divider()
    st.markdown("**公司区分规则**")
    st.caption("Company_Master.company_role = XTX 表示芯天下；Competitor 表示友商。自动检索只使用 Competitor。")

maintenance_bytes = maintenance_file.getvalue() if maintenance_file is not None else None
maintenance = read_maintenance_xlsx(maintenance_bytes)
company_df = ensure_company_columns(maintenance["Company_Master"])
xtx_df = ensure_xtx_columns(maintenance["XTX_Product_Library"])
weights_df = clean_columns(maintenance["Match_Weights"])
history_df = read_history_xlsx(history_file)

if "enabled" in company_df.columns:
    company_df = company_df[normalize_bool_series(company_df["enabled"])]
if "enabled" in weights_df.columns:
    weights_df = weights_df[normalize_bool_series(weights_df["enabled"])]

competitor_company_df = company_df[company_df["company_role"].astype(str).str.upper().eq("COMPETITOR")].copy()
xtx_company_df = company_df[company_df["company_role"].astype(str).str.upper().eq("XTX")].copy()

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

V2 建议在 `Company_Master` 增加 `model_prefixes` 字段，例如 Boya 填 `BY,BY25,BY26`，这样自动检索时会先按型号前缀锁定厂商，速度更快，也能避免误抓其他公司的 PDF。
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
        "下载维护数据库 XLSX 模板 V3",
        data=template_bytes,
        file_name="xtx_competitor_maintenance_template_v3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.subheader("1. 输入竞品信息")
col1, col2, col3 = st.columns(3)
with col1:
    competitor_options = ["自动识别/全部友商"] + sorted(competitor_company_df["company_name"].dropna().astype(str).unique().tolist())
    selected_company = st.selectbox("竞品厂商", competitor_options)
with col2:
    competitor_model = st.text_input("竞品型号", placeholder="例如：BY25D20AS / W25Q512JV / MX25L51245G")
with col3:
    product_type = st.selectbox("产品类型", PRODUCT_TYPE_OPTIONS, index=0)

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
    search_note = ""

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
            st.write("下载并校验 PDF...")
            try:
                pdf_bytes = download_pdf(pdf_url_input)
                found, page_no, method = scan_pdf_for_model_fast(pdf_bytes, competitor_model, max_pages=None)
                if not found:
                    st.warning("该 PDF 全文未快速定位到输入型号，请确认链接是否正确；仍会继续解析。")
                else:
                    st.write(f"PDF 型号校验通过：第 {page_no} 页，{method}")
                selected_pdf_url = pdf_url_input
            except Exception as e:
                st.error(f"PDF 下载失败：{e}")
                st.stop()

        else:
            if competitor_company_df.empty:
                st.error("友商官网白名单为空，请检查 Company_Master 中 company_role=Competitor 的记录。")
                st.stop()

            search_df, search_note = infer_company_df_by_model(competitor_company_df, competitor_model, selected_company)
            st.write(search_note)
            st.write(f"本次检索厂商数：{len(search_df)}")

            # V4：先查官网产品页。产品页表格命中后，直接用产品页数据，避免下载/扫描 PDF 导致等待过久。
            st.write("产品页优先检索目标型号...")
            product_page_df = crawl_product_pages_parallel(search_df, competitor_model, max_pages=max_pages)

            if product_page_first and not product_page_df.empty:
                best_page = product_page_df.iloc[0]
                spec = spec_from_product_page_row(best_page, selected_type=product_type, model=competitor_model)
                selected_pdf_url = str(best_page.get("datasheet_url", "") or best_page.get("page_url", ""))
                candidate_df = product_page_df
                st.write("已在官网产品页匹配到目标型号，跳过 PDF 解析。")
                st.write(str(best_page.get("page_url", "")))
            else:
                st.write("未在产品页命中，开始检索官网 PDF 候选链接...")
                candidate_df = crawl_companies_parallel(search_df, competitor_model, max_pages=max_pages)

                if candidate_df.empty:
                    if not product_page_df.empty:
                        best_page = product_page_df.iloc[0]
                        spec = spec_from_product_page_row(best_page, selected_type=product_type, model=competitor_model)
                        selected_pdf_url = str(best_page.get("datasheet_url", "") or best_page.get("page_url", ""))
                        candidate_df = product_page_df
                        st.warning("未找到 PDF，已改用官网产品页数据。")
                    else:
                        st.warning("未自动找到候选 PDF，也未在产品页匹配到型号。建议切换到“输入 PDF 链接”或“手动上传 PDF”。")
                        st.stop()
                else:
                    st.write("下载少量高分 PDF 并快速校验是否包含目标型号...")
                    scan_pages = None if int(pdf_scan_pages_input) == 0 else int(pdf_scan_pages_input)
                    pdf_bytes, selected_pdf_url, checked_df, err = choose_valid_pdf(
                        candidate_df, competitor_model, max_scan_pages=scan_pages, max_candidates=pdf_validate_limit
                    )
                    candidate_df = checked_df if not checked_df.empty else candidate_df
                    if err:
                        if not product_page_df.empty:
                            best_page = product_page_df.iloc[0]
                            spec = spec_from_product_page_row(best_page, selected_type=product_type, model=competitor_model)
                            selected_pdf_url = str(best_page.get("datasheet_url", "") or best_page.get("page_url", ""))
                            candidate_df = product_page_df
                            st.warning("PDF 未通过校验，已改用官网产品页数据。")
                        else:
                            st.error(err)
                            st.dataframe(candidate_df, use_container_width=True)
                            st.stop()
                    else:
                        st.write("已选择通过校验的 PDF：")
                        st.write(selected_pdf_url)
                        st.write("解析 PDF 文本...")
                        try:
                            pdf_text = extract_text_from_pdf_bytes(pdf_bytes, max_pages=pdf_parse_pages)
                        except Exception as e:
                            st.error(f"PDF 解析失败：{e}")
                            st.stop()
                        if not pdf_text.strip():
                            st.error("PDF 未提取到有效文本，可能是扫描版 PDF，需要后续增加 OCR。")
                            st.stop()
                        st.write("抽取产品类型、容量、电压、封装、温度...")
                        spec = analyze_spec_text(pdf_text, selected_type=product_type, model=competitor_model)

        if source_mode in ["手动上传 PDF", "输入 PDF 链接"]:
            st.write("解析 PDF 文本...")
            try:
                pdf_text = extract_text_from_pdf_bytes(pdf_bytes, max_pages=pdf_parse_pages)
            except Exception as e:
                st.error(f"PDF 解析失败：{e}")
                st.stop()

            if not pdf_text.strip():
                st.error("PDF 未提取到有效文本，可能是扫描版 PDF，需要后续增加 OCR。")
                st.stop()

            st.write("抽取产品类型、容量、电压、封装、温度...")
            spec = analyze_spec_text(pdf_text, selected_type=product_type, model=competitor_model)

        st.session_state["last_result"] = {
            "spec": spec,
            "candidate_df": candidate_df,
            "selected_pdf_url": selected_pdf_url,
            "competitor_model": competitor_model,
            "selected_company": selected_company,
            "product_type": product_type,
            "review_df": spec_to_review_df(spec),
            "search_note": search_note,
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

    st.divider()
    st.subheader("3. 自动解析结果")
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        st.metric("产品类型", spec.get("product_type", ""))
        st.caption(f"置信度：{spec.get('product_type_confidence', '')}")
    with m2:
        st.metric("容量", spec.get("capacity", ""))
        st.caption(f"置信度：{spec.get('capacity_confidence', '')}")
    with m3:
        st.metric("电压范围", spec.get("voltage_range", ""))
        st.caption(f"{spec.get('voltage_type', '')}，置信度：{spec.get('voltage_confidence', '')}")
    with m4:
        st.metric("封装形式", spec.get("package", ""))
        st.caption(f"尺寸：{spec.get('package_size', '')}，置信度：{spec.get('package_confidence', '')}")
    with m5:
        st.metric("温度范围", spec.get("temperature", ""))
        st.caption(f"{spec.get('temp_grade', '')}，置信度：{spec.get('temperature_confidence', '')}")

    if result.get("search_note"):
        st.info(result["search_note"])

    st.markdown("**规格书 / 产品页来源：**")
    if str(selected_pdf_url).startswith("http"):
        st.markdown(f"[打开规格书 PDF]({selected_pdf_url})")
    else:
        st.write(selected_pdf_url)

    if not candidate_df.empty:
        with st.expander("查看自动检索到的产品页 / PDF 候选链接和校验结果"):
            st.dataframe(candidate_df, use_container_width=True)

    st.divider()
    st.subheader("4. 人工确认 / 修正")
    st.caption("人工修改后，下方 XTX 对标型号推荐会自动刷新。核心字段：产品类型、容量、电压范围必须准确。")

    review_df = st.data_editor(
        result["review_df"],
        use_container_width=True,
        hide_index=True,
        column_config={
            "manual_value": st.column_config.TextColumn(
                "manual_value",
                help="例如：产品类型填 SPI NOR；容量填 512Mb；电压填 2.7V-3.6V；温度填 -40~85。",
            ),
            "confirm_status": st.column_config.SelectboxColumn(
                "confirm_status",
                options=["未确认", "已确认", "需FAE确认", "不适用"],
                required=True,
            ),
        },
        disabled=["field_key", "field_name", "extracted_value", "confidence"],
        key="review_editor_v2",
    )

    confirmed_spec, confirm_map = build_confirmed_spec(spec, review_df)

    with st.expander("查看用于推荐的最终规格", expanded=False):
        st.json({
            "product_type": confirmed_spec.get("product_type"),
            "capacity": confirmed_spec.get("capacity"),
            "density_mb": confirmed_spec.get("density_mb"),
            "voltage_range": confirmed_spec.get("voltage_range"),
            "vcc_min": confirmed_spec.get("vcc_min"),
            "vcc_max": confirmed_spec.get("vcc_max"),
            "package": confirmed_spec.get("package"),
            "package_size": confirmed_spec.get("package_size"),
            "temperature": confirmed_spec.get("temperature"),
        })

    st.divider()
    st.subheader("5. XTX 对标型号推荐")
    st.caption("V3 推荐规则：只有产品类型、容量、电压范围三项核心条件全部一致，才会进入推荐列表；不再把 SPI NAND 或其他容量/电压型号混入推荐。")
    recommend_df, rec_warnings = recommend_xtx(confirmed_spec, xtx_df, weights_df, top_n=10)
    for w in rec_warnings:
        st.warning(w)

    if recommend_df.empty:
        st.info("当前没有满足核心条件的 XTX 对标型号。请检查：产品类型、容量、电压范围是否识别/填写正确，或补充 XTX 产品库。")
    else:
        st.dataframe(recommend_df, use_container_width=True, height=360)
        best = recommend_df.iloc[0]
        st.success(
            f"优先推荐对标型号：{best['推荐XTX型号']}；"
            f"匹配度：{best['匹配度']}；匹配分：{best['匹配分']}；匹配率：{best['匹配率']}。"
        )

    operator_note = st.text_area("本次分析备注", placeholder="例如：封装后缀需确认；客户要求 QE=1；需确认是否车规等级。")

    best_dict = recommend_df.iloc[0].to_dict() if not recommend_df.empty else {}
    new_row = {
        "run_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "competitor_company": selected_company_result,
        "competitor_model": competitor_model_result,
        "product_type": confirmed_spec.get("product_type"),
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
        "recommended_xtx_model": best_dict.get("推荐XTX型号", ""),
        "match_score": best_dict.get("匹配分", ""),
        "match_percent": best_dict.get("匹配率", ""),
        "match_level": best_dict.get("匹配度", ""),
        "risk_points": best_dict.get("风险点", ""),
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
    file_name="xtx_competitor_maintenance_current_v3.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("使用说明 / 限制"):
    st.markdown(
        """
### V4 关键变化

1. **自动检索速度优化**  
   - 根据 `Company_Master.model_prefixes` 或内置型号前缀规则，先锁定可能的友商。  
   - 多家公司并发检索。  
   - 每家公司默认只检索 8 页，可在侧边栏调整。

2. **避免错抓规格书**  
   - 候选 PDF 必须来自对应友商的白名单域名。  
   - 下载 PDF 后，会优先解析官网产品页；只有需要使用 PDF 时才下载少量高分候选并扫描型号；不包含则跳过。  
   - 因此不会再把 BY25D20AS 匹配到其他公司的 GT25Q32A PDF。

3. **推荐规则更严格**  
   - 产品类型、容量、电压范围三项必须一致，才进入推荐列表。  
   - 封装和温度用于风险提示和排序，不再让其他容量、其他电压或 SPI NAND 混入。

4. **人工确认后自动刷新推荐**  
   - 在 `manual_value` 填入修正值后，页面会自动 rerun，下方推荐结果即时更新。  
   - 产品类型可以填：SPI NOR、SPI NAND、PPI NOR、PPI NAND、eMMC、DDR3、DDR4 等。

### 维护文件建议

`Company_Master` 建议字段：

- `company_name`：公司名称，例如 Boya。  
- `company_role`：`XTX` 或 `Competitor`。  
- `domain`：官网域名，例如 boyamicro.com。  
- `base_url`：官网首页。  
- `search_url_template`：官网搜索地址，可为空；如支持可写 `https://xxx.com/search?keyword={model}`。  
- `model_prefixes`：型号前缀，例如 Boya 填 `BY,BY25,BY26`。  
- `enabled`：1 启用，0 关闭。

### 当前限制

- 部分官网使用动态 JS 或下载权限限制时，自动检索仍可能失败。保留“输入 PDF 链接”和“手动上传 PDF”作为兜底。  
- 扫描版 PDF 暂不支持 OCR。  
- 封装可能会列出多个候选，最终仍建议人工确认具体订购后缀。  
"""
    )
