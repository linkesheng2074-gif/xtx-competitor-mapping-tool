"""
XTX 竞品规格书解析与对标推荐工具 V8
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
import json
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
    page_title="竞品规格书解析与 XTX 对标推荐工具 V8",
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


# V6 内置当前维护库：避免在 Streamlit Cloud 未上传 XLSX 时，只使用很小的示例产品库，导致 2Mbit 等型号无法推荐。
# 如果左侧上传维护数据库 XLSX，则以上传文件为准。
EMBEDDED_COMPANY_MASTER_JSON = r'[{"company_name": "XTX", "company_role": "XTX", "domain": "xtxtech.com", "base_url": "https://www.xtxtech.com", "search_url_template": null, "model_prefixes": "XT,XT25,XT26,XT27,XT28", "enabled": 1, "note": "我司官网，只展示，不参与竞品 PDF 检索"}, {"company_name": "Winbond", "company_role": "Competitor", "domain": "winbond.com", "base_url": "https://www.winbond.com", "search_url_template": null, "model_prefixes": "W,W25,W25Q,W25N,W29,W39", "enabled": 1, "note": "友商官网"}, {"company_name": "Macronix", "company_role": "Competitor", "domain": "macronix.com", "base_url": "https://www.macronix.com", "search_url_template": null, "model_prefixes": "MX,MX25,MX30,MX35,MX66,MX29", "enabled": 1, "note": "友商官网"}, {"company_name": "GigaDevice", "company_role": "Competitor", "domain": "gigadevice.com", "base_url": "https://www.gigadevice.com.cn", "search_url_template": null, "model_prefixes": "GD,GD25,GD5F,GD55,GD32", "enabled": 1, "note": "友商官网"}, {"company_name": "ISSI", "company_role": "Competitor", "domain": "issi.com", "base_url": "https://www.issi.com", "search_url_template": null, "model_prefixes": "IS,IS25,IS26,IS29", "enabled": 1, "note": "友商官网"}, {"company_name": "Puya", "company_role": "Competitor", "domain": "puyasemi.com", "base_url": "https://www.puyasemi.com", "search_url_template": null, "model_prefixes": "P25,PY,PY25,PY26", "enabled": 1, "note": "友商官网"}, {"company_name": "Boya", "company_role": "Competitor", "domain": "boyamicro.com", "base_url": "https://www.boyamicro.com", "search_url_template": "https://www.boyamicro.com/?zh/products/2", "model_prefixes": "BY,BY25,BY26,BY27", "enabled": 1, "note": "友商官网"}, {"company_name": "Zbit", "company_role": "Competitor", "domain": "zbitsemi.com", "base_url": "https://www.zbitsemi.com", "search_url_template": null, "model_prefixes": "ZB,ZB25,ZB26", "enabled": 1, "note": "友商官网"}, {"company_name": "XMC", "company_role": "Competitor", "domain": "xmcwh.com", "base_url": "https://www.xmcwh.com", "search_url_template": null, "model_prefixes": "XM,XM25,XM26,XM29", "enabled": 1, "note": "友商官网"}, {"company_name": "tsingtengms", "company_role": "Competitor", "domain": "tsingtengms.com", "base_url": "https://www.tsingtengms.com", "search_url_template": null, "model_prefixes": "GT,GT25,GT26", "enabled": 1, "note": "友商官网"}, {"company_name": "DS", "company_role": "Competitor", "domain": "dosilicon.com", "base_url": "https://www.dosilicon.com", "search_url_template": null, "model_prefixes": "DS,DS25,DS26,FM25", "enabled": 1, "note": "友商官网"}, {"company_name": "ESMT", "company_role": "Competitor", "domain": "esmt.com.tw", "base_url": "https://www.esmt.com.tw", "search_url_template": null, "model_prefixes": "F,F25,F59", "enabled": 1, "note": "友商官网"}, {"company_name": "giantec-semi", "company_role": "Competitor", "domain": "giantec-semi.com", "base_url": "https://www.giantec-semi.com", "search_url_template": null, "model_prefixes": "GT,GT25,GT26", "enabled": 1, "note": "友商官网"}, {"company_name": "Micron", "company_role": "Competitor", "domain": "micron.cn", "base_url": "https://www.micron.cn", "search_url_template": null, "model_prefixes": "MT,MT25,N25,MT29", "enabled": 1, "note": "友商官网"}, {"company_name": "longsys", "company_role": "Competitor", "domain": "longsys.com", "base_url": "https://cn.longsys.com", "search_url_template": null, "model_prefixes": "FORESEE,FS,F25,LGS", "enabled": 1, "note": "友商官网"}, {"company_name": "infineon", "company_role": "Competitor", "domain": "infineon.com", "base_url": "https://www.infineon.com", "search_url_template": null, "model_prefixes": "S25,S26,S29,SL", "enabled": 1, "note": "友商官网"}, {"company_name": "Kioxia", "company_role": "Competitor", "domain": "kioxia.com", "base_url": "https://www.kioxia.com", "search_url_template": null, "model_prefixes": "TC58,TH58", "enabled": 1, "note": "友商官网"}, {"company_name": "Samsung", "company_role": "Competitor", "domain": "semiconductor.samsung.com", "base_url": "https://semiconductor.samsung.com", "search_url_template": null, "model_prefixes": "K9,KLM,KLMBG,KLMAG", "enabled": 1, "note": "友商官网"}, {"company_name": "SK hynix", "company_role": "Competitor", "domain": "skhynix.com", "base_url": "https://www.skhynix.com", "search_url_template": null, "model_prefixes": "H25,H26,H27", "enabled": 1, "note": "友商官网"}]'
EMBEDDED_XTX_PRODUCT_LIBRARY_JSON = r'[{"xtx_model": "XT25Q512FWLIGT", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WLCSP", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "512Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27Q04EBSIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W32ASSIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "32Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W32ASOIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "32Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BQ128FSSIGT", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "Tape & Reel", "note": "128Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W64ASSIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "Tape & Reel", "note": "64Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W64ASOIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "64Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W02ADUIGT", "product_type": "SPI NOR", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "DFN6", "package_size": "1.2x1.2x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "2Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04AFAIGT", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "FODFN6", "package_size": "1.2x1.2x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04AFBIGT", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "FODFN8", "package_size": "1.5x0.75x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W64AWOIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "64Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W64AFTIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "FODFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "64Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W64ASOIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "64Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q18DWSIGA-B", "product_type": "SPI NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "8Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "8Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSIGA-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q128FFEIGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "FODFN8", "package_size": "3x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 24000.0, "inner_box_qty": 3000.0, "packing": "Tape & Reel", "note": "128Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W64ASSIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "64Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W64A-WLHA", "product_type": "NOR Wafer", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": NaN, "moq": NaN, "carton_qty": NaN, "inner_box_qty": NaN, "packing": "Cassette", "note": "64Mbit 1.65~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "64Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT26G11MWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W02AFUIGT", "product_type": "SPI NOR", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "FODFN6", "package_size": "1.2x0.7x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 5000.0, "moq": 5000.0, "carton_qty": 200000.0, "inner_box_qty": 50000.0, "packing": "Tape & Reel", "note": "2Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04ASOIGT", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W02ASOIGT", "product_type": "SPI NOR", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "2Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W02ATSIGT", "product_type": "SPI NOR", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "TSSOP8", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "2Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04ATSIGT", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "TSSOP8", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04AFCIGT", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "FODFN8", "package_size": "1.5x1.5x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W02ADTIGT", "product_type": "SPI NOR", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "2Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT28EG64GA2YUECGA", "product_type": "eMMC", "density_mb": 524288.0, "capacity_display": "64GByte", "voltage_type": "其他", "vcc_min": 1.8, "vcc_max": 3.3, "voltage_range": "1.8V/3.3V", "package": "BGA153", "package_size": "11.5x13x1.0mm", "temp_min": -25.0, "temp_max": 85.0, "temperature": "-25°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 1520.0, "moq": 1520.0, "carton_qty": 9120.0, "inner_box_qty": 1520.0, "packing": "Tray", "note": "64GByte 1.8V/3.3V eMMC", "original_product_category": "eMMC", "original_capacity": "64GByte", "original_voltage": "1.8V/3.3V", "original_temperature": "-25°C-+85°C"}, {"xtx_model": "26G01FWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT63M4G8E6MU-CDBIA", "product_type": "MCP", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "多电源", "vcc_min": 0.6, "vcc_max": 1.8, "voltage_range": "1.8/1.8/1.1/0.6", "package": "BGA149", "package_size": "8x9.5x0.8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 8400.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "4G4GD4X 1.8/1.8/1.1/0.6 NAND MCP", "original_product_category": "NAND MCP", "original_capacity": "4G4GD4X", "original_voltage": "1.8/1.8/1.1/0.6", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q256FWOIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "256Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT52L08G16ABEWGA", "product_type": "SDRAM", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "其他", "vcc_min": 1.2, "vcc_max": 2.5, "voltage_range": "1.2/2.5V", "package": "BGA96", "package_size": "13x7.5mm", "temp_min": -95.0, "temp_max": 0.0, "temperature": "0-95°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 209.0, "moq": 209.0, "carton_qty": 12540.0, "inner_box_qty": 2090.0, "packing": "Tray", "note": "8Gbit 1.2/2.5V SDRAM", "original_product_category": "SDRAM", "original_capacity": "8Gbit", "original_voltage": "1.2/2.5V", "original_temperature": "0-95℃"}, {"xtx_model": "25Q512FSFHGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XTD25W04ADTIGT", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G14DWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q256FBGIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04ASSIGA", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W02ASOIGU", "product_type": "SPI NOR", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "2Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04ASOIGU", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "4Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64FDXIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "DFN8", "package_size": "4x3x0.55mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "64Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26G08DWSIGA", "product_type": "SPI NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "8Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "8Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "27Q08ABSHGA", "product_type": "PPI NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": NaN, "moq": NaN, "carton_qty": NaN, "inner_box_qty": NaN, "packing": "Tray", "note": "8Gbit 1.7~1.9V PPI NAND_KA", "original_product_category": "PPI NAND_KA", "original_capacity": "8Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25W32FDTIGT", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 120000.0, "packing": "Tape & Reel", "note": "32Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25Q64FSSIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "64Mbit 1.65~2.0V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W32FSOIGT-01", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "32Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BQ256FWOIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "256Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25W32FSOIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 60000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "32Mbit 1.65~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q32FSOIGT", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "32Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTD25W04A-WLHA", "product_type": "NOR Wafer", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 149917.0, "moq": 149917.0, "carton_qty": 3747925.0, "inner_box_qty": 149917.0, "packing": "Cassette", "note": "4Mbit 1.65~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "4Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XTD25W02A-WLHA", "product_type": "NOR Wafer", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 221537.0, "moq": 221537.0, "carton_qty": 5538425.0, "inner_box_qty": 221537.0, "packing": "Cassette", "note": "2Mbit 1.65~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XTD25W02A-WLIA", "product_type": "NOR Wafer", "density_mb": 2.0, "capacity_display": "2Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 221537.0, "moq": 221537.0, "carton_qty": 5538425.0, "inner_box_qty": 221537.0, "packing": "Cassette", "note": "2Mbit 1.65~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "2Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSIGT-W01", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "Tape & Reel", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSOIGT-01", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOIGT-01", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F08FSOIGT-01", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "8Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25F256BWSIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BQ512FWSIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BQ256FWSIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "256Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W32FSOIGT-00", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "32Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSSIGA", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W512BWSIGT", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "512Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W512BWSIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BF256BSFIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 1500.0, "moq": 1500.0, "carton_qty": 15000.0, "inner_box_qty": 3000.0, "packing": "Tape & Reel", "note": "256Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27Q04A8BFIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA67", "package_size": "8x6.5x0.89mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3240.0, "moq": 3240.0, "carton_qty": 19440.0, "inner_box_qty": 3240.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27Q02A8BFIGA", "product_type": "PPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA67", "package_size": "8x6.5x0.89mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3240.0, "moq": 3240.0, "carton_qty": 19440.0, "inner_box_qty": 3240.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSHGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "Tape & Reel", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25F128FFEIGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "FODFN8", "package_size": "3x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 24000.0, "inner_box_qty": 3000.0, "packing": "Tape & Reel", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "H7A5CM25GCIX", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "H7A5CM2AGAIX", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "Tape & Reel", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G08DWSIGA", "product_type": "SPI NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "8Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "8Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q18DWSIGA", "product_type": "SPI NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "8Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "8Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32GSSIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT29F512ATSHGA", "product_type": "PPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "TSOP56", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 960.0, "moq": 960.0, "carton_qty": 5760.0, "inner_box_qty": 960.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V PPI NOR", "original_product_category": "PPI NOR", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25F64GSSIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64GSOIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q14DWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT55Q2GFBGIGA", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT55Q1GFWSIGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26Q04DWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q12DWSIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G01FWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F08FSSIGU", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "8Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G02DLAIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "LGA8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G01DLAIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "LGA8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F16FSSIGT-S", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "16Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q01DWSIGT", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "1Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "27Q01ABSIGA", "product_type": "PPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "1Gbit 1.7~1.9V PPI NAND_KA", "original_product_category": "PPI NAND_KA", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOIGT-00", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSOIGT-00", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q512F-WSHA", "product_type": "NOR Wafer", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 1951.0, "moq": 1951.0, "carton_qty": 48775.0, "inner_box_qty": 48775.0, "packing": "Cassette", "note": "512Mbit 1.7~2.0V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25Q08FSOIGU", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "8Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q08FDTIGT", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "8Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q256FSFIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q256F-WSIA", "product_type": "NOR Wafer", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4413.0, "moq": 4413.0, "carton_qty": 110325.0, "inner_box_qty": 110325.0, "packing": "Cassette", "note": "256Mbit 1.65~2.0V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q256FWOIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "256Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q256FWSIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "256Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q16FDTIGT", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "16Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q16FSOIGU", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "16Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64FSSIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "64Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F16FSSIGU-S", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "16Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q04DWSIGT", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "4Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G12DWSIGA-B", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F08FSSIGT", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "8Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27Q08ABSIGA", "product_type": "PPI NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "8Gbit 1.7~1.9V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "8Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q04DBGIGA-B", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25F16FDTHGT", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "SSS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "16Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "16Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT26G04CWSIGA-B", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G04DWSIGA-B", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G01DBGIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G11DWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G02DBGIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64FDTIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "FODFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "64Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64FSOIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "64Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDQ02GLAIGA", "product_type": "SD NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "LGA8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V SD NAND", "original_product_category": "SD NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDQ01GLAIGA", "product_type": "SD NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "LGA8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "1Gbit 1.7~1.9V SD NAND", "original_product_category": "SD NAND", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G12DWSIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q02DBGIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q02DLAIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "LGA8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q04DBGIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FWOIGT-S", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64FSOIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "64Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q128FSSIGU-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "128Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G01CBGIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q04DWSIGT-B", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "4Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q02DWSIGT-B", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "2Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q01DWSIGT-B", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "Tape & Reel", "note": "1Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q01DLAIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "LGA8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "1Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q02DWSIGA-B", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q01DWSIGA-B", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G04DWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G02DWSIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT28EG64GA2TNECGA", "product_type": "eMMC", "density_mb": 524288.0, "capacity_display": "64GByte", "voltage_type": "其他", "vcc_min": 1.8, "vcc_max": 3.3, "voltage_range": "1.8V/3.3V", "package": "BGA153", "package_size": "11.5x13x1.0mm", "temp_min": -25.0, "temp_max": 85.0, "temperature": "-25°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1520.0, "moq": 1520.0, "carton_qty": 9120.0, "inner_box_qty": 1520.0, "packing": "Tray", "note": "64GByte 1.8V/3.3V eMMC", "original_product_category": "eMMC", "original_capacity": "64GByte", "original_voltage": "1.8V/3.3V", "original_temperature": "-25°C-+85°C"}, {"xtx_model": "XT25F128F-WSHA", "product_type": "NOR Wafer", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9108.0, "moq": 9108.0, "carton_qty": 227700.0, "inner_box_qty": 227700.0, "packing": "Cassette", "note": "128Mbit 2.7~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25F16FDTIGT-S", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 3000.0, "packing": "T&R", "note": "16Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F16FSOIGT-S", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "16Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F16FSOIGU-S", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "16Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W08FSOHGU", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "8Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25F08FSOHGU", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "8Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25F08FSOIGT", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 3000.0, "packing": "T&R", "note": "8Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDQ08GWSIGA", "product_type": "SD NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "8Gbit 1.7~1.9V SD NAND", "original_product_category": "SD NAND", "original_capacity": "8Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDG08GWSIGA", "product_type": "SD NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "8Gbit 2.7~3.6V SD NAND", "original_product_category": "SD NAND", "original_capacity": "8Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDG04GWSIGA", "product_type": "SD NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SD NAND", "original_product_category": "SD NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDG02GWSIGA", "product_type": "SD NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SD NAND", "original_product_category": "SD NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSSIGT-S", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSSIGU-S", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSSIGT-S", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSSIGU-S", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q128FWOIGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "128Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q128FSSIGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "128Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSOIGT-S", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOIGT-S", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOIGU-S", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSOIGU-S", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q02DWSIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q04DWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q512FWSIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W1GBSFIGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "1Gbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDQ04GWSIGA", "product_type": "SD NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SD NAND", "original_product_category": "SD NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDQ02GWSIGA", "product_type": "SD NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~1.9V SD NAND", "original_product_category": "SD NAND", "original_capacity": "2Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F16F-WSHA", "product_type": "NOR Wafer", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 44024.0, "moq": 44024.0, "carton_qty": 1100600.0, "inner_box_qty": 1100600.0, "packing": "Cassette", "note": "16Mbit 2.3~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "16Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "25Q128FSSIGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 1.65~2.0V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "128Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25Q512FBGIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25Q512FWSIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "55Q1GFSFHGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.65~2.0V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "1Gbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "25Q512FSFIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSSIGA", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BF128FSSIGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64FSSIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "64Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64FWOIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "64Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT55Q1GFSFIGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q512FBGIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q512FSFIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BF128FSSIGU-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOHGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "25F128FSFHGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "25F128FSSHGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "27G14ABSIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V PPI NAND_KA", "original_product_category": "PPI NAND_KA", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "27Q14ABSIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V PPI NAND_KA", "original_product_category": "PPI NAND_KA", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSHGU-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "25W512BSFIGA-D", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.65~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "512Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BF256BWSIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FWOIGA", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4550.0, "moq": 4550.0, "carton_qty": 27300.0, "inner_box_qty": 4550.0, "packing": "Tray", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FDTIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDQ01GWSIGA", "product_type": "SD NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~1.9V SD NAND", "original_product_category": "SD NAND", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XTSDG01GWSIGA", "product_type": "SD NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SD NAND", "original_product_category": "SD NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSIGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F256BWSIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25F128FWOHGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "25F128FWOIGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25F128FSSIGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSIGU-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G11CWSIGA-B", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSSIGT", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G01DWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26G04CWSIGA-G", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26G02CWSIGA-G", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOIGA", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "27Q08ABSIGA", "product_type": "PPI NAND", "density_mb": 8192.0, "capacity_display": "8Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "8Gbit 1.7~1.9V PPI NAND_KA", "original_product_category": "PPI NAND_KA", "original_capacity": "8Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W08F-WSHA", "product_type": "NOR Wafer", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 61230.0, "moq": 61230.0, "carton_qty": 1530750.0, "inner_box_qty": 1530750.0, "packing": "Cassette", "note": "8Mbit 1.65~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "8Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25F08FSOIGU", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "8Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F08FDTIGT", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "8Mbit 2.3~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27G04A8BFIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA67", "package_size": "8x6.5x0.89mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3240.0, "moq": 3240.0, "carton_qty": 19440.0, "inner_box_qty": 3240.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25W1GBSFIGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.65~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "1Gbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25W512BSFIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.65~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "512Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W512BSFIGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "512Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25F256BSFIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1760.0, "moq": 1760.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q64F-WSIA", "product_type": "NOR Wafer", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 16985.0, "moq": 16985.0, "carton_qty": 424625.0, "inner_box_qty": 424625.0, "packing": "Cassette", "note": "64Mbit 1.65~2.0V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "64Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSIGT", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSOIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26G04CWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26G02CWSIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26G00CWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25BF256BWSIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "256Mbit 2.7~3.6V SPI NOR_KA", "original_product_category": "SPI NOR_KA", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q01DWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26Q04CWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "26G01CWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND_KA", "original_product_category": "SPI NAND_KA", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G01CWSIGT", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G11CWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSSIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W32FSOIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "32Mbit 1.65~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOIGT", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT61M2G8C2TM-B8BEA", "product_type": "MCP", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA162", "package_size": "8x10.5x1.0mm", "temp_min": -25.0, "temp_max": 85.0, "temperature": "-25°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2420.0, "moq": 2420.0, "carton_qty": 14520.0, "inner_box_qty": 2420.0, "packing": "Tray", "note": "2G1GD2 1.7~1.9V NAND MCP", "original_product_category": "NAND MCP", "original_capacity": "2G1GD2", "original_voltage": "1.7~1.9V", "original_temperature": "-25°C-+85°C"}, {"xtx_model": "XT27Q04ABSIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32F-WHIA", "product_type": "NOR Wafer", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 27580.0, "moq": 27580.0, "carton_qty": 689500.0, "inner_box_qty": 689500.0, "packing": "Cassette", "note": "32Mbit 2.7~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27G04ATSIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "TSOP48", "package_size": "12x20x1.2mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 960.0, "moq": 960.0, "carton_qty": 5760.0, "inner_box_qty": 960.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128F-WHHA", "product_type": "NOR Wafer", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 9184.0, "moq": 9184.0, "carton_qty": 229600.0, "inner_box_qty": 229600.0, "packing": "Cassette", "note": "128Mbit 2.7~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25F128F-WHIA", "product_type": "NOR Wafer", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9184.0, "moq": 9184.0, "carton_qty": 229600.0, "inner_box_qty": 229600.0, "packing": "Cassette", "note": "128Mbit 2.7~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FSSIGU", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FWOIGT", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F128FWOIGT-W", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "128Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G04CWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q08DSOIGT", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "8Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F256BSFIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 1500.0, "moq": 1500.0, "carton_qty": 15000.0, "inner_box_qty": 3000.0, "packing": "T&R", "note": "256Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25W32F-WHIA", "product_type": "NOR Wafer", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 27580.0, "moq": 27580.0, "carton_qty": 689500.0, "inner_box_qty": 689500.0, "packing": "Cassette", "note": "32Mbit 1.65~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "32Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q16D-WHIE", "product_type": "NOR Wafer", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 28006.0, "moq": 28006.0, "carton_qty": 700150.0, "inner_box_qty": 700150.0, "packing": "Cassette", "note": "16Mbit 1.65~2.0V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "16Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q16DSOIGT", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "16Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT61M1G8C2TM-B8BEA", "product_type": "MCP", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA162", "package_size": "8x10.5x1.0mm", "temp_min": -25.0, "temp_max": 85.0, "temperature": "-25°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2420.0, "moq": 2420.0, "carton_qty": 14520.0, "inner_box_qty": 2420.0, "packing": "Tray", "note": "1G1GD2 1.7~1.9V NAND MCP", "original_product_category": "NAND MCP", "original_capacity": "1G1GD2", "original_voltage": "1.7~1.9V", "original_temperature": "-25°C-+85°C"}, {"xtx_model": "XT25Q32F-WHIA", "product_type": "NOR Wafer", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 26462.0, "moq": 26462.0, "carton_qty": 661550.0, "inner_box_qty": 661550.0, "packing": "Cassette", "note": "32Mbit 1.65~1.95V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "32Mbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q32FSOIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "32Mbit 1.65~1.95V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q32FDTIGT", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "32Mbit 1.65~1.95V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q128F-WHIA", "product_type": "NOR Wafer", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 9184.0, "moq": 9184.0, "carton_qty": 229600.0, "inner_box_qty": 229600.0, "packing": "Cassette", "note": "128Mbit 1.65~2.0V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "128Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q128FSSIGT", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "128Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "128Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64F-WHIA", "product_type": "NOR Wafer", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 16214.0, "moq": 16214.0, "carton_qty": 405350.0, "inner_box_qty": 405350.0, "packing": "Cassette", "note": "64Mbit 2.7~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FWOIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSSIGT", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2000.0, "moq": 2000.0, "carton_qty": 20000.0, "inner_box_qty": 4000.0, "packing": "T&R", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSSIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 9000.0, "moq": 9000.0, "carton_qty": 72000.0, "inner_box_qty": 9000.0, "packing": "Tube", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FDTIGT", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32F-WHHA", "product_type": "NOR Wafer", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 27580.0, "moq": 27580.0, "carton_qty": 689500.0, "inner_box_qty": 689500.0, "packing": "Cassette", "note": "32Mbit 2.7~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25Q16DSOHGU", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "16Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25Q16D-WHHE", "product_type": "NOR Wafer", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 28006.0, "moq": 28006.0, "carton_qty": 700150.0, "inner_box_qty": 700150.0, "packing": "Cassette", "note": "16Mbit 1.65~2.0V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "16Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT25Q08D-WHIE", "product_type": "NOR Wafer", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 37186.0, "moq": 37186.0, "carton_qty": 929650.0, "inner_box_qty": 929650.0, "packing": "Cassette", "note": "8Mbit 1.65~2.0V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "8Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F32FSOIGU", "product_type": "SPI NOR", "density_mb": 32.0, "capacity_display": "32Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "32Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "32Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F64FSOIGU", "product_type": "SPI NOR", "density_mb": 64.0, "capacity_display": "64Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "64Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "64Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F256BWSIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "256Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25BF256BWSIGT", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 30000.0, "inner_box_qty": 6000.0, "packing": "T&R", "note": "256Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26Q04CWSIGA", "product_type": "SPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "4Gbit 1.7~1.9V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "4Gbit", "original_voltage": "1.7~1.9V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27G04ABSIGA", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA63", "package_size": "11x9x0.88mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2100.0, "moq": 2100.0, "carton_qty": 12600.0, "inner_box_qty": 2100.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F256BSFIGU", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2640.0, "moq": 2640.0, "carton_qty": 21120.0, "inner_box_qty": 2640.0, "packing": "Tube", "note": "256Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F256B-WHIA", "product_type": "NOR Wafer", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SSS", "min_package_qty": 3325.0, "moq": 3325.0, "carton_qty": 83125.0, "inner_box_qty": 83125.0, "packing": "Cassette", "note": "256Mbit 2.7~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G02CWSIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT26G01CWSIGA", "product_type": "SPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q16DSOIGU", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "16Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q16DDTIGT", "product_type": "SPI NOR", "density_mb": 16.0, "capacity_display": "16Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "16Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "16Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q08DSOIGU", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP8", "package_size": "150mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 10000.0, "moq": 10000.0, "carton_qty": 80000.0, "inner_box_qty": 10000.0, "packing": "Tube", "note": "8Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25Q08DDTIGT", "product_type": "SPI NOR", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "DFN8", "package_size": "2x3x0.4mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "Tape & Reel", "note": "8Mbit 1.65~2.0V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "8Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT25F08F-WSHA", "product_type": "NOR Wafer", "density_mb": 8.0, "capacity_display": "8Mbit", "voltage_type": "宽压/3V", "vcc_min": 2.3, "vcc_max": 3.6, "voltage_range": "2.3~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 61230.0, "moq": 61230.0, "carton_qty": 1530750.0, "inner_box_qty": 1530750.0, "packing": "Cassette", "note": "8Mbit 2.3~3.6V NOR Wafer", "original_product_category": "NOR Wafer", "original_capacity": "8Mbit", "original_voltage": "2.3~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "XT26G02ELGIGA", "product_type": "SPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "LGA8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3770.0, "moq": 3770.0, "carton_qty": 22620.0, "inner_box_qty": 3770.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V SPI NAND", "original_product_category": "SPI NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27G02BTSIGA", "product_type": "PPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "TSOP48", "package_size": "12x20x1.2mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 960.0, "moq": 960.0, "carton_qty": 5760.0, "inner_box_qty": 960.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27G01ATSIGA", "product_type": "PPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "TSOP48", "package_size": "12x20x1.2mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 960.0, "moq": 960.0, "carton_qty": 5760.0, "inner_box_qty": 960.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT27G02ATSIGA", "product_type": "PPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "TSOP48", "package_size": "12x20x1.2mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 960.0, "moq": 960.0, "carton_qty": 5760.0, "inner_box_qty": 960.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET6GS0A-HS(Z2SH23)", "product_type": "NAND Wafer", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": NaN, "temp_max": NaN, "temperature": null, "frequency_mhz": NaN, "status": "SS", "min_package_qty": 44000.0, "moq": 44000.0, "carton_qty": 44000.0, "inner_box_qty": 44000.0, "packing": "Cassette", "note": "4Gbit 1.65~1.95V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "4Gbit", "original_voltage": "1.65~1.95V", "original_temperature": null}, {"xtx_model": "ET6GS0A-HS(Z2SH22)", "product_type": "NAND Wafer", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 44000.0, "moq": 44000.0, "carton_qty": 44000.0, "inner_box_qty": 44000.0, "packing": "Cassette", "note": "4Gbit 1.65~1.95V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "4Gbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM3A-HS(Z1SH2C)", "product_type": "NAND Wafer", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 61000.0, "moq": 61000.0, "carton_qty": 61000.0, "inner_box_qty": 61000.0, "packing": "Cassette", "note": "2Gbit 2.7~3.6V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM3A-HS(Z1SH26)", "product_type": "NAND Wafer", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 61000.0, "moq": 61000.0, "carton_qty": 61000.0, "inner_box_qty": 61000.0, "packing": "Cassette", "note": "2Gbit 2.7~3.6V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM3A-HS(Z1SH22)", "product_type": "NAND Wafer", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 61000.0, "moq": 61000.0, "carton_qty": 61000.0, "inner_box_qty": 61000.0, "packing": "Cassette", "note": "2Gbit 1.65~1.95V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "2Gbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM3A-HS(Z1SH28)", "product_type": "NAND Wafer", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 61000.0, "moq": 61000.0, "carton_qty": 61000.0, "inner_box_qty": 61000.0, "packing": "Cassette", "note": "2Gbit 1.65~1.95V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "2Gbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM2B-HS(Z0SH26)", "product_type": "NAND Wafer", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 107800.0, "moq": 107800.0, "carton_qty": 107800.0, "inner_box_qty": 107800.0, "packing": "Cassette", "note": "1Gbit 2.7~3.6V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM2B-HS(Z0SH2C)", "product_type": "NAND Wafer", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 107800.0, "moq": 107800.0, "carton_qty": 107800.0, "inner_box_qty": 107800.0, "packing": "Cassette", "note": "1Gbit 2.7~3.6V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM2B-HS(Z0SH29)", "product_type": "NAND Wafer", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 107800.0, "moq": 107800.0, "carton_qty": 107800.0, "inner_box_qty": 107800.0, "packing": "Cassette", "note": "1Gbit 1.65~1.95V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "1Gbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM2B-HS(Z0SH28)", "product_type": "NAND Wafer", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 107800.0, "moq": 107800.0, "carton_qty": 107800.0, "inner_box_qty": 107800.0, "packing": "Cassette", "note": "1Gbit 1.65~1.95V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "1Gbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "ET5HM2B-HS(Z0SH22)", "product_type": "NAND Wafer", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 1.95, "voltage_range": "1.65~1.95V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 107800.0, "moq": 107800.0, "carton_qty": 107800.0, "inner_box_qty": 107800.0, "packing": "Cassette", "note": "1Gbit 1.65~1.95V NAND Wafer", "original_product_category": "NAND Wafer", "original_capacity": "1Gbit", "original_voltage": "1.65~1.95V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT61M2G8D2TA-B8BEA", "product_type": "MCP", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA162", "package_size": "8x10.5x1.0mm", "temp_min": -30.0, "temp_max": 85.0, "temperature": "-30°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2420.0, "moq": 2420.0, "carton_qty": 14520.0, "inner_box_qty": 2420.0, "packing": "Tray", "note": "2G2GD2 1.7~1.9V NAND MCP", "original_product_category": "NAND MCP", "original_capacity": "2G2GD2", "original_voltage": "1.7~1.9V", "original_temperature": "-30°C-+85°C"}, {"xtx_model": "XT25F04CDFIGT", "product_type": "SPI NOR", "density_mb": 4.0, "capacity_display": "4Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "DFN8", "package_size": "2x3x0.55mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3000.0, "moq": 3000.0, "carton_qty": 120000.0, "inner_box_qty": 30000.0, "packing": "T&R", "note": "4Mbit 2.7~3.6V SPI NOR", "original_product_category": "SPI NOR", "original_capacity": "4Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "XT61M4G8D2TA-B8BEA", "product_type": "MCP", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 1.9, "voltage_range": "1.7~1.9V", "package": "BGA162", "package_size": "8x10.5x1.0mm", "temp_min": -30.0, "temp_max": 85.0, "temperature": "-30°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 2420.0, "moq": 2420.0, "carton_qty": 14520.0, "inner_box_qty": 2420.0, "packing": "Tray", "note": "4G2GD2 1.7~1.9V NAND MCP", "original_product_category": "NAND MCP", "original_capacity": "4G2GD2", "original_voltage": "1.7~1.9V", "original_temperature": "-30°C-+85°C"}, {"xtx_model": "PN27G04ABGITG", "product_type": "PPI NAND", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3770.0, "moq": 3770.0, "carton_qty": 22620.0, "inner_box_qty": 3770.0, "packing": "Tray", "note": "4Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "4Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "PN27G02BBGITG", "product_type": "PPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3770.0, "moq": 3770.0, "carton_qty": 22620.0, "inner_box_qty": 3770.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "PN27G02ABGITG", "product_type": "PPI NAND", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3770.0, "moq": 3770.0, "carton_qty": 22620.0, "inner_box_qty": 3770.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "PN27G01ABGITG", "product_type": "PPI NAND", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "MP", "min_package_qty": 3770.0, "moq": 3770.0, "carton_qty": 22620.0, "inner_box_qty": 3770.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V Parallel NAND", "original_product_category": "Parallel NAND", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BWSQGA-X", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "宽压", "vcc_min": 1.65, "vcc_max": 3.6, "voltage_range": "1.65~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.65~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.65~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFWSQGA-X", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FWSQGA-X", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q256FWOJGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 5700.0, "moq": 5700.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q256FWSJGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q256FWSIGA-Z", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256FWSIGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 4800.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256FWOIGA-Z", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 5700.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256FWOIGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 5700.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FSFQGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M29F512BTSJGA", "product_type": "PPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "TSOP56", "package_size": null, "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 960.0, "moq": 96.0, "carton_qty": 5760.0, "inner_box_qty": 960.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V PPI NOR_M", "original_product_category": "PPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M29F512BBGJGA", "product_type": "PPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA64", "package_size": null, "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 4800.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V PPI NOR_M", "original_product_category": "PPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M29F02GBBGJGA", "product_type": "PPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA64", "package_size": null, "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 4800.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 2.7~3.6V PPI NOR_M", "original_product_category": "PPI NOR_M", "original_capacity": "2Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M29F01GBBGJGA", "product_type": "PPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA64", "package_size": null, "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 4800.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V PPI NOR_M", "original_product_category": "PPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M29F01GBTSJGA", "product_type": "PPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "TSOP56", "package_size": null, "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 960.0, "moq": 96.0, "carton_qty": 5760.0, "inner_box_qty": 960.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V PPI NOR_M", "original_product_category": "PPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FWSJGA-C", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FSFQGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFBGIGA-C", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F1GBSFJGA-B1", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BSFJGA-B1", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFBGJGA-B1", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFSFJGA-B1", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FWSJGA-B1", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FBGJGA-B1", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FSFJGA-B1", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBBGJGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BBGJGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BBGJGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M98Q9Q9Q9BGJGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit*3 1.7~2.0V Multi-channel SPI NOR", "original_product_category": "Multi-channel SPI NOR", "original_capacity": "512Mbit*3", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q256FBGIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256FWSIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256FSFIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256FWOIGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 570.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "256Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q2GFBGJGA-Z", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFBGJGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FWSJGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q2GFBGIGA-Z", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q1GFBGIGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q1GFSFIGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FBGIGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256BWSIGA-Z", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F1GBBGJGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBSFJGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BSFJGA-Z", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBBGIGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F1GBSFIGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BBGIGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BSFIGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BSFIGA-Z", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q2GFBGJGA-Y", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFBGJGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFWSJGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFSFJGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FBGJGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FWSJGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FSFJGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q256BWSJGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q2GFBGIGA-Y", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q1GFBGIGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q1GFWSIGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q1GFSFIGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FBGIGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256BSFIGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F1GBBGJGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBSFJGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BBGJGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BWSJGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BSFJGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BWSJGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBBGIGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F1GBSFIGA-Y", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BBGIGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BSFIGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BSFIGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256BWSIGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FWSIGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BWSIGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BWSIGA-Z", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BSFJGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M28EG64GD2YJEIGA", "product_type": "eMMC", "density_mb": 524288.0, "capacity_display": "64GByte", "voltage_type": "其他", "vcc_min": 1.8, "vcc_max": 3.3, "voltage_range": "1.8V/3.3V", "package": "BGA153", "package_size": "11.5x13x1.0mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 152.0, "moq": 152.0, "carton_qty": 1520.0, "inner_box_qty": 1520.0, "packing": "Tray", "note": "64GByte 1.8V/3.3V eMMC_M", "original_product_category": "eMMC_M", "original_capacity": "64GByte", "original_voltage": "1.8V/3.3V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M28EG32GD1YJEPGA", "product_type": "eMMC", "density_mb": 262144.0, "capacity_display": "32GByte", "voltage_type": "其他", "vcc_min": 1.8, "vcc_max": 3.3, "voltage_range": "1.8V/3.3V", "package": "BGA153", "package_size": "11.5x13x1.0mm", "temp_min": -55.0, "temp_max": 85.0, "temperature": "-55°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 152.0, "moq": 152.0, "carton_qty": 1520.0, "inner_box_qty": 1520.0, "packing": "Tray", "note": "32GByte 1.8V/3.3V eMMC_M", "original_product_category": "eMMC_M", "original_capacity": "32GByte", "original_voltage": "1.8V/3.3V", "original_temperature": "-55°C-+85°C"}, {"xtx_model": "M28EG16GD1YJEPGA", "product_type": "eMMC", "density_mb": 131072.0, "capacity_display": "16GByte", "voltage_type": "其他", "vcc_min": 1.8, "vcc_max": 3.3, "voltage_range": "1.8V/3.3V", "package": "BGA153", "package_size": "11.5x13x1.0mm", "temp_min": -55.0, "temp_max": 85.0, "temperature": "-55°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 152.0, "moq": 152.0, "carton_qty": 1520.0, "inner_box_qty": 1520.0, "packing": "Tray", "note": "16GByte 1.8V/3.3V eMMC_M", "original_product_category": "eMMC_M", "original_capacity": "16GByte", "original_voltage": "1.8V/3.3V", "original_temperature": "-55°C-+85°C"}, {"xtx_model": "M28EG08GD1YJEPGA", "product_type": "eMMC", "density_mb": 65536.0, "capacity_display": "8GByte", "voltage_type": "其他", "vcc_min": 1.8, "vcc_max": 3.3, "voltage_range": "1.8V/3.3V", "package": "BGA153", "package_size": "11.5x13x1.0mm", "temp_min": -55.0, "temp_max": 85.0, "temperature": "-55°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 152.0, "moq": 152.0, "carton_qty": 1520.0, "inner_box_qty": 1520.0, "packing": "Tray", "note": "8GByte 1.8V/3.3V eMMC_M", "original_product_category": "eMMC_M", "original_capacity": "8GByte", "original_voltage": "1.8V/3.3V", "original_temperature": "-55°C-+85°C"}, {"xtx_model": "M55Q1GFWSIGA-Z", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FWSIGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BWSIGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FSFIGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FSFIGA-Y", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BWSIGA-Y", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BWSJGA-Z", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BWSUGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C"}, {"xtx_model": "M25Q512FSFJGA-Z", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M98QA00QABGJGA", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "UD", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 480.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V Multi-channel SPI NOR", "original_product_category": "Multi-channel SPI NOR", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M42K04G16ABEJGA-F", "product_type": "DDR3", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "其他", "vcc_min": 1.35, "vcc_max": 1.5, "voltage_range": "1.35V/1.5V", "package": "BGA96", "package_size": "13x7.5mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 198.0, "moq": 198.0, "carton_qty": 1980.0, "inner_box_qty": 1980.0, "packing": "Tray", "note": "4Gbit 1.35V/1.5V DDR3L", "original_product_category": "DDR3L", "original_capacity": "4Gbit", "original_voltage": "1.35V/1.5V", "original_temperature": "-55°C-+125°C"}, {"xtx_model": "M25Q128FWOIGA-Z", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 570.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q128FSSIGA-Z", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 300.0, "moq": 300.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F128FWOIGA-Z", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 570.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F128F-WHIA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 9184.0, "moq": 9184.0, "carton_qty": 229600.0, "inner_box_qty": 9184.0, "packing": "Cassette", "note": "128Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q128F-WHIA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 9106.0, "moq": 9106.0, "carton_qty": 227650.0, "inner_box_qty": 9106.0, "packing": "Cassette", "note": "128Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256B-WHHA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 3325.0, "moq": 3325.0, "carton_qty": 83125.0, "inner_box_qty": 3325.0, "packing": "Cassette", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "M25Q512F-WSHA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "wafer", "package_size": null, "temp_min": -40.0, "temp_max": 105.0, "temperature": "-40°C-+105°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 1953.0, "moq": 1953.0, "carton_qty": 48825.0, "inner_box_qty": 1953.0, "packing": "Cassette", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+105°C"}, {"xtx_model": "M25F128FSSIGA-Z", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 300.0, "moq": 300.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M42K04G16ABEWGA-F", "product_type": "DDR3", "density_mb": 4096.0, "capacity_display": "4Gbit", "voltage_type": "其他", "vcc_min": 1.35, "vcc_max": 1.5, "voltage_range": "1.35V/1.5V", "package": "BGA96", "package_size": "13x7.5mm", "temp_min": -40.0, "temp_max": 95.0, "temperature": "-40°C-+95°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 198.0, "moq": 198.0, "carton_qty": 1980.0, "inner_box_qty": 1980.0, "packing": "Tray", "note": "4Gbit 1.35V/1.5V DDR3L", "original_product_category": "DDR3L", "original_capacity": "4Gbit", "original_voltage": "1.35V/1.5V", "original_temperature": "-40°C-+95°C"}, {"xtx_model": "M55Q2GFBGIGA-X", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q1GFBGIGA-X", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FWSIGA-X", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256BWSIGA-X", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F512BSFJGA-X", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BWSIGA-X", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "SS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BWSIGA-X", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q128FWOIGA-X", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 570.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F128FWOIGA-X", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 570.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q128FSSIGA-X", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 300.0, "moq": 300.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 1.7~2.0V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F128FSSIGA-X", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 300.0, "moq": 300.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256BSFJGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BWSJGA-F", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BSFJGA-F", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BSFJGA-D", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFWSJGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55BQ1GFWSJGA-E", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BWSJGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BWSJGA-B", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55BQ2GFBGJGA-E", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25BF1GBSFJGA-E", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "CS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55BQ1GFBGJGA-E", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BSFJGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BWSJGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BSFJGA-C", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBSFJGA-C", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q256BWSJGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FSFJGA-C", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFSFJGA-C", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFBGJGA-C", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q2GFBGJGA-C", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFBGJGA-B", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBSFJGA-B", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFSFJGA-B", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BSFJGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FSFJGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FWSJGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FBGJGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q128FWOIGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 570.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q128FSSIGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 300.0, "moq": 300.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "128Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F128FWOIGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "6x5mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 570.0, "moq": 570.0, "carton_qty": 34200.0, "inner_box_qty": 5700.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F128FSSIGA", "product_type": "SPI NOR", "density_mb": 128.0, "capacity_display": "128Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP8", "package_size": "208mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 300.0, "moq": 300.0, "carton_qty": 18000.0, "inner_box_qty": 3000.0, "packing": "Tray", "note": "128Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "128Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M55Q2GFBGJGA", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFBGJGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q1GFSFJGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FBGJGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FWSJGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25Q512FSFJGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F1GBSFJGA", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "1Gbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "1Gbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BWSJGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F512BSFJGA", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "512Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BWSJGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M25F256BSFJGA", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C+SAT", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M0", "original_product_category": "SPI NOR_M0", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-55°C-+125°C+SAT"}, {"xtx_model": "M55Q2GFBGUGA-E", "product_type": "SPI NOR", "density_mb": 2048.0, "capacity_display": "2Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "2Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "2Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C"}, {"xtx_model": "M55Q1GFBGUGA-E", "product_type": "SPI NOR", "density_mb": 1024.0, "capacity_display": "1Gbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -55.0, "temp_max": 125.0, "temperature": "-55°C-+125°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "1Gbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "1Gbit", "original_voltage": "1.7~2.0V", "original_temperature": "-55°C-+125°C"}, {"xtx_model": "M25Q256BSFIGA-D", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FSFIGA-C", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "ES", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BSFIGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q256BWSIGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "1.8V", "vcc_min": 1.65, "vcc_max": 2.0, "voltage_range": "1.65~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 1.65~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "1.65~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25F256BWSIGA-C", "product_type": "SPI NOR", "density_mb": 256.0, "capacity_display": "256Mbit", "voltage_type": "3.3V", "vcc_min": 2.7, "vcc_max": 3.6, "voltage_range": "2.7~3.6V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "256Mbit 2.7~3.6V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "256Mbit", "original_voltage": "2.7~3.6V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FWSIGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "WSON8", "package_size": "8x6mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FBGIGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "BGA24", "package_size": "6x8mm", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 480.0, "moq": 480.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "M25Q512FSFIGA-B", "product_type": "SPI NOR", "density_mb": 512.0, "capacity_display": "512Mbit", "voltage_type": "1.8V", "vcc_min": 1.7, "vcc_max": 2.0, "voltage_range": "1.7~2.0V", "package": "SOP16", "package_size": "300mil", "temp_min": -40.0, "temp_max": 85.0, "temperature": "-40°C-+85°C", "frequency_mhz": NaN, "status": "QS", "min_package_qty": 176.0, "moq": 176.0, "carton_qty": 10560.0, "inner_box_qty": 1760.0, "packing": "Tray", "note": "512Mbit 1.7~2.0V SPI NOR_M3", "original_product_category": "SPI NOR_M3", "original_capacity": "512Mbit", "original_voltage": "1.7~2.0V", "original_temperature": "-40°C-+85°C"}, {"xtx_model": "25F08FSOUGA", "product_type": "SPI NOR", "density_mb": NaN, "capacity_display": null, "voltage_type": null, "vcc_min": NaN, "vcc_max": NaN, "voltage_range": null, "package": "SOP8", "package_size": "150mil", "temp_min": NaN, "temp_max": NaN, "temperature": null, "frequency_mhz": NaN, "status": "CS", "min_package_qty": 4800.0, "moq": 4800.0, "carton_qty": 28800.0, "inner_box_qty": 4800.0, "packing": "Tray", "note": "SPI NOR_M", "original_product_category": "SPI NOR_M", "original_capacity": null, "original_voltage": null, "original_temperature": null}]'
EMBEDDED_MATCH_WEIGHTS_JSON = r'[{"criterion": "product_type", "weight": 30, "enabled": 1, "description": "核心条件：产品类型必须一致才推荐"}, {"criterion": "density", "weight": 35, "enabled": 1, "description": "核心条件：容量必须一致才推荐"}, {"criterion": "voltage", "weight": 35, "enabled": 1, "description": "核心条件：电压范围必须一致才推荐"}, {"criterion": "package", "weight": 15, "enabled": 1, "description": "二级条件：封装名称/尺寸匹配用于排序和风险提示"}, {"criterion": "temperature", "weight": 15, "enabled": 1, "description": "二级条件：温度范围覆盖用于排序和风险提示"}]'

try:
    DEFAULT_COMPANY_MASTER = pd.DataFrame(json.loads(EMBEDDED_COMPANY_MASTER_JSON))
    DEFAULT_XTX_PRODUCT_LIBRARY = pd.DataFrame(json.loads(EMBEDDED_XTX_PRODUCT_LIBRARY_JSON))
    DEFAULT_MATCH_WEIGHTS = pd.DataFrame(json.loads(EMBEDDED_MATCH_WEIGHTS_JSON))
except Exception:
    pass


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
    """读取维护库。

    优先级：
    1. 左侧上传的 XLSX；
    2. 仓库根目录中的 xtx_competitor_maintenance_current.xlsx；
    3. 仓库根目录中的 xtx_competitor_maintenance_template.xlsx；
    4. 代码内置的当前维护库。
    """
    sheets = None
    if file_bytes:
        sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    else:
        for local_path in ["xtx_competitor_maintenance_current.xlsx", "xtx_competitor_maintenance_template.xlsx"]:
            if os.path.exists(local_path):
                try:
                    sheets = pd.read_excel(local_path, sheet_name=None)
                    break
                except Exception:
                    sheets = None

    if sheets is None:
        return {
            "Company_Master": ensure_company_columns(DEFAULT_COMPANY_MASTER.copy()),
            "XTX_Product_Library": ensure_xtx_columns(DEFAULT_XTX_PRODUCT_LIBRARY.copy()),
            "Match_Weights": clean_columns(DEFAULT_MATCH_WEIGHTS.copy()),
            "History_Log": DEFAULT_HISTORY_LOG.copy(),
            "PDF_Library": pd.DataFrame(columns=["company_name", "model", "product_type", "pdf_file_name", "pdf_source_url", "version", "last_modified", "confirm_status", "note"]),
        }

    return {
        "Company_Master": ensure_company_columns(sheets.get("Company_Master", DEFAULT_COMPANY_MASTER.copy())),
        "XTX_Product_Library": ensure_xtx_columns(sheets.get("XTX_Product_Library", DEFAULT_XTX_PRODUCT_LIBRARY.copy())),
        "Match_Weights": clean_columns(sheets.get("Match_Weights", DEFAULT_MATCH_WEIGHTS.copy())),
        "History_Log": clean_columns(sheets.get("History_Log", DEFAULT_HISTORY_LOG.copy())),
        "PDF_Library": clean_columns(sheets.get("PDF_Library", pd.DataFrame(columns=["company_name", "model", "product_type", "pdf_file_name", "pdf_source_url", "version", "last_modified", "confirm_status", "note"]))),
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
    if not allowed:
        return False
    # GigaDevice 中英文站点经常在 .com 与 .com.cn 之间跳转，视为同一官网来源。
    gd_aliases = {"gigadevice.com", "gigadevice.com.cn"}
    if current in gd_aliases and allowed in gd_aliases:
        return True
    return current == allowed or current.endswith("." + allowed)


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
    """从官网产品页 HTML 表格中抽取包含目标型号的行。

    V5 优化：
    1. 优先用 pandas.read_html 解析标准表格，避免把整页文本当成一行，导致 7 + 4Mb 被误识别成 74Mb。
    2. 再用 BeautifulSoup 兜底解析非标准表格。
    3. 仅返回包含完整目标型号的表格行；不把普通页面文本用于容量/电压抽取。
    """
    model_norm = normalize_model(model)
    rows_out: list[dict] = []
    if not model_norm:
        return rows_out

    # 方式一：pandas.read_html 对 Boya 这类标准产品表格更稳定。
    try:
        html_tables = pd.read_html(io.StringIO(html))
        for table_idx, df in enumerate(html_tables):
            if df.empty:
                continue
            df = df.copy()
            df.columns = [normalize_header_name(str(c)) for c in df.columns]
            # 如果表头没有识别到型号列，也保留所有列用于 row_text 匹配。
            for row_idx, r in df.iterrows():
                values = ["" if pd.isna(v) else str(v).strip() for v in r.tolist()]
                row_text = " ".join(values)
                if model_norm not in normalize_model(row_text):
                    continue
                row = {
                    "source_url": current_url,
                    "table_index": table_idx,
                    "row_index": int(row_idx),
                    "row_text": row_text,
                }
                for col, val in zip(df.columns, values):
                    key = normalize_header_name(col)
                    if not key:
                        continue
                    if key in row and str(row[key]).strip():
                        row[key] = f"{row[key]} | {val}"
                    else:
                        row[key] = val
                rows_out.append(row)
    except Exception:
        pass

    if rows_out:
        return rows_out

    # 方式二：BeautifulSoup 兜底。
    soup = BeautifulSoup(html, "html.parser")
    for table_idx, table in enumerate(soup.find_all("table")):
        trs = table.find_all("tr")
        if not trs:
            continue

        headers = []
        header_tr_index = 0
        for i, tr in enumerate(trs[:5]):
            cells = tr.find_all(["th", "td"])
            texts = [c.get_text(" ", strip=True) for c in cells]
            normalized = [normalize_header_name(t) for t in texts]
            if any(t in ["density", "part_no", "vcc", "package", "temperature", "status"] for t in normalized):
                headers = normalized
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
                key = normalize_header_name(key)
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




def product_model_slug_candidates(model: str, company: str = "") -> list[str]:
    """生成产品详情页 URL 的可能 slug。部分官网产品详情页比用户输入型号多一个后缀，例如 GD25LE256 -> gd25le256h。"""
    model_norm = normalize_model(model).lower()
    if not model_norm:
        return []
    slugs = [model_norm]
    company_l = str(company or "").lower()
    # GigaDevice 官网常见详情页 slug 可能带 H/E 等后缀，用户通常只输入主体型号。
    if "giga" in company_l or model_norm.startswith("gd"):
        suffixes = ["h", "e", "d", "f", "g", "c", "b", "a"]
        if re.search(r"\d$", model_norm):
            slugs.extend([model_norm + s for s in suffixes])
    return list(dict.fromkeys(slugs))


def build_direct_product_urls(company: str, base_url: str, model: str) -> list[str]:
    """构造官网产品中心/产品详情页直达 URL。直达 URL 比爬全站快，尤其适合 GigaDevice 这类产品详情页。"""
    company_l = str(company or "").lower()
    urls: list[str] = []
    slugs = product_model_slug_candidates(model, company)

    if "giga" in company_l:
        hosts = ["https://www.gigadevice.com", "https://www.gigadevice.com.cn"]
        paths = [
            "/product/flash/spi-nor-flash",
            "/product/flash/spi-nand-flash",
            "/product/flash/parallel-nand-flash",
            "/product/flash/parallel-nor-flash",
        ]
        for host in hosts:
            for path in paths:
                urls.append(host + path)
                for slug in slugs:
                    urls.append(host + path + "/" + slug)

    # 其它厂商也先放入常见产品中心页面，避免从首页开始爬。
    if "boya" in company_l:
        urls.extend([
            "https://www.boyamicro.com/?zh/products/2",
            "https://www.boyamicro.com/zh/products/2",
        ])
    if "winbond" in company_l:
        urls.append("https://www.winbond.com/hq/product/code-storage-flash-memory/serial-nor-flash/")
    if "macronix" in company_l:
        urls.append("https://www.macronix.com/en-us/products/NOR-Flash/Pages/default.aspx")
    if "issi" in company_l:
        urls.append("https://www.issi.com/US/product-flash.shtml")
    if "zbit" in company_l:
        urls.append("https://www.zbitsemi.com/product")
    if "puy" in company_l:
        urls.append("https://www.puyasemi.com/product.html")

    # V8：所有友商都按“产品中心/站内搜索/产品详情页”优先规则尝试。
    base = str(base_url or "").rstrip("/")
    model_q = quote_plus(model)
    if base:
        common_paths = [
            "/product", "/products", "/product-center", "/product_center", "/products-center",
            "/product.html", "/products.html", "/en/products", "/en/product", "/zh/products", "/zh/product",
            f"/search?keyword={model_q}", f"/search?keywords={model_q}", f"/search?q={model_q}",
            f"/search.aspx?key={model_q}", f"/search.aspx?keyword={model_q}", f"/search.html?keyword={model_q}",
            f"/?s={model_q}", f"/?search={model_q}",
        ]
        for path in common_paths:
            urls.append(base + path)
        for slug in slugs:
            for path in ["/product", "/products", "/product/detail", "/products/detail", "/en/product", "/en/products", "/zh/product", "/zh/products"]:
                urls.append(base + path + "/" + slug)

    return [u for u in dict.fromkeys(urls) if valid_url(u)]


def line_value_after(lines: list[str], label_patterns: list[str], stop_patterns: list[str] | None = None) -> str:
    """在详情页文本中，读取某个标签下一行的值。适合 GigaDevice 详情页：Voltage\n1.65V~2.0V。"""
    stop_patterns = stop_patterns or []
    labels = [re.compile(p, flags=re.IGNORECASE) for p in label_patterns]
    stops = [re.compile(p, flags=re.IGNORECASE) for p in stop_patterns]
    for i, line in enumerate(lines):
        clean = str(line).strip()
        if not clean:
            continue
        if any(p.fullmatch(clean) or p.search(clean) for p in labels):
            vals = []
            for nxt in lines[i + 1:i + 5]:
                nxt = str(nxt).strip()
                if not nxt:
                    continue
                if any(s.search(nxt) for s in stops):
                    break
                # 遇到下一个明显标签则停止。
                if re.fullmatch(r"Status|Voltage|Density|Temperature.*|I/O Bus|Frequency.*|Features|Packages|Documentation|Development Tools|Datasheet|Buy now", nxt, flags=re.IGNORECASE):
                    break
                vals.append(nxt)
                # 大部分字段下一行就是值；Features/Packages 可能较长，也最多取几行。
                if len(vals) >= 2:
                    break
            return ",".join(vals).strip()
    return ""


def product_type_from_url_or_text(url: str, text: str, selected_type: str, model: str) -> str:
    if selected_type and selected_type != "未指定":
        return selected_type
    u = str(url).lower()
    t = str(text).lower()
    if "spi-nor" in u or "spi nor" in t:
        return "SPI NOR"
    if "spi-nand" in u or "spi nand" in t:
        return "SPI NAND"
    if "parallel-nor" in u or "parallel nor" in t:
        return "PPI NOR"
    if "parallel-nand" in u or "parallel nand" in t:
        return "PPI NAND"
    return extract_product_type(text, selected_type=selected_type, model=model).get("product_type", "")


def product_detail_row_from_page(html: str, current_url: str, model: str, company: str, selected_type: str = "未指定") -> dict | None:
    """从官网产品详情页抽取结构化字段。优先用于 GigaDevice 等没有传统 HTML 表格、但页面文本包含 Features 字段的网站。"""
    model_norm = normalize_model(model)
    soup = BeautifulSoup(html, "html.parser")
    page_text = soup.get_text("\n", strip=True)
    text_norm = normalize_model(page_text)
    if model_norm not in text_norm:
        return None
    lines = [x.strip() for x in page_text.splitlines() if x.strip()]

    part_no = ""
    for line in lines[:80]:
        m = re.search(rf"\b({re.escape(str(model).upper())}[A-Z0-9]*)\b", line.upper())
        if m:
            part_no = m.group(1)
            break
    if not part_no:
        part_no = model

    stop_labels = [r"Status", r"Voltage", r"Density", r"Temperature", r"I/O Bus", r"Frequency", r"Features", r"Packages", r"Documentation", r"Development Tools"]
    status = line_value_after(lines, [r"Status"], stop_labels)
    voltage = line_value_after(lines, [r"Voltage"], stop_labels)
    density = line_value_after(lines, [r"Density"], stop_labels)
    temperature = line_value_after(lines, [r"Temperature.*"], stop_labels)
    frequency = line_value_after(lines, [r"Frequency.*"], stop_labels)
    feature = line_value_after(lines, [r"Features"], stop_labels)
    package = line_value_after(lines, [r"Packages"], stop_labels)

    # 如果标签提取失败，再用正则兜底。
    flat = re.sub(r"\s+", " ", page_text)
    if not density:
        m = re.search(r"Density\s+([0-9]+\s*(?:Kb|Mb|Gb|Kbit|Mbit|Gbit))", flat, flags=re.IGNORECASE)
        density = m.group(1) if m else ""
    if not voltage:
        m = re.search(r"Voltage\s+([0-9.]+\s*V?\s*[~\-–]\s*[0-9.]+\s*V)", flat, flags=re.IGNORECASE)
        voltage = m.group(1) if m else ""
    if not package:
        m = re.search(r"Packages\s+(.+?)(?:Documentation|Development Tools|Datasheet|Buy now|$)", flat, flags=re.IGNORECASE)
        package = m.group(1).strip() if m else ""

    if not any([density, voltage, temperature, frequency, feature, package]):
        return None

    row_text = " ".join([part_no, density, voltage, frequency, feature, package, temperature, status])
    return {
        "part_no": part_no,
        "density": density,
        "vcc": voltage,
        "frequency": frequency,
        "feature": feature,
        "package": package,
        "temperature": temperature,
        "status": status,
        "source_url": current_url,
        "page_url": current_url,
        "row_text": row_text,
        "source_kind": "官网产品详情页",
        "page_score": 650 + score_pdf_candidate(current_url, row_text, model),
    }

def crawl_company_for_product_page(company_row: pd.Series, model: str, max_pages: int = 3) -> pd.DataFrame:
    """V8：只检索官网产品中心/产品详情页，不再在自动模式下下载或扫描 PDF。"""
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

        current_page_rows = []

        # 1) 标准产品选择器表格：例如 Boya 的产品列表。
        for row in table_rows_containing_model(html, model, final_url):
            row.update({
                "company_name": company,
                "page_url": final_url,
                "datasheet_url": best_pdf,
                "source_kind": "官网产品页表格",
                "page_score": 500 + score_pdf_candidate(final_url, row.get("row_text", ""), model),
            })
            current_page_rows.append(row)

        # 2) 产品详情页标签字段：例如 GigaDevice 的 GD25LE256H 详情页。
        if page_has_model and not current_page_rows:
            detail_row = product_detail_row_from_page(html, final_url, model, company)
            if detail_row:
                detail_row.update({
                    "company_name": company,
                    "datasheet_url": best_pdf,
                })
                current_page_rows.append(detail_row)

        if current_page_rows:
            results.extend(current_page_rows)
            # 已经命中产品中心/详情页，不继续深爬，避免慢。
            continue

        # 3) 页面文本包含型号，但没有结构化字段：只保留来源，不做字段误抽取。
        if page_has_model:
            results.append({
                "company_name": company,
                "page_url": final_url,
                "source_url": final_url,
                "datasheet_url": best_pdf,
                "source_kind": "官网产品页文本",
                "row_text": page_text[:5000],
                "page_score": 250 + score_pdf_candidate(final_url, page_text[:1000], model),
            })
            continue

        # 4) 没命中时，只继续少量高相关产品中心链接；不进入 PDF 下载。
        links = extract_links_from_html(html, final_url)
        for item in links:
            link_url = item["url"]
            link_text = item["text"]
            if not same_domain_or_subdomain(link_url, domain):
                continue
            link_url_l = link_url.lower()
            combined = normalize_model(link_url + " " + link_text)
            if model_norm in combined or any(k in link_url_l for k in ["product", "products", "flash", "nor", "nand", "spi"]):
                if link_url not in visited and len(queue) < max_pages * 3:
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



def get_first_non_empty(row: pd.Series, keys: list[str]) -> str:
    """从产品页表格行中按多个可能字段名取第一个非空值。"""
    for key in keys:
        if key in row.index:
            val = row.get(key, "")
            if val is not None and str(val).strip() not in ["", "nan", "None"]:
                return str(val).strip()
    return ""


def spec_from_product_page_row(row: pd.Series, selected_type: str, model: str) -> dict:
    """把官网产品页表格行转换成统一 spec。

    V5：只有 source_kind=官网产品页表格 时，才从行字段抽取容量/电压/封装/温度；
    如果只是普通页面文本命中型号，为避免误识别，容量/电压/封装/温度留空，交给选型下拉修正。
    """
    row_text = str(row.get("row_text", ""))
    source_kind = str(row.get("source_kind", ""))
    is_table_source = "表格" in source_kind or bool(get_first_non_empty(row, ["part_no", "density", "vcc", "package", "temperature"]))

    product_type = extract_product_type(row_text, selected_type=selected_type, model=model)

    density_mb, density_display = None, ""
    voltage = build_voltage_dict(None, None, "低")
    package = {"package": "", "package_size": "", "confidence": "低"}
    temp = build_temperature_dict(None, None, "低")

    if is_table_source:
        density_text = get_first_non_empty(row, ["density", "capacity", "memory"]) or row_text
        vcc_text = get_first_non_empty(row, ["vcc", "voltage"])
        package_text = get_first_non_empty(row, ["package", "pkg"])
        temp_text = get_first_non_empty(row, ["temperature", "tem", "temp"])

        density_mb, density_display = parse_density_from_text(density_text)
        # 如果密度列缺失，再从整行找带单位容量，但 parse 函数不会乱取纯数字。
        if density_mb is None:
            density_mb, density_display = parse_density_from_text(row_text)

        voltage = infer_voltage_from_text(vcc_text or row_text)
        package = extract_package(package_text) if package_text else {"package": "", "package_size": "", "confidence": "低"}
        temp = extract_all_temperature_ranges(temp_text) if temp_text else build_temperature_dict(None, None, "低")

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

    # V8：优先尝试产品中心/产品详情页直达链接，避免从首页慢速爬取。
    seeds.extend(build_direct_product_urls(company, base_url, model))
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




def collect_uploaded_pdf_library(uploaded_files) -> list[dict]:
    """读取侧边栏上传的 PDF 库文件。"""
    items = []
    for f in uploaded_files or []:
        try:
            items.append({"file_name": f.name, "source": "上传PDF库", "pdf_bytes": f.getvalue()})
        except Exception:
            continue
    return items


def collect_repo_pdf_library(repo_dir: str = "pdf_library") -> list[dict]:
    """读取部署仓库中的 pdf_library 文件夹。用户可把常用竞品规格书提交到 GitHub 的 pdf_library/ 目录。"""
    items = []
    if not os.path.isdir(repo_dir):
        return items
    for root, _, files in os.walk(repo_dir):
        for name in files:
            if not name.lower().endswith(".pdf"):
                continue
            path = os.path.join(root, name)
            try:
                with open(path, "rb") as f:
                    items.append({"file_name": name, "source": "仓库pdf_library", "path": path, "pdf_bytes": f.read()})
            except Exception:
                continue
    return items




def collect_sheet_pdf_library(pdf_library_df: pd.DataFrame) -> list[dict]:
    """读取维护表 PDF_Library 中的 PDF 索引。支持 pdf_source_url，也支持 pdf_file_name 对应仓库 pdf_library/ 文件。"""
    items = []
    if pdf_library_df is None or pdf_library_df.empty:
        return items
    df = clean_columns(pdf_library_df)
    for _, r in df.iterrows():
        file_name = str(r.get("pdf_file_name", "") or "").strip()
        url = str(r.get("pdf_source_url", "") or "").strip()
        model = str(r.get("model", "") or "").strip()
        if not file_name and not url:
            continue
        item = {
            "file_name": file_name or url.split("/")[-1],
            "source": "维护表PDF_Library",
            "pdf_url": url if valid_url(url) else "",
            "metadata_model": model,
            "metadata_company": str(r.get("company_name", "") or "").strip(),
            "metadata_product_type": str(r.get("product_type", "") or "").strip(),
            "note": str(r.get("note", "") or "").strip(),
        }
        # 如果仓库中存在同名文件，直接读取，不需要联网。
        if file_name:
            local_path = os.path.join("pdf_library", file_name)
            if os.path.exists(local_path):
                try:
                    with open(local_path, "rb") as f:
                        item["pdf_bytes"] = f.read()
                    item["path"] = local_path
                except Exception:
                    pass
        items.append(item)
    return items


def ensure_pdf_item_bytes(item: dict) -> dict:
    """确保 PDF 库 item 中有 pdf_bytes。维护表只有URL时，命中后才下载，避免启动时慢。"""
    if item.get("pdf_bytes"):
        return item
    url = str(item.get("pdf_url", "") or "").strip()
    if url:
        item = dict(item)
        item["pdf_bytes"] = download_pdf(url)
    return item

def search_pdf_library(pdf_items: list[dict], model: str, max_scan_pages: int | None = None, max_files: int = 50) -> tuple[dict | None, pd.DataFrame]:
    """在 PDF 库中搜索目标型号。优先文件名命中；否则快速扫描 PDF 文本，找到型号即停止。"""
    model_norm = normalize_model(model)
    rows = []
    if not model_norm or not pdf_items:
        return None, pd.DataFrame()

    # 维护表型号 / 文件名命中最快。
    for item in pdf_items:
        file_name = str(item.get("file_name", ""))
        meta_model = str(item.get("metadata_model", ""))
        if model_norm in normalize_model(file_name) or (meta_model and model_norm in normalize_model(meta_model)):
            row = {
                "file_name": file_name,
                "source": item.get("source", "PDF库"),
                "match_method": "维护表型号/文件名包含型号",
                "model_found": True,
                "model_page": "",
                "score": 1000,
            }
            rows.append(row)
            hit = ensure_pdf_item_bytes(dict(item))
            hit.update(row)
            return hit, pd.DataFrame(rows)

    # 文件名没命中时，再扫描少量 PDF。
    scan_items = pdf_items[:max_files]
    for item in scan_items:
        file_name = str(item.get("file_name", ""))
        try:
            item = ensure_pdf_item_bytes(dict(item))
            found, page_no, method = scan_pdf_for_model_fast(item.get("pdf_bytes", b""), model, max_pages=max_scan_pages)
            row = {
                "file_name": file_name,
                "source": item.get("source", "PDF库"),
                "match_method": method if found else "全文/限页扫描未命中",
                "model_found": bool(found),
                "model_page": page_no or "",
                "score": 800 if found else 0,
            }
            rows.append(row)
            if found:
                hit = ensure_pdf_item_bytes(dict(item))
                hit.update(row)
                return hit, pd.DataFrame(rows)
        except Exception as e:
            rows.append({
                "file_name": file_name,
                "source": item.get("source", "PDF库"),
                "match_method": f"扫描失败：{e}",
                "model_found": False,
                "model_page": "",
                "score": 0,
            })
    return None, pd.DataFrame(rows)

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
    """解析容量文本，避免把表格行号和容量粘连成错误容量。

    典型输入：512Kb、4Mb、4Mbit、1Gb、8Gbit、32Mbit。
    不再简单删除所有空格，否则 "7 4Mb" 会被误判成 74Mb。
    """
    if value is None or str(value).strip() == "":
        return None, ""
    raw = str(value).strip()
    s = raw.upper().replace("Ｍ", "M").replace("Ｇ", "G").replace("Ｋ", "K")
    s = s.replace("‐", "-").replace("‑", "-").replace("–", "-").replace("—", "-")

    # 优先匹配有明确单位的容量。要求数字和单位之间最多允许少量空格，不跨越其它数字。
    patterns = [
        (r"(?<![A-Z0-9])(\d+(?:\.\d+)?)\s*(G\s*BIT|GBIT|G-BIT|G\b)", 1024, "Gb"),
        (r"(?<![A-Z0-9])(\d+(?:\.\d+)?)\s*(M\s*BIT|MBIT|M-BIT|M\b|MB\b)", 1, "Mb"),
        (r"(?<![A-Z0-9])(\d+(?:\.\d+)?)\s*(K\s*BIT|KBIT|K-BIT|K\b|KB\b)", 1/1024, "Kb"),
    ]
    for pattern, factor, unit in patterns:
        m = re.search(pattern, s)
        if m:
            num = float(m.group(1))
            density_mb = num * factor
            if unit == "Gb":
                return int(density_mb) if density_mb.is_integer() else density_mb, f"{num:g}Gb"
            if unit == "Mb":
                return int(density_mb) if float(density_mb).is_integer() else density_mb, f"{density_mb:g}Mb"
            return density_mb, f"{num:g}Kb"

    # 没有单位时，仅当整个字符串基本就是一个数字才按 Mb 处理；不要从完整表格行中随便取第一个数字。
    compact = re.sub(r"\s+", "", s)
    if re.fullmatch(r"\d+(?:\.\d+)?", compact):
        density = float(compact)
        return int(density) if density.is_integer() else density, f"{density:g}Mb"
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
    """解析电压范围，必须匹配带 V 的范围，避免把表格 ID / 容量 / 型号数字误判为电压。"""
    if value is None or str(value).strip() == "":
        return None, None
    s = str(value).replace("–", "-").replace("—", "-").replace("~", "-").replace("～", "-").replace("至", "-")
    s = re.sub(r"\s+", "", s)

    # 典型：2.7-3.6V、2.7V-3.6V、Vcc2.7-3.6V
    patterns = [
        r"([0-9](?:\.\d{1,3})?)V?-([0-9](?:\.\d{1,3})?)V",
        r"VCC[:：=]?([0-9](?:\.\d{1,3})?)-([0-9](?:\.\d{1,3})?)V?",
        r"([0-9](?:\.\d{1,3})?)V(?:TO|-)([0-9](?:\.\d{1,3})?)V",
    ]
    for pattern in patterns:
        m = re.search(pattern, s, flags=re.IGNORECASE)
        if not m:
            continue
        v1 = safe_float(m.group(1))
        v2 = safe_float(m.group(2))
        if v1 is None or v2 is None:
            continue
        vmin, vmax = min(v1, v2), max(v1, v2)
        if 1.0 <= vmin <= 5.5 and 1.0 <= vmax <= 5.5:
            return vmin, vmax
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




def normalize_display_text(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ["nan", "none", "null"]:
        return ""
    return s.replace("℃", "°C").replace("～", "~").replace("–", "-").replace("—", "-")


def infer_voltage_from_text(value: str) -> dict:
    raw = normalize_display_text(value)
    if not raw:
        return build_voltage_dict(None, None, "低")
    vmin, vmax = parse_voltage_from_text(raw)
    if vmin is not None and vmax is not None:
        return build_voltage_dict(vmin, vmax, "产品页")

    s = raw.upper().replace(" ", "")
    display = raw
    vtype = ""
    if re.search(r"1\.8\s*V|1V8|1\.8VONLY", s):
        vtype = "1.8V"
    elif re.search(r"3\.3\s*V|3V3", s):
        vtype = "3.3V"
    elif re.search(r"\b3\s*V\b|[^0-9]3V|^3V", s):
        vtype = "3V"
    elif "WV" in s or "WIDE" in s or "宽压" in raw:
        vtype = "宽压"
    elif s:
        vtype = raw
    return {"vcc_min": None, "vcc_max": None, "display": display, "voltage_type": vtype, "confidence": "产品页"}


def extract_all_temperature_ranges(text: str) -> dict:
    raw = normalize_display_text(text)
    if not raw:
        return build_temperature_dict(None, None, "低")
    t = raw.replace("TO", "to")
    patterns = [
        r"(-?\d{1,3})\s*°?\s*C\s*(?:to|-|~)\s*(-?\d{1,3})\s*°?\s*C",
        r"(-?\d{1,3})\s*(?:to|-|~)\s*(-?\d{1,3})\s*°?\s*C",
    ]
    ranges = []
    seen = set()
    for pattern in patterns:
        for m in re.finditer(pattern, t, flags=re.IGNORECASE):
            a, b = int(m.group(1)), int(m.group(2))
            tmin, tmax = min(a, b), max(a, b)
            if -65 <= tmin <= 25 and 70 <= tmax <= 150:
                key = (tmin, tmax)
                if key not in seen:
                    seen.add(key)
                    ranges.append(key)
    if not ranges:
        return build_temperature_dict(None, None, "低")
    disp = " / ".join([f"{a}°C~{b}°C" for a, b in ranges])
    min_t = min(a for a, _ in ranges)
    max_t = max(b for _, b in ranges)
    base = build_temperature_dict(min_t, max_t, "产品页")
    base["display"] = disp
    return base


def render_spec_card(label: str, value: str, caption: str = ""):
    value = normalize_display_text(value)
    caption = normalize_display_text(caption)
    html = (
        f'<div style="min-height:118px; padding:8px 0 10px 0;">'
        f'<div style="font-size:13px; color:#374151; margin-bottom:8px;">{label}</div>'
        f'<div style="font-size:27px; line-height:1.22; font-weight:500; color:#111827; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">{value or "&nbsp;"}</div>'
        f'<div style="font-size:12px; color:#6b7280; margin-top:10px; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">{caption or "&nbsp;"}</div>'
        f'</div>'
    )
    st.markdown(html, unsafe_allow_html=True)

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


def display_density_from_row(row: pd.Series) -> str:
    val = row.get("capacity_display", "")
    if val is not None and str(val).strip() not in ["", "nan", "None"]:
        return str(val).strip()
    density = safe_float(row.get("density_mb"))
    if density is None:
        return ""
    if density >= 1024 and abs(density / 1024 - round(density / 1024)) < 1e-9:
        return f"{density / 1024:g}Gb"
    return f"{density:g}Mb"


def display_voltage_from_row(row: pd.Series) -> str:
    val = row.get("voltage_range", "")
    if val is not None and str(val).strip() not in ["", "nan", "None"]:
        return str(val).strip().replace("~", "~")
    vmin = safe_float(row.get("vcc_min"))
    vmax = safe_float(row.get("vcc_max"))
    if vmin is None or vmax is None:
        return ""
    return f"{vmin:g}~{vmax:g}V"


def display_temperature_from_row(row: pd.Series) -> str:
    val = row.get("temperature", "")
    if val is not None and str(val).strip() not in ["", "nan", "None"]:
        return str(val).strip()
    tmin = safe_float(row.get("temp_min"))
    tmax = safe_float(row.get("temp_max"))
    if tmin is None or tmax is None:
        return ""
    return f"{int(tmin)}°C~{int(tmax)}°C"


def unique_non_empty(values) -> list[str]:
    out = []
    seen = set()
    for v in values:
        if v is None:
            continue
        sv = str(v).strip()
        if sv in ["", "nan", "None"]:
            continue
        key = normalize_model(sv)
        if key not in seen:
            out.append(sv)
            seen.add(key)
    return out


def build_selection_options(xtx_df: pd.DataFrame) -> dict[str, list[str]]:
    df = xtx_df.copy() if not xtx_df.empty else DEFAULT_XTX_PRODUCT_LIBRARY.copy()
    product_types = unique_non_empty([normalize_product_type(v) for v in df.get("product_type", pd.Series(dtype=str)).tolist()])
    product_types = unique_non_empty(product_types + [v for v in PRODUCT_TYPE_OPTIONS if v != "未指定"])
    capacities = unique_non_empty([display_density_from_row(row) for _, row in df.iterrows()])
    voltages = unique_non_empty([display_voltage_from_row(row) for _, row in df.iterrows()])
    packages = unique_non_empty(df.get("package", pd.Series(dtype=str)).tolist())
    package_sizes = unique_non_empty(df.get("package_size", pd.Series(dtype=str)).tolist())
    temperatures = unique_non_empty([display_temperature_from_row(row) for _, row in df.iterrows()])
    return {
        "product_type": [""] + product_types,
        "capacity": [""] + capacities,
        "voltage": [""] + voltages,
        "package": [""] + packages,
        "package_size": [""] + package_sizes,
        "temperature": [""] + temperatures,
    }


def find_option_index(options: list[str], value: str, kind: str = "text") -> int:
    if not options:
        return 0
    if value is None or str(value).strip() == "":
        return 0
    val = str(value).strip()
    if kind == "capacity":
        d0, _ = parse_density_from_text(val)
        for i, opt in enumerate(options):
            d1, _ = parse_density_from_text(opt)
            if d0 is not None and d1 is not None and abs(float(d0) - float(d1)) <= 0.001:
                return i
    elif kind == "voltage":
        v0 = parse_voltage_from_text(val)
        for i, opt in enumerate(options):
            v1 = parse_voltage_from_text(opt)
            if v0[0] is not None and v1[0] is not None and values_match(v0[0], v1[0]) and values_match(v0[1], v1[1]):
                return i
    elif kind == "temperature":
        t0 = parse_temperature_from_text(val)
        for i, opt in enumerate(options):
            t1 = parse_temperature_from_text(opt)
            if t0[0] is not None and t1[0] is not None and int(t0[0]) == int(t1[0]) and int(t0[1]) == int(t1[1]):
                return i
    else:
        key0 = normalize_product_type(val) if kind == "product_type" else normalize_model(val)
        for i, opt in enumerate(options):
            key1 = normalize_product_type(opt) if kind == "product_type" else normalize_model(opt)
            if key0 and key0 == key1:
                return i
    return 0


def build_spec_from_selection(original_spec: dict, selection: dict) -> tuple[dict, dict]:
    """用下拉选型结果覆盖自动解析结果，作为推荐输入。"""
    spec = original_spec.copy()
    confirm_map = {}

    pt = normalize_product_type(selection.get("product_type", ""))
    if pt:
        spec["product_type"] = pt
        spec["product_type_confidence"] = "选型确认"
        confirm_map["product_type"] = "已确认"
    else:
        confirm_map["product_type"] = "未确认"

    cap = selection.get("capacity", "")
    density_mb, display = parse_density_from_text(cap)
    if density_mb is not None:
        spec["density_mb"] = density_mb
        spec["capacity"] = display or cap
        spec["capacity_confidence"] = "选型确认"
        confirm_map["capacity"] = "已确认"
    else:
        confirm_map["capacity"] = "未确认"

    voltage = selection.get("voltage", "")
    vmin, vmax = parse_voltage_from_text(voltage)
    if vmin is not None and vmax is not None:
        vd = build_voltage_dict(vmin, vmax, "选型确认")
        spec.update({
            "vcc_min": vd["vcc_min"], "vcc_max": vd["vcc_max"],
            "voltage_range": vd["display"], "voltage_type": vd["voltage_type"],
            "voltage_confidence": vd["confidence"],
        })
        confirm_map["voltage"] = "已确认"
    else:
        confirm_map["voltage"] = "未确认"

    pkg = str(selection.get("package", "") or "").strip()
    if pkg:
        spec["package"] = pkg
        spec["package_confidence"] = "选型确认"
        confirm_map["package"] = "已确认"
    else:
        confirm_map["package"] = "未确认"

    pkg_size = str(selection.get("package_size", "") or "").strip()
    if pkg_size:
        spec["package_size"] = pkg_size
        spec["package_confidence"] = "选型确认"
        confirm_map["package_size"] = "已确认"
    else:
        confirm_map["package_size"] = "未确认"

    temp = selection.get("temperature", "")
    tmin, tmax = parse_temperature_from_text(temp)
    if tmin is not None and tmax is not None:
        td = build_temperature_dict(tmin, tmax, "选型确认")
        spec.update({
            "temp_min": td["temp_min"], "temp_max": td["temp_max"],
            "temperature": td["display"], "temp_grade": td["temp_grade"],
            "temperature_confidence": td["confidence"],
        })
        confirm_map["temperature"] = "已确认"
    else:
        confirm_map["temperature"] = "未确认"
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
    """V5 核心推荐规则：只要求 产品类型 + 容量 一致。

    电压范围、封装、封装尺寸、温度范围不再作为进入推荐列表的硬条件，
    而是在后续评分和风险点中体现，避免因为竞品电压识别错误导致没有推荐结果。
    """
    warnings = []
    if xtx_df.empty:
        return pd.DataFrame(), ["XTX 产品库为空。"]

    spec_pt = normalize_product_type(spec.get("product_type", ""))
    spec_density = spec.get("density_mb")

    missing = []
    if not spec_pt:
        missing.append("产品类型")
    if spec_density is None or spec_density == "":
        missing.append("容量")
    if missing:
        return pd.DataFrame(), [f"核心字段缺失：{'、'.join(missing)}。请在选型确认区补充后再推荐。"]

    df = xtx_df.copy()
    df["_norm_product_type"] = df["product_type"].apply(normalize_product_type)
    df["_density_mb"] = df["density_mb"].apply(safe_float)

    mask = (
        df["_norm_product_type"].eq(spec_pt)
        & df["_density_mb"].apply(lambda x: x is not None and abs(float(x) - float(spec_density)) <= 0.001)
    )
    matched = df[mask].drop(columns=["_norm_product_type", "_density_mb"], errors="ignore")
    if matched.empty:
        warnings.append(f"没有找到满足核心条件的 XTX 型号：产品类型={spec_pt}，容量={spec_density}Mb。")
    return matched, warnings


def calc_secondary_score(spec: dict, xtx_row: pd.Series, weights: dict[str, float]) -> tuple[float, float, str, list[str]]:
    # 产品类型 + 容量已经过滤，这里只做电压/封装/温度的排序评分和风险提示。
    core_score = weights.get("product_type", 20) + weights.get("density", 35)
    score = float(core_score)
    max_score = float(sum(weights.values())) if weights else 100.0
    risks = []

    spec_vmin = spec.get("vcc_min")
    spec_vmax = spec.get("vcc_max")
    xtx_vmin = safe_float(xtx_row.get("vcc_min"))
    xtx_vmax = safe_float(xtx_row.get("vcc_max"))

    if weights.get("voltage", 0) > 0:
        if spec_vmin is not None and spec_vmax is not None and xtx_vmin is not None and xtx_vmax is not None:
            if values_match(xtx_vmin, spec_vmin) and values_match(xtx_vmax, spec_vmax):
                score += weights["voltage"]
            elif xtx_vmin <= spec_vmin and xtx_vmax >= spec_vmax:
                score += weights["voltage"] * 0.7
                risks.append("XTX 电压范围覆盖竞品，但需确认系统电压")
            else:
                risks.append("电压范围不一致，需确认是否可替代")
        else:
            risks.append("电压范围未选择/未识别，需人工确认")

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
            risks.append("封装未选择/未识别，需人工确认")

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
            risks.append("温度范围未选择/未识别，需人工确认")

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

st.title("🔎 竞品规格书解析与 XTX 对标推荐工具 V8")
st.caption("V8：自动模式优先读取友商官网产品中心/详情页；PDF 只作为链接或本地 PDF 库备用；manual_value 默认清空，空值时使用 extracted_value 推荐。")

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
    max_pages = st.slider("每家官网产品中心最大检索页数", min_value=1, max_value=8, value=3, step=1, help="建议 2~3。V8 自动模式只查产品中心/产品详情页，不再下载或扫描官网 PDF。")
    pdf_library_files = st.file_uploader(
        "上传竞品 PDF 库（可多选，可选）",
        type=["pdf"],
        accept_multiple_files=True,
        help="官网产品中心搜不到时，工具会在这里上传的 PDF 库中按型号搜索。也可以把 PDF 提交到 GitHub 仓库的 pdf_library/ 文件夹。",
    )
    pdf_library_scan_pages_input = st.slider("PDF库型号扫描页数", min_value=0, max_value=300, value=0, step=20, help="0 表示扫描全文，找到型号即停止；PDF库很大时可设 80~120 提速。")
    pdf_library_max_files = st.slider("PDF库最多扫描文件数", min_value=5, max_value=200, value=50, step=5, help="PDF库文件很多时限制扫描数量，避免等待过久。文件名包含型号会优先命中，不受此限制。")
    pdf_parse_pages = st.slider("PDF详细解析页数", min_value=5, max_value=80, value=20, step=5, help="仅用于手动上传、输入PDF链接或PDF库命中后的字段抽取；官网产品中心命中时不解析PDF。")
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
pdf_library_df = clean_columns(maintenance.get("PDF_Library", pd.DataFrame()))

if "enabled" in company_df.columns:
    company_df = company_df[normalize_bool_series(company_df["enabled"])]
if "enabled" in weights_df.columns:
    weights_df = weights_df[normalize_bool_series(weights_df["enabled"])]

competitor_company_df = company_df[company_df["company_role"].astype(str).str.upper().eq("COMPETITOR")].copy()

# PDF 库：包含 GitHub 仓库 pdf_library/ 目录中的 PDF，以及侧边栏上传的多份 PDF。
pdf_library_items = collect_sheet_pdf_library(pdf_library_df) + collect_repo_pdf_library() + collect_uploaded_pdf_library(pdf_library_files)
xtx_company_df = company_df[company_df["company_role"].astype(str).str.upper().eq("XTX")].copy()

with st.expander("下载/查看维护文件格式", expanded=False):
    st.markdown(
        """
维护数据库 XLSX 建议包含五个 Sheet：

| Sheet | 用途 | 是否必须 |
|---|---|---|
| Company_Master | 公司官网白名单，区分芯天下和友商 | 必须 |
| XTX_Product_Library | 我司可对标产品库 | 必须 |
| Match_Weights | 匹配规则权重配置 | 建议 |
| History_Log | 历史分析记录 | 可选 |
| PDF_Library | 竞品规格书 PDF 库索引；实际 PDF 可放入仓库 pdf_library/ 或网页左侧上传 | 可选 |

V8 建议在 `Company_Master` 增加 `model_prefixes` 字段，例如 Boya 填 `BY,BY25,BY26`，这样自动检索时会先按型号前缀锁定厂商，速度更快。V8 自动模式不下载/扫描官网 PDF，只把官网 PDF 作为下载链接。
"""
    )
    template_bytes = to_excel_bytes(
        {
            "Company_Master": DEFAULT_COMPANY_MASTER,
            "XTX_Product_Library": DEFAULT_XTX_PRODUCT_LIBRARY,
            "Match_Weights": DEFAULT_MATCH_WEIGHTS,
            "History_Log": DEFAULT_HISTORY_LOG,
            "PDF_Library": pd.DataFrame(columns=["company_name", "model", "product_type", "density_mb", "capacity_display", "voltage_range", "package", "package_size", "temperature", "pdf_file_name", "pdf_source_url", "version", "last_modified", "confirm_status", "note"]),
        }
    )
    st.download_button(
        "下载维护数据库 XLSX 模板 V8",
        data=template_bytes,
        file_name="xtx_competitor_maintenance_template_v7.xlsx",
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

st.subheader("2. 选择信息来源")
source_mode = st.radio(
    "信息来源",
    ["自动从友商官网产品中心查找", "PDF库搜索", "输入 PDF 链接", "手动上传 PDF"],
    horizontal=True,
)

pdf_url_input = ""
uploaded_pdf = None
if source_mode == "输入 PDF 链接":
    pdf_url_input = st.text_input("PDF 链接", placeholder="https://...pdf")
elif source_mode == "手动上传 PDF":
    uploaded_pdf = st.file_uploader("上传竞品规格书 PDF", type=["pdf"])

run = st.button("开始解析/匹配", type="primary")

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

        elif source_mode in ["自动从友商官网产品中心查找", "自动从友商官网白名单查找"]:
            if competitor_company_df.empty:
                st.error("友商官网白名单为空，请检查 Company_Master 中 company_role=Competitor 的记录。")
                st.stop()

            search_df, search_note = infer_company_df_by_model(competitor_company_df, competitor_model, selected_company)
            st.write(search_note)
            st.write(f"本次检索厂商数：{len(search_df)}")

            st.write("正在检索友商官网产品中心 / 产品详情页...")
            product_page_df = crawl_product_pages_parallel(search_df, competitor_model, max_pages=max_pages)

            if not product_page_df.empty:
                best_page = product_page_df.iloc[0]
                spec = spec_from_product_page_row(best_page, selected_type=product_type, model=competitor_model)
                selected_pdf_url = str(best_page.get("datasheet_url", "") or best_page.get("page_url", ""))
                candidate_df = product_page_df
                st.write("已在官网产品中心/详情页匹配到目标型号；字段解析来自官网产品页，不下载、不扫描官网 PDF。")
                st.write(str(best_page.get("page_url", "")))
                if str(best_page.get("datasheet_url", "")):
                    st.write("检测到官网规格书链接，仅作为下载入口：")
                    st.write(str(best_page.get("datasheet_url", "")))
            else:
                st.warning("未在官网产品中心/详情页匹配到型号，开始搜索 PDF 库。")
                scan_pages = None if int(pdf_library_scan_pages_input) == 0 else int(pdf_library_scan_pages_input)
                hit, lib_df = search_pdf_library(pdf_library_items, competitor_model, max_scan_pages=scan_pages, max_files=pdf_library_max_files)
                candidate_df = lib_df
                if not hit:
                    st.warning("官网产品中心未命中，PDF库也未命中。建议：1）确认型号完整后重试；2）上传该型号规格书到 PDF库；3）手动输入 PDF 链接或手动上传 PDF。")
                    if not lib_df.empty:
                        st.dataframe(lib_df, use_container_width=True)
                    st.stop()
                pdf_bytes = hit.get("pdf_bytes", b"")
                if not pdf_bytes:
                    st.error("PDF库索引已命中，但未找到实际PDF文件，也没有可下载的 pdf_source_url。请把PDF放入仓库 pdf_library/ 文件夹、左侧上传，或在 PDF_Library 中填写可访问的 pdf_source_url。")
                    st.stop()
                selected_pdf_url = f"PDF库：{hit.get('file_name', '')}"
                st.write(f"已在 PDF库命中：{hit.get('file_name', '')}")
                st.write("解析 PDF库命中的 PDF 文本...")
                try:
                    pdf_text = extract_text_from_pdf_bytes(pdf_bytes, max_pages=pdf_parse_pages)
                except Exception as e:
                    st.error(f"PDF 解析失败：{e}")
                    st.stop()
                if not pdf_text.strip():
                    st.error("PDF 未提取到有效文本，可能是扫描版 PDF，需要后续增加 OCR。")
                    st.stop()
                spec = analyze_spec_text(pdf_text, selected_type=product_type, model=competitor_model)

        elif source_mode == "PDF库搜索":
            scan_pages = None if int(pdf_library_scan_pages_input) == 0 else int(pdf_library_scan_pages_input)
            hit, lib_df = search_pdf_library(pdf_library_items, competitor_model, max_scan_pages=scan_pages, max_files=pdf_library_max_files)
            candidate_df = lib_df
            if not hit:
                st.error("PDF库未命中目标型号。请先在左侧上传PDF库，或把PDF提交到GitHub仓库的 pdf_library/ 文件夹。")
                if not lib_df.empty:
                    st.dataframe(lib_df, use_container_width=True)
                st.stop()
            pdf_bytes = hit.get("pdf_bytes", b"")
            if not pdf_bytes:
                st.error("PDF库索引已命中，但未找到实际PDF文件，也没有可下载的 pdf_source_url。请把PDF放入仓库 pdf_library/ 文件夹、左侧上传，或在 PDF_Library 中填写可访问的 pdf_source_url。")
                st.stop()
            selected_pdf_url = f"PDF库：{hit.get('file_name', '')}"
            st.write(f"已在 PDF库命中：{hit.get('file_name', '')}")


        if source_mode in ["手动上传 PDF", "输入 PDF 链接", "PDF库搜索"]:
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

        analysis_id = f"{normalize_model(competitor_model)}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
        st.session_state["last_result"] = {
            "analysis_id": analysis_id,
            "spec": spec,
            "candidate_df": candidate_df,
            "selected_pdf_url": selected_pdf_url,
            "competitor_model": competitor_model,
            "selected_company": selected_company,
            "product_type": product_type,
            "review_df": spec_to_review_df(spec),
            "search_note": search_note,
        }
        status.update(label="解析/匹配完成", state="complete", expanded=False)

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
    m1, m2, m3, m4, m5 = st.columns([1.0, 1.0, 1.15, 1.7, 1.9])
    with m1:
        render_spec_card("产品类型", spec.get("product_type", ""), f"置信度：{spec.get('product_type_confidence', '')}")
    with m2:
        render_spec_card("容量", spec.get("capacity", ""), f"置信度：{spec.get('capacity_confidence', '')}")
    with m3:
        render_spec_card("电压范围", spec.get("voltage_range", ""), f"{spec.get('voltage_type', '')}，置信度：{spec.get('voltage_confidence', '')}")
    with m4:
        render_spec_card("封装形式", spec.get("package", ""), f"尺寸：{spec.get('package_size', '')}，置信度：{spec.get('package_confidence', '')}")
    with m5:
        render_spec_card("温度范围", spec.get("temperature", ""), f"{spec.get('temp_grade', '')}，置信度：{spec.get('temperature_confidence', '')}")

    if result.get("search_note"):
        st.info(result["search_note"])

    st.markdown("**信息来源：**")
    if str(selected_pdf_url).startswith("http"):
        st.markdown(f"[打开信息来源]({selected_pdf_url})")
    else:
        st.write(selected_pdf_url)

    if not candidate_df.empty:
        with st.expander("查看自动检索到的产品页 / PDF 候选链接和校验结果"):
            st.dataframe(candidate_df, use_container_width=True)

    st.divider()
    st.subheader("4. 选型确认 / 修正")
    st.caption("自动解析结果仅作为参考。manual_value 默认清空；为空时按 extracted_value 推荐，有选择时按 manual_value 推荐。下拉选项来自 XTX_Product_Library。")

    options = build_selection_options(xtx_df)
    review_rows = spec_to_review_df(spec).to_dict("records")

    option_map = {
        "product_type": (options["product_type"], "product_type", spec.get("product_type", "")),
        "capacity": (options["capacity"], "capacity", spec.get("capacity", "")),
        "voltage": (options["voltage"], "voltage", spec.get("voltage_range", "")),
        "package": (options["package"], "text", spec.get("package", "")),
        "package_size": (options["package_size"], "text", spec.get("package_size", "")),
        "temperature": (options["temperature"], "temperature", spec.get("temperature", "")),
    }

    # 表格样式：保留上一版 field_key/field_name/extracted_value/manual_value/confidence/confirm_status 的结构，
    # 但 manual_value 这一列改成每行不同的下拉选项。
    st.markdown(
        """
        <style>
        .xtx-manual-header {font-weight:600; color:#374151; background:#f3f4f6; padding:8px 6px; border-top:1px solid #e5e7eb; border-bottom:1px solid #e5e7eb;}
        .xtx-manual-cell {padding:8px 6px; border-bottom:1px solid #eef2f7; min-height:42px;}
        .xtx-small {font-size:13px; color:#374151; word-break:break-all;}
        </style>
        """,
        unsafe_allow_html=True,
    )

    hcols = st.columns([1.1, 1.4, 2.2, 2.3, 1.4, 1.1])
    for col, title in zip(hcols, ["field_key", "field_name", "extracted_value", "manual_value", "confidence", "confirm_status"]):
        col.markdown(f'<div class="xtx-manual-header">{title}</div>', unsafe_allow_html=True)

    analysis_id = result.get("analysis_id", normalize_model(competitor_model_result))
    manual_values = {}
    confirm_map = {}
    for row in review_rows:
        key = str(row.get("field_key", "")).strip()
        row_options, kind, current_value = option_map.get(key, ([""], "text", ""))
        # V8：manual_value 每次新解析默认清空。空值代表使用 extracted_value；只有用户显式选择后才覆盖。
        default_idx = 0
        rcols = st.columns([1.1, 1.4, 2.2, 2.3, 1.4, 1.1])
        rcols[0].markdown(f'<div class="xtx-manual-cell xtx-small">{key}</div>', unsafe_allow_html=True)
        rcols[1].markdown(f'<div class="xtx-manual-cell xtx-small">{row.get("field_name", "")}</div>', unsafe_allow_html=True)
        rcols[2].markdown(f'<div class="xtx-manual-cell xtx-small">{row.get("extracted_value", "")}</div>', unsafe_allow_html=True)
        with rcols[3]:
            selected = st.selectbox(
                label=f"manual_value_{key}",
                options=row_options,
                index=default_idx,
                key=f"manual_value_v8_{analysis_id}_{key}",
                label_visibility="collapsed",
            )
        manual_values[key] = selected
        status = "已确认" if str(selected).strip() else "未确认"
        confirm_map[key] = status
        rcols[4].markdown(f'<div class="xtx-manual-cell xtx-small">{row.get("confidence", "")}</div>', unsafe_allow_html=True)
        rcols[5].markdown(f'<div class="xtx-manual-cell xtx-small">{status}</div>', unsafe_allow_html=True)

    selection = {
        "product_type": manual_values.get("product_type", ""),
        "capacity": manual_values.get("capacity", ""),
        "voltage": manual_values.get("voltage", ""),
        "package": manual_values.get("package", ""),
        "package_size": manual_values.get("package_size", ""),
        "temperature": manual_values.get("temperature", ""),
    }
    confirmed_spec, auto_confirm_map = build_spec_from_selection(spec, selection)
    # 使用表格中显示的确认状态覆盖 build_spec_from_selection 生成的状态，便于导出历史。
    confirm_map.update(auto_confirm_map)

    with st.expander("查看用于推荐的最终规格", expanded=False):
        st.json({
            "product_type": confirmed_spec.get("product_type"),
            "capacity": confirmed_spec.get("capacity"),
            "density_mb": confirmed_spec.get("density_mb"),
            "voltage_range": confirmed_spec.get("voltage_range"),
            "package": confirmed_spec.get("package"),
            "package_size": confirmed_spec.get("package_size"),
            "temperature": confirmed_spec.get("temperature"),
            "xtx_product_library_rows": int(len(xtx_df)),
        })

    st.divider()
    st.subheader("5. XTX 对标型号推荐")
    st.caption("V8 推荐规则：XTX 对标型号只参考 XTX_Product_Library；manual_value 为空时使用 extracted_value，有内容时使用 manual_value；产品类型、容量两项核心条件一致即进入推荐列表。")
    recommend_df, rec_warnings = recommend_xtx(confirmed_spec, xtx_df, weights_df, top_n=10)
    for w in rec_warnings:
        st.warning(w)

    if recommend_df.empty:
        st.info(f"当前没有满足核心条件的 XTX 对标型号。请检查：是否已上传/加载最新版维护库、产品类型和容量是否选择正确。当前 XTX_Product_Library 行数：{len(xtx_df)}。")
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

tab1, tab2, tab3, tab4, tab5 = st.tabs(["公司白名单", "XTX产品库", "PDF库索引", "匹配权重", "历史记录"])
with tab1:
    st.markdown("**芯天下/XTX 公司记录**")
    st.dataframe(xtx_company_df, use_container_width=True)
    st.markdown("**友商/竞品公司记录**")
    st.dataframe(competitor_company_df, use_container_width=True)
with tab2:
    st.dataframe(xtx_df, use_container_width=True)
with tab3:
    st.dataframe(pdf_library_df, use_container_width=True)
    st.caption(f"当前可搜索 PDF库文件数：{len(pdf_library_items)}。来源包括维护表 PDF_Library、仓库 pdf_library/ 文件夹、左侧上传PDF。")
with tab4:
    st.dataframe(weights_df, use_container_width=True)
with tab5:
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
            "PDF_Library": pdf_library_df,
        }
    ),
    file_name="xtx_competitor_maintenance_current_v7.xlsx",
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
