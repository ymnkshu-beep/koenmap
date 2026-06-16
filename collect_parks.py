#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
collect_parks.py — 全国自治体公園一覧オープンデータ収集スクリプト
"""

import csv
import hashlib
import io
import json
import logging
import os
import re
import time
import unicodedata
from datetime import datetime, timezone
from typing import Optional
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import chardet
import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ─── 設定 ────────────────────────────────────────────────────────────────────
UA = "park-data-collector/1.0"
REQUEST_INTERVAL = 1.2          # 秒
REQUEST_TIMEOUT  = 30
OUTPUT_CSV       = "parks_japan_all.csv"
LOG_CSV          = "collection_log.csv"
REPORT_MD        = "collection_report.md"

OUT_COLS = [
    "source_id", "pref_name", "city_name", "park_name", "address",
    "latitude", "longitude", "park_type", "area_m2",
    "source_org", "source_url", "license", "retrieved_at", "original_filename",
]

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

session = requests.Session()
session.headers.update({"User-Agent": UA})
session.verify = False
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

robots_cache: dict[str, RobotFileParser] = {}
last_request: dict[str, float] = {}


# ─── ユーティリティ ───────────────────────────────────────────────────────────
def normalize_text(s) -> str:
    if not isinstance(s, str):
        s = "" if s is None or (isinstance(s, float) and s != s) else str(s)
    s = s.strip()
    s = unicodedata.normalize("NFKC", s)           # 全角→半角、全角スペース→半角
    s = re.sub(r"[０-９]", lambda m: chr(ord(m.group()) - 0xFEE0), s)  # 全角数字→半角
    return s


def robots_allowed(url: str) -> bool:
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    if base not in robots_cache:
        rp = RobotFileParser()
        rp.set_url(f"{base}/robots.txt")
        try:
            rp.read()
        except Exception:
            try:
                import ssl, urllib.request
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                with urllib.request.urlopen(f"{base}/robots.txt", context=ctx, timeout=10) as resp:
                    rp.parse(resp.read().decode("utf-8", errors="ignore").splitlines())
            except Exception:
                rp = None
        robots_cache[base] = rp
    rp = robots_cache[base]
    if rp is None:
        return True
    return rp.can_fetch(UA, url)


def polite_get(url: str, **kwargs) -> Optional[requests.Response]:
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    elapsed = time.time() - last_request.get(base, 0)
    if elapsed < REQUEST_INTERVAL:
        time.sleep(REQUEST_INTERVAL - elapsed)
    last_request[base] = time.time()
    try:
        r = session.get(url, timeout=REQUEST_TIMEOUT, **kwargs)
        r.raise_for_status()
        return r
    except Exception as e:
        log.warning(f"GET failed {url}: {e}")
        return None


def detect_encoding(raw: bytes) -> str:
    det = chardet.detect(raw)
    enc = det.get("encoding") or "utf-8"
    for candidate in [enc, "utf-8", "shift_jis", "euc-jp"]:
        try:
            raw.decode(candidate)
            return candidate
        except Exception:
            continue
    return "utf-8"


def read_csv_bytes(raw: bytes) -> Optional[pd.DataFrame]:
    enc = detect_encoding(raw)
    for e in [enc, "utf-8-sig", "utf-8", "shift_jis", "cp932", "euc-jp"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=e, dtype=str, low_memory=False)
        except Exception:
            continue
    return None


def read_excel_bytes(raw: bytes, filename: str = "") -> Optional[pd.DataFrame]:
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        # 「公園」「park」を優先
        preferred = [s for s in sheet_names
                     if re.search(r"公園|park", s, re.IGNORECASE)]
        target = preferred[0] if preferred else sheet_names[0]
        df = pd.read_excel(io.BytesIO(raw), sheet_name=target, dtype=str, engine="openpyxl")
        return df
    except Exception as e:
        log.warning(f"Excel read error {filename}: {e}")
    try:
        return pd.read_excel(io.BytesIO(raw), dtype=str, engine="xlrd")
    except Exception:
        return None


def read_geojson_bytes(raw: bytes) -> Optional[pd.DataFrame]:
    try:
        gj = json.loads(raw.decode(detect_encoding(raw)))
        features = gj.get("features", [])
        rows = []
        for f in features:
            props = f.get("properties") or {}
            geom = f.get("geometry") or {}
            coords = geom.get("coordinates")
            if coords and geom.get("type") == "Point":
                props["__lon"] = coords[0]
                props["__lat"] = coords[1]
            rows.append(props)
        return pd.DataFrame(rows) if rows else None
    except Exception as e:
        log.warning(f"GeoJSON read error: {e}")
        return None


def guess_columns(df: pd.DataFrame, col_type: str) -> Optional[str]:
    """列名からpark_name / address候補を推定"""
    patterns = {
        "park_name": [
            r"公園名|park.*name|名称|施設名|こうえんめい|park_name|kouen|名前",
        ],
        "address": [
            r"住所|所在地|address|addr|所在|ちゅうしょ|location",
        ],
        "latitude": [
            r"緯度|lat(?:itude)?|^y$|y座標",
        ],
        "longitude": [
            r"経度|lon(?:gitude)?|lng|^x$|x座標",
        ],
        "park_type": [
            r"公園種別|種別|park.*type|kind|分類",
        ],
        "area": [
            r"面積|area|㎡|m2",
        ],
    }
    pat_list = patterns.get(col_type, [])
    for col in df.columns:
        for pat in pat_list:
            if re.search(pat, str(col), re.IGNORECASE):
                return col
    return None


def df_to_park_rows(
    df: pd.DataFrame,
    pref_name: str,
    city_name: str,
    source_id: str,
    source_org: str,
    source_url: str,
    license_: str,
    retrieved_at: str,
    original_filename: str,
) -> list[dict]:
    """DataFrame → 正規化済みの公園行リスト"""
    # 列名正規化
    df.columns = [normalize_text(str(c)) for c in df.columns]

    name_col = guess_columns(df, "park_name")
    addr_col = guess_columns(df, "address")

    if not name_col or not addr_col:
        log.warning(f"  必須列 missing: name={name_col} addr={addr_col} (cols={list(df.columns)[:10]})")
        return []

    lat_col  = guess_columns(df, "latitude")
    lon_col  = guess_columns(df, "longitude")
    type_col = guess_columns(df, "park_type")
    area_col = guess_columns(df, "area")

    rows = []
    for _, row in df.iterrows():
        pname = normalize_text(row.get(name_col, ""))
        addr  = normalize_text(row.get(addr_col, ""))
        if not pname or not addr:
            continue

        entry = {
            "source_id":        source_id,
            "pref_name":        pref_name,
            "city_name":        city_name,
            "park_name":        pname,
            "address":          addr,
            "latitude":         normalize_text(row.get(lat_col, ""))  if lat_col  else "",
            "longitude":        normalize_text(row.get(lon_col, ""))  if lon_col  else "",
            "park_type":        normalize_text(row.get(type_col, "")) if type_col else "",
            "area_m2":          normalize_text(row.get(area_col, "")) if area_col else "",
            "source_org":       source_org,
            "source_url":       source_url,
            "license":          license_,
            "retrieved_at":     retrieved_at,
            "original_filename": original_filename,
        }
        rows.append(entry)
    return rows


# ─── データソース定義 ─────────────────────────────────────────────────────────
SOURCES = [
    # ── 方式A: 直接URL取得 ──────────────────────────────────────────────────
    {
        "source_id": "A01_kanagawa",
        "pref": "神奈川県", "city": "（全33市町村）",
        "org": "神奈川県",
        "url": "https://www.pref.kanagawa.jp/docs/b8k/cnt/f536260/index.html",
        "license": "CC BY 4.0",
        "type": "html_link",
        "link_pattern": r"\.(csv|xlsx?)$",
    },
    {
        "source_id": "A02_tokyo_catalog",
        "pref": "東京都", "city": "（複数区市町村）",
        "org": "東京都",
        "url": "https://catalog.data.metro.tokyo.lg.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92%E4%B8%80%E8%A6%A7&rows=50",
        "license": "CC BY",
        "type": "ckan_api",
        "ckan_base": "https://catalog.data.metro.tokyo.lg.jp",
    },
    {
        "source_id": "A03_osaka_pref",
        "pref": "大阪府", "city": "（府営公園）",
        "org": "大阪府",
        "url": "https://data.bodik.jp/dataset/270008_park_list",
        "license": "CC BY 4.0",
        "type": "bodik_dataset",
    },
    {
        "source_id": "A04_kumamoto_pref",
        "pref": "熊本県", "city": "",
        "org": "熊本県",
        "url": "https://data.bodik.jp/dataset/430005_00174",
        "license": "CC BY 4.0",
        "type": "bodik_dataset",
    },
    {
        "source_id": "A05_nagoya",
        "pref": "愛知県", "city": "名古屋市",
        "org": "名古屋市",
        "url": "https://data.bodik.jp/dataset/231002_171310000_001",
        "license": "CC BY 4.0",
        "type": "bodik_dataset",
    },
    {
        "source_id": "A06_yokohama",
        "pref": "神奈川県", "city": "横浜市",
        "org": "横浜市",
        "url": "https://www.city.yokohama.lg.jp/kurashi/machizukuri-kankyo/midori-koen/koen/kouen.html",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"\.(xlsx?)$",
    },
    {
        "source_id": "A07_saitama",
        "pref": "埼玉県", "city": "さいたま市",
        "org": "さいたま市",
        "url": "https://www.city.saitama.lg.jp/004/006/003/003/p095465.html",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"\.csv$",
    },
    {
        "source_id": "A08_chiba",
        "pref": "千葉県", "city": "千葉市",
        "org": "千葉市",
        "url": "https://www.city.chiba.jp/sogoseisaku/shichokoshitsu/kohokocho/map_opendata.html",
        "license": "CC 2.1 Japan",
        "type": "html_link",
        "link_pattern": r"公園.*\.csv$",
    },
    {
        "source_id": "A09_fukuoka",
        "pref": "福岡県", "city": "福岡市",
        "org": "福岡市",
        "url": "https://data.bodik.jp/dataset/401307_tosikouen",
        "license": "CC BY",
        "type": "bodik_dataset",
    },
    {
        "source_id": "A10_kitakyushu",
        "pref": "福岡県", "city": "北九州市",
        "org": "北九州市",
        "url": "https://data.bodik.jp/dataset/401005_toshikoendaicholist",
        "license": "CC BY",
        "type": "bodik_dataset",
    },
    {
        "source_id": "A11_sakai",
        "pref": "大阪府", "city": "堺市",
        "org": "堺市",
        "url": "https://data.bodik.jp/dataset/271403_sakai_park",
        "license": "CC BY 4.0",
        "type": "bodik_dataset",
    },
    {
        "source_id": "A12_himeji",
        "pref": "兵庫県", "city": "姫路市",
        "org": "姫路市",
        "url": "https://city.himeji.gkan.jp/gkan/api/3/action/package_show?id=himejikoen",
        "license": "CC BY 4.0",
        "type": "ckan_package_api",
        "ckan_base": "https://city.himeji.gkan.jp",
    },
    {
        "source_id": "A13_hamamatsu",
        "pref": "静岡県", "city": "浜松市",
        "org": "浜松市",
        "url": "https://opendata.pref.shizuoka.jp/dataset/11964.html",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"\.(csv|xlsx?)$",
    },
    {
        "source_id": "A14_shizuoka",
        "pref": "静岡県", "city": "静岡市",
        "org": "静岡市",
        "url": "https://opendata.pref.shizuoka.jp/dataset/12366.html",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"\.(csv|xlsx?)$",
    },
    {
        "source_id": "A15_okazaki",
        "pref": "愛知県", "city": "岡崎市",
        "org": "岡崎市",
        "url": "https://www.city.okazaki.lg.jp/shisei/opendata/1005809.html",
        "license": "CC BY 4.0",
        "type": "html_link",
        "link_pattern": r"公園.*\.csv$",
    },
    {
        "source_id": "A16_nara",
        "pref": "奈良県", "city": "奈良市",
        "org": "奈良市",
        "url": "https://www.city.nara.lg.jp/soshiki/115/50044.html",
        "license": "CC BY 2.1 Japan",
        "type": "html_link",
        "link_pattern": r"\.csv$",
    },
    {
        "source_id": "A17_utsunomiya",
        "pref": "栃木県", "city": "宇都宮市",
        "org": "宇都宮市",
        "url": "https://catalog.city.utsunomiya.tochigi.jp/api/3/action/package_show?id=kouenryokuti",
        "license": "CC BY",
        "type": "ckan_package_api",
        "ckan_base": "https://catalog.city.utsunomiya.tochigi.jp",
    },
    {
        "source_id": "A18_nasushiobara",
        "pref": "栃木県", "city": "那須塩原市",
        "org": "那須塩原市",
        "url": "https://opendata-nasu.opendatastack.jp/api/3/action/package_show?id=shisetu-kouenn-tosisetibi",
        "license": "CC BY 2.1 Japan",
        "type": "ckan_package_api",
        "ckan_base": "https://opendata-nasu.opendatastack.jp",
    },
    {
        "source_id": "A19_niigata",
        "pref": "新潟県", "city": "新潟市",
        "org": "新潟市",
        "url": "https://www.city.niigata.lg.jp/shisei/seisaku/it/open-data/opendata-gis/od-gis_kurashibosai/od-gis_park.html",
        "license": "CC BY 2.1 Japan",
        "type": "html_link",
        "link_pattern": r"\.(csv|geojson)$",
    },
    {
        "source_id": "A20_funabashi",
        "pref": "千葉県", "city": "船橋市",
        "org": "船橋市",
        "url": "https://odcs.bodik.jp/122041/",
        "license": "CC BY 2.1 JP",
        "type": "odcs_page",
        "keyword": "公園",
    },
    {
        "source_id": "A21_ichikawa",
        "pref": "千葉県", "city": "市川市",
        "org": "市川市",
        "url": "https://www.city.ichikawa.lg.jp/pla01/opendata.html",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"公園.*\.csv$",
    },
    {
        "source_id": "A22_sendai",
        "pref": "宮城県", "city": "仙台市",
        "org": "仙台市",
        "url": "https://miyagi.dataeye.jp/api/package_show/515",
        "license": "CC BY 4.0",
        "type": "dataeye_api",
        "fallback_html": "https://miyagi.dataeye.jp/datasets/515",
    },
    {
        "source_id": "A23_nagahama",
        "pref": "滋賀県", "city": "長浜市",
        "org": "長浜市",
        "url": "https://data.bodik.jp/dataset/252034_park",
        "license": "CC BY",
        "type": "bodik_dataset",
    },
    {
        "source_id": "A24_kakogawa",
        "pref": "兵庫県", "city": "加古川市",
        "org": "加古川市",
        "url": "https://opendata-api-kakogawa.jp/ckan/api/3/action/package_show?id=urbanpark",
        "license": "CC BY",
        "type": "ckan_package_api",
        "ckan_base": "https://opendata-api-kakogawa.jp",
    },
    {
        "source_id": "A25_hiroshima",
        "pref": "広島県", "city": "広島市",
        "org": "広島市",
        "url": "https://hiroshima-opendata.dataeye.jp/api/package_search?q=%E5%85%AC%E5%9C%92&organization=hiroshima",
        "license": "CC BY 2.1 Japan",
        "type": "dataeye_api",
        "fallback_html": "https://hiroshima-opendata.dataeye.jp/",
    },
    {
        "source_id": "A26_okayama",
        "pref": "岡山県", "city": "岡山市",
        "org": "岡山市",
        "url": "https://www.okayama-opendata.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&fq=organization:130009",
        "license": "CC BY 4.0",
        "type": "ckan_api",
        "ckan_base": "https://www.okayama-opendata.jp",
    },
    {
        "source_id": "A27_matsuyama",
        "pref": "愛媛県", "city": "松山市",
        "org": "松山市",
        "url": "https://www.city.matsuyama.ehime.jp/shisei/opendata/metadata/kouen.html",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"\.csv$",
    },
    {
        "source_id": "A28_takamatsu",
        "pref": "香川県", "city": "高松市",
        "org": "高松市",
        "url": "https://opendata.smartcity-takamatsu.jp/ckan/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&fq=organization:takamatsu&rows=20",
        "license": "CC BY 4.0",
        "type": "ckan_api",
        "ckan_base": "https://opendata.smartcity-takamatsu.jp",
    },
    {
        "source_id": "A29_kochi",
        "pref": "高知県", "city": "高知市",
        "org": "高知市",
        "url": "https://www.city.kochi.kochi.jp/soshiki/80/kochicity-opendata.html",
        "license": "CC BY 4.0",
        "type": "html_link",
        "link_pattern": r"公園.*\.csv$",
    },
    {
        "source_id": "A30_kumamoto_city",
        "pref": "熊本県", "city": "熊本市",
        "org": "熊本市",
        "url": "https://www.city.kumamoto.jp/dynamic/opendata/pub/detail.aspx?c_id=38&id=37",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"\.csv$",
    },
    {
        "source_id": "A31_naha",
        "pref": "沖縄県", "city": "那覇市",
        "org": "那覇市",
        "url": "https://odcs.bodik.jp/472018/",
        "license": "CC BY",
        "type": "odcs_page",
        "keyword": "公園",
    },
    {
        "source_id": "A32_okinawa_city",
        "pref": "沖縄県", "city": "沖縄市",
        "org": "沖縄市",
        "url": "https://www.city.okinawa.okinawa.jp/opendata/index.php?p=1_1&asc=od_data_type&displayedresults=100",
        "license": "CC BY",
        "type": "html_link",
        "link_pattern": r"公園.*\.csv$",
    },
    {
        "source_id": "A33_sapporo",
        "pref": "北海道", "city": "札幌市",
        "org": "札幌市",
        "url": "https://ckan.pf-sapporo.jp/api/3/action/package_show?id=kouen_ryokuchi",
        "license": "CC BY 4.0",
        "type": "ckan_package_api",
        "ckan_base": "https://ckan.pf-sapporo.jp",
    },
    {
        "source_id": "A34_osaka_city",
        "pref": "大阪府", "city": "大阪市",
        "org": "大阪市",
        "url": "https://www.geospatial.jp/ckan/api/3/action/package_show?id=mapnavi-city-osaka",
        "license": "CC BY",
        "type": "ckan_package_api",
        "ckan_base": "https://www.geospatial.jp",
        "resource_filter": "公園",
    },
    {
        "source_id": "A35_nagasaki",
        "pref": "長崎県", "city": "長崎市",
        "org": "長崎市",
        "url": "https://data.bodik.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&fq=organization:422011&rows=20",
        "license": "CC BY 4.0",
        "type": "ckan_api",
        "ckan_base": "https://data.bodik.jp",
    },
    {
        "source_id": "A36_kyoto",
        "pref": "京都府", "city": "京都市",
        "org": "京都市",
        "url": "https://data.city.kyoto.lg.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&rows=20",
        "license": "CC BY",
        "type": "ckan_api",
        "ckan_base": "https://data.city.kyoto.lg.jp",
    },
    # ── 方式B: プラットフォームクロール ────────────────────────────────────────
    {
        "source_id": "B01_bodik_ckan",
        "pref": "", "city": "",
        "org": "BODIK CKAN（全国自治体）",
        "url": "https://data.bodik.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92+%E4%B8%80%E8%A6%A7&res_format=CSV&rows=100&start=0",
        "license": "CC BY",
        "type": "bodik_ckan_bulk",
    },
    {
        "source_id": "B02_bodik_ckan_xlsx",
        "pref": "", "city": "",
        "org": "BODIK CKAN（全国自治体・XLSX）",
        "url": "https://data.bodik.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92+%E4%B8%80%E8%A6%A7&res_format=XLSX&rows=100&start=0",
        "license": "CC BY",
        "type": "bodik_ckan_bulk",
    },
    {
        "source_id": "B03_searchckan",
        "pref": "", "city": "",
        "org": "データカタログ横断検索（全国）",
        "url": "https://search.ckan.jp/backend/api/package_search?q=%E9%83%BD%E5%B8%82%E5%85%AC%E5%9C%92%E4%B8%80%E8%A6%A7&rows=50",
        "license": "CC BY",
        "type": "ckan_api_bulk",
        "ckan_base": "https://search.ckan.jp",
    },
    {
        "source_id": "B04_geospatial",
        "pref": "", "city": "",
        "org": "G空間情報センター（全国）",
        "url": "https://www.geospatial.jp/ckan/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&tags=%E9%83%BD%E5%B8%82%E5%85%AC%E5%9C%92&res_format=CSV&rows=50",
        "license": "CC BY",
        "type": "ckan_api_bulk",
        "ckan_base": "https://www.geospatial.jp",
    },
    {
        "source_id": "B05_miyagi",
        "pref": "宮城県", "city": "",
        "org": "宮城県共同ポータル",
        "url": "https://miyagi.dataeye.jp/api/package_search?q=%E5%85%AC%E5%9C%92&limit=50",
        "license": "CC BY 4.0",
        "type": "dataeye_api",
        "fallback_html": "https://miyagi.dataeye.jp/datasets",
    },
    {
        "source_id": "B06_shizuoka_pref",
        "pref": "静岡県", "city": "",
        "org": "ふじのくにオープンデータ（静岡県）",
        "url": "https://opendata.pref.shizuoka.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&rows=50",
        "license": "CC BY",
        "type": "ckan_api_bulk",
        "ckan_base": "https://opendata.pref.shizuoka.jp",
    },
    {
        "source_id": "B07_gifu",
        "pref": "岐阜県", "city": "",
        "org": "岐阜県オープンデータカタログ",
        "url": "https://gifu-opendata.pref.gifu.lg.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&rows=50",
        "license": "CC BY",
        "type": "ckan_api_bulk",
        "ckan_base": "https://gifu-opendata.pref.gifu.lg.jp",
    },
    {
        "source_id": "B08_yamaguchi",
        "pref": "山口県", "city": "",
        "org": "山口県オープンデータカタログ",
        "url": "https://yamaguchi-opendata.jp/ckan/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&rows=50",
        "license": "CC BY",
        "type": "ckan_api_bulk",
        "ckan_base": "https://yamaguchi-opendata.jp",
    },
    {
        "source_id": "B09_okayama_pref",
        "pref": "岡山県", "city": "",
        "org": "おかやまオープンデータカタログ",
        "url": "https://www.okayama-opendata.jp/api/3/action/package_search?q=%E5%85%AC%E5%9C%92&rows=50",
        "license": "CC BY 4.0",
        "type": "ckan_api_bulk",
        "ckan_base": "https://www.okayama-opendata.jp",
    },
    {
        "source_id": "B10_hiroshima_pref",
        "pref": "広島県", "city": "",
        "org": "広島広域都市圏オープンデータ",
        "url": "https://hiroshima-opendata.dataeye.jp/api/package_search?q=%E5%85%AC%E5%9C%92&limit=50",
        "license": "CC BY 2.1 Japan",
        "type": "dataeye_api",
        "fallback_html": "https://hiroshima-opendata.dataeye.jp/",
    },
]

# ─── ダウンローダ ─────────────────────────────────────────────────────────────
def download_file(url: str) -> Optional[tuple[bytes, str]]:
    """(bytes, filename) or None"""
    if not robots_allowed(url):
        log.warning(f"robots.txt disallow: {url}")
        return None
    r = polite_get(url)
    if r is None:
        return None
    fname = urlparse(url).path.split("/")[-1] or "data"
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r'filename[^;=\n]*=(["\']?)([^"\'\n;]+)\1', cd)
    if m:
        fname = m.group(2).strip()
    return r.content, fname


def parse_file(raw: bytes, fname: str) -> Optional[pd.DataFrame]:
    ext = os.path.splitext(fname)[1].lower()
    if ext in (".xlsx", ".xls"):
        return read_excel_bytes(raw, fname)
    if ext == ".geojson" or fname.endswith(".geojson"):
        return read_geojson_bytes(raw)
    # CSV or unknown → try CSV
    return read_csv_bytes(raw)


def collect_from_url(url: str, source: dict) -> list[dict]:
    """単一URLからファイルDL→パース→行リスト"""
    result = download_file(url)
    if not result:
        return []
    raw, fname = result
    df = parse_file(raw, fname)
    if df is None or df.empty:
        log.warning(f"  parse failed or empty: {fname}")
        return []
    log.info(f"  parsed {fname}: {len(df)} rows × {len(df.columns)} cols")
    return df_to_park_rows(
        df,
        pref_name=source["pref"],
        city_name=source["city"],
        source_id=source["source_id"],
        source_org=source["org"],
        source_url=url,
        license_=source["license"],
        retrieved_at=datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        original_filename=fname,
    )


def extract_download_links(html: str, base_url: str, pattern: str) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        full = urljoin(base_url, href)
        if re.search(pattern, full, re.IGNORECASE):
            links.append(full)
    return links


def process_ckan_resources(resources: list[dict], source: dict) -> list[dict]:
    """CKAN resource リストからCSV/XLSX/GeoJSONを取得"""
    rows = []
    resource_filter = source.get("resource_filter", "")
    for res in resources:
        fmt  = (res.get("format") or "").lower()
        rurl = res.get("url") or res.get("download_url") or ""
        rname = res.get("name") or ""
        if not rurl:
            continue
        if fmt not in ("csv", "xlsx", "xls", "excel", "geojson", "json") and \
           not re.search(r"\.(csv|xlsx?|geojson)$", rurl, re.IGNORECASE):
            continue
        if resource_filter and resource_filter not in rname and resource_filter not in rurl:
            continue
        log.info(f"    DL resource: {rname} [{fmt}] {rurl[:80]}")
        rows.extend(collect_from_url(rurl, source))
    return rows


# ─── ソース別ハンドラ ────────────────────────────────────────────────────────
def handle_bodik_dataset(source: dict, collected_rows: list, log_rows: list):
    """BODIK の個別データセットページからCSV/XLSXリンクを抽出"""
    page_url = source["url"]
    log.info(f"[{source['source_id']}] bodik_dataset {page_url}")
    r = polite_get(page_url)
    if r is None:
        log_rows.append(make_log(source, "error", "ページ取得失敗", 0))
        return
    # CKAN API /api/3/action/package_show を利用
    parsed = urlparse(page_url)
    pkg_id = parsed.path.rstrip("/").split("/")[-1]
    api_url = f"{parsed.scheme}://{parsed.netloc}/api/3/action/package_show?id={pkg_id}"
    ar = polite_get(api_url)
    if ar:
        try:
            data = ar.json()
            resources = data["result"]["resources"]
            rows = process_ckan_resources(resources, source)
            if rows:
                collected_rows.extend(rows)
                log_rows.append(make_log(source, "success", "CKAN API", len(rows)))
                return
        except Exception as e:
            log.warning(f"  CKAN API failed: {e}")
    # フォールバック: HTML解析
    links = extract_download_links(r.text, page_url, r"\.(csv|xlsx?)$")
    rows = []
    for lnk in links[:5]:
        rows.extend(collect_from_url(lnk, source))
    if rows:
        collected_rows.extend(rows)
        log_rows.append(make_log(source, "success", "HTML link", len(rows)))
    else:
        log_rows.append(make_log(source, "error", "リソース取得失敗", 0))


def handle_html_link(source: dict, collected_rows: list, log_rows: list):
    log.info(f"[{source['source_id']}] html_link {source['url']}")
    r = polite_get(source["url"])
    if r is None:
        log_rows.append(make_log(source, "error", "ページ取得失敗", 0))
        return
    links = extract_download_links(r.text, source["url"], source.get("link_pattern", r"\.(csv|xlsx?)$"))
    # link_pattern が緩い場合: 正規化してフィルタ
    if not links:
        links = extract_download_links(r.text, source["url"], r"\.(csv|xlsx?|geojson)$")
    rows = []
    for lnk in links[:10]:
        log.info(f"  DL: {lnk}")
        rows.extend(collect_from_url(lnk, source))
    if rows:
        collected_rows.extend(rows)
        log_rows.append(make_log(source, "success", "HTML link", len(rows)))
    else:
        log_rows.append(make_log(source, "skipped", f"ダウンロードリンクなし ({len(links)} links found)", 0))


def handle_ckan_api(source: dict, collected_rows: list, log_rows: list):
    log.info(f"[{source['source_id']}] ckan_api {source['url']}")
    r = polite_get(source["url"])
    if r is None:
        log_rows.append(make_log(source, "error", "API取得失敗", 0))
        return
    try:
        data = r.json()
        results = data.get("result", {}).get("results", data.get("results", []))
        if not results and "result" in data and isinstance(data["result"], dict):
            results = data["result"].get("results", [])
    except Exception as e:
        log_rows.append(make_log(source, "error", f"JSON解析失敗: {e}", 0))
        return
    rows = []
    for pkg in results:
        resources = pkg.get("resources", [])
        pkg_source = dict(source)
        pkg_source["city"] = source.get("city") or pkg.get("organization", {}).get("title", "")
        pkg_source["pref"] = source.get("pref", "")
        rows.extend(process_ckan_resources(resources, pkg_source))
    if rows:
        collected_rows.extend(rows)
        log_rows.append(make_log(source, "success", "CKAN search API", len(rows)))
    else:
        log_rows.append(make_log(source, "skipped", "対象リソース無し", 0))


def handle_ckan_package_api(source: dict, collected_rows: list, log_rows: list):
    log.info(f"[{source['source_id']}] ckan_package_api {source['url']}")
    r = polite_get(source["url"])
    if r is None:
        log_rows.append(make_log(source, "error", "API取得失敗", 0))
        return
    try:
        data = r.json()
        resources = data["result"]["resources"]
    except Exception as e:
        log_rows.append(make_log(source, "error", f"JSON解析失敗: {e}", 0))
        return
    rows = process_ckan_resources(resources, source)
    if rows:
        collected_rows.extend(rows)
        log_rows.append(make_log(source, "success", "CKAN package API", len(rows)))
    else:
        log_rows.append(make_log(source, "skipped", "対象リソース無し", 0))


def handle_bodik_ckan_bulk(source: dict, collected_rows: list, log_rows: list):
    """BODIK CKAN 横断検索（重複除去のため source_id に pkg_id を付与）"""
    log.info(f"[{source['source_id']}] bodik_ckan_bulk {source['url']}")
    r = polite_get(source["url"])
    if r is None:
        log_rows.append(make_log(source, "error", "API取得失敗", 0))
        return
    try:
        data = r.json()
        results = data["result"]["results"]
    except Exception as e:
        log_rows.append(make_log(source, "error", f"JSON解析失敗: {e}", 0))
        return
    total_rows = 0
    already = {s["source_id"] for s in SOURCES}
    for pkg in results:
        pkg_id = pkg.get("name", "")
        org = pkg.get("organization", {}).get("title", "不明")
        pref = ""
        city = org
        sid = f"{source['source_id']}_{pkg_id}"
        sub = dict(source)
        sub["source_id"] = sid
        sub["pref"] = pref
        sub["city"] = city
        sub["org"] = org
        sub["license"] = (pkg.get("license_title") or source["license"])
        rows = process_ckan_resources(pkg.get("resources", []), sub)
        total_rows += len(rows)
        collected_rows.extend(rows)
    log_rows.append(make_log(source, "success", f"BODIK bulk {len(results)} pkgs", total_rows))


def handle_ckan_api_bulk(source: dict, collected_rows: list, log_rows: list):
    handle_ckan_api(source, collected_rows, log_rows)


def handle_dataeye_api(source: dict, collected_rows: list, log_rows: list):
    log.info(f"[{source['source_id']}] dataeye_api {source['url']}")
    r = polite_get(source["url"])
    rows = []
    if r:
        try:
            data = r.json()
            # dataeye API は構造が異なる場合あり
            results = data if isinstance(data, list) else \
                      data.get("results", data.get("datasets", []))
            for pkg in results:
                resources = pkg.get("resources", pkg.get("files", []))
                for res in resources:
                    rurl = res.get("url") or res.get("download_url", "")
                    if rurl and re.search(r"\.(csv|xlsx?)$", rurl, re.IGNORECASE):
                        rows.extend(collect_from_url(rurl, source))
        except Exception as e:
            log.warning(f"  dataeye JSON parse error: {e}")
    # フォールバック
    if not rows and source.get("fallback_html"):
        rf = polite_get(source["fallback_html"])
        if rf:
            links = extract_download_links(rf.text, source["fallback_html"], r"\.(csv|xlsx?)$")
            keyword = source.get("keyword", "公園")
            links = [l for l in links if keyword in l or keyword in rf.text[:100]]
            for lnk in links[:5]:
                rows.extend(collect_from_url(lnk, source))
    if rows:
        collected_rows.extend(rows)
        log_rows.append(make_log(source, "success", "dataeye API", len(rows)))
    else:
        log_rows.append(make_log(source, "skipped", "対象リソース無し", 0))


def handle_odcs_page(source: dict, collected_rows: list, log_rows: list):
    """BODIK ODCS のページから公園CSVを探す"""
    log.info(f"[{source['source_id']}] odcs_page {source['url']}")
    r = polite_get(source["url"])
    if r is None:
        log_rows.append(make_log(source, "error", "ページ取得失敗", 0))
        return
    keyword = source.get("keyword", "公園")
    links = extract_download_links(r.text, source["url"], r"\.(csv|xlsx?)$")
    # キーワードを含むリンクのみ
    soup = BeautifulSoup(r.text, "lxml")
    park_links = []
    for a in soup.find_all("a", href=True):
        ctx = (a.get_text() + a["href"])
        if keyword in ctx:
            full = urljoin(source["url"], a["href"])
            if re.search(r"\.(csv|xlsx?)$", full, re.IGNORECASE):
                park_links.append(full)
    rows = []
    for lnk in (park_links or links)[:5]:
        rows.extend(collect_from_url(lnk, source))
    if rows:
        collected_rows.extend(rows)
        log_rows.append(make_log(source, "success", "ODCS page", len(rows)))
    else:
        log_rows.append(make_log(source, "skipped", "対象リソース無し", 0))


HANDLERS = {
    "bodik_dataset":    handle_bodik_dataset,
    "html_link":        handle_html_link,
    "ckan_api":         handle_ckan_api,
    "ckan_package_api": handle_ckan_package_api,
    "bodik_ckan_bulk":  handle_bodik_ckan_bulk,
    "ckan_api_bulk":    handle_ckan_api_bulk,
    "dataeye_api":      handle_dataeye_api,
    "odcs_page":        handle_odcs_page,
}


# ─── ログ補助 ────────────────────────────────────────────────────────────────
def make_log(source: dict, status: str, reason: str, count: int) -> dict:
    return {
        "org_name":     source["org"],
        "data_url":     source["url"],
        "status":       status,
        "reason":       reason,
        "record_count": count,
        "retrieved_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }


# ─── 重複排除 ────────────────────────────────────────────────────────────────
def dedup(rows: list[dict]) -> list[dict]:
    seen = set()
    out = []
    for r in rows:
        key = (r["pref_name"], r["city_name"], r["park_name"], r["address"])
        h = hashlib.md5("||".join(key).encode()).hexdigest()
        if h not in seen:
            seen.add(h)
            out.append(r)
    return out


# ─── メイン ──────────────────────────────────────────────────────────────────
def main():
    collected_rows: list[dict] = []
    log_rows: list[dict] = []

    total = len(SOURCES)
    for i, src in enumerate(SOURCES, 1):
        log.info(f"=== [{i}/{total}] {src['source_id']} — {src['org']} ===")
        handler = HANDLERS.get(src["type"])
        if handler is None:
            log.error(f"Unknown handler: {src['type']}")
            log_rows.append(make_log(src, "error", f"不明なtype: {src['type']}", 0))
            continue
        try:
            handler(src, collected_rows, log_rows)
        except Exception as e:
            log.error(f"  uncaught: {e}")
            log_rows.append(make_log(src, "error", str(e), 0))

    # 重複排除
    before = len(collected_rows)
    collected_rows = dedup(collected_rows)
    after = len(collected_rows)
    log.info(f"重複排除: {before} → {after} 件")

    # ─── parks_japan_all.csv 出力 ───────────────────────────────────────────
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=OUT_COLS)
        writer.writeheader()
        for row in collected_rows:
            writer.writerow({c: row.get(c, "") for c in OUT_COLS})
    log.info(f"出力: {OUTPUT_CSV} ({after} 件)")

    # ─── collection_log.csv 出力 ────────────────────────────────────────────
    log_cols = ["org_name", "data_url", "status", "reason", "record_count", "retrieved_at"]
    with open(LOG_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=log_cols)
        writer.writeheader()
        for row in log_rows:
            writer.writerow({c: row.get(c, "") for c in log_cols})
    log.info(f"出力: {LOG_CSV} ({len(log_rows)} 件)")

    # ─── collection_report.md 出力 ──────────────────────────────────────────
    success = [r for r in log_rows if r["status"] == "success"]
    skipped = [r for r in log_rows if r["status"] == "skipped"]
    errors  = [r for r in log_rows if r["status"] == "error"]
    orgs    = len({r["org_name"] for r in log_rows})
    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    report = f"""# 公園オープンデータ収集レポート

## 概要
- **収集日時**: {now_str}
- **対象ソース数**: {total}
- **対象自治体/組織数**: {orgs}
- **成功**: {len(success)} ソース
- **スキップ**: {len(skipped)} ソース
- **エラー**: {len(errors)} ソース
- **総レコード数（重複排除後）**: {after:,} 件
- **重複排除前**: {before:,} 件（{before - after:,} 件削除）

## 成功したソース（{len(success)} 件）

| org | records | url |
|-----|--------:|-----|
"""
    for r in sorted(success, key=lambda x: -x["record_count"]):
        report += f'| {r["org_name"]} | {r["record_count"]:,} | {r["data_url"][:80]} |\n'

    report += f"""
## スキップしたソース（{len(skipped)} 件）

| org | 理由 |
|-----|------|
"""
    for r in skipped:
        report += f'| {r["org_name"]} | {r["reason"]} |\n'

    report += f"""
## エラーが発生したソース（{len(errors)} 件）

| org | 理由 |
|-----|------|
"""
    for r in errors:
        report += f'| {r["org_name"]} | {r["reason"]} |\n'

    with open(REPORT_MD, "w", encoding="utf-8") as f:
        f.write(report)
    log.info(f"出力: {REPORT_MD}")

    # ターミナルにレポート表示
    print("\n" + "=" * 60)
    print(report)


if __name__ == "__main__":
    main()
