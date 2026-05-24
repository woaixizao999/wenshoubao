# -*- coding: utf-8 -*-
"""
稳收宝：低风险债基自动筛选工具。

第一版使用天天基金/东方财富公开数据做本地筛选、评分和报告输出。
程序不提供交易、不下单，输出仅作为研究筛选线索。
"""

from __future__ import annotations

import argparse
import datetime as dt
import difflib
import hashlib
import html
import json
import math
import os
import queue
import re
import sys
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from lxml import html as lxml_html


APP_NAME = "稳收宝"
APP_VERSION = "V26.5.23"
DEFAULT_DEEP_LIMIT = 300
DEFAULT_SEARCH_LIMIT = 20
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
)


def _workspace_root() -> Path:
    env_root = os.environ.get("WENSHOUBAO_HOME")
    if env_root:
        return Path(env_root)
    preferred = Path("E:/codex/investment")
    if preferred.exists():
        return preferred
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent.parent
        return exe_dir
    return Path(__file__).resolve().parent.parent


def _program_runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


WORKSPACE_ROOT = _workspace_root()
PROGRAM_DIR = Path(__file__).resolve().parent
PROGRAM_RUNTIME_DIR = _program_runtime_dir()
OUTPUT_DIR = PROGRAM_RUNTIME_DIR / "output"
CACHE_DIR = PROGRAM_RUNTIME_DIR / "winout"
LOCAL_RESEARCH_DIR = WORKSPACE_ROOT / "fund_research"
ASSETS_DIR = PROGRAM_DIR / "assets"
APP_ICON_RELATIVE = Path("assets") / "win.ico"


Progress = Callable[[str], None]


def bundled_resource(relative_path: Path) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return PROGRAM_DIR / relative_path


def open_with_windows_default(path: Path) -> None:
    target = path.resolve()
    if os.name != "nt":
        webbrowser.open(target.as_uri())
        return

    import ctypes

    verb = "explore" if target.is_dir() else "open"
    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        verb,
        str(target),
        None,
        str(target.parent),
        1,
    )
    if result <= 32:
        raise OSError(f"Windows 默认应用打开失败，错误码：{result}")


def now_text() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = html.unescape(str(value))
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def compact_text(value: str) -> str:
    return re.sub(r"\s+", "", html.unescape(value or ""))


def to_float(value: Any, default: float | None = None) -> float | None:
    if value is None:
        return default
    text = str(value).replace(",", "").replace("%", "").replace("％", "").strip()
    if not text or text in {"--", "---", "nan", "None"}:
        return default
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return default
    try:
        return float(match.group(0))
    except ValueError:
        return default


def annualized_return(total_pct: float | None, years: float) -> float:
    if total_pct is None:
        return 0.0
    if total_pct <= -99.0:
        return -99.0
    return ((1.0 + total_pct / 100.0) ** (1.0 / years) - 1.0) * 100.0


def pct_text(value: Any) -> str:
    number = to_float(value)
    if number is None:
        return "--"
    return f"{number:.2f}%"


def safe_filename_key(value: str) -> str:
    return hashlib.sha1(value.encode("utf-8")).hexdigest()


class FetchClient:
    def __init__(self, cache_dir: Path, rate_limit_seconds: float = 0.25):
        self.cache_dir = cache_dir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.rate_limit_seconds = rate_limit_seconds
        self._last_request = 0.0
        self.used_stale_cache = False

    def get(
        self,
        url: str,
        referer: str = "https://fund.eastmoney.com/",
        ttl_hours: float = 12.0,
        use_cache: bool = True,
    ) -> str:
        cache_file = self.cache_dir / f"{safe_filename_key(url)}.txt"
        if use_cache and cache_file.exists():
            age_hours = (time.time() - cache_file.stat().st_mtime) / 3600.0
            if age_hours <= ttl_hours:
                return cache_file.read_text(encoding="utf-8", errors="ignore")

        wait = self.rate_limit_seconds - (time.time() - self._last_request)
        if wait > 0:
            time.sleep(wait)
        self._last_request = time.time()

        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "*/*",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
                "Referer": referer,
                "Connection": "close",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
            text = self._decode(raw)
            if use_cache:
                cache_file.write_text(text, encoding="utf-8")
            return text
        except Exception:
            if cache_file.exists():
                self.used_stale_cache = True
                return cache_file.read_text(encoding="utf-8", errors="ignore")
            raise

    @staticmethod
    def _decode(raw: bytes) -> str:
        for enc in ("utf-8-sig", "utf-8", "gb18030"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="ignore")


def fetch_fund_code_map(fetcher: FetchClient) -> dict[str, dict[str, str]]:
    text = fetcher.get(
        "https://fund.eastmoney.com/js/fundcode_search.js",
        referer="https://fund.eastmoney.com/",
        ttl_hours=24 * 7,
    )
    start = text.find("[")
    end = text.rfind("]")
    if start < 0 or end <= start:
        raise ValueError("无法解析 fundcode_search.js")
    data = json.loads(text[start : end + 1])
    result: dict[str, dict[str, str]] = {}
    for item in data:
        if len(item) >= 4:
            result[str(item[0]).zfill(6)] = {
                "pinyin": str(item[1]),
                "name": str(item[2]),
                "subtype": str(item[3]),
            }
    return result


def fetch_rank_data(fetcher: FetchClient, code_map: dict[str, dict[str, str]]) -> pd.DataFrame:
    today = dt.date.today()
    start = today - dt.timedelta(days=366)
    params = {
        "op": "ph",
        "dt": "kf",
        "ft": "zq",
        "rs": "",
        "gs": "0",
        "sc": "3nzf",
        "st": "desc",
        "sd": start.isoformat(),
        "ed": today.isoformat(),
        "qdii": "",
        "tabSubtype": ",,,,,",
        "pi": "1",
        "pn": "10000",
        "dx": "1",
        "v": str(time.time()),
    }
    url = "https://fund.eastmoney.com/data/rankhandler.aspx?" + urllib.parse.urlencode(params)
    text = fetcher.get(url, referer="https://fund.eastmoney.com/data/fundranking.html", ttl_hours=6)
    if "无访问权限" in text:
        raise PermissionError("天天基金排行接口返回无访问权限")
    match = re.search(r"datas:\[(.*?)\]\s*,\s*allRecords", text, re.S)
    if not match:
        raise ValueError("无法解析天天基金排行数据")
    items = re.findall(r'"(.*?)"', match.group(1), re.S)
    rows: list[dict[str, Any]] = []
    for item in items:
        fields = item.split(",")
        if len(fields) < 17:
            continue
        code = fields[0].zfill(6)
        meta = code_map.get(code, {})
        row = {
            "Code": code,
            "Name": fields[1],
            "Pinyin": fields[2],
            "Date": fields[3],
            "UnitNav": fields[4],
            "AccumNav": fields[5],
            "Daily": to_float(fields[6], 0.0),
            "Week": to_float(fields[7], 0.0),
            "Month": to_float(fields[8], 0.0),
            "ThreeMonth": to_float(fields[9], 0.0),
            "SixMonth": to_float(fields[10], 0.0),
            "OneYear": to_float(fields[11]),
            "TwoYear": to_float(fields[12]),
            "ThreeYear": to_float(fields[13]),
            "YTD": to_float(fields[14]),
            "SinceStart": to_float(fields[15]),
            "StartDate": fields[16],
            "OriginalFee": fields[19] if len(fields) > 19 else "",
            "DiscountFee": fields[20] if len(fields) > 20 else "",
            "Subtype": meta.get("subtype", ""),
            "CodeMapName": meta.get("name", ""),
        }
        row["PerfScorePre"] = (
            0.45 * annualized_return(row["ThreeYear"], 3)
            + 0.25 * annualized_return(row["TwoYear"], 2)
            + 0.20 * (row["OneYear"] or 0.0)
            + 0.10 * (row["YTD"] or 0.0)
        )
        rows.append(row)
    return pd.DataFrame(rows)


BAD_BASE_TERMS = (
    "可转债",
    "转债",
    "可交换",
    "混合一级",
    "混合二级",
    "混合型",
    "偏债",
    "股票",
    "权益",
    "固收+",
    "增强",
    "双债",
    "FOF",
    "QDII",
)
ALLOWED_TYPE_TERMS = ("债券型-长债", "债券型-中短债", "债券型-短债", "指数型-固收")
ALLOWED_NAME_HINTS = ("纯债", "短债", "中短债", "利率债", "政金债", "政策性金融债", "国开", "农发", "进出", "定开")


def base_filter_rank(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    kept: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []
    for row in df.to_dict("records"):
        name = row.get("Name", "")
        subtype = row.get("Subtype", "")
        full = f"{name} {subtype}"
        reason = ""
        if any(term in full for term in BAD_BASE_TERMS):
            reason = "初筛剔除：含可转债/混合债/固收+等非保守债基特征"
        elif not any(term in subtype for term in ALLOWED_TYPE_TERMS) and not any(term in name for term in ALLOWED_NAME_HINTS):
            reason = "初筛剔除：基金类型不在短债、纯债、中短债、利率债、政金债、定开债范围"
        elif (row.get("OneYear") or 0) > 15 or (row.get("TwoYear") or 0) > 25 or (row.get("ThreeYear") or 0) > 25:
            reason = "初筛剔除：历史收益异常偏高，疑似非低风险债基或特殊样本"

        if reason:
            item = dict(row)
            item["RejectReason"] = reason
            rejected.append(item)
        else:
            kept.append(row)
    return pd.DataFrame(kept), pd.DataFrame(rejected)


def fund_category(row: dict[str, Any]) -> str:
    name = clean_text(row.get("Name", ""))
    fund_type = clean_text(row.get("FundType") or row.get("Subtype") or "")
    full = f"{name} {fund_type}"
    if any(term in full for term in ("短债", "中短债", "60天", "90天", "月添利")):
        return "短债/中短债"
    if any(term in full for term in ("利率债", "政金债", "政策性金融债", "国开", "农发", "进出")) or "指数型-固收" in fund_type:
        return "利率债/政金债"
    if any(term in full for term in ("定开", "定期开放", "封闭", "持有")):
        return "定开/持有期债"
    return "纯债长债"


def preselect_score(row: dict[str, Any]) -> float:
    score = annualized_return(to_float(row.get("ThreeYear")), 3) * 0.38
    score += annualized_return(to_float(row.get("TwoYear")), 2) * 0.24
    score += (to_float(row.get("OneYear"), 0.0) or 0.0) * 0.18
    score += (to_float(row.get("YTD"), 0.0) or 0.0) * 0.08

    name = clean_text(row.get("Name", ""))
    subtype = clean_text(row.get("Subtype", ""))
    full = f"{name} {subtype}"
    if any(term in full for term in ("利率债", "政金债", "政策性金融债", "短债", "中短债")):
        score += 1.0
    elif "纯债" in full or "债券型-长债" in subtype:
        score += 0.6
    if any(term in full for term in ("定开", "定期开放", "持有")):
        score -= 0.4
    if any(term in full for term in BAD_BASE_TERMS):
        score -= 5.0
    return score


def add_preselect_scores(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df["FundCategory"] = [fund_category(row) for row in df.to_dict("records")]
    df["PreselectScore"] = [preselect_score(row) for row in df.to_dict("records")]
    if "PerfScorePre" not in df.columns:
        df["PerfScorePre"] = df["PreselectScore"]
    else:
        df["PerfScorePre"] = df["PreselectScore"]
    return df


def load_local_enriched_lookup() -> dict[str, dict[str, Any]]:
    path = LOCAL_RESEARCH_DIR / "r2_bond_fund_300_samples_with_company.csv"
    if not path.exists():
        return {}
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    lookup: dict[str, dict[str, Any]] = {}
    for row in df.to_dict("records"):
        lookup[str(row.get("Code", "")).zfill(6)] = row
    return lookup


def load_local_universe() -> pd.DataFrame:
    path = LOCAL_RESEARCH_DIR / "r2_bond_fund_universe_clean.csv"
    if not path.exists():
        raise FileNotFoundError("没有可用的联网数据，也没有找到本地样本基金池")
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    rename = {
        "Code": "Code",
        "Name": "Name",
        "Subtype": "Subtype",
        "Date": "Date",
        "OneYear": "OneYear",
        "TwoYear": "TwoYear",
        "ThreeYear": "ThreeYear",
        "YTD": "YTD",
        "SinceStart": "SinceStart",
        "StartDate": "StartDate",
    }
    df = df.rename(columns=rename)
    for col in ("OneYear", "TwoYear", "ThreeYear", "YTD", "SinceStart"):
        df[col] = df[col].map(to_float)
    df["PerfScorePre"] = (
        0.45 * df["ThreeYear"].map(lambda x: annualized_return(x, 3))
        + 0.25 * df["TwoYear"].map(lambda x: annualized_return(x, 2))
        + 0.20 * df["OneYear"].fillna(0)
        + 0.10 * df["YTD"].fillna(0)
    )
    df["OriginalFee"] = ""
    df["DiscountFee"] = ""
    df["SourceFallback"] = "本地历史样本"
    return df


def parse_basic_page(text: str) -> dict[str, Any]:
    flat = compact_text(lxml_html.fromstring(text).text_content())
    result: dict[str, Any] = {}

    match = re.search(
        r"成立日期：(?P<start>\d{4}-\d{2}-\d{2})基金经理：(?P<manager>.*?)类型：(?P<type>.*?)管理人：(?P<company>.*?)净资产规模：(?P<scale>.*?)(?:（截止至：(?P<scale_date>\d{4}-\d{2}-\d{2})）|$)",
        flat,
    )
    if match:
        result.update(
            {
                "StartDateDetail": match.group("start"),
                "Manager": clean_text(match.group("manager")),
                "FundType": clean_text(match.group("type")),
                "FundCompany": clean_text(match.group("company")),
                "ScaleText": clean_text(match.group("scale")),
                "ScaleDate": match.group("scale_date") or "",
            }
        )
        result["Scale"] = to_float(result["ScaleText"])

    status = re.search(r"交易状态：(?P<status>.*?)购买手续费：", flat)
    result["PurchaseStatus"] = clean_text(status.group("status")) if status else ""
    limit = re.search(r"单日累计购买上限(?P<num>\d+(?:\.\d+)?)(?P<unit>万元|亿元|元)", flat)
    if limit:
        number = float(limit.group("num"))
        unit = limit.group("unit")
        if unit == "亿元":
            result["DailyPurchaseLimitWan"] = number * 10000
        elif unit == "万元":
            result["DailyPurchaseLimitWan"] = number
        else:
            result["DailyPurchaseLimitWan"] = number / 10000
    else:
        result["DailyPurchaseLimitWan"] = None
    return result


def parse_fee_page(text: str) -> dict[str, Any]:
    flat = compact_text(lxml_html.fromstring(text).text_content())
    result: dict[str, Any] = {}
    op = re.search(r"管理费率(?P<mgmt>[^托]+)托管费率(?P<custody>[^销]+)销售服务费率(?P<sales>[^注]+)", flat)
    if op:
        result["MgmtFee"] = clean_text(op.group("mgmt"))
        result["CustodyFee"] = clean_text(op.group("custody"))
        result["SalesServiceFee"] = clean_text(op.group("sales"))
    sub = re.search(r"申购费率.*?小于100万元.*?(?P<orig>\d+(?:\.\d+)?%).*?\|.*?(?P<disc>\d+(?:\.\d+)?%)", flat)
    if sub:
        result["SubscriptionFee"] = sub.group("disc")
        result["OriginalSubscriptionFee"] = sub.group("orig")
    redemption = re.search(r"赎回费率.*?小于7天(?P<fee>\d+(?:\.\d+)?%)", flat)
    if redemption:
        result["RedemptionFee"] = redemption.group("fee")
    return result


def parse_manager_page(text: str) -> dict[str, Any]:
    doc = lxml_html.fromstring(text)
    result: dict[str, Any] = {}
    for table in doc.xpath("//table"):
        table_text = compact_text(table.text_content())
        if "任职回报" not in table_text or "起始期" not in table_text:
            continue
        rows = table.xpath(".//tbody/tr")
        if not rows:
            continue
        cells = [clean_text(" ".join(td.xpath(".//text()"))) for td in rows[0].xpath("./td")]
        if len(cells) >= 5:
            result["CurrentMgrStart"] = cells[0]
            result["CurrentMgrEnd"] = cells[1]
            result["Manager"] = cells[2] or result.get("Manager", "")
            result["CurrentMgrTenure"] = cells[3]
            result["CurrentMgrReturn"] = cells[4]
        break
    return result


SAFE_BOND_TERMS = ("国债", "附息国债", "特别国债", "国开", "农发", "进出", "政策性", "政金", "地方债")
CONVERTIBLE_TERMS = ("转债", "可转债", "可交换", "EB")


def parse_bond_holdings(text: str) -> dict[str, Any]:
    match = re.search(r'content:"(?P<content>.*?)",arryear', text, re.S)
    if not match:
        return {"Top5BondConcentration": None, "Top5Bonds": "", "HasConvertibleBond": False, "AllTop5SafeBonds": False}
    content = match.group("content").replace('\\"', '"').replace("\\/", "/")
    content = html.unescape(content)
    doc = lxml_html.fromstring(content)
    bonds: list[dict[str, Any]] = []
    for tr in doc.xpath("//tbody/tr"):
        cells = [clean_text(" ".join(td.xpath(".//text()"))) for td in tr.xpath("./td")]
        if len(cells) >= 5:
            bonds.append(
                {
                    "rank": cells[0],
                    "code": cells[1],
                    "name": cells[2],
                    "ratio": to_float(cells[3], 0.0) or 0.0,
                    "market_value": cells[4],
                }
            )
    top5 = bonds[:5]
    concentration = sum(item["ratio"] for item in top5) if top5 else None
    top5_text = "；".join(f"{item['name']} {item['ratio']:.2f}%" for item in top5)
    has_convertible = any(any(term in item["name"] for term in CONVERTIBLE_TERMS) for item in bonds)
    all_safe = bool(top5) and all(any(term in item["name"] for term in SAFE_BOND_TERMS) for item in top5)
    return {
        "Top5BondConcentration": concentration,
        "Top5Bonds": top5_text,
        "HasConvertibleBond": has_convertible,
        "AllTop5SafeBonds": all_safe,
        "BondCount": len(bonds),
    }


def empty_nav_metrics() -> dict[str, Any]:
    return {
        "MaxDrawdown1Y": None,
        "MaxDrawdown3Y": None,
        "MonthlyLossCount1Y": None,
        "MonthlyLossCount3Y": None,
    }


def calc_window_nav_metrics(values: list[tuple[dt.date, float]], days: int) -> tuple[float | None, int | None]:
    if len(values) < 2:
        return None, None
    cutoff = values[-1][0] - dt.timedelta(days=days)
    window = [(day, value) for day, value in values if day >= cutoff]
    if len(window) < 2:
        return None, None

    peak = window[0][1]
    max_dd = 0.0
    for _, value in window:
        if value > peak:
            peak = value
        if peak > 0:
            max_dd = max(max_dd, (peak - value) / peak * 100.0)

    month_first_last: dict[str, list[float]] = {}
    for day, value in window:
        month = day.strftime("%Y-%m")
        month_first_last.setdefault(month, [value, value])
        month_first_last[month][1] = value
    monthly_loss = sum(1 for first, last in month_first_last.values() if last < first)
    return max_dd, monthly_loss


def calc_nav_metrics(values: list[tuple[dt.date, float]]) -> dict[str, Any]:
    values = sorted(values, key=lambda x: x[0])
    if len(values) < 2:
        return empty_nav_metrics()
    mdd_1y, loss_1y = calc_window_nav_metrics(values, 370)
    mdd_3y, loss_3y = calc_window_nav_metrics(values, 365 * 3 + 10)
    return {
        "MaxDrawdown1Y": mdd_1y,
        "MaxDrawdown3Y": mdd_3y,
        "MonthlyLossCount1Y": loss_1y,
        "MonthlyLossCount3Y": loss_3y,
    }


def parse_nav_history(text: str) -> dict[str, Any]:
    data = json.loads(text)
    data_part = data.get("Data") or {}
    items = data_part.get("LSJZList", [])
    if not items:
        return empty_nav_metrics()
    chronological = list(reversed(items))
    values: list[tuple[dt.date, float]] = []
    for item in chronological:
        nav = to_float(item.get("LJJZ")) or to_float(item.get("DWJZ"))
        if not nav:
            continue
        try:
            day = dt.date.fromisoformat(str(item.get("FSRQ", ""))[:10])
        except ValueError:
            continue
        values.append((day, nav))
    return calc_nav_metrics(values)


def parse_pingzhong_nav_history(text: str) -> dict[str, Any]:
    match = re.search(r"var Data_ACWorthTrend = (\[.*?\]);", text, re.S)
    if not match:
        match = re.search(r"var Data_netWorthTrend = (\[.*?\]);", text, re.S)
    if not match:
        return empty_nav_metrics()

    raw = json.loads(match.group(1))
    values: list[tuple[dt.date, float]] = []
    for item in raw:
        if isinstance(item, list) and len(item) >= 2:
            ts, nav = item[0], item[1]
        elif isinstance(item, dict):
            ts, nav = item.get("x"), item.get("y")
        else:
            continue
        nav_value = to_float(nav)
        if ts is None or nav_value is None:
            continue
        day = dt.datetime.fromtimestamp(float(ts) / 1000.0).date()
        values.append((day, nav_value))

    return calc_nav_metrics(values)


def apply_local_detail_fallback(detail: dict[str, Any], code: str, local_lookup: dict[str, dict[str, Any]]) -> dict[str, Any]:
    local = local_lookup.get(code)
    if not local:
        return detail
    mapping = {
        "FundType": "FundType",
        "ScaleText": "Scale",
        "Manager": "Manager",
        "CurrentMgrStart": "CurrentMgrStart",
        "CurrentMgrTenure": "CurrentMgrTenure",
        "CurrentMgrReturn": "CurrentMgrReturn",
        "MgmtFee": "MgmtFee",
        "CustodyFee": "CustodyFee",
        "SalesServiceFee": "SalesServiceFee",
        "SubscriptionFee": "SubscriptionFee",
        "RedemptionFee": "RedemptionFee",
        "Top5BondConcentration": "Top5BondConcentration",
        "Top5Bonds": "Top5Bonds",
        "FundCompany": "FundCompany",
    }
    for target, source in mapping.items():
        if not detail.get(target) and local.get(source):
            detail[target] = local.get(source)
    if not detail.get("Scale") and detail.get("ScaleText"):
        detail["Scale"] = to_float(detail["ScaleText"])
    if not detail.get("HasConvertibleBond") and detail.get("Top5Bonds"):
        detail["HasConvertibleBond"] = any(term in str(detail["Top5Bonds"]) for term in CONVERTIBLE_TERMS)
    return detail


def fetch_detail_for_fund(
    row: dict[str, Any],
    fetcher: FetchClient,
    local_lookup: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    code = row["Code"]
    detail: dict[str, Any] = {
        "DetailSource": "天天基金公开页面",
        "FundPage": f"https://fund.eastmoney.com/{code}.html",
        "BasicPage": f"https://fundf10.eastmoney.com/jbgk_{code}.html",
        "FeePage": f"https://fundf10.eastmoney.com/jjfl_{code}.html",
        "ManagerPage": f"https://fundf10.eastmoney.com/jjjl_{code}.html",
        "BondPage": f"https://fundf10.eastmoney.com/ccmx1_{code}.html",
        "CompanyVerifyPage": "https://fund.eastmoney.com/company/default.html",
        "MaxDrawdown1Y": None,
        "MaxDrawdown3Y": None,
        "MonthlyLossCount1Y": None,
        "MonthlyLossCount3Y": None,
    }
    errors: list[str] = []
    try:
        basic = fetcher.get(detail["BasicPage"], referer=detail["FundPage"], ttl_hours=24)
        detail.update(parse_basic_page(basic))
    except Exception as exc:
        errors.append(f"基本概况失败：{exc}")
    try:
        fee = fetcher.get(detail["FeePage"], referer=detail["FundPage"], ttl_hours=24 * 3)
        detail.update(parse_fee_page(fee))
    except Exception as exc:
        errors.append(f"费率失败：{exc}")
    try:
        manager = fetcher.get(detail["ManagerPage"], referer=detail["FundPage"], ttl_hours=24 * 3)
        detail.update(parse_manager_page(manager))
    except Exception as exc:
        errors.append(f"经理失败：{exc}")
    try:
        bond_url = f"https://fundf10.eastmoney.com/FundArchivesDatas.aspx?type=zqcc&code={code}&year=&rt={time.time()}"
        bonds = fetcher.get(bond_url, referer=detail["BondPage"], ttl_hours=24 * 7)
        detail.update(parse_bond_holdings(bonds))
    except Exception as exc:
        errors.append(f"债券持仓失败：{exc}")
    try:
        trend_url = f"https://fund.eastmoney.com/pingzhongdata/{code}.js?v={dt.datetime.now().strftime('%Y%m%d%H')}"
        nav = fetcher.get(trend_url, referer=detail["FundPage"], ttl_hours=12)
        detail.update(parse_pingzhong_nav_history(nav))
    except Exception as exc:
        errors.append(f"历史净值失败：{exc}")

    detail = apply_local_detail_fallback(detail, code, local_lookup)
    if errors:
        detail["DetailWarnings"] = "；".join(errors)
        if detail.get("DetailSource") == "天天基金公开页面":
            detail["DetailSource"] = "天天基金公开页面+本地样本兜底"
    else:
        detail["DetailWarnings"] = ""
    return detail


def detail_reject_reason(row: dict[str, Any]) -> str:
    name = clean_text(row.get("Name", ""))
    fund_type = clean_text(row.get("FundType") or row.get("Subtype") or "")
    if any(term in f"{name} {fund_type}" for term in BAD_BASE_TERMS):
        return "深度剔除：基金类型/名称显示含可转债、混合债或固收+特征"
    scale = to_float(row.get("Scale"))
    if scale is None:
        return "深度剔除：无法确认基金规模"
    if scale < 2:
        return "深度剔除：规模低于2亿元"
    status = row.get("PurchaseStatus", "")
    if any(term in status for term in ("暂停申购", "暂停购买", "封闭期")):
        return "深度剔除：暂停申购或处于封闭期"
    limit = row.get("DailyPurchaseLimitWan")
    if limit is not None and limit < 1:
        return "深度剔除：日累计申购上限低于1万元"
    if row.get("HasConvertibleBond"):
        return "深度剔除：债券持仓中出现转债/可交换债"
    concentration = to_float(row.get("Top5BondConcentration"))
    if concentration is not None and concentration > 80 and not row.get("AllTop5SafeBonds"):
        return "深度剔除：前五大债券集中度超过80%，且不是清一色国债/政金债"
    return ""


def tenure_days(text: Any) -> int | None:
    s = str(text or "")
    if not s:
        return None
    total = 0
    year = re.search(r"(\d+)年", s)
    day = re.search(r"(\d+)天", s)
    if year:
        total += int(year.group(1)) * 365
    if day:
        total += int(day.group(1))
    if total:
        return total
    if s.isdigit():
        return int(s)
    return None


def score_scale(scale: float | None) -> float:
    if scale is None:
        return 0.0
    if 10 <= scale <= 80:
        return 10.0
    if 5 <= scale < 10 or 80 < scale <= 150:
        return 8.0
    if 2 <= scale < 5 or 150 < scale <= 300:
        return 6.0
    if scale > 300:
        return 5.0
    return 0.0


def score_fee(row: dict[str, Any]) -> float:
    mgmt = to_float(row.get("MgmtFee"), 0.3) or 0.3
    custody = to_float(row.get("CustodyFee"), 0.1) or 0.1
    sales = to_float(row.get("SalesServiceFee"), 0.0) or 0.0
    annual = mgmt + custody + sales
    if annual <= 0.20:
        score = 8
    elif annual <= 0.35:
        score = 7
    elif annual <= 0.45:
        score = 6
    elif annual <= 0.60:
        score = 4
    else:
        score = 2
    if sales >= 0.20:
        score -= 1.5
    return max(0.0, float(score))


def score_manager(row: dict[str, Any]) -> float:
    days = tenure_days(row.get("CurrentMgrTenure"))
    if days is None and row.get("CurrentMgrStart"):
        try:
            start = dt.date.fromisoformat(str(row["CurrentMgrStart"])[:10])
            days = (dt.date.today() - start).days
        except Exception:
            days = None
    if days is None:
        score = 5.0
    elif days >= 1095:
        score = 12.0
    elif days >= 730:
        score = 10.0
    elif days >= 365:
        score = 8.0
    elif days >= 180:
        score = 5.5
    else:
        score = 3.0
    mgr_ret = to_float(row.get("CurrentMgrReturn"))
    if mgr_ret is not None and mgr_ret < 0:
        score -= 2
    return max(0.0, min(12.0, score))


def missing_fields(row: dict[str, Any]) -> list[str]:
    required = {
        "Scale": "规模",
        "Manager": "基金经理",
        "FundCompany": "基金公司",
        "Top5Bonds": "债券持仓",
        "MgmtFee": "管理费",
        "CustodyFee": "托管费",
        "MaxDrawdown1Y": "近1年回撤",
        "MaxDrawdown3Y": "近3年回撤",
    }
    missing: list[str] = []
    for col, label in required.items():
        value = row.get(col)
        if value is None or clean_text(value) in {"", "--", "---", "nan"}:
            missing.append(label)
    return missing


def missing_data_penalty(row: dict[str, Any]) -> float:
    missing = missing_fields(row)
    penalty = 0.0
    severe = {"规模", "债券持仓", "近1年回撤", "近3年回撤"}
    for item in missing:
        penalty += 1.2 if item in severe else 0.6
    if row.get("DetailWarnings"):
        penalty += 0.8
    return min(8.0, penalty)


def score_drawdown(row: dict[str, Any]) -> float:
    mdd_1y = to_float(row.get("MaxDrawdown1Y"))
    mdd_3y = to_float(row.get("MaxDrawdown3Y"))
    loss_1y = to_float(row.get("MonthlyLossCount1Y"), 0) or 0
    loss_3y = to_float(row.get("MonthlyLossCount3Y"), 0) or 0
    if mdd_1y is None and mdd_3y is None:
        return 3.0
    score = 15.0
    if mdd_1y is not None:
        score -= mdd_1y * 2.4
    if mdd_3y is not None:
        score -= mdd_3y * 0.8
    score -= loss_1y * 0.30
    score -= loss_3y * 0.05
    return max(0.0, min(15.0, score))


def score_holding(row: dict[str, Any]) -> float:
    concentration = to_float(row.get("Top5BondConcentration"))
    if concentration is None:
        return 5.0
    if row.get("HasConvertibleBond"):
        return 0.0
    if concentration <= 35:
        score = 15.0
    elif concentration <= 55:
        score = 13.0
    elif concentration <= 70:
        score = 10.0
    elif concentration <= 80:
        score = 7.0
    else:
        score = 5.0 if row.get("AllTop5SafeBonds") else 2.0
    if row.get("AllTop5SafeBonds") and concentration <= 80:
        score += 1.0
    return max(0.0, min(15.0, score))


def score_term(row: dict[str, Any]) -> float:
    name = clean_text(row.get("Name", ""))
    status = clean_text(row.get("PurchaseStatus", ""))
    if "暂停" in status:
        return 0.0
    if any(term in name for term in ("66个月", "87个月", "三年", "5年")):
        return 1.5
    if any(term in name for term in ("定开", "定期开放", "一年持有", "12个月", "18个月", "两年")):
        return 2.5
    if any(term in name for term in ("6个月", "六个月", "3个月", "三个月")):
        return 3.5
    return 5.0


def score_risk_boundary(row: dict[str, Any]) -> float:
    fund_type = clean_text(row.get("FundType") or row.get("Subtype") or "")
    name = clean_text(row.get("Name", ""))
    full = f"{name} {fund_type}"
    if any(term in full for term in BAD_BASE_TERMS):
        return 0.0
    if "指数型-固收" in fund_type or any(term in name for term in ("利率债", "政金债", "政策性金融债", "国开", "农发")):
        return 15.0
    if any(term in fund_type for term in ("债券型-中短债", "债券型-短债")):
        return 14.0
    if "债券型-长债" in fund_type or "纯债" in name:
        return 13.0
    return 10.0


def score_disclosure(row: dict[str, Any]) -> float:
    required = ("Scale", "Manager", "FundCompany", "Top5Bonds", "MgmtFee", "CustodyFee", "MaxDrawdown1Y", "MaxDrawdown3Y")
    present = sum(1 for col in required if clean_text(row.get(col)) not in ("", "--", "---", "nan"))
    score = 2.0 * present / len(required)
    if row.get("DetailWarnings"):
        score -= 0.8
    return max(0.0, min(2.0, score))


def normalize_fund_key(name: str) -> str:
    key = re.sub(r"\(.*?\)", "", clean_text(name))
    key = re.sub(r"（.*?）", "", key)
    key = re.sub(r"(A/B|A|B|C|D|E|I|Y|Z|O|R)$", "", key)
    key = re.sub(r"(债券|纯债|短债)(A/B|A|B|C|D|E|I|Y|Z|O|R)$", r"\1", key)
    return key.strip()


def score_candidates(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    if "FundCategory" not in df.columns:
        df["FundCategory"] = [fund_category(row) for row in df.to_dict("records")]
    perf_rank = df.groupby("FundCategory")["PerfScorePre"].rank(method="average", pct=True)
    small_group = df.groupby("FundCategory")["PerfScorePre"].transform("count") < 5
    perf_rank = perf_rank.where(~small_group, df["PerfScorePre"].rank(method="average", pct=True))
    df["ScoreRiskBoundary"] = 0.0
    df["ScorePerformance"] = 0.0
    df["ScoreDrawdown"] = 0.0
    df["ScoreHolding"] = 0.0
    df["ScoreManager"] = 0.0
    df["ScoreScale"] = 0.0
    df["ScoreFee"] = 0.0
    df["ScoreTerm"] = 0.0
    df["ScoreDisclosure"] = 0.0
    df["ScoreMissingPenalty"] = 0.0
    df["TotalScore"] = 0.0
    df["MainNotes"] = ""
    for idx, row in df.iterrows():
        item = row.to_dict()
        risk = score_risk_boundary(item)
        performance = float(perf_rank.loc[idx]) * 18.0
        if to_float(item.get("ThreeYear")) is None:
            performance *= 0.60
        if to_float(item.get("TwoYear")) is None:
            performance *= 0.75
        drawdown = score_drawdown(item)
        holding = score_holding(item)
        manager = score_manager(item)
        scale = score_scale(to_float(item.get("Scale")))
        fee = score_fee(item)
        term = score_term(item)
        disclosure = score_disclosure(item)
        penalty = missing_data_penalty(item)
        total = risk + performance + drawdown + holding + manager + scale + fee + term + disclosure - penalty

        notes: list[str] = []
        missing = missing_fields(item)
        if missing:
            notes.append(f"关键数据缺失扣分：{','.join(missing)}")
        concentration = to_float(item.get("Top5BondConcentration"))
        if concentration is not None and concentration > 70:
            notes.append(f"前五债券集中度较高({concentration:.2f}%)")
        if item.get("AllTop5SafeBonds") and concentration is not None and concentration > 80:
            notes.append("集中度高但前五主要为国债/政金债")
        mdd = to_float(item.get("MaxDrawdown1Y"))
        if mdd is not None and mdd > 1:
            notes.append(f"近一年最大回撤{mdd:.2f}%")
        mdd_3y = to_float(item.get("MaxDrawdown3Y"))
        if mdd_3y is not None and mdd_3y > 3:
            notes.append(f"近三年最大回撤{mdd_3y:.2f}%")
        if score_term(item) < 5:
            notes.append("流动性受持有期/定开影响")
        if score_manager(item) < 6:
            notes.append("当前经理任期偏短")
        if item.get("DetailWarnings"):
            notes.append("部分字段使用缓存或本地样本兜底")

        df.loc[idx, "ScoreRiskBoundary"] = round(risk, 2)
        df.loc[idx, "ScorePerformance"] = round(performance, 2)
        df.loc[idx, "ScoreDrawdown"] = round(drawdown, 2)
        df.loc[idx, "ScoreHolding"] = round(holding, 2)
        df.loc[idx, "ScoreManager"] = round(manager, 2)
        df.loc[idx, "ScoreScale"] = round(scale, 2)
        df.loc[idx, "ScoreFee"] = round(fee, 2)
        df.loc[idx, "ScoreTerm"] = round(term, 2)
        df.loc[idx, "ScoreDisclosure"] = round(disclosure, 2)
        df.loc[idx, "ScoreMissingPenalty"] = round(penalty, 2)
        df.loc[idx, "TotalScore"] = round(max(0.0, total), 2)
        df.loc[idx, "MainNotes"] = "；".join(notes) if notes else "风险边界清晰，综合表现较稳"
    df["FundKey"] = df["Name"].map(normalize_fund_key)
    return df.sort_values("TotalScore", ascending=False)


def deduplicate_top_funds(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    if df.empty:
        return df
    share_priority = {"A": 0.40, "D": 0.35, "E": 0.30, "B": 0.10, "C": -0.50}

    def bonus(name: str) -> float:
        m = re.search(r"(A/B|A|B|C|D|E|I|Y|Z|O|R)$", name or "")
        if not m:
            return 0.0
        return share_priority.get(m.group(1), 0.0)

    tmp = df.copy()
    tmp["_SelectScore"] = tmp["TotalScore"] + tmp["Name"].map(bonus)
    tmp = tmp.sort_values(["FundKey", "_SelectScore"], ascending=[True, False])
    tmp = tmp.drop_duplicates("FundKey", keep="first")
    tmp = tmp.sort_values("TotalScore", ascending=False).head(top_n)
    return tmp.drop(columns=["_SelectScore"], errors="ignore")


@dataclass
class AnalysisResult:
    top10: pd.DataFrame
    scored: pd.DataFrame
    rejected: pd.DataFrame
    output_dir: Path
    top_csv: Path
    full_xlsx: Path
    report_html: Path
    reject_csv: Path
    log_file: Path
    logs: list[str]


def append_log(logs: list[str], message: str, progress: Progress | None = None) -> None:
    line = f"[{now_text()}] {message}"
    logs.append(line)
    if progress:
        progress(message)


def load_rank_source(
    fetcher: FetchClient,
    logs: list[str],
    progress: Progress | None = None,
    offline: bool = False,
) -> tuple[pd.DataFrame, str]:
    try:
        if offline:
            raise RuntimeError("用户指定离线样本模式")
        code_map = fetch_fund_code_map(fetcher)
        return fetch_rank_data(fetcher, code_map), "online"
    except Exception as exc:
        append_log(logs, f"联网排行抓取失败，改用本地历史样本：{exc}", progress)
        return load_local_universe(), "local"


def run_analysis(
    deep_limit: int = DEFAULT_DEEP_LIMIT,
    progress: Progress | None = None,
    offline: bool = False,
) -> AnalysisResult:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    logs: list[str] = []
    local_lookup = load_local_enriched_lookup()
    fetcher = FetchClient(CACHE_DIR)

    append_log(logs, "开始抓取基金排行和基金代码表", progress)
    rank_df, source_mode = load_rank_source(fetcher, logs, progress, offline)

    append_log(logs, f"原始基金数量：{len(rank_df)}", progress)
    rank_df = add_preselect_scores(rank_df)
    base_kept, base_rejected = base_filter_rank(rank_df)
    append_log(logs, f"初筛保留：{len(base_kept)}；初筛剔除：{len(base_rejected)}", progress)
    if base_kept.empty:
        raise RuntimeError("初筛后没有可分析基金")

    base_kept = base_kept.sort_values("PreselectScore", ascending=False).head(deep_limit).copy()
    append_log(logs, f"进入深度核验候选：{len(base_kept)}", progress)

    enriched_rows: list[dict[str, Any]] = []
    deep_rejected: list[dict[str, Any]] = []
    for index, row in enumerate(base_kept.to_dict("records"), start=1):
        code = row["Code"]
        name = row["Name"]
        append_log(logs, f"深度核验 {index}/{len(base_kept)}：{code} {name}", progress)
        detail: dict[str, Any]
        if source_mode == "local":
            detail = apply_local_detail_fallback(
                {
                    "DetailSource": "本地历史样本",
                    "FundPage": f"https://fund.eastmoney.com/{code}.html",
                    "BasicPage": f"https://fundf10.eastmoney.com/jbgk_{code}.html",
                    "FeePage": f"https://fundf10.eastmoney.com/jjfl_{code}.html",
                    "ManagerPage": f"https://fundf10.eastmoney.com/jjjl_{code}.html",
                    "BondPage": f"https://fundf10.eastmoney.com/ccmx1_{code}.html",
                    "CompanyVerifyPage": "https://fund.eastmoney.com/company/default.html",
                    "MaxDrawdown1Y": None,
                    "MaxDrawdown3Y": None,
                    "MonthlyLossCount1Y": None,
                    "MonthlyLossCount3Y": None,
                    "DetailWarnings": "离线模式未重新抓取网页",
                },
                code,
                local_lookup,
            )
        else:
            detail = fetch_detail_for_fund(row, fetcher, local_lookup)
        merged = {**row, **detail}
        if not merged.get("Scale") and merged.get("ScaleText"):
            merged["Scale"] = to_float(merged["ScaleText"])
        reject = detail_reject_reason(merged)
        if reject:
            merged["RejectReason"] = reject
            deep_rejected.append(merged)
        else:
            enriched_rows.append(merged)

    scored = score_candidates(pd.DataFrame(enriched_rows))
    top10 = deduplicate_top_funds(scored, 10)
    rejected = pd.concat(
        [
            base_rejected,
            pd.DataFrame(deep_rejected),
        ],
        ignore_index=True,
        sort=False,
    )
    append_log(logs, f"深度通过：{len(scored)}；深度剔除：{len(deep_rejected)}；前十名单生成完成", progress)
    if fetcher.used_stale_cache:
        append_log(logs, "注意：部分请求使用了过期缓存", progress)

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    top_csv = OUTPUT_DIR / "前十名.csv"
    full_xlsx = OUTPUT_DIR / "完整评分表.xlsx"
    reject_csv = OUTPUT_DIR / "剔除清单.csv"
    report_html = OUTPUT_DIR / "稳收宝报告.html"
    log_file = OUTPUT_DIR / "运行日志.txt"

    top_csv, full_xlsx, reject_csv, report_html = write_outputs(
        top10, scored, rejected, top_csv, full_xlsx, reject_csv, report_html, logs
    )
    actual_output_dir = top_csv.parent
    archive_top_csv = actual_output_dir / f"前十名_{stamp}.csv"
    archive_xlsx = actual_output_dir / f"完整评分表_{stamp}.xlsx"
    archive_report = actual_output_dir / f"稳收宝报告_{stamp}.html"
    top10.to_csv(archive_top_csv, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(archive_xlsx, engine="openpyxl") as writer:
        top10.to_excel(writer, index=False, sheet_name="前十名")
        scored.to_excel(writer, index=False, sheet_name="完整评分表")
        rejected.to_excel(writer, index=False, sheet_name="剔除清单")
    archive_report.write_text(report_html.read_text(encoding="utf-8"), encoding="utf-8")
    log_file = safe_text_write(actual_output_dir / log_file.name, "\n".join(logs), logs)
    return AnalysisResult(top10, scored, rejected, actual_output_dir, top_csv, full_xlsx, report_html, reject_csv, log_file, logs)


def normalize_search_text(value: Any) -> str:
    return compact_text(str(value or "")).lower()


def fund_match_score(row: dict[str, Any], query: str) -> float:
    q = normalize_search_text(query)
    if not q:
        return 0.0
    code = normalize_search_text(row.get("Code"))
    name = normalize_search_text(row.get("Name"))
    pinyin = normalize_search_text(row.get("Pinyin"))
    code_map_name = normalize_search_text(row.get("CodeMapName"))
    names = [text for text in (name, code_map_name) if text]

    score = 0.0
    if q == code:
        score = max(score, 120.0)
    elif q and q in code:
        score = max(score, 100.0)
    for text in names:
        if q == text:
            score = max(score, 115.0)
        elif q in text:
            score = max(score, 95.0 + min(15.0, len(q)))
        else:
            score = max(score, difflib.SequenceMatcher(None, q, text).ratio() * 90.0)
    if pinyin:
        if q in pinyin:
            score = max(score, 85.0)
        else:
            score = max(score, difflib.SequenceMatcher(None, q, pinyin).ratio() * 75.0)
    return score


def fuzzy_search_funds(rank_df: pd.DataFrame, query: str, limit: int = DEFAULT_SEARCH_LIMIT) -> pd.DataFrame:
    if rank_df.empty:
        return rank_df
    df = rank_df.copy()
    df["MatchScore"] = [fund_match_score(row, query) for row in df.to_dict("records")]
    q = normalize_search_text(query)
    threshold = 35.0 if len(q) <= 2 else 45.0
    matched = df[df["MatchScore"] >= threshold].copy()
    if matched.empty:
        return matched
    return matched.sort_values(["MatchScore", "PerfScorePre"], ascending=[False, False]).head(limit)


def run_fund_search(
    query: str,
    search_limit: int = DEFAULT_SEARCH_LIMIT,
    progress: Progress | None = None,
    offline: bool = False,
) -> AnalysisResult:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    logs: list[str] = []
    local_lookup = load_local_enriched_lookup()
    fetcher = FetchClient(CACHE_DIR)

    query = clean_text(query)
    if not query:
        raise RuntimeError("请输入基金名称或基金代码")

    append_log(logs, f"开始搜索基金：{query}", progress)
    rank_df, source_mode = load_rank_source(fetcher, logs, progress, offline)
    rank_df = add_preselect_scores(rank_df)
    append_log(logs, f"可搜索基金数量：{len(rank_df)}", progress)
    matched = fuzzy_search_funds(rank_df, query, search_limit)
    append_log(logs, f"模糊匹配结果：{len(matched)} 只", progress)
    if matched.empty:
        raise RuntimeError(f"没有找到与“{query}”相近的基金")

    enriched_rows: list[dict[str, Any]] = []
    warning_rows: list[dict[str, Any]] = []
    for index, row in enumerate(matched.to_dict("records"), start=1):
        code = row["Code"]
        name = row["Name"]
        append_log(logs, f"搜索核验 {index}/{len(matched)}：{code} {name}", progress)
        if source_mode == "local":
            detail = apply_local_detail_fallback(
                {
                    "DetailSource": "本地历史样本",
                    "FundPage": f"https://fund.eastmoney.com/{code}.html",
                    "BasicPage": f"https://fundf10.eastmoney.com/jbgk_{code}.html",
                    "FeePage": f"https://fundf10.eastmoney.com/jjfl_{code}.html",
                    "ManagerPage": f"https://fundf10.eastmoney.com/jjjl_{code}.html",
                    "BondPage": f"https://fundf10.eastmoney.com/ccmx1_{code}.html",
                    "CompanyVerifyPage": "https://fund.eastmoney.com/company/default.html",
                    "MaxDrawdown1Y": None,
                    "MaxDrawdown3Y": None,
                    "MonthlyLossCount1Y": None,
                    "MonthlyLossCount3Y": None,
                    "DetailWarnings": "离线模式未重新抓取网页",
                },
                code,
                local_lookup,
            )
        else:
            detail = fetch_detail_for_fund(row, fetcher, local_lookup)
        merged = {**row, **detail}
        if not merged.get("Scale") and merged.get("ScaleText"):
            merged["Scale"] = to_float(merged["ScaleText"])
        warning = detail_reject_reason(merged)
        if warning:
            merged["SearchWarning"] = warning.replace("深度剔除：", "")
            warning_rows.append({**merged, "RejectReason": f"搜索提示：{merged['SearchWarning']}"})
        enriched_rows.append(merged)

    scored = score_candidates(pd.DataFrame(enriched_rows))
    if not scored.empty:
        for idx, row in scored.iterrows():
            warning = clean_text(row.get("SearchWarning"))
            if warning:
                current = str(row.get("MainNotes") or "")
                scored.loc[idx, "MainNotes"] = f"搜索提示：{warning}；{current}" if current else f"搜索提示：{warning}"
        scored = scored.sort_values(["MatchScore", "TotalScore"], ascending=[False, False])
    top = scored.head(10).copy()
    rejected = pd.DataFrame(warning_rows)
    append_log(logs, f"搜索评分完成：展示 {len(top)} 只；提示项 {len(rejected)} 条", progress)
    if fetcher.used_stale_cache:
        append_log(logs, "注意：部分请求使用了过期缓存", progress)

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    top_csv = OUTPUT_DIR / "搜索结果.csv"
    full_xlsx = OUTPUT_DIR / "搜索评分表.xlsx"
    reject_csv = OUTPUT_DIR / "搜索提示清单.csv"
    report_html = OUTPUT_DIR / "搜索评分报告.html"
    log_file = OUTPUT_DIR / "运行日志.txt"
    top_csv, full_xlsx, reject_csv, report_html = write_outputs(
        top, scored, rejected, top_csv, full_xlsx, reject_csv, report_html, logs
    )
    actual_output_dir = top_csv.parent
    top.to_csv(actual_output_dir / f"搜索结果_{stamp}.csv", index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(actual_output_dir / f"搜索评分表_{stamp}.xlsx", engine="openpyxl") as writer:
        top.to_excel(writer, index=False, sheet_name="搜索结果")
        scored.to_excel(writer, index=False, sheet_name="完整评分表")
        rejected.to_excel(writer, index=False, sheet_name="搜索提示清单")
    (actual_output_dir / f"搜索评分报告_{stamp}.html").write_text(report_html.read_text(encoding="utf-8"), encoding="utf-8")
    log_file = safe_text_write(actual_output_dir / log_file.name, "\n".join(logs), logs)
    return AnalysisResult(top, scored, rejected, actual_output_dir, top_csv, full_xlsx, report_html, reject_csv, log_file, logs)


REPORT_COLUMNS = [
    "Code",
    "Name",
    "FundCompany",
    "FundType",
    "Scale",
    "TotalScore",
    "OneYear",
    "TwoYear",
    "ThreeYear",
    "MaxDrawdown1Y",
    "MaxDrawdown3Y",
    "Manager",
    "CurrentMgrTenure",
    "MgmtFee",
    "CustodyFee",
    "SalesServiceFee",
    "SubscriptionFee",
    "Top5BondConcentration",
    "Top5Bonds",
    "MainNotes",
    "FundPage",
    "FeePage",
    "BondPage",
]


def write_outputs(
    top10: pd.DataFrame,
    scored: pd.DataFrame,
    rejected: pd.DataFrame,
    top_csv: Path,
    full_xlsx: Path,
    reject_csv: Path,
    report_html: Path,
    logs: list[str],
) -> tuple[Path, Path, Path, Path]:
    export_top = top10.copy()
    for col in REPORT_COLUMNS:
        if col not in export_top.columns:
            export_top[col] = ""
    top_csv = safe_csv_write(export_top[REPORT_COLUMNS], top_csv, logs)
    reject_csv = safe_csv_write(rejected, reject_csv, logs)
    full_xlsx = safe_excel_write(export_top, scored, rejected, full_xlsx, logs)
    report_html = safe_text_write(report_html, render_report_html(export_top, scored, rejected, logs), logs)
    return top_csv, full_xlsx, reject_csv, report_html


def timestamped_sibling(path: Path) -> Path:
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{stamp}{path.suffix}")


def fallback_output_dirs() -> list[Path]:
    dirs: list[Path] = []
    if getattr(sys, "frozen", False):
        dirs.append(Path(sys.executable).resolve().parent / "output")
    else:
        dirs.append(PROGRAM_DIR / "output")
    base = os.environ.get("LOCALAPPDATA")
    if base:
        dirs.append(Path(base) / "WenShouBao" / "output")
    dirs.append(Path.home() / "WenShouBao" / "output")
    dirs.append(Path("C:/tmp/WenShouBao/output"))
    return dirs


def write_candidates(path: Path) -> list[Path]:
    candidates = [
        path,
        timestamped_sibling(path),
    ]
    for fallback in fallback_output_dirs():
        candidates.append(fallback / path.name)
        candidates.append(timestamped_sibling(fallback / path.name))
    return candidates


def safe_csv_write(df: pd.DataFrame, path: Path, logs: list[str]) -> Path:
    last_error: OSError | None = None
    for index, candidate in enumerate(write_candidates(path)):
        try:
            candidate.parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(candidate, index=False, encoding="utf-8-sig")
            if index > 0:
                logs.append(f"[{now_text()}] 默认输出位置不可写或文件被占用，已写入：{candidate}")
            return candidate
        except OSError as exc:
            last_error = exc
    raise last_error or PermissionError(str(path))


def safe_excel_write(top10: pd.DataFrame, scored: pd.DataFrame, rejected: pd.DataFrame, path: Path, logs: list[str]) -> Path:
    last_error: OSError | None = None
    for index, candidate in enumerate(write_candidates(path)):
        try:
            candidate.parent.mkdir(parents=True, exist_ok=True)
            with pd.ExcelWriter(candidate, engine="openpyxl") as writer:
                top10[REPORT_COLUMNS].to_excel(writer, index=False, sheet_name="前十名")
                scored.to_excel(writer, index=False, sheet_name="完整评分表")
                rejected.to_excel(writer, index=False, sheet_name="剔除清单")
            if index > 0:
                logs.append(f"[{now_text()}] 默认输出位置不可写或文件被占用，已写入：{candidate}")
            return candidate
        except OSError as exc:
            last_error = exc
    raise last_error or PermissionError(str(path))


def safe_text_write(path: Path, content: str, logs: list[str]) -> Path:
    last_error: OSError | None = None
    for index, candidate in enumerate(write_candidates(path)):
        try:
            candidate.parent.mkdir(parents=True, exist_ok=True)
            candidate.write_text(content, encoding="utf-8")
            if index > 0:
                logs.append(f"[{now_text()}] 默认输出位置不可写或文件被占用，已写入：{candidate}")
            return candidate
        except OSError as exc:
            last_error = exc
    raise last_error or PermissionError(str(path))


def render_report_html(top10: pd.DataFrame, scored: pd.DataFrame, rejected: pd.DataFrame, logs: list[str]) -> str:
    rows = []
    for rank, (_, row) in enumerate(top10.iterrows(), start=1):
        code = html.escape(str(row.get("Code", "")))
        name = html.escape(str(row.get("Name", "")))
        fund_url = html.escape(str(row.get("FundPage", "")))
        fee_url = html.escape(str(row.get("FeePage", "")))
        bond_url = html.escape(str(row.get("BondPage", "")))
        rows.append(
            "<tr>"
            f"<td>{rank}</td>"
            f"<td><a href='{fund_url}'>{code}</a></td>"
            f"<td>{name}</td>"
            f"<td>{html.escape(str(row.get('FundCompany', '')))}</td>"
            f"<td>{html.escape(str(row.get('FundType', row.get('Subtype', ''))))}</td>"
            f"<td>{to_float(row.get('Scale')) or ''}</td>"
            f"<td class='score'>{to_float(row.get('TotalScore'), 0):.2f}</td>"
            f"<td>{pct_text(row.get('OneYear'))}</td>"
            f"<td>{pct_text(row.get('TwoYear'))}</td>"
            f"<td>{pct_text(row.get('ThreeYear'))}</td>"
            f"<td>{pct_text(row.get('MaxDrawdown1Y'))}</td>"
            f"<td>{pct_text(row.get('MaxDrawdown3Y'))}</td>"
            f"<td>{html.escape(str(row.get('Manager', '')))}</td>"
            f"<td>{html.escape(str(row.get('CurrentMgrTenure', '')))}</td>"
            f"<td>{html.escape(str(row.get('MgmtFee', '')))} / {html.escape(str(row.get('CustodyFee', '')))} / {html.escape(str(row.get('SalesServiceFee', '')))}</td>"
            f"<td>{pct_text(row.get('Top5BondConcentration'))}</td>"
            f"<td>{html.escape(str(row.get('Top5Bonds', '')))}</td>"
            f"<td>{html.escape(str(row.get('MainNotes', '')))}</td>"
            f"<td><a href='{fee_url}'>费率</a> · <a href='{bond_url}'>持仓</a></td>"
            "</tr>"
        )
    generated_at = now_text()
    log_tail = "\n".join(logs[-20:])
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>稳收宝报告</title>
<style>
body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 24px; color: #1f2933; background: #f7f8fa; }}
h1 {{ margin: 0 0 8px; font-size: 28px; }}
.meta {{ color: #5b6472; margin-bottom: 20px; }}
.notice {{ background: #fff7e6; border: 1px solid #ffd591; padding: 12px 14px; border-radius: 6px; margin-bottom: 18px; }}
.summary {{ display: flex; gap: 12px; margin: 16px 0; }}
.summary div {{ background: #fff; border: 1px solid #e5e7eb; border-radius: 6px; padding: 10px 14px; min-width: 150px; }}
table {{ border-collapse: collapse; width: 100%; background: #fff; font-size: 13px; }}
th, td {{ border: 1px solid #e5e7eb; padding: 8px; vertical-align: top; }}
th {{ background: #edf2f7; position: sticky; top: 0; }}
a {{ color: #0f62fe; text-decoration: none; }}
.score {{ font-weight: 700; color: #b42318; }}
pre {{ white-space: pre-wrap; background: #111827; color: #e5e7eb; padding: 12px; border-radius: 6px; }}
</style>
</head>
<body>
<h1>稳收宝低风险债基筛选报告</h1>
<div class="meta">生成时间：{generated_at}；数据来源：天天基金/东方财富公开页面与本地缓存。</div>
<div class="notice">本报告仅用于基金筛选研究，不构成任何投资建议或收益承诺。购买前请到天天基金、基金公司官网、销售平台再次核对费率、经理、持仓、限购和风险等级。</div>
<div class="summary">
  <div>前十名：<strong>{len(top10)}</strong></div>
  <div>评分候选：<strong>{len(scored)}</strong></div>
  <div>剔除记录：<strong>{len(rejected)}</strong></div>
</div>
<table>
<thead>
<tr>
<th>排名</th><th>代码</th><th>基金名称</th><th>基金公司</th><th>类型</th><th>规模(亿)</th><th>总分</th>
<th>近1年</th><th>近2年</th><th>近3年</th><th>近1年回撤</th><th>近3年回撤</th><th>经理</th><th>经理任期</th>
<th>费率(管/托/销)</th><th>前五集中度</th><th>前五债券</th><th>主要提示</th><th>核验</th>
</tr>
</thead>
<tbody>
{''.join(rows)}
</tbody>
</table>
<h2>最近运行日志</h2>
<pre>{html.escape(log_tail)}</pre>
</body>
</html>
"""


def run_cli(args: argparse.Namespace) -> int:
    def progress(message: str) -> None:
        print(f"[{now_text()}] {message}", flush=True)

    result = run_analysis(deep_limit=args.deep_limit, offline=args.offline, progress=progress)
    print("\n前十名：")
    display_cols = ["Code", "Name", "FundCompany", "FundType", "Scale", "TotalScore", "OneYear", "ThreeYear", "MaxDrawdown1Y", "MaxDrawdown3Y", "MainNotes"]
    print(result.top10[[c for c in display_cols if c in result.top10.columns]].to_string(index=False))
    print(f"\n输出目录：{result.output_dir}")
    return 0


def launch_gui() -> None:
    import tkinter as tk
    from tkinter import messagebox, ttk

    class App(tk.Tk):
        def __init__(self) -> None:
            super().__init__()
            self.title("稳收宝 - 低风险债基筛选")
            self.geometry("1180x760")
            self.minsize(980, 620)
            self.result: AnalysisResult | None = None
            self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
            self._build_ui()
            self.after(150, self._drain_events)

        def _build_ui(self) -> None:
            top = ttk.Frame(self, padding=12)
            top.pack(fill=tk.X)
            title = ttk.Label(top, text="稳收宝", font=("Microsoft YaHei", 22, "bold"))
            title.pack(side=tk.LEFT)
            subtitle = ttk.Label(top, text="低风险债券基金自动筛选、评分与报告输出")
            subtitle.pack(side=tk.LEFT, padx=14)

            buttons = ttk.Frame(self, padding=(12, 0))
            buttons.pack(fill=tk.X)
            self.run_btn = ttk.Button(buttons, text="抓取并分析", command=self._start_run)
            self.run_btn.pack(side=tk.LEFT, padx=(0, 8))
            self.report_btn = ttk.Button(buttons, text="打开前十报告", command=lambda: self._open_path("report"), state=tk.DISABLED)
            self.report_btn.pack(side=tk.LEFT, padx=4)
            self.xlsx_btn = ttk.Button(buttons, text="打开完整评分表", command=lambda: self._open_path("xlsx"), state=tk.DISABLED)
            self.xlsx_btn.pack(side=tk.LEFT, padx=4)
            self.reject_btn = ttk.Button(buttons, text="查看剔除清单", command=lambda: self._open_path("reject"), state=tk.DISABLED)
            self.reject_btn.pack(side=tk.LEFT, padx=4)
            self.out_btn = ttk.Button(buttons, text="打开输出目录", command=lambda: os.startfile(OUTPUT_DIR), state=tk.DISABLED)
            self.out_btn.pack(side=tk.LEFT, padx=4)

            self.status = tk.StringVar(value="准备就绪。点击“抓取并分析”开始。")
            status_label = ttk.Label(self, textvariable=self.status, padding=(12, 8))
            status_label.pack(fill=tk.X)

            columns = ("rank", "code", "name", "company", "type", "scale", "score", "one", "three", "mdd", "manager", "notes")
            self.tree = ttk.Treeview(self, columns=columns, show="headings", height=15)
            headings = {
                "rank": "排名",
                "code": "代码",
                "name": "基金名称",
                "company": "基金公司",
                "type": "类型",
                "scale": "规模(亿)",
                "score": "总分",
                "one": "近1年",
                "three": "近3年",
                "mdd": "近1年回撤",
                "manager": "经理",
                "notes": "主要提示",
            }
            widths = {
                "rank": 48,
                "code": 70,
                "name": 180,
                "company": 100,
                "type": 110,
                "scale": 80,
                "score": 70,
                "one": 70,
                "three": 70,
                "mdd": 90,
                "manager": 90,
                "notes": 260,
            }
            for col in columns:
                self.tree.heading(col, text=headings[col])
                self.tree.column(col, width=widths[col], anchor=tk.W)
            self.tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

            log_frame = ttk.LabelFrame(self, text="运行日志", padding=8)
            log_frame.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0, 12))
            self.log_text = tk.Text(log_frame, height=9, wrap=tk.WORD)
            self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
            scroll.pack(side=tk.RIGHT, fill=tk.Y)
            self.log_text.configure(yscrollcommand=scroll.set)

        def _start_run(self) -> None:
            self.run_btn.configure(state=tk.DISABLED)
            self.report_btn.configure(state=tk.DISABLED)
            self.xlsx_btn.configure(state=tk.DISABLED)
            self.reject_btn.configure(state=tk.DISABLED)
            self.out_btn.configure(state=tk.DISABLED)
            self.status.set("正在抓取并分析，请稍候...")
            self.log_text.delete("1.0", tk.END)
            for item in self.tree.get_children():
                self.tree.delete(item)
            thread = threading.Thread(target=self._worker, daemon=True)
            thread.start()

        def _worker(self) -> None:
            try:
                result = run_analysis(progress=lambda msg: self.events.put(("log", msg)))
                self.events.put(("done", result))
            except Exception as exc:
                self.events.put(("error", f"{exc}\n\n{traceback.format_exc()}"))

        def _drain_events(self) -> None:
            try:
                while True:
                    kind, payload = self.events.get_nowait()
                    if kind == "log":
                        self.status.set(str(payload))
                        self.log_text.insert(tk.END, f"[{now_text()}] {payload}\n")
                        self.log_text.see(tk.END)
                    elif kind == "done":
                        self.result = payload
                        self._show_result(payload)
                    elif kind == "error":
                        self.run_btn.configure(state=tk.NORMAL)
                        self.status.set("运行失败，请查看错误信息。")
                        self.log_text.insert(tk.END, str(payload))
                        messagebox.showerror("稳收宝运行失败", str(payload).splitlines()[0])
            except queue.Empty:
                pass
            self.after(150, self._drain_events)

        def _show_result(self, result: AnalysisResult) -> None:
            self.run_btn.configure(state=tk.NORMAL)
            self.report_btn.configure(state=tk.NORMAL)
            self.xlsx_btn.configure(state=tk.NORMAL)
            self.reject_btn.configure(state=tk.NORMAL)
            self.out_btn.configure(state=tk.NORMAL)
            self.status.set(f"完成：前十名 {len(result.top10)} 只；完整评分 {len(result.scored)} 只；剔除 {len(result.rejected)} 条。")
            for rank, (_, row) in enumerate(result.top10.iterrows(), start=1):
                self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        rank,
                        row.get("Code", ""),
                        row.get("Name", ""),
                        row.get("FundCompany", ""),
                        row.get("FundType", row.get("Subtype", "")),
                        f"{to_float(row.get('Scale'), 0):.2f}",
                        f"{to_float(row.get('TotalScore'), 0):.2f}",
                        pct_text(row.get("OneYear")),
                        pct_text(row.get("ThreeYear")),
                        pct_text(row.get("MaxDrawdown1Y")),
                        row.get("Manager", ""),
                        row.get("MainNotes", ""),
                    ),
                )
            messagebox.showinfo("稳收宝完成", f"分析完成，报告已生成到：\n{result.output_dir}")

        def _open_path(self, kind: str) -> None:
            if not self.result:
                return
            path_map = {
                "report": self.result.report_html,
                "xlsx": self.result.full_xlsx,
                "reject": self.result.reject_csv,
            }
            path = path_map[kind]
            if path.exists():
                os.startfile(path)
            else:
                messagebox.showwarning("文件不存在", str(path))

    App().mainloop()


def launch_gui() -> None:
    from PySide6.QtCore import QObject, QThread, Qt, Signal, Slot
    from PySide6.QtGui import QAction, QIcon
    from PySide6.QtWidgets import (
        QAbstractItemView,
        QApplication,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMenu,
        QMessageBox,
        QPushButton,
        QTableWidget,
        QTableWidgetItem,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )

    class Worker(QObject):
        log = Signal(str)
        done = Signal(object)
        error = Signal(str)

        def __init__(self, task: str, query: str = "") -> None:
            super().__init__()
            self.task = task
            self.query = query

        @Slot()
        def run(self) -> None:
            try:
                if self.task == "search":
                    result = run_fund_search(self.query, progress=lambda msg: self.log.emit(msg))
                else:
                    result = run_analysis(progress=lambda msg: self.log.emit(msg))
                self.done.emit(result)
            except Exception:
                self.error.emit(traceback.format_exc())

    class Window(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle(f"稳收宝 - 低风险债基筛选 {APP_VERSION}")
            icon_path = bundled_resource(APP_ICON_RELATIVE)
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
            self.resize(1180, 760)
            self.result: AnalysisResult | None = None
            self.thread: QThread | None = None
            self.worker: Worker | None = None
            self.current_task = "analysis"
            self._build_ui()

        def _build_ui(self) -> None:
            root = QWidget()
            root.setObjectName("root")
            root.setStyleSheet(
                """
                #root {
                    background-color: #15181e;
                }
                QLabel {
                    background-color: transparent;
                    color: #d8e0ea;
                }
                QPushButton {
                    background-color: #222a34;
                    color: #edf5ff;
                    border: 1px solid #3a4656;
                    border-radius: 6px;
                    padding: 7px 14px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #2b3542;
                    border-color: #5a718b;
                }
                QPushButton:pressed {
                    background-color: #19212b;
                }
                QPushButton:disabled {
                    background-color: #1b2028;
                    color: #6b7685;
                    border-color: #2c3440;
                }
                QPushButton#primaryButton {
                    background-color: #0f766e;
                    color: #f8fffd;
                    border-color: #14b8a6;
                }
                QPushButton#primaryButton:hover {
                    background-color: #10857c;
                }
                QLineEdit, QTextEdit {
                    background-color: #20252d;
                    color: #edf5ff;
                    border: 1px solid #3a4656;
                    border-radius: 6px;
                    padding: 6px 8px;
                    selection-background-color: #155e75;
                    selection-color: #ffffff;
                }
                QLineEdit:focus, QTextEdit:focus {
                    border-color: #22d3ee;
                }
                QMenu {
                    background-color: #20252d;
                    color: #edf5ff;
                    border: 1px solid #3a4656;
                }
                QMenu::item:selected {
                    background-color: #155e75;
                }
                """
            )
            layout = QVBoxLayout(root)
            layout.setContentsMargins(14, 14, 14, 14)
            layout.setSpacing(10)

            title_bar = QHBoxLayout()
            title = QLabel("稳收宝")
            title.setStyleSheet("font-size: 30px; font-weight: 800; color: #f8fafc;")
            version_box = QVBoxLayout()
            version_box.setSpacing(0)
            title_mark = QLabel("我爱洗澡")
            title_mark.setStyleSheet("font-size: 12px; color: #7dd3fc;")
            version_box.addWidget(title_mark)
            version = QLabel(APP_VERSION)
            version.setStyleSheet("font-size: 13px; color: #a7b4c4;")
            version_box.addWidget(version)
            title_bar.addWidget(title)
            title_bar.addStretch()
            title_bar.addLayout(version_box)
            subtitle = QLabel("低风险债券基金自动筛选、评分与报告输出。仅供参考，不构成投资建议。")
            subtitle.setStyleSheet("color: #8fa0b3;")
            layout.addLayout(title_bar)
            layout.addWidget(subtitle)

            button_bar = QHBoxLayout()
            self.run_btn = QPushButton("抓取并分析")
            self.run_btn.setObjectName("primaryButton")
            self.run_btn.clicked.connect(self.start_run)
            self.report_btn = QPushButton("打开前十报告")
            self.report_btn.clicked.connect(lambda: self.open_path("report"))
            self.xlsx_btn = QPushButton("显示完整评分表")
            self.xlsx_btn.clicked.connect(self.show_full_scores)
            self.reject_btn = QPushButton("查看剔除清单")
            self.reject_btn.clicked.connect(lambda: self.open_path("reject"))
            self.output_btn = QPushButton("打开输出目录")
            self.output_btn.clicked.connect(self.open_output_dir)
            button_bar.addWidget(self.run_btn)
            for btn in (self.report_btn, self.xlsx_btn, self.reject_btn, self.output_btn):
                button_bar.addWidget(btn)
            button_bar.addStretch()
            layout.addLayout(button_bar)

            search_bar = QHBoxLayout()
            search_label = QLabel("基金搜索")
            search_label.setStyleSheet("color: #93c5fd; font-weight: 600;")
            self.search_input = QLineEdit()
            self.search_input.setPlaceholderText("输入基金名称或代码，支持模糊搜索")
            self.search_input.returnPressed.connect(self.start_search)
            self.search_btn = QPushButton("搜索基金并评分")
            self.search_btn.clicked.connect(self.start_search)
            search_bar.addWidget(search_label)
            search_bar.addWidget(self.search_input, 1)
            search_bar.addWidget(self.search_btn)
            layout.addLayout(search_bar)

            self.status = QLabel("准备就绪。点击“抓取并分析”开始。")
            self.status.setStyleSheet("padding: 8px 0; color: #e5edf8; font-weight: 600;")
            layout.addWidget(self.status)

            self.table = QTableWidget(0, 0)
            self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
            self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            self.table.setAlternatingRowColors(True)
            self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self.table.customContextMenuRequested.connect(self.show_table_menu)
            self.table.setStyleSheet(
                """
                QTableWidget {
                    background-color: #1b2028;
                    alternate-background-color: #202733;
                    color: #eef6ff;
                    gridline-color: #344154;
                    selection-background-color: #155e75;
                    selection-color: #ffffff;
                    border: 1px solid #3a4656;
                    border-radius: 6px;
                }
                QTableWidget::viewport {
                    background-color: #1b2028;
                }
                QTableWidget::item:selected {
                    background-color: #155e75;
                    color: #ffffff;
                }
                QHeaderView::section {
                    background: #243140;
                    color: #eaf6ff;
                    padding: 7px;
                    border: 1px solid #344154;
                    font-weight: 700;
                }
                QTableCornerButton::section {
                    background: #243140;
                    border: 1px solid #344154;
                }
                """
            )
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
            self.table.horizontalHeader().setStretchLastSection(True)
            layout.addWidget(self.table, 1)

            self.log_box = QTextEdit()
            self.log_box.setReadOnly(True)
            self.log_box.setPlaceholderText("运行日志")
            self.log_box.setMinimumHeight(150)
            self.log_box.setStyleSheet(
                """
                QTextEdit {
                    background-color: #20252d;
                    color: #edf5ff;
                    border: 1px solid #3a4656;
                    border-radius: 6px;
                    padding: 6px 8px;
                    font-family: Consolas, "Microsoft YaHei";
                }
                """
            )
            self.log_box.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self.log_box.customContextMenuRequested.connect(self.show_log_menu)
            layout.addWidget(self.log_box)
            self.setCentralWidget(root)
            self.set_result_buttons(False)

            self.search_input.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self.search_input.customContextMenuRequested.connect(self.show_search_menu)

        def set_result_buttons(self, enabled: bool) -> None:
            self.report_btn.setEnabled(enabled)
            self.xlsx_btn.setEnabled(enabled)
            self.reject_btn.setEnabled(enabled)
            self.output_btn.setEnabled(enabled)

        def start_run(self) -> None:
            self.start_task("analysis")

        def start_search(self) -> None:
            query = self.search_input.text().strip()
            if not query:
                QMessageBox.information(self, "请输入基金", "请输入基金名称或基金代码后再搜索。")
                return
            self.start_task("search", query)

        def start_task(self, task: str, query: str = "") -> None:
            if self.thread is not None and self.thread.isRunning():
                return
            self.current_task = task
            self.run_btn.setEnabled(False)
            self.search_btn.setEnabled(False)
            self.search_input.setEnabled(False)
            self.set_result_buttons(False)
            self.table.setRowCount(0)
            self.log_box.clear()
            if task == "search":
                self.status.setText(f"正在搜索并评分：{query}")
            else:
                self.status.setText("正在抓取并分析，请稍候...")
            self.clear_busy_cursor()

            self.thread = QThread()
            self.worker = Worker(task, query)
            self.worker.moveToThread(self.thread)
            self.thread.started.connect(self.worker.run)
            self.worker.log.connect(self.on_log)
            self.worker.done.connect(self.thread.quit)
            self.worker.error.connect(self.thread.quit)
            self.worker.done.connect(self.worker.deleteLater)
            self.worker.error.connect(self.worker.deleteLater)
            self.worker.done.connect(self.on_done)
            self.worker.error.connect(self.on_error)
            self.thread.finished.connect(self.on_thread_finished)
            self.thread.finished.connect(self.thread.deleteLater)
            self.thread.start()

        def clear_busy_cursor(self) -> None:
            while QApplication.overrideCursor() is not None:
                QApplication.restoreOverrideCursor()

        @Slot()
        def on_thread_finished(self) -> None:
            self.clear_busy_cursor()
            self.worker = None
            self.thread = None

        @Slot(str)
        def on_log(self, message: str) -> None:
            self.status.setText(message)
            self.log_box.append(f"[{now_text()}] {message}")

        @Slot(object)
        def on_done(self, result: AnalysisResult) -> None:
            self.clear_busy_cursor()
            self.result = result
            self.run_btn.setEnabled(True)
            self.search_btn.setEnabled(True)
            self.search_input.setEnabled(True)
            self.set_result_buttons(True)
            label = "搜索结果" if self.current_task == "search" else "前十名"
            self.status.setText(f"完成：{label} {len(result.top10)} 只；完整评分 {len(result.scored)} 只；提示/剔除 {len(result.rejected)} 条。")
            self.populate_top_table(result.top10)
            self.log_box.append(f"[{now_text()}] 分析完成，报告已生成到：{result.output_dir}")

        @Slot(str)
        def on_error(self, error_text: str) -> None:
            self.clear_busy_cursor()
            self.run_btn.setEnabled(True)
            self.search_btn.setEnabled(True)
            self.search_input.setEnabled(True)
            self.status.setText("运行失败，请查看日志。")
            self.log_box.append(error_text)
            QMessageBox.critical(self, "稳收宝运行失败", error_text.splitlines()[-1] if error_text else "未知错误")

        def populate_top_table(self, df: pd.DataFrame) -> None:
            headers = ["排名", "代码", "基金名称", "基金公司", "类型", "规模(亿)", "总分", "近1年", "近3年", "近1年回撤", "近3年回撤", "经理", "主要提示"]
            self.table.clear()
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)
            self.table.setRowCount(len(df))
            for rank, (_, row) in enumerate(df.iterrows(), start=1):
                values = [
                    rank,
                    row.get("Code", ""),
                    row.get("Name", ""),
                    row.get("FundCompany", ""),
                    row.get("FundType", row.get("Subtype", "")),
                    f"{to_float(row.get('Scale'), 0):.2f}",
                    f"{to_float(row.get('TotalScore'), 0):.2f}",
                    pct_text(row.get("OneYear")),
                    pct_text(row.get("ThreeYear")),
                    pct_text(row.get("MaxDrawdown1Y")),
                    pct_text(row.get("MaxDrawdown3Y")),
                    row.get("Manager", ""),
                    row.get("MainNotes", ""),
                ]
                for col, value in enumerate(values):
                    self.table.setItem(rank - 1, col, QTableWidgetItem(str(value)))

        def show_full_scores(self) -> None:
            if not self.result:
                return
            df = self.result.scored.copy()
            if df.empty:
                QMessageBox.information(self, "完整评分表", "当前没有可显示的完整评分数据。")
                return
            self.populate_dataframe(df)
            self.status.setText(f"已在主界面显示完整评分表：{len(df)} 只基金，{len(df.columns)} 个字段。")

        def populate_dataframe(self, df: pd.DataFrame) -> None:
            labels = {
                "Code": "代码",
                "Name": "基金名称",
                "Pinyin": "拼音缩写",
                "FundCompany": "基金公司",
                "FundType": "类型",
                "Subtype": "类型",
                "Date": "净值日期",
                "UnitNav": "单位净值",
                "AccumNav": "累计净值",
                "Daily": "日涨幅",
                "Week": "近1周",
                "Month": "近1月",
                "ThreeMonth": "近3月",
                "SixMonth": "近6月",
                "StartDate": "成立日",
                "OriginalFee": "原申购费",
                "DiscountFee": "优惠申购费",
                "CodeMapName": "代码表名称",
                "Scale": "规模(亿)",
                "ScaleText": "规模文本",
                "TotalScore": "总分",
                "PerfScore": "历史表现分",
                "PerfScorePre": "深度核验预排序分",
                "PreselectScore": "综合预筛分",
                "MatchScore": "搜索匹配分",
                "FundCategory": "基金类别",
                "Group": "样本分组",
                "ScoreRiskBoundary": "风险边界分",
                "ScorePerformance": "业绩分",
                "ScoreDrawdown": "回撤波动分",
                "ScoreHolding": "持仓质量分",
                "ScoreManager": "经理团队分",
                "ScoreScale": "规模流动性分",
                "ScoreFee": "费率份额分",
                "ScoreTerm": "期限匹配分",
                "ScoreDisclosure": "披露异常分",
                "ScoreMissingPenalty": "数据缺失扣分",
                "OneYear": "近1年",
                "TwoYear": "近2年",
                "ThreeYear": "近3年",
                "YTD": "今年以来",
                "SinceStart": "成立以来",
                "MaxDrawdown1Y": "近1年回撤",
                "MaxDrawdown3Y": "近3年回撤",
                "MonthlyLossCount1Y": "近1年月亏损次数",
                "MonthlyLossCount3Y": "近3年月亏损次数",
                "DetailSource": "明细数据来源",
                "SourceFallback": "数据兜底来源",
                "DetailWarnings": "明细抓取提示",
                "Manager": "经理",
                "CurrentMgrStart": "当前经理起始日",
                "CurrentMgrTenure": "经理任期",
                "CurrentMgrReturn": "经理任内收益",
                "MgmtFee": "管理费",
                "CustodyFee": "托管费",
                "SalesServiceFee": "销售服务费",
                "SubscriptionFee": "申购费",
                "RedemptionFee": "赎回费",
                "Top5BondConcentration": "前五债券集中度",
                "Top5Bonds": "前五债券",
                "HasConvertibleBond": "是否含转债",
                "AllTop5SafeBonds": "前五是否均为国债/政金债",
                "BondCount": "债券持仓数量",
                "MainNotes": "主要提示",
                "FundPage": "基金页",
                "BasicPage": "概况页",
                "FeePage": "费率页",
                "ManagerPage": "经理页",
                "BondPage": "持仓页",
                "CompanyVerifyPage": "基金公司核验页",
                "FundKey": "同基金去重键",
                "SearchWarning": "搜索风险提示",
                "RejectReason": "剔除/提示原因",
            }
            columns = list(df.columns)
            self.table.clear()
            self.table.setColumnCount(len(columns))
            self.table.setHorizontalHeaderLabels([labels.get(col, col) for col in columns])
            self.table.setRowCount(len(df))
            pct_columns = {"OneYear", "TwoYear", "ThreeYear", "YTD", "SinceStart", "MaxDrawdown1Y", "MaxDrawdown3Y", "Top5BondConcentration"}
            score_columns = {col for col in columns if col.startswith("Score") or col == "TotalScore"}
            for row_index, (_, row) in enumerate(df.iterrows()):
                for col_index, col in enumerate(columns):
                    value = row.get(col, "")
                    if col in pct_columns:
                        text = pct_text(value)
                    elif col in score_columns or col in {"Scale", "PerfScorePre", "MatchScore"}:
                        number = to_float(value)
                        text = "" if number is None else f"{number:.2f}"
                    elif value is None or (isinstance(value, float) and math.isnan(value)):
                        text = ""
                    else:
                        text = str(value)
                    self.table.setItem(row_index, col_index, QTableWidgetItem(text))

        def show_table_menu(self, pos: Any) -> None:
            row = self.table.rowAt(pos.y())
            col = self.table.columnAt(pos.x())
            if row >= 0:
                self.table.selectRow(row)
            menu = QMenu(self)
            copy_cell = QAction("复制单元格", self)
            copy_row = QAction("复制整行", self)
            copy_selected = QAction("复制选中内容", self)
            copy_headers = QAction("复制表头", self)
            copy_cell.setEnabled(row >= 0 and col >= 0)
            copy_row.setEnabled(row >= 0)
            copy_selected.setEnabled(bool(self.table.selectedIndexes()))
            copy_cell.triggered.connect(lambda: self.copy_table_cell(row, col))
            copy_row.triggered.connect(lambda: self.copy_table_row(row))
            copy_selected.triggered.connect(self.copy_table_selection)
            copy_headers.triggered.connect(self.copy_table_headers)
            menu.addAction(copy_cell)
            menu.addAction(copy_row)
            menu.addAction(copy_selected)
            menu.addSeparator()
            menu.addAction(copy_headers)
            menu.exec(self.table.viewport().mapToGlobal(pos))

        def copy_table_cell(self, row: int, col: int) -> None:
            item = self.table.item(row, col)
            QApplication.clipboard().setText(item.text() if item else "")

        def copy_table_row(self, row: int) -> None:
            values = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                values.append(item.text() if item else "")
            QApplication.clipboard().setText("\t".join(values))

        def copy_table_headers(self) -> None:
            headers = []
            for col in range(self.table.columnCount()):
                item = self.table.horizontalHeaderItem(col)
                headers.append(item.text() if item else "")
            QApplication.clipboard().setText("\t".join(headers))

        def copy_table_selection(self) -> None:
            indexes = sorted(self.table.selectedIndexes(), key=lambda idx: (idx.row(), idx.column()))
            if not indexes:
                return
            rows: dict[int, dict[int, str]] = {}
            for index in indexes:
                item = self.table.item(index.row(), index.column())
                rows.setdefault(index.row(), {})[index.column()] = item.text() if item else ""
            lines = []
            for row in sorted(rows):
                lines.append("\t".join(rows[row].get(col, "") for col in sorted(rows[row])))
            QApplication.clipboard().setText("\n".join(lines))

        def show_log_menu(self, pos: Any) -> None:
            menu = QMenu(self)
            copy_action = QAction("复制", self)
            select_all_action = QAction("全选", self)
            clear_action = QAction("清空日志", self)
            copy_action.triggered.connect(self.log_box.copy)
            select_all_action.triggered.connect(self.log_box.selectAll)
            clear_action.triggered.connect(self.log_box.clear)
            menu.addAction(copy_action)
            menu.addAction(select_all_action)
            menu.addSeparator()
            menu.addAction(clear_action)
            menu.exec(self.log_box.viewport().mapToGlobal(pos))

        def show_search_menu(self, pos: Any) -> None:
            menu = QMenu(self)
            cut_action = QAction("剪切", self)
            copy_action = QAction("复制", self)
            paste_action = QAction("粘贴", self)
            select_all_action = QAction("全选", self)
            clear_action = QAction("清空", self)
            cut_action.triggered.connect(self.search_input.cut)
            copy_action.triggered.connect(self.search_input.copy)
            paste_action.triggered.connect(self.search_input.paste)
            select_all_action.triggered.connect(self.search_input.selectAll)
            clear_action.triggered.connect(self.search_input.clear)
            has_selection = self.search_input.hasSelectedText()
            cut_action.setEnabled(has_selection)
            copy_action.setEnabled(has_selection)
            menu.addAction(cut_action)
            menu.addAction(copy_action)
            menu.addAction(paste_action)
            menu.addSeparator()
            menu.addAction(select_all_action)
            menu.addAction(clear_action)
            menu.exec(self.search_input.mapToGlobal(pos))

        def open_path(self, kind: str) -> None:
            if not self.result:
                return
            paths = {
                "report": self.result.report_html,
                "reject": self.result.reject_csv,
            }
            self.open_file(paths[kind])

        def open_output_dir(self) -> None:
            path = self.result.output_dir if self.result else OUTPUT_DIR
            self.open_file(path)

        def open_file(self, path: Path) -> None:
            if not path.exists():
                QMessageBox.warning(self, "文件不存在", str(path))
                return
            try:
                open_with_windows_default(path)
            except Exception as exc:
                QMessageBox.warning(self, "打开失败", f"无法打开：\n{path}\n\n{exc}")

    app = QApplication(sys.argv[:1])
    icon_path = bundled_resource(APP_ICON_RELATIVE)
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = Window()
    window.show()
    app.exec()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="稳收宝：低风险债基自动筛选工具")
    parser.add_argument("--cli", action="store_true", help="使用命令行模式运行")
    parser.add_argument("--offline", action="store_true", help="使用本地历史样本，不联网抓取")
    parser.add_argument("--deep-limit", type=int, default=DEFAULT_DEEP_LIMIT, help="深度核验候选数量")
    args = parser.parse_args(argv)
    if args.cli or args.offline:
        return run_cli(args)
    launch_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
