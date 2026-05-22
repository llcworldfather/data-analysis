# -*- coding: utf-8 -*-
"""物料行情搜索：多后端策略，百度千帆优先，公开 HTML 搜索兜底。"""
from __future__ import annotations

import html
import os
import re
import threading
import time
import urllib.parse
import urllib.request
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any

import requests

# 搜索结果 TTL 缓存：同一物料名 10 分钟内复用上次结果，避免重复 HTTP 请求
_SEARCH_CACHE: dict[str, tuple[float, list]] = {}   # query → (timestamp, results)
_SEARCH_CACHE_TTL = 600.0                             # 秒
_SEARCH_CACHE_MAX = 200                               # 最多缓存条数
_SEARCH_CACHE_LOCK = threading.Lock()


def _write_search_cache(query: str, results: list) -> None:
    """写入搜索缓存；超容量时淘汰最旧的 20 条。"""
    with _SEARCH_CACHE_LOCK:
        _SEARCH_CACHE[query] = (time.monotonic(), results)
        if len(_SEARCH_CACHE) > _SEARCH_CACHE_MAX:
            oldest = sorted(_SEARCH_CACHE.items(), key=lambda kv: kv[1][0])[:20]
            for k, _ in oldest:
                _SEARCH_CACHE.pop(k, None)

SearchHit = dict[str, Any]


def strip_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_search_url(url: str) -> str:
    """DuckDuckGo 等结果常包一层跳转链接，还原真实地址。"""
    url = html.unescape(url or "")
    if "uddg=" in url:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        if qs.get("uddg"):
            return qs["uddg"][0]
    if url.startswith("//"):
        return "https:" + url
    return url


def fetch_page_summary(url: str) -> dict[str, str]:
    """抓取搜索结果页的标题和 meta 描述。"""
    if not url or not url.startswith(("http://", "https://")):
        return {"title": "", "description": ""}
    try:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                )
            },
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            page = resp.read(180000).decode("utf-8", errors="ignore")
    except Exception:
        return {"title": "", "description": ""}

    title = ""
    m_title = re.search(r"<title[^>]*>(.*?)</title>", page, re.S | re.I)
    if m_title:
        title = strip_html(m_title.group(1))

    description = ""
    m_desc = re.search(
        r'<meta[^>]+(?:name|property)=["\'](?:description|og:description)["\'][^>]+content=["\']([^"\']+)["\']',
        page,
        re.S | re.I,
    )
    if not m_desc:
        m_desc = re.search(
            r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+(?:name|property)=["\'](?:description|og:description)["\']',
            page,
            re.S | re.I,
        )
    if m_desc:
        description = strip_html(m_desc.group(1))

    return {"title": title[:160], "description": description[:260]}


_DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}


def _post_json(
    url: str,
    payload: dict,
    headers: dict,
    *,
    timeout: tuple[float, float] | float,
) -> dict:
    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=timeout)
        resp.raise_for_status()
        return resp.json()
    except requests.HTTPError as exc:
        detail = (exc.response.text if exc.response is not None else str(exc))[:300]
        code = exc.response.status_code if exc.response is not None else "?"
        raise RuntimeError(f"API 调用失败（HTTP {code}）：{detail}") from exc


class MaterialSearchBackend(ABC):
    """单一搜索源。"""

    @abstractmethod
    def search(self, query: str, *, max_results: int = 3) -> list[SearchHit]:
        ...


class BaiduQianfanBackend(MaterialSearchBackend):
    """百度千帆 AI 搜索 API。"""

    def search(self, query: str, *, max_results: int = 3) -> list[SearchHit]:
        api_key = os.environ.get("BAIDU_SEARCH_API_KEY", "").strip()
        if not api_key:
            return []

        short_query = query.strip()
        if len(short_query) > 60:
            short_query = short_query[:60]

        payload = {
            "messages": [{"role": "user", "content": short_query}],
            "search_source": "baidu_search_v2",
            "resource_type_filter": [
                {"type": "web", "top_k": max(1, min(max_results, 10))}
            ],
            "search_filter": {
                "range": {"page_time": {"gte": "now-3M/d", "lte": "now/d"}}
            },
            "sort": {"priority": "auto"},
        }
        headers = {
            "Content-Type": "application/json",
            "X-Appbuilder-Authorization": f"Bearer {api_key}",
        }
        try:
            data = _post_json(
                "https://qianfan.baidubce.com/v2/ai_search/web_search",
                payload,
                headers,
                timeout=(10.0, 30.0),
            )
        except RuntimeError as exc:
            return [{
                "title": "百度千帆搜索失败",
                "snippet": str(exc)[:240],
                "url": "",
                "source": "baidu_qianfan",
            }]

        if data.get("code") or data.get("message"):
            return [{
                "title": "百度千帆搜索失败",
                "snippet": f"{data.get('code', '')} {data.get('message', '')}".strip(),
                "url": "",
                "source": "baidu_qianfan",
            }]

        results: list[SearchHit] = []
        for item in data.get("references", [])[:max_results]:
            title = str(item.get("title") or item.get("web_anchor") or "").strip()
            snippet = str(item.get("snippet") or item.get("content") or "").strip()
            url = str(item.get("url") or "").strip()
            date = str(item.get("date") or "").strip()
            if title or snippet or url:
                results.append({
                    "title": title,
                    "snippet": snippet,
                    "url": url,
                    "date": date,
                    "source": "baidu_qianfan",
                })
        return results


class _HtmlSearchBackend(MaterialSearchBackend):
    """基于 HTML 抓取的通用后端。"""

    def __init__(
        self,
        *,
        build_url: Any,
        parse_blocks: Any,
        enrich_first_n: int = 2,
    ) -> None:
        self._build_url = build_url
        self._parse_blocks = parse_blocks
        self._enrich_first_n = enrich_first_n

    def search(self, query: str, *, max_results: int = 3) -> list[SearchHit]:
        url = self._build_url(query)
        try:
            req = urllib.request.Request(url, headers=_DEFAULT_HEADERS)
            with urllib.request.urlopen(req, timeout=15) as resp:
                page = resp.read().decode("utf-8", errors="ignore")
        except Exception:
            return []

        results: list[SearchHit] = []
        for block in self._parse_blocks(page):
            title = block.get("title", "")
            link = block.get("url", "")
            snippet = block.get("snippet", "")
            if not title:
                continue
            if len(results) < self._enrich_first_n and link:
                page_summary = fetch_page_summary(link)
                if not snippet:
                    snippet = page_summary.get("description", "")
                if page_summary.get("title") and page_summary["title"] not in title:
                    title = f"{title} | {page_summary['title']}"
            results.append({"title": title, "snippet": snippet, "url": link})
            if len(results) >= max_results:
                break
        return results[:max_results]


def _duckduckgo_blocks(page: str) -> list[dict[str, str]]:
    blocks = re.split(r'<div[^>]+class="[^"]*result[^"]*"[^>]*>', page)
    out: list[dict[str, str]] = []
    for block in blocks:
        title_match = re.search(
            r'class="result__a"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', block, re.S
        )
        if not title_match:
            continue
        snippet_match = re.search(
            r'class="result__snippet"[^>]*>(.*?)</(?:a|div)>', block, re.S
        )
        out.append({
            "url": normalize_search_url(title_match.group(1)),
            "title": strip_html(title_match.group(2)),
            "snippet": strip_html(snippet_match.group(1)) if snippet_match else "",
        })
    return out


def _baidu_html_blocks(page: str) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for block in re.findall(
        r'<div[^>]+(?:class="[^"]*result[^"]*"|tpl="[^"]+")[^>]*>(.*?)</div>\s*</div>',
        page,
        re.S | re.I,
    ):
        title_match = re.search(
            r"<h3[^>]*>.*?<a[^>]+href=\"([^\"]+)\"[^>]*>(.*?)</a>.*?</h3>",
            block,
            re.S | re.I,
        )
        if not title_match:
            continue
        snippet_match = re.search(
            r'<span[^>]+class="[^"]*(?:content-right|c-color-text)[^"]*"[^>]*>(.*?)</span>',
            block,
            re.S | re.I,
        )
        if not snippet_match:
            snippet_match = re.search(
                r'<div[^>]+class="[^"]*(?:c-abstract|c-span-last)[^"]*"[^>]*>(.*?)</div>',
                block,
                re.S | re.I,
            )
        out.append({
            "url": normalize_search_url(title_match.group(1)),
            "title": strip_html(title_match.group(2)),
            "snippet": strip_html(snippet_match.group(1)) if snippet_match else "",
        })
    return out


def _bing_blocks(page: str) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for block in re.findall(r'<li[^>]+class="b_algo"[^>]*>(.*?)</li>', page, re.S | re.I):
        title_match = re.search(
            r"<h2[^>]*>.*?<a[^>]+href=\"([^\"]+)\"[^>]*>(.*?)</a>.*?</h2>",
            block,
            re.S | re.I,
        )
        if not title_match:
            continue
        snippet_match = re.search(r"<p[^>]*>(.*?)</p>", block, re.S | re.I)
        out.append({
            "url": normalize_search_url(title_match.group(1)),
            "title": strip_html(title_match.group(2)),
            "snippet": strip_html(snippet_match.group(1)) if snippet_match else "",
        })
    return out


class StaticFallbackBackend(MaterialSearchBackend):
    """搜索引擎均不可用时的行业入口。"""

    def __init__(self, material_name: str) -> None:
        self._material_name = material_name

    def search(self, query: str, *, max_results: int = 3) -> list[SearchHit]:
        name = self._material_name
        lower = name.lower()
        fallback: list[SearchHit] = [
            {
                "title": f"{name} 1688 报价/批发价格入口",
                "snippet": "用于观察同类物料现货报价区间、供应商报价密度和价格离散度，适合作为采购比价参考。",
                "url": "https://www.1688.com/",
            }
        ]
        if any(k in lower for k in ["ppr", "pp-r", "pp", "管件", "管材", "给水管"]):
            fallback.extend([
                {
                    "title": "PP / PPR 管材上游原料行情入口",
                    "snippet": "PPR 管件和管材主要受 PP 原料、丙烯、原油、装置检修、下游开工和房地产/基建需求影响。",
                    "url": "https://plas.oilchem.net/",
                },
                {
                    "title": "生意社 PP 商品指数与价格行情入口",
                    "snippet": "可用于跟踪 PP 原料价格趋势、上游成本支撑和短期市场情绪。",
                    "url": "https://www.100ppi.com/",
                },
            ])
        if any(k in lower for k in ["pe", "聚乙烯", "包装", "膜", "opp"]):
            fallback.extend([
                {
                    "title": "PE / 包装膜上游原料行情入口",
                    "snippet": "PE、OPP、包装膜价格通常受原油、乙烯/丙烯、聚烯烃供应、薄膜需求和库存影响。",
                    "url": "https://plas.oilchem.net/",
                },
                {
                    "title": "生意社聚乙烯/聚丙烯行情入口",
                    "snippet": "可用于观察聚烯烃原料价格趋势和成本传导压力。",
                    "url": "https://www.100ppi.com/",
                },
            ])
        if any(k in lower for k in ["钛白", "钛白粉"]):
            fallback.append({
                "title": "钛白粉行情与上游钛矿/硫酸成本入口",
                "snippet": "钛白粉价格通常受钛矿、硫酸、出口订单、行业开工和库存变化影响。",
                "url": "https://www.100ppi.com/",
            })
        if any(k in lower for k in ["碳酸钙", "钙粉"]):
            fallback.append({
                "title": "碳酸钙/填充母料行情与矿石能源成本入口",
                "snippet": "碳酸钙价格通常受矿石、能源、运输、环保限产和下游塑料填充需求影响。",
                "url": "https://www.100ppi.com/",
            })
        return fallback[:max_results]


class MaterialPriceSearch:
    """按优先级串联多个后端。"""

    def __init__(self, material_name: str) -> None:
        self._material_name = material_name
        # enrich_first_n=1：只对首条结果额外抓页，减少并发 HTTP 请求数
        self._html_backends = [
            _HtmlSearchBackend(
                build_url=lambda q: "https://duckduckgo.com/html/?"
                + urllib.parse.urlencode({"q": q}),
                parse_blocks=_duckduckgo_blocks,
                enrich_first_n=1,
            ),
            _HtmlSearchBackend(
                build_url=lambda q: "https://www.baidu.com/s?"
                + urllib.parse.urlencode({"wd": q}),
                parse_blocks=_baidu_html_blocks,
                enrich_first_n=0,
            ),
            _HtmlSearchBackend(
                build_url=lambda q: "https://www.bing.com/search?"
                + urllib.parse.urlencode({"q": q}),
                parse_blocks=_bing_blocks,
                enrich_first_n=1,
            ),
        ]

    def search(self, *, max_results: int = 3) -> list[SearchHit]:
        _year = datetime.now().year
        query = f"{self._material_name} {_year} 价格 走势 行情 报价 原材料"

        # TTL 缓存：相同查询 10 分钟内直接返回缓存结果
        with _SEARCH_CACHE_LOCK:
            cached = _SEARCH_CACHE.get(query)
            if cached is not None:
                ts, results = cached
                if time.monotonic() - ts < _SEARCH_CACHE_TTL:
                    return results
                del _SEARCH_CACHE[query]

        baidu = BaiduQianfanBackend()
        hits = baidu.search(query, max_results=max_results)
        if hits and not all(item.get("title") == "百度千帆搜索失败" for item in hits):
            _write_search_cache(query, hits)
            return hits

        for backend in self._html_backends:
            hits = backend.search(query, max_results=max_results)
            if hits:
                _write_search_cache(query, hits)
                return hits

        results = StaticFallbackBackend(self._material_name).search(
            query, max_results=max_results
        )
        _write_search_cache(query, results)
        return results


def web_search_material_price(material_name: str, *, max_results: int = 3) -> list[SearchHit]:
    """公开入口：物料近期价格/行情证据搜索。"""
    return MaterialPriceSearch(material_name).search(max_results=max_results)
