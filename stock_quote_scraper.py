"""Fetch latest stock prices and intraday changes for selected A-share and HK stocks.

The script scrapes Eastmoney mobile quote pages, respecting a >30 second gap
between consecutive requests. Output is printed as a simple table.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass
from typing import Dict, Iterable

import requests

# Minimal headers to look like a regular browser session.
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/129.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

# Regex that extracts the embedded quotedata JSON blob from the page.
QUOTEDATA_PATTERN = re.compile(r"var\s+quotedata\s*=\s*(\{.*?\});", re.DOTALL)

SHANGHAI_PREFIXES = ("60", "68", "5")  # includes ETFs such as 512890
SHENZHEN_PREFIXES = ("00", "30")


@dataclass(frozen=True)
class StockTarget:
    """Configuration for a stock to scrape."""

    secid: str  # e.g. "0.000333"
    url: str    # e.g. "https://wap.eastmoney.com/quote/stock/0.000333.html"


STOCKS: tuple[StockTarget, ...] = (
    StockTarget(
        secid="0.000333",
        url="https://wap.eastmoney.com/quote/stock/0.000333.html",
    ),
    StockTarget(
        secid="1.600938",
        url="https://wap.eastmoney.com/quote/stock/1.600938.html",
    ),
    StockTarget(
        secid="116.00883",
        url="https://wap.eastmoney.com/quote/stock/116.00883.html",
    ),
)

SLEEP_SECONDS = 11  # Ensure the gap exceeds 10 seconds as requested.


class QuoteParseError(RuntimeError):
    """Raised when the quotedata JSON blob is missing or malformed."""


def build_stock_url(stock_code: str) -> str:
    """Build Eastmoney mobile quote URL from stock code."""
    if len(stock_code) == 5:  # Hong Kong stock
        return f"https://wap.eastmoney.com/quote/stock/116.{stock_code}.html"
    elif len(stock_code) == 6:
        if stock_code.startswith(SHANGHAI_PREFIXES):  # Shanghai stock & ETFs
            return f"https://wap.eastmoney.com/quote/stock/1.{stock_code}.html"
        elif stock_code.startswith(SHENZHEN_PREFIXES):  # Shenzhen stock
            return f"https://wap.eastmoney.com/quote/stock/0.{stock_code}.html"
        else:
            raise ValueError(f"Unsupported A-share stock code: {stock_code}")
    else:
        raise ValueError(f"Invalid stock code format: {stock_code}")


def build_stock_secid(stock_code: str) -> str:
    """Build secid from stock code."""
    if len(stock_code) == 5:  # Hong Kong stock
        return f"116.{stock_code}"
    elif len(stock_code) == 6:
        if stock_code.startswith(SHANGHAI_PREFIXES):  # Shanghai stock & ETFs
            return f"1.{stock_code}"
        elif stock_code.startswith(SHENZHEN_PREFIXES):  # Shenzhen stock
            return f"0.{stock_code}"
        else:
            raise ValueError(f"Unsupported A-share stock code: {stock_code}")
    else:
        raise ValueError(f"Invalid stock code format: {stock_code}")


def fetch_quote_by_code(stock_code: str, verbose: bool = False) -> Dict[str, float | str]:
    """Fetch quote information for a single stock by its code."""
    try:
        url = build_stock_url(stock_code)
        secid = build_stock_secid(stock_code)
        target = StockTarget(secid=secid, url=url)
        
        result = fetch_quote(target, verbose=verbose)
        result["stock_code"] = stock_code  # Add original stock code for convenience
        return result
        
    except ValueError as e:
        raise QuoteParseError(str(e))


def fetch_quotes_by_codes(stock_codes: list[str], verbose: bool = False, sleep_seconds: int = SLEEP_SECONDS) -> list[Dict[str, float | str]]:
    """Fetch quotes for multiple stocks by their codes, respecting delay between requests."""
    results = []
    
    for index, code in enumerate(stock_codes):
        if index > 0:  # Sleep before all requests except the first
            time.sleep(sleep_seconds)
        
        try:
            quote = fetch_quote_by_code(code, verbose=verbose)
            results.append(quote)
        except QuoteParseError as e:
            if verbose:
                print(f"Failed to fetch {code}: {e}")
            # Add a placeholder result with error info
            results.append({
                "stock_code": code,
                "name": "ERROR",
                "latest_price": None,
                "price_change": None,
                "pct_change": None,
                "error": str(e)
            })
    
    return results


def get_display_width(text: str) -> int:
    """Calculate display width of a string, counting CJK characters as 2 and ASCII as 1."""
    width = 0
    for char in text:
        # CJK Unified Ideographs and CJK symbols occupy 2 columns
        if '\u4e00' <= char <= '\u9fff' or '\u3000' <= char <= '\u303f':
            width += 2
        else:
            width += 1
    return width


def pad_string_to_width(text: str, target_width: int, align: str = 'left') -> str:
    """Pad a string to a target display width, accounting for CJK characters.
    
    Args:
        text: The string to pad
        target_width: The desired display width
        align: 'left' or 'right' alignment
    """
    current_width = get_display_width(text)
    padding_needed = max(0, target_width - current_width)
    
    if align == 'left':
        return text + ' ' * padding_needed
    else:  # right
        return ' ' * padding_needed + text


def fetch_quote(target: StockTarget, verbose: bool = False) -> Dict[str, float | str]:
    """Return quote information for a single stock target."""

    response = requests.get(target.url, headers=DEFAULT_HEADERS, timeout=15)
    response.raise_for_status()

    # print("--- Response Content Start ---", file=sys.stderr)  # Debug output
    # print(response.text, file=sys.stderr)  # Debug output
    # print("--- Response Content End ---", file=sys.stderr)  # Debug output

    match = QUOTEDATA_PATTERN.search(response.text)
    if not match:
        raise QuoteParseError(f"Could not locate quotedata in {target.url}")

    quotedata = json.loads(match.group(1))

    price_scale = 10 ** quotedata.get("decimal59", 2)
    pct_scale = 10 ** quotedata.get("decimal152", 2)

    def scaled(value: int, scale: int) -> float:
        return value / scale if scale else float(value)

    latest_price = scaled(quotedata["price"], price_scale)
    price_change = scaled(quotedata.get("zde", 0), price_scale)
    pct_change = scaled(quotedata.get("zdf", 0), pct_scale)
    
    if verbose:
        code = target.secid.split('.')[-1]
        name = quotedata.get('name', target.secid)
        print(f"Fetched quote for {code:>6} {pad_string_to_width(name, 12, 'left')} - Price:{latest_price:>8.2f}, Change:{price_change:>8.2f}, Pct:{pct_change:8.2f}%")

    return {
        "name": quotedata.get("name", target.secid),
        "secid": target.secid,
        "latest_price": latest_price,
        "price_change": price_change,
        "pct_change": pct_change,
        "price_digits": quotedata.get("decimal59", 2),
        "pct_digits": quotedata.get("decimal152", 2),
    }


def fetch_all_quotes(stocks: Iterable[StockTarget], verbose: bool = False) -> list[Dict[str, float | str]]:
    """Fetch quotes sequentially, respecting the required delay."""

    results: list[Dict[str, float | str]] = []

    for index, target in enumerate(stocks):
        if index:
            time.sleep(SLEEP_SECONDS)
        results.append(fetch_quote(target, verbose=verbose))

    return results


def render_table(rows: Iterable[Dict[str, float | str]]) -> str:
    """Create a human-readable text table for the gathered quotes."""

    lines = [
        f"{'Name':<12} {'SecID':<10} {'Price':>12} {'Change':>12} {'Pct':>10}",
        "-" * 60,
    ]

    for row in rows:
        price_format = f"{{:.{int(row['price_digits'])}f}}"
        pct_format = f"{{:+.{int(row['pct_digits'])}f}}%"
        lines.append(
            f"{row['name']:<12} "
            f"{row['secid']:<10} "
            f"{price_format.format(row['latest_price']):>12} "
            f"{price_format.format(row['price_change']):>12} "
            f"{pct_format.format(row['pct_change']) :>10}"
        )

    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch stock quotes from Eastmoney")
    parser.add_argument("codes", nargs="*", help="Stock codes to fetch (e.g., 000333, 600938, 00883)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output")
    parser.add_argument("-s", "--sleep", type=int, default=SLEEP_SECONDS, help=f"Sleep seconds between requests (default: {SLEEP_SECONDS})")
    
    args = parser.parse_args()
    
    try:
        if args.codes:
            # Fetch quotes for specific stock codes
            quotes = fetch_quotes_by_codes(args.codes, verbose=args.verbose, sleep_seconds=args.sleep)
            print(render_table(quotes))
        else:
            # Default behavior: fetch predefined stocks
            quotes = fetch_all_quotes(STOCKS, verbose=args.verbose)
            print(render_table(quotes))
    except requests.RequestException as exc:
        raise SystemExit(f"Network error while fetching quotes: {exc}") from exc
    except QuoteParseError as exc:
        raise SystemExit(str(exc)) from exc


if __name__ == "__main__":
    main()
