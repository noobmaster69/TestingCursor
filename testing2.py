from __future__ import annotations

import argparse
import csv
import re
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import parse_qs, unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?<!\d)(?:\+?1[\s\-.\(\)]*)?\(?\d{3}\)?[\s\-.\(\)]*\d{3}[\s\-.\(\)]*\d{4}(?!\d)")
BLOCKED_HOST_TOKENS = (
    "google.",
    "gstatic.com",
    "bing.com",
    "microsoft.com",
    "duckduckgo.com",
    "wikipedia.org",
    "yelp.com",
)
CONTACT_PATH_HINTS = ("/contact", "/contact-us", "/about", "/about-us")


@dataclass
class Lead:
    query: str = ""
    company_name: str = ""
    website: str = ""
    contact_page: str = ""
    email: str = ""
    phone: str = ""
    source: str = ""
    notes: str = ""


def log(message: str) -> None:
    print(f"[{time.strftime('%H:%M:%S')}] {message}")


def unique(values: Iterable[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = (value or "").strip()
        if not cleaned:
            continue
        key = cleaned.lower()
        if key in seen:
            continue
        out.append(cleaned)
        seen.add(key)
    return out


def normalize_phone(raw: str) -> str:
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    return ""


def build_session(retries_count: int = 2) -> requests.Session:
    session = requests.Session()
    retries = Retry(
        total=retries_count,
        connect=retries_count,
        read=retries_count,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET", "POST"]),
        backoff_factor=0.7,
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"})
    return session


def host_of(url: str) -> str:
    try:
        host = urlparse(url).netloc.lower()
    except Exception:
        return ""
    if host.startswith("www."):
        host = host[4:]
    return host


def looks_like_business_site(url: str) -> bool:
    host = host_of(url)
    if not host:
        return False
    return not any(token in host for token in BLOCKED_HOST_TOKENS)


def parse_ddg_redirect(raw_url: str) -> str:
    parsed = urlparse(raw_url)
    if parsed.path == "/l/" and (
        parsed.netloc.lower().endswith("duckduckgo.com") or parsed.netloc == ""
    ):
        return unquote((parse_qs(parsed.query).get("uddg") or [""])[0]).strip()
    return raw_url


def parse_bing_redirect(raw_url: str) -> str:
    parsed = urlparse(raw_url)
    if "bing.com" in parsed.netloc.lower() and parsed.path.startswith("/ck/a"):
        return unquote((parse_qs(parsed.query).get("u") or [""])[0]).strip()
    return raw_url


def clean_url(raw_url: str, source_base: str = "") -> str:
    url = (raw_url or "").strip()
    if not url:
        return ""
    if source_base and (url.startswith("/") or url.startswith("?")):
        url = urljoin(source_base, url)
    url = parse_ddg_redirect(url)
    url = parse_bing_redirect(url)
    if url.startswith("//"):
        url = f"https:{url}"
    if not url.startswith("http"):
        return ""
    return url


def search_duckduckgo(session: requests.Session, query: str, limit: int, timeout: int) -> list[str]:
    resp = session.post("https://html.duckduckgo.com/html/", data={"q": query}, timeout=timeout)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    urls: list[str] = []
    for a in soup.select("a.result__a, a[href]"):
        cleaned = clean_url(a.get("href", ""), source_base=str(resp.url))
        if not looks_like_business_site(cleaned):
            continue
        urls.append(cleaned)
        if len(unique(urls)) >= limit:
            break
    return unique(urls)[:limit]


def search_bing(session: requests.Session, query: str, limit: int, timeout: int) -> list[str]:
    resp = session.get("https://www.bing.com/search", params={"q": query}, timeout=timeout)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    urls: list[str] = []
    for a in soup.select("li.b_algo h2 a, a[href]"):
        cleaned = clean_url(a.get("href", ""), source_base=str(resp.url))
        if not looks_like_business_site(cleaned):
            continue
        urls.append(cleaned)
        if len(unique(urls)) >= limit:
            break
    return unique(urls)[:limit]


def search_yahoo(session: requests.Session, query: str, limit: int, timeout: int) -> list[str]:
    resp = session.get("https://search.yahoo.com/search", params={"p": query}, timeout=timeout)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    urls: list[str] = []
    for a in soup.select("div#web a[href], h3 a[href], a[href]"):
        cleaned = clean_url(a.get("href", ""), source_base=str(resp.url))
        if not looks_like_business_site(cleaned):
            continue
        urls.append(cleaned)
        if len(unique(urls)) >= limit:
            break
    return unique(urls)[:limit]


def extract_company_name(soup: BeautifulSoup, fallback_url: str) -> str:
    if soup.title and soup.title.text.strip():
        title = soup.title.text.strip()
        title = re.sub(r"\s*[-|]\s*.*$", "", title).strip()
        if title:
            return title
    host = host_of(fallback_url)
    return host.split(".")[0].capitalize() if host else ""


def extract_contacts_from_html(html_text: str) -> tuple[str, str]:
    emails = unique([m.lower().strip(".,;:()[]{}<>") for m in EMAIL_RE.findall(html_text or "")])
    phones = unique([normalize_phone(m) for m in PHONE_RE.findall(html_text or "") if normalize_phone(m)])
    return (emails[0] if emails else "", phones[0] if phones else "")


def find_contact_page(base_url: str, soup: BeautifulSoup) -> str:
    for a in soup.select("a[href]"):
        href = (a.get("href") or "").strip()
        text = (a.get_text(" ", strip=True) or "").lower()
        absolute = urljoin(base_url, href)
        low = absolute.lower()
        if "contact" in text or "about" in text:
            return absolute
        if any(token in low for token in CONTACT_PATH_HINTS):
            return absolute
    return ""


def fetch_page(session: requests.Session, url: str, timeout: int) -> tuple[str, BeautifulSoup]:
    resp = session.get(url, timeout=timeout)
    resp.raise_for_status()
    html_text = resp.text or ""
    return html_text, BeautifulSoup(html_text, "html.parser")


def build_lead_from_site(
    session: requests.Session,
    query: str,
    site_url: str,
    timeout: int,
) -> Lead | None:
    lead = Lead(query=query, website=site_url, source="requests")
    try:
        main_html, main_soup = fetch_page(session, site_url, timeout)
    except Exception as exc:
        log(f"Failed {site_url}: {exc}")
        return None

    lead.company_name = extract_company_name(main_soup, site_url)
    email, phone = extract_contacts_from_html(main_html)
    lead.email = email
    lead.phone = phone

    contact_url = find_contact_page(site_url, main_soup)
    if contact_url:
        lead.contact_page = contact_url
        try:
            contact_html, _ = fetch_page(session, contact_url, timeout)
            c_email, c_phone = extract_contacts_from_html(contact_html)
            if not lead.email:
                lead.email = c_email
            if not lead.phone:
                lead.phone = c_phone
        except Exception as exc:
            lead.notes = f"contact page failed: {exc}"

    return lead


def search_web(session: requests.Session, query: str, per_query: int, timeout: int) -> list[str]:
    results = []
    for engine_name, fn in (
        ("DuckDuckGo", search_duckduckgo),
        ("Bing", search_bing),
        ("Yahoo", search_yahoo),
    ):
        try:
            found = fn(session, query, per_query, timeout)
            log(f"{engine_name}: found {len(found)} site(s)")
            results.extend(found)
        except Exception as exc:
            log(f"{engine_name} failed: {exc}")
    deduped = unique(results)
    if deduped:
        return deduped[:per_query]

    # Query fallback: if no sites returned, try a "business website" focused variant.
    fallback_query = f"{query} business website contact"
    if fallback_query != query:
        log("Primary search returned 0 sites, trying fallback query...")
        for engine_name, fn in (("DuckDuckGo", search_duckduckgo), ("Bing", search_bing)):
            try:
                found = fn(session, fallback_query, per_query, timeout)
                log(f"{engine_name} fallback: found {len(found)} site(s)")
                results.extend(found)
            except Exception as exc:
                log(f"{engine_name} fallback failed: {exc}")
    return unique(results)[:per_query]


def scrape(
    queries: list[str],
    per_query: int,
    timeout: int,
    sleep_between: float,
) -> list[Lead]:
    session = build_session()
    all_leads: list[Lead] = []
    seen_sites: set[str] = set()

    for query in queries:
        log(f"Searching web for: {query}")
        candidate_sites = search_web(session, query, per_query, timeout)
        for site in candidate_sites:
            key = site.lower()
            if key in seen_sites:
                continue
            seen_sites.add(key)
            lead = build_lead_from_site(session, query, site, timeout)
            if lead is None:
                continue
            all_leads.append(lead)
            log(
                f"Lead: {lead.company_name or host_of(lead.website)} "
                f"email={bool(lead.email)} phone={bool(lead.phone)}"
            )
            if sleep_between > 0:
                time.sleep(sleep_between)
    return all_leads


def export_csv(leads: list[Lead], output_file: Path) -> Path:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with output_file.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["query", "company_name", "website", "contact_page", "email", "phone", "source", "notes"],
        )
        writer.writeheader()
        for lead in leads:
            writer.writerow(asdict(lead))
    return output_file


def parse_queries(raw: str) -> list[str]:
    return unique([part.strip() for part in (raw or "").split(";") if part.strip()])


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="General business lead scraper (requests + BeautifulSoup).")
    parser.add_argument("--queries", default="construction companies near Rockville MD", help="Semicolon-separated search queries.")
    parser.add_argument("--per-query", type=int, default=10, help="Max sites to process per query.")
    parser.add_argument("--timeout", type=int, default=20, help="HTTP/browser timeout in seconds.")
    parser.add_argument("--sleep", type=float, default=0.2, help="Delay between site requests in seconds.")
    parser.add_argument("--output", default="testing2_results.csv", help="Output CSV path.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    queries = parse_queries(args.queries)
    if not queries:
        raise SystemExit("No queries provided.")

    leads = scrape(
        queries=queries,
        per_query=max(1, args.per_query),
        timeout=max(5, args.timeout),
        sleep_between=max(0.0, args.sleep),
    )
    output_path = export_csv(leads, Path(args.output))
    log(f"Saved {len(leads)} lead(s) to {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

