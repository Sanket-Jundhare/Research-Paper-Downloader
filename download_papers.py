"""
Open-access research paper downloader.

Reads `Papers.xlsx`, looks up legal open-access PDF locations
via Crossref / OpenAlex / Unpaywall (and Europe PMC when applicable),
downloads validated PDFs into publisher folders, and logs outcomes.

Does not bypass paywalls, authentication, CAPTCHAs, DRM, or access controls.
Does not modify the Excel input file or PDF contents.
"""

from __future__ import annotations

import argparse
import csv
import difflib
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Optional
from urllib.parse import quote, urlparse

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import requests

# ---------------------------------------------------------------------------
# Paths and constants
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
EXCEL_NAME = "Papers.xlsx"
DOWNLOADED_DIR = BASE_DIR / "Downloaded_Papers"
LOGS_DIR = BASE_DIR / "Logs"
TEMP_DIR = BASE_DIR / "temp"

DOWNLOAD_LOG = LOGS_DIR / "download_log.csv"
FAILED_LOG = LOGS_DIR / "failed_downloads.csv"
ACCESS_LOG = LOGS_DIR / "access_required.csv"

PUBLISHER_DIRS = (
    "IEEE",
    "Elsevier",
    "Springer",
    "PubMed",
    "Wiley",
    "MDPI",
    "ACM",
    "Scopus",
    "Scholar",
    "Other",
)

DOWNLOAD_LOG_FIELDS = [
    "paper_number",
    "title",
    "doi",
    "source",
    "publisher",
    "year",
    "destination_folder",
    "filename",
    "pdf_url",
    "status",
    "error",
    "timestamp",
]

FAILED_LOG_FIELDS = [
    "paper_number",
    "title",
    "doi",
    "source",
    "reason",
    "error",
]

ACCESS_LOG_FIELDS = [
    "paper_number",
    "title",
    "doi",
    "publisher",
    "source",
    "publisher_url",
    "reason",
]

SUCCESS_STATUSES = frozenset({"downloaded", "already_exists"})
RETRYABLE_STATUSES = frozenset(
    {"failed", "access_required", "no_open_access", "invalid_doi"}
)

DOI_RE = re.compile(r"^10\.\d{4,9}/[-._;()/:A-Z0-9]+$", re.IGNORECASE)
INVALID_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
MULTI_UNDERSCORE_RE = re.compile(r"_+")

USER_AGENT_TEMPLATE = (
    "PaperDownloader/1.0 (systematic-review; mailto:{email}; "
    "+https://github.com/local/paper_downloader)"
)

REQUEST_TIMEOUT = 25
MAX_RETRIES = 3
RETRY_BACKOFF_BASE = 2  # 2s, 4s, 8s
INTER_REQUEST_DELAY = 0.25
MIN_PDF_BYTES = 1_000
MAX_FILENAME_LEN = 180

DOI_PREFIX_MAP = (
    ("10.1109/", "IEEE"),
    ("10.1016/", "Elsevier"),
    ("10.1007/", "Springer"),
    ("10.1038/", "Springer"),
    ("10.1111/", "Wiley"),
    ("10.1002/", "Wiley"),
    ("10.3390/", "MDPI"),
    ("10.1145/", "ACM"),
)

SOURCE_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("IEEE", ("ieee", "ieee xplore", "ieeexplore")),
    ("Elsevier", ("elsevier", "sciencedirect")),
    ("Springer", ("springer", "springer nature")),
    ("PubMed", ("pubmed", "pmc", "pubmed central", "europe pmc")),
    ("Wiley", ("wiley", "onlinelibrary.wiley")),
    ("MDPI", ("mdpi",)),
    ("ACM", ("acm", "dl.acm")),
    ("Scopus", ("scopus",)),
    ("Scholar", ("google scholar", "scholar", "gs extract", "scholar extract")),
]

HOST_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("IEEE", ("ieeexplore.ieee.org", "ieee.org")),
    ("Elsevier", ("sciencedirect.com", "elsevier.com")),
    ("Springer", ("link.springer.com", "springer.com", "nature.com")),
    ("PubMed", ("pubmed.ncbi.nlm.nih.gov", "pmc.ncbi.nlm.nih.gov", "ncbi.nlm.nih.gov")),
    ("Wiley", ("onlinelibrary.wiley.com", "wiley.com")),
    ("MDPI", ("mdpi.com",)),
    ("ACM", ("dl.acm.org", "acm.org")),
]

OA_HOST_PRIORITY = {
    "publisher": 100,
    "repository": 60,
    "other": 20,
}

logger = logging.getLogger("paper_downloader")


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------


@dataclass
class ColumnMap:
    """Detected Excel column names for key fields."""

    title: Optional[str]
    doi: Optional[str]
    source: Optional[str]
    publisher: Optional[str]
    authors: Optional[str]
    year: Optional[str]
    paper_number: Optional[str]
    doi_url: Optional[str] = None
    pmid: Optional[str] = None


@dataclass
class OACandidate:
    """A candidate open-access PDF location."""

    url: str
    source_api: str
    host_type: str = "other"
    version: str = ""
    license: str = ""
    score: int = 0
    is_direct_pdf: bool = False


@dataclass
class PaperRecord:
    """One row from the Excel workbook."""

    paper_number: int
    title: str
    doi_raw: str
    doi: str
    source: str
    publisher: str
    authors: str
    year: str
    row_index: int
    pmid: str = ""


@dataclass
class ProcessResult:
    """Outcome of processing a single paper."""

    status: str
    destination_folder: str = ""
    filename: str = ""
    pdf_url: str = ""
    publisher: str = ""
    error: str = ""
    oa_source: str = ""


@dataclass
class TitleResolution:
    """Best-effort identity resolved from title metadata (for Scholar rows, etc.)."""

    doi: str = ""
    pdf_url: str = ""
    source_api: str = ""
    publisher: str = ""
    score: float = 0.0
    core_score: float = 0.0
    paper_id: str = ""
    matched_title: str = ""


@dataclass
class Summary:
    """Aggregate counters for the run."""

    total: int = 0
    downloaded: int = 0
    already_exists: int = 0
    no_oa: int = 0
    failed: int = 0
    invalid_doi: int = 0
    access_required: int = 0
    skipped: int = 0
    not_found: int = 0


# ---------------------------------------------------------------------------
# Setup helpers
# ---------------------------------------------------------------------------


def configure_workspace(base: Path) -> None:
    """Point all I/O at an isolated workspace (used by the web app)."""
    global BASE_DIR, DOWNLOADED_DIR, LOGS_DIR, TEMP_DIR
    global DOWNLOAD_LOG, FAILED_LOG, ACCESS_LOG
    BASE_DIR = Path(base).resolve()
    DOWNLOADED_DIR = BASE_DIR / "Downloaded_Papers"
    LOGS_DIR = BASE_DIR / "Logs"
    TEMP_DIR = BASE_DIR / "temp"
    DOWNLOAD_LOG = LOGS_DIR / "download_log.csv"
    FAILED_LOG = LOGS_DIR / "failed_downloads.csv"
    ACCESS_LOG = LOGS_DIR / "access_required.csv"


def ensure_directories() -> None:
    """Create project folders if they do not already exist."""
    DOWNLOADED_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    for name in PUBLISHER_DIRS:
        (DOWNLOADED_DIR / name).mkdir(parents=True, exist_ok=True)


def _write_header_if_needed(path: Path, fields: list[str]) -> None:
    """Create a CSV with headers, or rewrite if only an outdated header exists."""
    if not path.exists():
        with path.open("w", newline="", encoding="utf-8") as handle:
            csv.DictWriter(handle, fieldnames=fields).writeheader()
        return

    with path.open("r", newline="", encoding="utf-8") as handle:
        rows = list(csv.reader(handle))

    if not rows:
        with path.open("w", newline="", encoding="utf-8") as handle:
            csv.DictWriter(handle, fieldnames=fields).writeheader()
        return

    header = rows[0]
    data_rows = rows[1:]
    if header == fields:
        return

    # Only rewrite when there is no data yet (safe migration of empty logs).
    if not data_rows or all(not any(cell.strip() for cell in row) for row in data_rows):
        with path.open("w", newline="", encoding="utf-8") as handle:
            csv.DictWriter(handle, fieldnames=fields).writeheader()
        logger.info("Updated empty log headers: %s", path.name)
    else:
        logger.warning(
            "Log %s has existing data with different headers; appending with new schema.",
            path.name,
        )


def ensure_log_files() -> None:
    """Ensure log CSVs exist with the expected headers."""
    _write_header_if_needed(DOWNLOAD_LOG, DOWNLOAD_LOG_FIELDS)
    _write_header_if_needed(FAILED_LOG, FAILED_LOG_FIELDS)
    _write_header_if_needed(ACCESS_LOG, ACCESS_LOG_FIELDS)


def utc_timestamp() -> str:
    """Return an ISO-8601 UTC timestamp."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ---------------------------------------------------------------------------
# Excel loading and column detection
# ---------------------------------------------------------------------------


def _normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def _score_column(name: str, patterns: Iterable[str]) -> int:
    norm = _normalize_header(name)
    if not norm or norm.startswith("unnamed"):
        return 0
    best = 0
    for pattern in patterns:
        if norm == pattern:
            best = max(best, 100)
        elif norm.startswith(pattern + " ") or norm.endswith(" " + pattern):
            best = max(best, 90)
        elif pattern in norm:
            best = max(best, 70)
    return best


def detect_columns(columns: Iterable[Any], *, verbose: bool = True) -> ColumnMap:
    """
    Intelligently map workbook columns to title/DOI/source/etc.

    Uses fuzzy header matching so slight naming differences still work.
    """
    cols = [str(c) for c in columns]
    if verbose:
        print("Detected Excel columns:")
        for col in cols:
            print(f"  - {col}")

    def pick(patterns: tuple[str, ...], exclude: Optional[set[str]] = None) -> Optional[str]:
        exclude = exclude or set()
        scored = []
        for col in cols:
            if col in exclude:
                continue
            score = _score_column(col, patterns)
            if score:
                scored.append((score, col))
        if not scored:
            return None
        scored.sort(key=lambda item: (-item[0], item[1]))
        return scored[0][1]

    title = pick(("title", "article title", "document title", "paper title", "gs title"))
    doi = pick(("doi", "digital object identifier", "doi link", "doi url"))
    doi_url = pick(("doi url", "doi link", "article url", "url"), exclude={doi} if doi else set())
    pmid = pick(("pmid", "pubmed id"))
    source = pick(("source", "database", "source title", "index", "collection"))
    publisher = pick(
        ("publisher", "publisher name", "publ", "venue", "journal", "publication")
    )
    authors = pick(
        (
            "authors",
            "author",
            "author names",
            "author(s)",
            "first author",
            "first authors",
            "gs authors",
        )
    )
    year = pick(
        ("year", "publication year", "pub year", "cover date", "gs year", "pubdate")
    )
    paper_number = pick(("paper_number", "paper number", "record number", "row number", "serial number", "sr no", "serial"))

    # First unnamed integer-like index column often holds the paper number.
    if paper_number is None:
        for col in cols:
            if _normalize_header(col).startswith("unnamed"):
                paper_number = col
                break

    mapping = ColumnMap(
        title=title,
        doi=doi,
        source=source,
        publisher=publisher,
        authors=authors,
        year=year,
        paper_number=paper_number,
        doi_url=doi_url,
        pmid=pmid,
    )

    if verbose:
        print("\nMapped columns:")
        print(f"  Title:        {mapping.title}")
        print(f"  DOI:          {mapping.doi}")
        print(f"  DOI URL:      {mapping.doi_url}")
        print(f"  PMID:         {mapping.pmid}")
        print(f"  Source:       {mapping.source}")
        print(f"  Publisher:    {mapping.publisher}")
        print(f"  Authors:      {mapping.authors}")
        print(f"  Year:         {mapping.year}")
        print(f"  Paper number: {mapping.paper_number}")
        print()
    return mapping


def load_excel(path: Path, *, verbose: bool = True) -> tuple[pd.DataFrame, ColumnMap]:
    """Load an Excel (.xlsx/.xlsm/.xls) or CSV workbook without modifying it."""
    if not path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {path}\n"
            f"Place '{EXCEL_NAME}' in {BASE_DIR}"
        )
    suffix = path.suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(path)
    elif suffix == ".xls":
        try:
            df = pd.read_excel(path, engine="xlrd")
        except ImportError as exc:
            raise ValueError(
                "Legacy .xls files require 'xlrd'. Please save the file as .xlsx or .csv and upload again."
            ) from exc
    else:
        df = pd.read_excel(path, engine="openpyxl")

    # Drop completely empty rows so trailing blank lines in Excel don't count as papers.
    df = df.dropna(how="all").reset_index(drop=True)
    columns = detect_columns(df.columns, verbose=verbose)
    if not columns.doi and not columns.doi_url and not columns.title:
        raise ValueError("Could not detect Title or DOI columns in the spreadsheet.")
    return df, columns


def _cell_str(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def clean_doi(raw: str) -> str:
    """Normalize a DOI string by stripping common prefixes and whitespace."""
    doi = (raw or "").strip()
    if not doi:
        return ""
    doi = doi.replace("\\", "/")
    # Take first token if multiple values somehow appear.
    doi = doi.split()[0]
    lower = doi.lower()
    for prefix in (
        "https://doi.org/",
        "http://doi.org/",
        "https://dx.doi.org/",
        "http://dx.doi.org/",
        "doi:",
    ):
        if lower.startswith(prefix):
            doi = doi[len(prefix) :]
            lower = doi.lower()
            break
    doi = doi.strip().strip("<>").strip(".,;")
    return doi


def is_valid_doi(doi: str) -> bool:
    """Return True if the string looks like a DOI."""
    if not doi:
        return False
    if not DOI_RE.match(doi):
        # Allow slightly broader but still DOI-like forms.
        if not doi.lower().startswith("10."):
            return False
        if "/" not in doi:
            return False
    return True


def dataframe_to_papers(df: pd.DataFrame, columns: ColumnMap) -> list[PaperRecord]:
    """Convert the workbook rows into PaperRecord objects."""
    papers: list[PaperRecord] = []
    seen_numbers: set[int] = set()
    for idx, row in df.iterrows():
        title = _cell_str(row[columns.title]) if columns.title else ""
        doi_raw = _cell_str(row[columns.doi]) if columns.doi else ""
        if not doi_raw and columns.doi_url:
            doi_url_val = _cell_str(row[columns.doi_url])
            if "10." in doi_url_val:
                doi_raw = doi_url_val
        source = _cell_str(row[columns.source]) if columns.source else ""
        publisher = _cell_str(row[columns.publisher]) if columns.publisher else ""
        authors = _cell_str(row[columns.authors]) if columns.authors else ""
        year = _cell_str(row[columns.year]) if columns.year else ""
        if year.endswith(".0") and year[:-2].isdigit():
            year = year[:-2]
        pmid = _cell_str(row[columns.pmid]) if columns.pmid else ""
        if pmid.endswith(".0") and pmid[:-2].isdigit():
            pmid = pmid[:-2]

        if not title and not doi_raw and not pmid:
            continue

        if columns.paper_number:
            num_raw = _cell_str(row[columns.paper_number])
            try:
                paper_number = int(float(num_raw)) if num_raw else len(papers) + 1
            except ValueError:
                paper_number = len(papers) + 1
        else:
            paper_number = len(papers) + 1

        if paper_number in seen_numbers:
            paper_number = len(papers) + 1
        while paper_number in seen_numbers:
            paper_number += 1
        seen_numbers.add(paper_number)

        papers.append(
            PaperRecord(
                paper_number=paper_number,
                title=title,
                doi_raw=doi_raw,
                doi=clean_doi(doi_raw),
                source=source,
                publisher=publisher,
                authors=authors,
                year=year,
                row_index=int(idx),
                pmid=pmid,
            )
        )
    return papers


# ---------------------------------------------------------------------------
# Classification and filenames
# ---------------------------------------------------------------------------


def _match_rules(text: str, rules: list[tuple[str, tuple[str, ...]]]) -> Optional[str]:
    hay = text.lower()
    for label, needles in rules:
        if any(needle in hay for needle in needles):
            return label
    return None


def classify_from_doi_prefix(doi: str) -> Optional[str]:
    """Map well-known DOI prefixes to publisher folders."""
    lower = doi.lower()
    for prefix, label in DOI_PREFIX_MAP:
        if lower.startswith(prefix):
            return label
    return None


def classify_from_host(url: str) -> Optional[str]:
    """Classify a publisher folder from a URL hostname."""
    if not url:
        return None
    host = urlparse(url).netloc.lower()
    for label, hosts in HOST_RULES:
        if any(h in host for h in hosts):
            return label
    return None


def classify_publisher(
    source: str = "",
    publisher: str = "",
    doi: str = "",
    url: str = "",
    meta_publisher: str = "",
) -> str:
    """
    Determine the destination publisher folder.

    Priority: explicit publisher text → source text → DOI prefix → URL host →
    metadata publisher → Scopus/Scholar fallback → Other.
    """
    for text in (publisher, meta_publisher, source):
        matched = _match_rules(text, SOURCE_RULES)
        # Keep index-only labels for later; prefer real publishers first.
        if matched and matched not in {"Scopus", "Scholar"}:
            return matched

    prefix = classify_from_doi_prefix(doi)
    if prefix:
        return prefix

    host = classify_from_host(url)
    if host:
        return host

    # Index / extract sources used when no publisher signal exists.
    if _match_rules(source, SOURCE_RULES) == "Scopus" or source.lower() == "scopus":
        return "Scopus"
    if _match_rules(source, SOURCE_RULES) == "Scholar" or "scholar" in source.lower():
        return "Scholar"

    matched_source = _match_rules(source, SOURCE_RULES)
    if matched_source:
        return matched_source

    return "Other"


def create_safe_filename(paper_number: int, title: str) -> str:
    """Build a Windows-safe PDF filename from paper number and title."""
    prefix = f"{paper_number:03d}_"
    suffix = ".pdf"
    budget = MAX_FILENAME_LEN - len(prefix) - len(suffix)

    cleaned = (title or "untitled").strip()
    cleaned = cleaned.replace("“", "").replace("”", "").replace("„", "")
    cleaned = cleaned.replace("‘", "'").replace("’", "'")
    cleaned = INVALID_FILENAME_RE.sub("", cleaned)
    cleaned = cleaned.replace(" ", "_")
    cleaned = MULTI_UNDERSCORE_RE.sub("_", cleaned)
    cleaned = cleaned.strip("._")
    if not cleaned:
        cleaned = "untitled"
    if len(cleaned) > budget:
        cleaned = cleaned[:budget].rstrip("._")

    return f"{prefix}{cleaned}{suffix}"


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------


def build_session(email: str) -> requests.Session:
    """Create a requests session with a polite User-Agent."""
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": USER_AGENT_TEMPLATE.format(email=email),
            "Accept": "application/json,application/pdf,*/*",
        }
    )
    return session


def request_with_retries(
    session: requests.Session,
    method: str,
    url: str,
    *,
    params: Optional[dict[str, Any]] = None,
    stream: bool = False,
    accept: Optional[str] = None,
    headers: Optional[dict[str, str]] = None,
) -> requests.Response:
    """HTTP request with exponential backoff retries."""
    req_headers: dict[str, str] = dict(headers or {})
    if accept and "Accept" not in req_headers:
        req_headers["Accept"] = accept

    last_error: Optional[Exception] = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.request(
                method,
                url,
                params=params,
                headers=req_headers or None,
                timeout=REQUEST_TIMEOUT,
                stream=stream,
                allow_redirects=True,
            )
            # Retry on transient server / rate-limit responses.
            if response.status_code in {429, 500, 502, 503, 504}:
                response.close()
                raise requests.HTTPError(
                    f"Transient HTTP {response.status_code} for {url}",
                    response=response,
                )
            return response
        except (requests.RequestException, requests.HTTPError) as exc:
            last_error = exc
            if attempt >= MAX_RETRIES:
                break
            wait = RETRY_BACKOFF_BASE ** attempt
            logger.warning(
                "Request failed (attempt %s/%s): %s; sleeping %ss",
                attempt,
                MAX_RETRIES,
                exc,
                wait,
            )
            time.sleep(wait)
    assert last_error is not None
    raise last_error


# ---------------------------------------------------------------------------
# Open-access discovery
# ---------------------------------------------------------------------------


def _score_oa_candidate(candidate: OACandidate) -> int:
    score = OA_HOST_PRIORITY.get(candidate.host_type.lower(), 10)
    url_l = candidate.url.lower()
    host = urlparse(candidate.url).netloc.lower()

    if candidate.is_direct_pdf or "pdf" in url_l or url_l.endswith(".pdf"):
        score += 45
    if candidate.host_type.lower() == "publisher":
        score += 25
    if "repository" in candidate.host_type.lower():
        if any(token in host for token in ("edu", "ac.uk", "repository", "handle.net", "eprints")):
            score += 25  # institutional repository (rarely WAF-blocked)
        if any(token in host for token in ("arxiv.org", "biorxiv.org", "medrxiv.org", "zenodo")):
            score += 22  # author / preprint repository
    if "europepmc.org" in host:
        score += 15
    # pmc.ncbi.nlm.nih.gov frequently serves JS bot-challenges to non-browser clients;
    # keep it as a lower-priority fallback behind direct publisher/institutional URLs.
    if "pmc.ncbi.nlm.nih.gov" in host or "/pmc/articles/" in url_l:
        score += 5

    version = (candidate.version or "").lower()
    if "published" in version:
        score += 10
    elif "accepted" in version:
        score += 5

    candidate.score = score
    return score


def _add_candidate(
    bag: dict[str, OACandidate],
    url: Optional[str],
    *,
    source_api: str,
    host_type: str = "other",
    version: str = "",
    license: str = "",
    is_direct_pdf: bool = False,
) -> None:
    if not url or not isinstance(url, str):
        return
    url = url.strip()
    if not url.startswith("http"):
        return
    key = url.split("?")[0].rstrip("/").lower()
    candidate = OACandidate(
        url=url,
        source_api=source_api,
        host_type=host_type or "other",
        version=version or "",
        license=license or "",
        is_direct_pdf=is_direct_pdf,
    )
    _score_oa_candidate(candidate)
    existing = bag.get(key)
    if existing is None or candidate.score > existing.score:
        bag[key] = candidate


def fetch_openalex(
    session: requests.Session, doi: str, email: str
) -> tuple[dict[str, Any], list[OACandidate]]:
    """Query OpenAlex for work metadata and OA PDF candidates."""
    print("Checking OpenAlex...")
    url = f"https://api.openalex.org/works/doi:{quote(doi, safe='')}"
    response = request_with_retries(
        session,
        "GET",
        url,
        params={"mailto": email},
        accept="application/json",
    )
    time.sleep(INTER_REQUEST_DELAY)
    if response.status_code == 404:
        return {}, []
    response.raise_for_status()
    data = response.json()
    bag: dict[str, OACandidate] = {}

    def absorb_location(loc: Any, label: str) -> None:
        if not isinstance(loc, dict):
            return
        source_obj = loc.get("source")
        host_type = loc.get("host_type") or "other"
        # OpenAlex uses source.type like "journal" / "repository".
        if isinstance(source_obj, dict):
            stype = str(source_obj.get("type") or "").lower()
            if stype == "repository":
                host_type = "repository"
            elif stype in {"journal", "ebook platform", "book series"}:
                host_type = "publisher"
        elif isinstance(source_obj, str) and source_obj:
            host_type = source_obj
        pdf = loc.get("pdf_url")
        landing = loc.get("landing_page_url")
        _add_candidate(
            bag,
            pdf,
            source_api=label,
            host_type=str(host_type),
            version=str(loc.get("version") or ""),
            license=str(loc.get("license") or ""),
            is_direct_pdf=True,
        )
        # Only keep landing pages that look like direct PDF endpoints.
        if isinstance(landing, str) and (
            ".pdf" in landing.lower() or "/pdf" in landing.lower()
        ):
            _add_candidate(
                bag,
                landing,
                source_api=label,
                host_type=str(host_type),
                version=str(loc.get("version") or ""),
                license=str(loc.get("license") or ""),
                is_direct_pdf=True,
            )

    absorb_location(data.get("best_oa_location"), "OpenAlex")
    absorb_location(data.get("primary_location"), "OpenAlex")
    for loc in data.get("locations") or []:
        absorb_location(loc, "OpenAlex")

    return data, list(bag.values())


def fetch_unpaywall(
    session: requests.Session, doi: str, email: str
) -> tuple[dict[str, Any], list[OACandidate]]:
    """Query Unpaywall for legal OA PDF locations."""
    print("Checking Unpaywall...")
    url = f"https://api.unpaywall.org/v2/{quote(doi, safe='')}"
    response = request_with_retries(
        session,
        "GET",
        url,
        params={"email": email},
        accept="application/json",
    )
    time.sleep(INTER_REQUEST_DELAY)
    if response.status_code == 404:
        return {}, []
    if response.status_code == 422:
        logger.warning(
            "Unpaywall rejected email %r (HTTP 422); use a valid non-example email.",
            email,
        )
        return {}, []
    response.raise_for_status()
    data = response.json()
    bag: dict[str, OACandidate] = {}

    def absorb(loc: Optional[dict[str, Any]]) -> None:
        if not loc:
            return
        _add_candidate(
            bag,
            loc.get("url_for_pdf"),
            source_api="Unpaywall",
            host_type=str(loc.get("host_type") or "other"),
            version=str(loc.get("version") or ""),
            license=str(loc.get("license") or ""),
            is_direct_pdf=True,
        )
        # Prefer PDF URL; landing URL only if it looks like a PDF link.
        landing = loc.get("url")
        if landing and (".pdf" in landing.lower() or "/pdf" in landing.lower()):
            _add_candidate(
                bag,
                landing,
                source_api="Unpaywall",
                host_type=str(loc.get("host_type") or "other"),
                version=str(loc.get("version") or ""),
                license=str(loc.get("license") or ""),
                is_direct_pdf=True,
            )

    absorb(data.get("best_oa_location"))
    for loc in data.get("oa_locations") or []:
        absorb(loc)
    return data, list(bag.values())


def fetch_crossref(
    session: requests.Session, doi: str, email: str
) -> tuple[dict[str, Any], list[OACandidate]]:
    """Query Crossref for metadata and any declared PDF links."""
    print("Checking Crossref...")
    url = f"https://api.crossref.org/works/{quote(doi, safe='')}"
    response = request_with_retries(
        session,
        "GET",
        url,
        params={"mailto": email},
        accept="application/json",
    )
    time.sleep(INTER_REQUEST_DELAY)
    if response.status_code == 404:
        return {}, []
    response.raise_for_status()
    message = response.json().get("message") or {}
    bag: dict[str, OACandidate] = {}
    for link in message.get("link") or []:
        content_type = str(link.get("content-type") or "").lower()
        href = str(link.get("URL") or "").strip()
        if not href:
            continue
        # Skip XML / HTML TDM endpoints (e.g. Elsevier API XML links).
        if "xml" in content_type or "html" in content_type or "plain" in content_type:
            continue
        if "pdf" in content_type or ".pdf" in href.lower() or "/pdf" in href.lower():
            _add_candidate(
                bag,
                href,
                source_api="Crossref",
                host_type="publisher",
                is_direct_pdf=True,
            )
    return message, list(bag.values())


def fetch_europe_pmc(
    session: requests.Session, doi: str = "", pmid: str = ""
) -> tuple[dict[str, Any], list[OACandidate]]:
    """Query Europe PMC for a PubMed Central PDF when available."""
    query = f'DOI:"{doi}"' if doi else (f"EXT_ID:{pmid} AND SRC:MED" if pmid else "")
    if not query:
        return {}, []
    print("Checking Europe PMC...")
    url = "https://www.ebi.ac.uk/europepmc/webservices/rest/search"
    response = request_with_retries(
        session,
        "GET",
        url,
        params={
            "query": query,
            "format": "json",
            "resultType": "core",
            "pageSize": 1,
        },
        accept="application/json",
    )
    time.sleep(INTER_REQUEST_DELAY)
    response.raise_for_status()
    data = response.json()
    results = ((data.get("resultList") or {}).get("result")) or []
    bag: dict[str, OACandidate] = {}
    if not results:
        return data, []
    hit = results[0]
    pmcid = hit.get("pmcid") or ""
    # Some records expose a direct PDF URL in fullTextUrlList.
    for field_name in ("fullTextUrlList",):
        block = hit.get(field_name) or {}
        urls = block.get("fullTextUrl") or []
        if isinstance(urls, dict):
            urls = [urls]
        for item in urls:
            if not isinstance(item, dict):
                continue
            if str(item.get("documentStyle") or "").lower() == "pdf":
                _add_candidate(
                    bag,
                    item.get("url"),
                    source_api="EuropePMC",
                    host_type="repository",
                    is_direct_pdf=True,
                )
    if pmcid:
        pdf_url = f"https://pmc.ncbi.nlm.nih.gov/articles/{pmcid}/pdf/"
        _add_candidate(
            bag,
            pdf_url,
            source_api="EuropePMC",
            host_type="repository",
            version="publishedVersion",
            is_direct_pdf=True,
        )
    return hit, list(bag.values())


def _is_known_oa_host(url: str) -> bool:
    """Return True for hosts that commonly serve legal open-access PDFs."""
    host = urlparse(url).netloc.lower()
    known = (
        "pmc.ncbi.nlm.nih.gov",
        "europepmc.org",
        "arxiv.org",
        "biorxiv.org",
        "medrxiv.org",
        "zenodo.org",
        "mdpi.com",
        "plos.org",
        "frontiersin.org",
        "hindawi.com",
        "biomedcentral.com",
        "nature.com",  # some Nature OA; still validated as PDF on download
        "nih.gov",
    )
    return any(token in host for token in known)


def fetch_semantic_scholar_by_doi(
    session: requests.Session, doi: str
) -> list[OACandidate]:
    """Query Semantic Scholar by DOI for a legal openAccessPdf URL."""
    url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{quote(doi, safe='')}"
    try:
        response = session.get(
            url,
            params={"fields": "isOpenAccess,openAccessPdf"},
            headers={"Accept": "application/json"},
            timeout=REQUEST_TIMEOUT,
        )
        time.sleep(INTER_REQUEST_DELAY)
        if response.status_code in {404, 429, 403}:
            return []
        response.raise_for_status()
        data = response.json()
    except Exception as exc:  # noqa: BLE001
        logger.debug("Semantic Scholar DOI lookup failed for %s: %s", doi, exc)
        return []

    oa = data.get("openAccessPdf") or {}
    if isinstance(oa, dict) and oa.get("url"):
        bag: dict[str, OACandidate] = {}
        _add_candidate(
            bag,
            oa.get("url"),
            source_api="SemanticScholar",
            host_type="repository",
            is_direct_pdf=True,
        )
        return list(bag.values())
    return []


def get_open_access_links(
    session: requests.Session, doi: str, email: str, *, pmid: str = ""
) -> tuple[list[OACandidate], dict[str, Any]]:
    """
    Collect and rank legitimate open-access PDF candidates.

    Preference order is encoded via scoring:
    publisher OA → institutional repo → author repo → Europe PMC → PMC → other.
    """
    meta: dict[str, Any] = {
        "publisher": "",
        "publisher_url": f"https://doi.org/{doi}" if doi else "",
        "is_oa": False,
    }
    bag: dict[str, OACandidate] = {}

    def absorb_list(candidates: list[OACandidate]) -> None:
        for c in candidates:
            _add_candidate(
                bag,
                c.url,
                source_api=c.source_api,
                host_type=c.host_type,
                version=c.version,
                license=c.license,
                is_direct_pdf=c.is_direct_pdf,
            )

    if doi:
        # OpenAlex
        try:
            oa_data, candidates = fetch_openalex(session, doi, email)
            absorb_list(candidates)
            if oa_data:
                open_access = oa_data.get("open_access")
                if isinstance(open_access, dict):
                    meta["is_oa"] = bool(open_access.get("is_oa")) or meta["is_oa"]

                primary = oa_data.get("primary_location")
                if isinstance(primary, dict):
                    source_obj = primary.get("source")
                    if isinstance(source_obj, dict):
                        host = (
                            source_obj.get("host_organization_name")
                            or source_obj.get("display_name")
                        )
                        if host:
                            meta["publisher"] = str(host)
                    landing = primary.get("landing_page_url")
                    if isinstance(landing, str) and landing:
                        meta["publisher_url"] = landing

                host_venue = oa_data.get("host_venue")
                if isinstance(host_venue, dict) and not meta["publisher"]:
                    pub = host_venue.get("publisher")
                    if pub:
                        meta["publisher"] = str(pub)
        except Exception as exc:  # noqa: BLE001 - continue with other sources
            logger.warning("OpenAlex lookup failed for %s: %s", doi, exc)
            print(f"  OpenAlex lookup error: {exc}")

        # Unpaywall
        try:
            upw_data, candidates = fetch_unpaywall(session, doi, email)
            absorb_list(candidates)
            if upw_data:
                meta["is_oa"] = bool(upw_data.get("is_oa")) or meta["is_oa"]
                if upw_data.get("publisher"):
                    meta["publisher"] = meta["publisher"] or str(upw_data.get("publisher"))
                if upw_data.get("doi_url"):
                    meta["publisher_url"] = str(upw_data.get("doi_url"))
                journal = upw_data.get("journal_name")
                if journal and not meta["publisher"]:
                    meta["publisher"] = str(journal)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Unpaywall lookup failed for %s: %s", doi, exc)
            print(f"  Unpaywall lookup error: {exc}")

        # Crossref (needed for publisher metadata when missing, or OA links if none yet)
        if not bag or not meta["publisher"]:
            try:
                cr_data, candidates = fetch_crossref(session, doi, email)
                absorb_list(candidates)
                if cr_data:
                    pub = cr_data.get("publisher")
                    if pub:
                        meta["publisher"] = meta["publisher"] or str(pub)
            except Exception as exc:  # noqa: BLE001
                logger.warning("Crossref lookup failed for %s: %s", doi, exc)
                print(f"  Crossref lookup error: {exc}")

    # Europe PMC (query when no candidates yet, or when PMID is provided)
    if not bag or pmid:
        try:
            _, candidates = fetch_europe_pmc(session, doi=doi, pmid=pmid)
            absorb_list(candidates)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Europe PMC lookup failed for %s: %s", doi or pmid, exc)
            print(f"  Europe PMC lookup error: {exc}")

    # Semantic Scholar fallback when still no candidates
    if doi and not bag:
        absorb_list(fetch_semantic_scholar_by_doi(session, doi))

    ranked = sorted(bag.values(), key=lambda c: c.score, reverse=True)
    # Prefer direct PDF candidates and PDF-shaped URLs.
    pdf_like = [
        c
        for c in ranked
        if c.is_direct_pdf
        or ".pdf" in c.url.lower()
        or "/pdf" in c.url.lower()
        or c.url.lower().endswith("/pdf/")
    ]
    return (pdf_like or ranked), meta


# ---------------------------------------------------------------------------
# Download / validation
# ---------------------------------------------------------------------------


def _starts_with_pdf_magic(raw: bytes) -> bool:
    """Return True if raw bytes begin with %PDF- (allowing minor leading whitespace/BOM)."""
    head = raw[:1024].lstrip(b"\xef\xbb\xbf \t\r\n")
    return head.startswith(b"%PDF-")


def validate_pdf(path: Path) -> tuple[bool, str]:
    """Validate that a file is a real PDF by magic bytes and size."""
    try:
        size = path.stat().st_size
    except OSError as exc:
        return False, f"stat failed: {exc}"
    if size < MIN_PDF_BYTES:
        return False, f"file too small ({size} bytes)"
    try:
        with path.open("rb") as handle:
            head = handle.read(1024)
    except OSError as exc:
        return False, f"read failed: {exc}"
    if not _starts_with_pdf_magic(head):
        return False, "missing %PDF- magic bytes (likely HTML or error page)"
    return True, ""


def download_pdf(
    session: requests.Session,
    url: str,
    dest: Path,
) -> tuple[bool, str]:
    """
    Download a URL to dest after validating it is a PDF.

    Writes via a temp file and only keeps the result when validation passes.
    """
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = TEMP_DIR / f"{dest.stem}.partial"
    parsed = urlparse(url)
    referer = f"{parsed.scheme}://{parsed.netloc}/" if parsed.scheme and parsed.netloc else ""
    try:
        last_error = ""
        headers = {
            "Accept": "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        }
        if referer:
            headers["Referer"] = referer

        response = request_with_retries(
            session,
            "GET",
            url,
            stream=True,
            headers=headers,
        )

        if response.status_code != 200:
            code = response.status_code
            response.close()
            if code in {401, 403}:
                return False, f"HTTP {code} (publisher blocked automated download)"
            if code == 404:
                return False, "HTTP 404 (PDF URL not found)"
            return False, f"HTTP {code}"

        content_type = (response.headers.get("Content-Type") or "").lower()
        chunks: list[bytes] = []
        total_bytes = 0
        # Read first chunk to fail fast on obvious HTML/XML pages.
        for chunk in response.iter_content(chunk_size=64 * 1024):
            if not chunk:
                continue
            if not chunks and content_type and "pdf" not in content_type and "octet-stream" not in content_type:
                if not _starts_with_pdf_magic(chunk):
                    response.close()
                    if "html" in content_type:
                        return False, "HTML page returned instead of PDF (bot protection or paywall)"
                    return False, f"non-PDF content-type: {content_type}"
            chunks.append(chunk)
            total_bytes += len(chunk)
            if total_bytes > 100 * 1024 * 1024:  # 100 MB safety cap
                response.close()
                return False, "PDF exceeds 100 MB safety limit"
        response.close()
        data = b"".join(chunks)

        if not _starts_with_pdf_magic(data):
            if "html" in content_type:
                return False, "HTML page returned instead of PDF (bot protection or paywall)"
            return False, last_error or "response body is not a PDF"

        if len(data) < MIN_PDF_BYTES:
            return False, f"PDF too small ({len(data)} bytes)"

        temp_path.write_bytes(data)
        ok, reason = validate_pdf(temp_path)
        if not ok:
            temp_path.unlink(missing_ok=True)
            return False, reason

        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.exists():
            dest.unlink()
        temp_path.replace(dest)
        return True, ""
    except Exception as exc:  # noqa: BLE001
        temp_path.unlink(missing_ok=True)
        return False, str(exc)
    finally:
        temp_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Logging / resume
# ---------------------------------------------------------------------------


def load_download_log() -> list[dict[str, str]]:
    """Load existing download log rows."""
    if not DOWNLOAD_LOG.exists():
        return []
    with DOWNLOAD_LOG.open("r", newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        return [dict(row) for row in reader]


def _log_matches(
    row: dict[str, str], doi: str, paper_number: int, title: str = ""
) -> bool:
    """Return True if a log row refers to the same paper."""
    row_doi = clean_doi(row.get("doi") or "").lower()
    doi_l = doi.lower()
    if doi_l and row_doi:
        return row_doi == doi_l
    if doi_l or row_doi:
        return False
    row_num = (row.get("paper_number") or "").strip()
    if row_num and row_num == str(paper_number):
        row_title = _normalize_title_for_match(row.get("title") or "")
        if not title or not row_title:
            return True
        return row_title == _normalize_title_for_match(title)
    return False


def already_successful(
    doi: str, paper_number: int, log_rows: list[dict[str, str]], title: str = ""
) -> bool:
    """Return True if this paper was already downloaded successfully."""
    for row in log_rows:
        status = (row.get("status") or "").strip().lower()
        if status in SUCCESS_STATUSES and _log_matches(row, doi, paper_number, title):
            return True
    return False


def already_processed(
    doi: str, paper_number: int, log_rows: list[dict[str, str]], title: str = ""
) -> bool:
    """Return True if any prior outcome was logged for this paper."""
    return any(_log_matches(row, doi, paper_number, title) for row in log_rows)


def find_existing_pdf(doi: str, paper_number: int, filename: str) -> Optional[Path]:
    """Search publisher folders for an already-saved PDF matching the exact filename."""
    for folder in PUBLISHER_DIRS:
        candidate = DOWNLOADED_DIR / folder / filename
        if candidate.exists() and candidate.stat().st_size >= MIN_PDF_BYTES:
            ok, _ = validate_pdf(candidate)
            if ok:
                return candidate
    return None


def append_csv(path: Path, fields: list[str], row: dict[str, Any]) -> None:
    """
    Append one row to a CSV, writing a header if the file is new/empty.

    Retries on Windows PermissionError (common when Excel has the CSV open).
    """
    payload = {key: row.get(key, "") for key in fields}
    last_error: Optional[Exception] = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            write_header = (not path.exists()) or path.stat().st_size == 0
            with path.open("a", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(
                    handle, fieldnames=fields, extrasaction="ignore"
                )
                if write_header:
                    writer.writeheader()
                writer.writerow(payload)
            return
        except PermissionError as exc:
            last_error = exc
            wait = RETRY_BACKOFF_BASE ** attempt
            logger.warning(
                "Cannot write %s (attempt %s/%s). Close the file if it is open "
                "in Excel. Retrying in %ss...",
                path.name,
                attempt,
                MAX_RETRIES,
                wait,
            )
            print(
                f"  Warning: cannot write {path.name} "
                f"(permission denied). Close it in Excel if open. "
                f"Retry {attempt}/{MAX_RETRIES} in {wait}s..."
            )
            time.sleep(wait)

    assert last_error is not None
    raise PermissionError(
        f"Permission denied writing {path}. "
        f"Close '{path.name}' in Excel/another program and re-run the script."
    ) from last_error


def save_log(
    paper: PaperRecord,
    result: ProcessResult,
) -> None:
    """Write outcomes to the appropriate log CSV files."""
    timestamp = utc_timestamp()
    append_csv(
        DOWNLOAD_LOG,
        DOWNLOAD_LOG_FIELDS,
        {
            "paper_number": paper.paper_number,
            "title": paper.title,
            "doi": paper.doi or paper.doi_raw,
            "source": paper.source,
            "publisher": result.publisher or paper.publisher,
            "year": paper.year,
            "destination_folder": result.destination_folder,
            "filename": result.filename,
            "pdf_url": result.pdf_url,
            "status": result.status,
            "error": result.error,
            "timestamp": timestamp,
        },
    )

    if result.status in {"failed", "invalid_doi"}:
        append_csv(
            FAILED_LOG,
            FAILED_LOG_FIELDS,
            {
                "paper_number": paper.paper_number,
                "title": paper.title,
                "doi": paper.doi or paper.doi_raw,
                "source": paper.source,
                "reason": result.status,
                "error": result.error,
            },
        )

    if result.status in {"access_required", "no_open_access"}:
        append_csv(
            ACCESS_LOG,
            ACCESS_LOG_FIELDS,
            {
                "paper_number": paper.paper_number,
                "title": paper.title,
                "doi": paper.doi or paper.doi_raw,
                "publisher": result.publisher or paper.publisher,
                "source": paper.source,
                "publisher_url": result.pdf_url or (
                    f"https://doi.org/{paper.doi}" if paper.doi else ""
                ),
                "reason": result.error or result.status,
            },
        )


# ---------------------------------------------------------------------------
# Per-paper processing
# ---------------------------------------------------------------------------


def _normalize_title_for_match(title: str) -> str:
    """Normalize a title for robust bibliographic matching."""
    text = (title or "").lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _first_author_surname(authors: str) -> str:
    """Extract a best-effort first-author surname from an Excel author field."""
    text = (authors or "").strip()
    if not text:
        return ""
    first = re.split(r"\s*(?:;|\||\band\b)\s*|\s*,\s*", text, maxsplit=1, flags=re.I)[0]
    first = re.sub(r"[^A-Za-zÀ-ÿ'\- ]+", " ", first).strip()
    parts = first.split()
    return parts[-1] if parts else ""


def _title_similarity(target: str, candidate: str) -> float:
    """Score how closely two normalized titles match."""
    if not target or not candidate:
        return 0.0
    title_score = difflib.SequenceMatcher(None, target, candidate).ratio()
    target_tokens = set(target.split())
    candidate_tokens = set(candidate.split())
    overlap = (
        len(target_tokens & candidate_tokens) / len(target_tokens)
        if target_tokens
        else 0.0
    )
    return 0.75 * title_score + 0.25 * overlap


def _year_bonus(expected_year: str, candidate_year: Any) -> float:
    """Small score boost when publication years agree."""
    expected = str(expected_year or "").strip()
    if not expected or not re.fullmatch(r"\d{4}", expected):
        return 0.0
    cand = str(candidate_year or "").strip()
    if re.fullmatch(r"\d{4}", cand) and cand == expected:
        return 0.06
    return 0.0


def _consider_resolution(
    best: TitleResolution,
    *,
    doi: str = "",
    pdf_url: str = "",
    source_api: str,
    publisher: str = "",
    score: float,
    core_score: float = 0.0,
    paper_id: str = "",
    matched_title: str = "",
) -> TitleResolution:
    """Keep the highest-scoring title resolution seen so far."""
    doi = clean_doi(doi) if doi else ""
    if doi and not is_valid_doi(doi):
        doi = ""
    if not doi and not pdf_url:
        return best
    if score > best.score:
        return TitleResolution(
            doi=doi,
            pdf_url=pdf_url or "",
            source_api=source_api,
            publisher=publisher or "",
            score=score,
            core_score=core_score,
            paper_id=paper_id or "",
            matched_title=matched_title or "",
        )
    # Prefer filling a missing DOI/PDF on an equally good match.
    if score >= best.score - 1e-9:
        if not best.doi and doi:
            best.doi = doi
        if not best.pdf_url and pdf_url:
            best.pdf_url = pdf_url
            best.source_api = best.source_api or source_api
        if not best.publisher and publisher:
            best.publisher = publisher
        if not best.paper_id and paper_id:
            best.paper_id = paper_id
        if not best.matched_title and matched_title:
            best.matched_title = matched_title
        if core_score > best.core_score:
            best.core_score = core_score
    return best


def _resolve_via_crossref(
    session: requests.Session,
    title: str,
    authors: str,
    year: str,
    email: str,
) -> TitleResolution:
    """Resolve identity with Crossref bibliographic search."""
    best = TitleResolution()
    author = _first_author_surname(authors)
    query = title.strip()
    if author:
        query = f"{query} {author}"

    params: dict[str, Any] = {
        "query.bibliographic": query,
        "rows": 8,
        "mailto": email,
    }
    if year and re.fullmatch(r"\d{4}", year.strip()):
        params["query.bibliographic"] = query  # year used only for scoring

    try:
        response = request_with_retries(
            session,
            "GET",
            "https://api.crossref.org/works",
            params=params,
            accept="application/json",
        )
        time.sleep(INTER_REQUEST_DELAY)
        response.raise_for_status()
        data = response.json()
    except Exception as exc:  # noqa: BLE001
        logger.warning("Crossref title lookup failed for %r: %s", title, exc)
        return best

    target = _normalize_title_for_match(title)
    for item in (data.get("message") or {}).get("items") or []:
        item_titles = item.get("title") or []
        if not item_titles:
            continue
        candidate_title = str(item_titles[0])
        candidate_norm = _normalize_title_for_match(candidate_title)
        core = _title_similarity(target, candidate_norm)
        score = core
        if author:
            candidate_authors = item.get("author") or []
            author_text = " ".join(
                str(a.get("family") or "")
                for a in candidate_authors
                if isinstance(a, dict)
            ).lower()
            if author.lower() in author_text:
                score += 0.08
        issued = ((item.get("issued") or {}).get("date-parts") or [[None]])[0]
        score += _year_bonus(year, issued[0] if issued else None)
        doi = clean_doi(str(item.get("DOI") or ""))
        best = _consider_resolution(
            best,
            doi=doi,
            source_api="Crossref",
            publisher=str(item.get("publisher") or ""),
            score=score,
            core_score=core,
            matched_title=candidate_title,
        )
    return best


def _resolve_via_openalex(
    session: requests.Session,
    title: str,
    authors: str,
    year: str,
    email: str,
) -> TitleResolution:
    """Resolve identity with OpenAlex title search (helpful for Scholar extracts)."""
    best = TitleResolution()
    target = _normalize_title_for_match(title)
    if not target:
        return best

    # Do not hard-filter by year: Scholar extracts are often off by a year, and
    # brand-new papers may not be indexed under the spreadsheet year yet.
    params: dict[str, Any] = {
        "search": title.strip(),
        "per_page": 8,
        "mailto": email,
    }

    try:
        response = request_with_retries(
            session,
            "GET",
            "https://api.openalex.org/works",
            params=params,
            accept="application/json",
        )
        time.sleep(INTER_REQUEST_DELAY)
        response.raise_for_status()
        data = response.json()
    except Exception as exc:  # noqa: BLE001
        logger.warning("OpenAlex title lookup failed for %r: %s", title, exc)
        return best

    author = _first_author_surname(authors)
    for item in data.get("results") or []:
        candidate_title = str(item.get("title") or "")
        candidate_norm = _normalize_title_for_match(candidate_title)
        core = _title_similarity(target, candidate_norm)
        score = core
        if author:
            authorships = item.get("authorships") or []
            names = []
            for auth in authorships:
                author_obj = auth.get("author") if isinstance(auth, dict) else None
                if isinstance(author_obj, dict):
                    names.append(str(author_obj.get("display_name") or "").lower())
            if any(author.lower() in name for name in names):
                score += 0.08
        score += _year_bonus(year, item.get("publication_year"))

        ids = item.get("ids") or {}
        doi = clean_doi(str(ids.get("doi") or item.get("doi") or ""))
        pdf_url = ""
        best_oa = item.get("best_oa_location") or {}
        if isinstance(best_oa, dict):
            pdf_url = str(best_oa.get("pdf_url") or "")
        host = ""
        primary = item.get("primary_location") or {}
        if isinstance(primary, dict):
            source_obj = primary.get("source")
            if isinstance(source_obj, dict):
                host = str(
                    source_obj.get("host_organization_name")
                    or source_obj.get("display_name")
                    or ""
                )

        best = _consider_resolution(
            best,
            doi=doi,
            pdf_url=pdf_url,
            source_api="OpenAlex",
            publisher=host,
            score=score,
            core_score=core,
            paper_id=str(item.get("id") or ""),
            matched_title=candidate_title,
        )
    return best


def _resolve_via_semantic_scholar(
    session: requests.Session,
    title: str,
    authors: str,
    year: str,
) -> TitleResolution:
    """
    Resolve identity with Semantic Scholar (covers many Google Scholar papers).

    Uses the public Academic Graph API only — does not scrape scholar.google.com.
    """
    best = TitleResolution()
    target = _normalize_title_for_match(title)
    if not target:
        return best

    params = {
        "query": title.strip(),
        "limit": 8,
        "fields": "title,year,authors,externalIds,openAccessPdf,venue",
    }
    try:
        # Single attempt: Semantic Scholar often 429s anonymous clients; don't stall the job.
        response = session.get(
            "https://api.semanticscholar.org/graph/v1/paper/search",
            params=params,
            headers={"Accept": "application/json"},
            timeout=REQUEST_TIMEOUT,
        )
        time.sleep(INTER_REQUEST_DELAY)
        if response.status_code in {429, 403}:
            logger.warning(
                "Semantic Scholar title lookup rate-limited (%s) for %r",
                response.status_code,
                title,
            )
            return best
        response.raise_for_status()
        data = response.json()
    except Exception as exc:  # noqa: BLE001
        logger.warning("Semantic Scholar title lookup failed for %r: %s", title, exc)
        return best

    author = _first_author_surname(authors)
    for item in data.get("data") or []:
        candidate_title = str(item.get("title") or "")
        candidate_norm = _normalize_title_for_match(candidate_title)
        core = _title_similarity(target, candidate_norm)
        score = core
        if author:
            names = " ".join(
                str(a.get("name") or "").lower()
                for a in (item.get("authors") or [])
                if isinstance(a, dict)
            )
            if author.lower() in names:
                score += 0.08
        score += _year_bonus(year, item.get("year"))

        external = item.get("externalIds") or {}
        doi = clean_doi(str(external.get("DOI") or ""))
        oa = item.get("openAccessPdf") or {}
        pdf_url = ""
        if isinstance(oa, dict):
            pdf_url = str(oa.get("url") or "")

        best = _consider_resolution(
            best,
            doi=doi,
            pdf_url=pdf_url,
            source_api="SemanticScholar",
            publisher=str(item.get("venue") or ""),
            score=score,
            core_score=core,
            paper_id=str(item.get("paperId") or ""),
            matched_title=candidate_title,
        )
    return best


def resolve_from_title(
    session: requests.Session,
    title: str,
    authors: str,
    email: str,
    year: str = "",
) -> TitleResolution:
    """
    Resolve a missing DOI / OA PDF from title metadata.

    Crossref + OpenAlex + Semantic Scholar are queried. This is the path used for
    Google Scholar extracts that often omit DOIs.
    """
    if not title or not title.strip():
        return TitleResolution()

    print("  Looking up title via Crossref / OpenAlex / Semantic Scholar...")
    candidates = [
        _resolve_via_crossref(session, title, authors, year, email),
        _resolve_via_openalex(session, title, authors, year, email),
        _resolve_via_semantic_scholar(session, title, authors, year),
    ]

    best = TitleResolution()
    for cand in candidates:
        best = _consider_resolution(
            best,
            doi=cand.doi,
            pdf_url=cand.pdf_url,
            source_api=cand.source_api,
            publisher=cand.publisher,
            score=cand.score,
            core_score=cand.core_score,
            paper_id=cand.paper_id,
            matched_title=cand.matched_title,
        )

    # Conservative acceptance: require a strong core title match.
    # Bonuses (author/year) can push composite score above 1.0.
    title_norm = _normalize_title_for_match(title)
    accept = (
        bool(title_norm)
        and len(title_norm.split()) >= 4
        and best.core_score >= 0.86
        and (
            best.score >= 0.90
            or (
                best.score >= 0.85
                and bool(year or authors)
                and bool(best.doi or best.pdf_url)
            )
        )
    )
    if not accept:
        logger.info(
            "No sufficiently strong title match: score=%.3f core=%.3f | %s",
            best.score,
            best.core_score,
            title,
        )
        return TitleResolution()

    logger.info(
        "Resolved from title via %s (score=%.3f core=%.3f): doi=%s pdf=%s | %s",
        best.source_api,
        best.score,
        best.core_score,
        best.doi or "-",
        "yes" if best.pdf_url else "no",
        title,
    )
    return best


def find_doi_by_title(
    session: requests.Session,
    title: str,
    authors: str,
    email: str,
    year: str = "",
) -> Optional[str]:
    """Backward-compatible DOI-only helper around resolve_from_title()."""
    resolved = resolve_from_title(session, title, authors, email, year=year)
    return resolved.doi or None


def process_paper(
    paper: PaperRecord,
    session: requests.Session,
    email: str,
    *,
    dry_run: bool,
    log_rows: list[dict[str, str]],
) -> ProcessResult:
    """Process a single paper through classify → OA lookup → download."""
    print("--------------------------------------------------")
    print(f"[{paper.paper_number}] Processing:")
    print(f"Title: {paper.title or '(no title)'}")
    print(f"DOI: {paper.doi or paper.doi_raw or '(none)'}")
    print(f"Source: {paper.source or '(none)'}")

    title_oa_seed: Optional[OACandidate] = None
    if not paper.doi or not is_valid_doi(paper.doi):
        if paper.pmid:
            try:
                epmc_hit, epmc_cands = fetch_europe_pmc(session, pmid=paper.pmid)
                epmc_doi = clean_doi(str(epmc_hit.get("doi") or ""))
                if epmc_doi and is_valid_doi(epmc_doi):
                    paper.doi = epmc_doi
                    print(f"[OK] DOI resolved from PMID {paper.pmid}: {paper.doi}")
                if epmc_cands:
                    title_oa_seed = epmc_cands[0]
            except Exception as exc:  # noqa: BLE001
                logger.warning("Europe PMC PMID lookup failed for %s: %s", paper.pmid, exc)

    if not paper.doi or not is_valid_doi(paper.doi):
        print("[!] DOI missing - searching by title (Crossref / OpenAlex / Semantic Scholar)...")

        resolved = resolve_from_title(
            session,
            paper.title,
            paper.authors,
            email,
            year=paper.year,
        )

        if resolved.doi:
            paper.doi = resolved.doi
            print(f"[OK] DOI found from title via {resolved.source_api}: {resolved.doi}")
        if resolved.pdf_url:
            title_oa_seed = OACandidate(
                url=resolved.pdf_url,
                source_api=resolved.source_api or "TitleLookup",
                host_type="other",
                is_direct_pdf=True,
            )
            print(f"[OK] OA PDF candidate from {resolved.source_api}")
        if resolved.publisher and not paper.publisher:
            paper.publisher = resolved.publisher

        if not paper.doi and not title_oa_seed:
            print("[X] Could not identify DOI or OA PDF from title")
            folder = classify_publisher(
                source=paper.source, publisher=paper.publisher
            )
            result = ProcessResult(
                status="invalid_doi",
                publisher=folder,
                error=(
                    "Missing or invalid DOI, and no matching paper was found by title search."
                ),
            )
            if not dry_run:
                save_log(paper, result)
            return result

    filename = create_safe_filename(paper.paper_number, paper.title)
    existing = find_existing_pdf(paper.doi, paper.paper_number, filename)
    if existing is not None or (
        paper.doi and already_successful(paper.doi, paper.paper_number, log_rows, paper.title)
    ):
        folder = existing.parent.name if existing else classify_publisher(
            source=paper.source, publisher=paper.publisher, doi=paper.doi
        )
        name = existing.name if existing else filename
        print("[OK] Already exists - skipping download")
        print(f"Saved to:\nDownloaded_Papers/{folder}/{name}")
        result = ProcessResult(
            status="already_exists",
            destination_folder=folder,
            filename=name,
            publisher=folder,
        )
        if not dry_run:
            # Avoid duplicate log spam if already logged as success.
            if paper.doi and not already_successful(paper.doi, paper.paper_number, log_rows, paper.title):
                save_log(paper, result)
        return result

    meta: dict[str, Any] = {
        "publisher": paper.publisher,
        "publisher_url": f"https://doi.org/{paper.doi}" if paper.doi else "",
        "is_oa": False,
    }
    candidates: list[OACandidate] = []

    if paper.doi and is_valid_doi(paper.doi):
        try:
            candidates, meta = get_open_access_links(
                session, paper.doi, email, pmid=paper.pmid
            )
        except Exception as exc:  # noqa: BLE001
            result = ProcessResult(
                status="failed",
                publisher=paper.publisher,
                error=f"Metadata/OA lookup failed: {exc}",
            )
            print(f"[X] OA lookup failed: {exc}")
            if not dry_run:
                save_log(paper, result)
            return result

    if title_oa_seed is not None:
        # Seed from title/PMID APIs (esp. Semantic Scholar / OpenAlex for Scholar rows).
        bag: dict[str, OACandidate] = {}
        for cand in candidates:
            _add_candidate(
                bag,
                cand.url,
                source_api=cand.source_api,
                host_type=cand.host_type,
                version=cand.version,
                license=cand.license,
                is_direct_pdf=cand.is_direct_pdf,
            )
        _add_candidate(
            bag,
            title_oa_seed.url,
            source_api=title_oa_seed.source_api,
            host_type=title_oa_seed.host_type,
            is_direct_pdf=True,
        )
        candidates = sorted(bag.values(), key=lambda c: c.score, reverse=True)
        meta["is_oa"] = True

    folder = classify_publisher(
        source=paper.source,
        publisher=paper.publisher,
        doi=paper.doi,
        url=meta.get("publisher_url", ""),
        meta_publisher=meta.get("publisher", ""),
    )
    dest = DOWNLOADED_DIR / folder / filename
    print(f"Destination folder: Downloaded_Papers/{folder}/")
    print(f"Filename: {filename}")

    # Crossref often lists subscriber-only PDF URLs (especially IEEE).
    # Keep Crossref candidates only when another API marked the work OA,
    # or when the URL host is a known open repository/publisher.
    candidates = [
        c
        for c in candidates
        if c.source_api != "Crossref"
        or meta.get("is_oa")
        or _is_known_oa_host(c.url)
    ]

    if not candidates:
        print("[X] No legal open-access PDF found")
        print(
            "Added to access_required.csv"
            if not dry_run
            else "(dry-run: would add to access_required.csv)"
        )
        reason = (
            "Marked open-access in index, but no downloadable PDF URL is exposed by OpenAlex/Unpaywall/Europe PMC"
            if meta.get("is_oa")
            else "Not open access / paywalled — no free legal PDF found"
        )
        result = ProcessResult(
            status="no_open_access",
            destination_folder=folder,
            filename=filename,
            publisher=meta.get("publisher") or folder,
            pdf_url=meta.get("publisher_url", ""),
            error=reason,
        )
        if not dry_run:
            save_log(paper, result)
        return result

    # Try candidates in ranked order.
    last_error = ""
    for candidate in candidates:
        print(
            f"Open access PDF candidate ({candidate.source_api}, "
            f"host={candidate.host_type}, score={candidate.score}):"
        )
        print(f"  {candidate.url}")
        if dry_run:
            print("[OK] Dry-run: would attempt download (not downloading)")
            return ProcessResult(
                status="downloaded",
                destination_folder=folder,
                filename=filename,
                pdf_url=candidate.url,
                publisher=meta.get("publisher") or folder,
                oa_source=candidate.source_api,
                error="dry-run",
            )

        print("Downloading...")
        ok, err = download_pdf(session, candidate.url, dest)
        time.sleep(INTER_REQUEST_DELAY)
        if ok:
            # Re-classify using the successful URL host if more specific.
            url_folder = classify_from_host(candidate.url)
            final_folder = folder
            if url_folder and url_folder != folder and url_folder != "Scopus":
                # Keep Excel/source classification unless it was generic.
                if folder in {"Other", "Scopus"}:
                    final_folder = url_folder
                    final_path = DOWNLOADED_DIR / final_folder / filename
                    final_path.parent.mkdir(parents=True, exist_ok=True)
                    if dest.exists() and final_path != dest:
                        dest.replace(final_path)

            print("[OK] Downloaded")
            print(f"Saved to:\nDownloaded_Papers/{final_folder}/{filename}")
            print(f"OA source: {candidate.source_api}")
            result = ProcessResult(
                status="downloaded",
                destination_folder=final_folder,
                filename=filename,
                pdf_url=candidate.url,
                publisher=meta.get("publisher") or final_folder,
                oa_source=candidate.source_api,
            )
            save_log(paper, result)
            return result

        last_error = err
        print(f"  Candidate failed validation/download: {err}")

    print("[X] No legal open-access PDF found")
    print("Added to access_required.csv")
    best_candidate_url = candidates[0].url if candidates else meta.get("publisher_url", "")
    result = ProcessResult(
        status="access_required",
        destination_folder=folder,
        filename=filename,
        publisher=meta.get("publisher") or folder,
        pdf_url=best_candidate_url or meta.get("publisher_url", ""),
        oa_source=candidates[0].source_api if candidates else "",
        error=last_error or "OA candidate URL did not return a valid PDF",
    )
    save_log(paper, result)
    return result


# ---------------------------------------------------------------------------
# CLI / main
# ---------------------------------------------------------------------------


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Download legally available open-access research PDFs."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Classify and look up OA links without downloading.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Process only the first N papers (0 = all).",
    )
    parser.add_argument(
        "--retry-failed",
        action="store_true",
        help="Retry papers previously logged as failed / access_required / no OA.",
    )
    parser.add_argument(
        "--email",
        type=str,
        default="",
        help="Contact email for Unpaywall/OpenAlex polite pool (or set UNPAYWALL_EMAIL).",
    )
    parser.add_argument(
        "--excel",
        type=str,
        default=str(BASE_DIR / EXCEL_NAME),
        help="Path to the input Excel file (read-only).",
    )
    return parser.parse_args(argv)


def resolve_email(cli_email: str, *, exit_on_missing: bool = True) -> str:
    """Resolve the contact email used for API courtesy parameters."""
    email = (
        cli_email
        or os.environ.get("UNPAYWALL_EMAIL")
        or os.environ.get("PAPER_DOWNLOADER_EMAIL")
        or ""
    ).strip()
    if not email:
        message = (
            "An email address is required for Unpaywall/OpenAlex.\n"
            "Pass --email you@example.com or set UNPAYWALL_EMAIL."
        )
        if exit_on_missing:
            raise SystemExit(message)
        raise ValueError(message)
    return email


ProgressCallback = Callable[[dict[str, Any]], None]


def _update_summary(summary: Summary, status: str) -> None:
    """Increment summary counters for a finished paper status."""
    if status == "downloaded":
        summary.downloaded += 1
    elif status == "already_exists":
        summary.already_exists += 1
    elif status == "no_open_access":
        summary.no_oa += 1
        summary.not_found += 1
    elif status == "access_required":
        summary.access_required += 1
        summary.not_found += 1
    elif status == "invalid_doi":
        summary.invalid_doi += 1
        summary.not_found += 1
    elif status == "failed":
        summary.failed += 1
        summary.not_found += 1


def _human_reason(row: dict[str, Any]) -> str:
    """Return a clear, user-friendly status reason for Excel/UI reports."""
    status = str(row.get("status") or "")
    err = str(row.get("error") or "").strip()
    if status == "downloaded":
        if err == "dry-run":
            return "Dry run: open-access PDF URL found"
        src = str(row.get("oa_source") or "").strip()
        return f"Downloaded via {src}" if src else "Downloaded successfully"
    if status == "already_exists":
        return "Already downloaded previously"
    if err:
        return err
    if status == "no_open_access":
        return "Not open access / paywalled — no free legal PDF found"
    if status == "access_required":
        return "Open-access PDF link returned non-PDF or blocked automated download"
    if status == "invalid_doi":
        return "Missing or invalid DOI, and title lookup found no reliable match"
    if status == "failed":
        return "Lookup or download error"
    return "Unknown status"


def generate_excel_report(
    results: list[dict[str, Any]],
    summary: dict[str, Any],
    output_path: Path,
    *,
    source_filename: str = "",
    email: str = "",
    dry_run: bool = False,
) -> Path:
    """
    Create a multi-sheet Excel (.xlsx) report of downloaded vs. not-downloaded papers
    and the exact reason why each unavailable paper was not downloaded.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color="10231C", end_color="10231C", fill_type="solid")
    header_font = Font(color="F4FFE8", bold=True, size=11)
    ok_fill = PatternFill(start_color="E7F6EC", end_color="E7F6EC", fill_type="solid")
    ok_font = Font(color="067647", bold=True)
    warn_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    warn_font = Font(color="92400E", bold=True)
    err_fill = PatternFill(start_color="FEE4E2", end_color="FEE4E2", fill_type="solid")
    err_font = Font(color="B42318", bold=True)

    # --- Sheet 1: Summary ---
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.append(["PaperFetch Download Summary", ""])
    ws_sum["A1"].font = Font(bold=True, size=14, color="10231C")
    ws_sum.append([])
    total = int(summary.get("total") or len(results))
    downloaded = int(summary.get("downloaded") or 0)
    already_exists = int(summary.get("already_exists") or 0)
    saved_total = downloaded + already_exists
    not_found = int(
        summary.get("not_found")
        if summary.get("not_found") is not None
        else (
            int(summary.get("no_oa") or 0)
            + int(summary.get("access_required") or 0)
            + int(summary.get("invalid_doi") or 0)
            + int(summary.get("failed") or 0)
        )
    )

    meta_rows = [
        ("Source Spreadsheet", source_filename or "Uploaded Excel"),
        ("Generated (UTC)", utc_timestamp()),
        ("Mode", "Dry run (no PDFs saved)" if dry_run else "Full PDF download"),
        ("Contact Email", email or "-"),
        ("", ""),
        ("Total Papers in Run", total),
        ("Downloaded (Total Saved)", saved_total),
        ("  - Newly Downloaded", downloaded),
        ("  - Already Existed", already_exists),
        ("Not Found / Not Downloaded", not_found),
        ("  - No Open Access (Paywalled)", int(summary.get("no_oa") or 0)),
        ("  - Blocked / Non-PDF Response", int(summary.get("access_required") or 0)),
        ("  - Missing / Invalid DOI", int(summary.get("invalid_doi") or 0)),
        ("  - Network / Lookup Failed", int(summary.get("failed") or 0)),
    ]
    for label, val in meta_rows:
        ws_sum.append([label, val])
    for row_idx in range(4, ws_sum.max_row + 1):
        cell_a = ws_sum.cell(row=row_idx, column=1)
        cell_b = ws_sum.cell(row=row_idx, column=2)
        if cell_a.value in {"Total Papers in Run", "Downloaded (Total Saved)", "Not Found / Not Downloaded"}:
            cell_a.font = Font(bold=True, size=11)
            cell_b.font = Font(bold=True, size=11)
    ws_sum.column_dimensions["A"].width = 34
    ws_sum.column_dimensions["B"].width = 48

    columns = [
        ("Paper #", 10),
        ("Title", 48),
        ("DOI", 28),
        ("Source", 18),
        ("Year", 8),
        ("Publisher", 26),
        ("Outcome", 16),
        ("Detailed Status", 18),
        ("Reason / Why Not", 46),
        ("Folder", 14),
        ("Saved Filename", 38),
        ("OA Source", 16),
        ("PDF / Publisher URL", 44),
    ]

    def populate_sheet(ws: Any, subset: list[dict[str, Any]]) -> None:
        ws.append([title for title, _ in columns])
        for col_idx, (_, width) in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, len(subset) + 1)}"

        for r_idx, row in enumerate(subset, start=2):
            status = str(row.get("status") or "")
            is_ok = status in SUCCESS_STATUSES
            outcome = "Downloaded" if is_ok else "Not Downloaded"
            reason = _human_reason(row)
            url = str(row.get("pdf_url") or "").strip()
            values = [
                row.get("paper_number", ""),
                row.get("title", ""),
                row.get("doi", ""),
                row.get("source", ""),
                row.get("year", ""),
                row.get("publisher", ""),
                outcome,
                status,
                reason,
                row.get("destination_folder", "") if is_ok else "",
                row.get("filename", "") if is_ok else "",
                row.get("oa_source", ""),
                url,
            ]
            ws.append(values)
            for c_idx in range(1, len(columns) + 1):
                ws.cell(row=r_idx, column=c_idx).alignment = Alignment(
                    vertical="top", wrap_text=(c_idx in {2, 6, 9, 11, 13})
                )
            outcome_cell = ws.cell(row=r_idx, column=7)
            if is_ok:
                outcome_cell.fill = ok_fill
                outcome_cell.font = ok_font
            elif status in {"no_open_access", "access_required"}:
                outcome_cell.fill = warn_fill
                outcome_cell.font = warn_font
            else:
                outcome_cell.fill = err_fill
                outcome_cell.font = err_font
            if url.startswith("http"):
                url_cell = ws.cell(row=r_idx, column=13)
                url_cell.hyperlink = url
                url_cell.font = Font(color="0563C1", underline="single")

    downloaded_rows = [r for r in results if str(r.get("status") or "") in SUCCESS_STATUSES]
    not_downloaded_rows = [r for r in results if str(r.get("status") or "") not in SUCCESS_STATUSES]

    ws_all = wb.create_sheet("All Papers")
    populate_sheet(ws_all, results)

    ws_dl = wb.create_sheet("Downloaded")
    populate_sheet(ws_dl, downloaded_rows)

    ws_not = wb.create_sheet("Not Downloaded")
    populate_sheet(ws_not, not_downloaded_rows)

    wb.save(output_path)
    return output_path


def run_download(
    excel_path: Path,
    email: str,
    *,
    dry_run: bool = False,
    limit: int = 0,
    retry_failed: bool = False,
    workspace: Optional[Path] = None,
    on_progress: Optional[ProgressCallback] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
    verbose: bool = False,
) -> tuple[Summary, list[dict[str, Any]]]:
    """
    Run a download job programmatically (used by the web UI).

    Returns the aggregate summary and a per-paper result list suitable for JSON.
    """
    if workspace is not None:
        configure_workspace(workspace)

    ensure_directories()
    ensure_log_files()

    excel_path = Path(excel_path).expanduser().resolve()
    email = resolve_email(email, exit_on_missing=False)

    df, columns = load_excel(excel_path, verbose=verbose)
    papers = dataframe_to_papers(df, columns)
    log_rows = load_download_log()

    to_process = select_papers(
        papers,
        log_rows,
        retry_failed=retry_failed,
        limit=limit,
    )
    summary = Summary(total=len(papers), skipped=max(0, len(papers) - len(to_process)))
    if limit:
        summary.skipped = 0
        summary.total = len(to_process)

    results: list[dict[str, Any]] = []

    if on_progress:
        on_progress(
            {
                "event": "started",
                "total_papers": summary.total,
                "to_process": len(to_process),
                "skipped": summary.skipped,
            }
        )

    if not to_process:
        if on_progress:
            on_progress({"event": "finished", "summary": summary.__dict__, "results": results})
        return summary, results

    session = build_session(email)

    for index, paper in enumerate(to_process, start=1):
        if cancel_check and cancel_check():
            summary.total = len(results)
            if on_progress:
                on_progress(
                    {
                        "event": "cancelled",
                        "summary": summary.__dict__,
                        "results": results,
                    }
                )
            return summary, results

        if on_progress:
            on_progress(
                {
                    "event": "paper_start",
                    "index": index,
                    "of": len(to_process),
                    "paper_number": paper.paper_number,
                    "title": paper.title,
                    "doi": paper.doi or paper.doi_raw,
                }
            )

        result = process_paper(
            paper,
            session,
            email,
            dry_run=dry_run,
            log_rows=log_rows,
        )

        if result.status == "downloaded" and not dry_run:
            _update_summary(summary, "downloaded")
            log_rows.append(
                {
                    "doi": paper.doi,
                    "paper_number": str(paper.paper_number),
                    "title": paper.title,
                    "status": "downloaded",
                }
            )
        elif result.status == "downloaded" and dry_run:
            _update_summary(summary, "downloaded")
        elif result.status == "already_exists":
            _update_summary(summary, "already_exists")
            log_rows.append(
                {
                    "doi": paper.doi,
                    "paper_number": str(paper.paper_number),
                    "title": paper.title,
                    "status": "already_exists",
                }
            )
        else:
            _update_summary(summary, result.status)

        row = {
            "paper_number": paper.paper_number,
            "title": paper.title,
            "doi": paper.doi or paper.doi_raw,
            "source": paper.source,
            "year": paper.year,
            "status": result.status,
            "destination_folder": result.destination_folder,
            "filename": result.filename,
            "pdf_url": result.pdf_url,
            "publisher": result.publisher or paper.publisher,
            "error": result.error,
            "oa_source": result.oa_source,
        }
        results.append(row)

        if on_progress:
            on_progress(
                {
                    "event": "paper_done",
                    "index": index,
                    "of": len(to_process),
                    "result": row,
                    "summary": summary.__dict__,
                }
            )

    if on_progress:
        on_progress({"event": "finished", "summary": summary.__dict__, "results": results})

    return summary, results


def select_papers(
    papers: list[PaperRecord],
    log_rows: list[dict[str, str]],
    *,
    retry_failed: bool,
    limit: int,
) -> list[PaperRecord]:
    """Select papers using resume/retry/limit only; no date filter is applied."""
    failed_dois: set[str] = set()
    failed_titles: set[str] = set()
    if retry_failed:
        for row in log_rows:
            if (row.get("status") or "").strip().lower() in RETRYABLE_STATUSES:
                d = clean_doi(row.get("doi") or "").lower()
                if d:
                    failed_dois.add(d)
                else:
                    t = _normalize_title_for_match(row.get("title") or "")
                    if t:
                        failed_titles.add(t)

    selected: list[PaperRecord] = []
    for paper in papers:
        if retry_failed:
            p_title = _normalize_title_for_match(paper.title)
            if (paper.doi and paper.doi.lower() in failed_dois) or (
                not paper.doi and p_title and p_title in failed_titles
            ):
                selected.append(paper)
            continue

        # Resume: skip anything already logged (success or prior OA/failure).
        # Use --retry-failed to revisit unsuccessful outcomes.
        if already_processed(paper.doi, paper.paper_number, log_rows, paper.title):
            continue
        selected.append(paper)

    if limit and limit > 0:
        selected = selected[:limit]
    return selected


def print_summary(summary: Summary) -> None:
    """Print the final download summary block."""
    print()
    print("==================================================")
    print("DOWNLOAD SUMMARY")
    print("==================================================")
    print()
    print(f"Total papers:       {summary.total}")
    print(f"Downloaded:         {summary.downloaded}")
    print(f"Already existed:    {summary.already_exists}")
    print(f"Not found total:    {summary.not_found}")
    print(f"  No OA available:  {summary.no_oa}")
    print(f"  Access required:  {summary.access_required}")
    print(f"  Invalid DOI:      {summary.invalid_doi}")
    print(f"  Failed:           {summary.failed}")
    if summary.skipped:
        print(f"Skipped (resume):   {summary.skipped}")
    print()
    print("==================================================")


def main(argv: Optional[list[str]] = None) -> int:
    """CLI entry point."""
    # Avoid Windows cp1252 crashes on console output.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001
        pass

    args = parse_args(argv)
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    ensure_directories()
    ensure_log_files()

    email = resolve_email(args.email)
    excel_path = Path(args.excel).expanduser().resolve()

    print("==================================================")
    print("Paper Downloader")
    print("==================================================")
    print()
    print(f"Excel: {excel_path}")
    print(f"Dry-run: {args.dry_run}")
    print(f"Limit: {args.limit or 'all'}")
    print(f"Retry failed: {args.retry_failed}")
    print("Date filter: NONE (all years are eligible)")
    print()

    df, columns = load_excel(excel_path)
    papers = dataframe_to_papers(df, columns)
    log_rows = load_download_log()

    print(f"Total papers: {len(papers)}")
    print()

    to_process = select_papers(
        papers,
        log_rows,
        retry_failed=args.retry_failed,
        limit=args.limit,
    )
    summary = Summary(total=len(papers), skipped=max(0, len(papers) - len(to_process)))
    if args.limit:
        summary.skipped = 0
        summary.total = len(to_process)

    if not to_process:
        print("Nothing to process (all matching papers already completed).")
        print_summary(summary)
        return 0

    session = build_session(email)
    results_for_report: list[tuple[PaperRecord, ProcessResult]] = []

    for index, paper in enumerate(to_process, start=1):
        print()
        print(f"[{index}/{len(to_process)}] ", end="")
        result = process_paper(
            paper,
            session,
            email,
            dry_run=args.dry_run,
            log_rows=log_rows,
        )
        results_for_report.append((paper, result))

        if result.status == "downloaded" and not args.dry_run:
            _update_summary(summary, "downloaded")
            log_rows.append(
                {
                    "doi": paper.doi,
                    "paper_number": str(paper.paper_number),
                    "title": paper.title,
                    "status": "downloaded",
                }
            )
        elif result.status == "downloaded" and args.dry_run:
            _update_summary(summary, "downloaded")
        elif result.status == "already_exists":
            _update_summary(summary, "already_exists")
            log_rows.append(
                {
                    "doi": paper.doi,
                    "paper_number": str(paper.paper_number),
                    "title": paper.title,
                    "status": "already_exists",
                }
            )
        else:
            _update_summary(summary, result.status)

    report_rows = [
        {
            "paper_number": paper.paper_number,
            "title": paper.title,
            "doi": paper.doi or paper.doi_raw,
            "source": paper.source,
            "year": paper.year,
            "status": result.status,
            "destination_folder": result.destination_folder,
            "filename": result.filename,
            "pdf_url": result.pdf_url,
            "publisher": result.publisher or paper.publisher,
            "error": result.error,
            "oa_source": result.oa_source,
        }
        for paper, result in results_for_report
    ]
    report_xlsx = LOGS_DIR / "report.xlsx"
    generate_excel_report(
        report_rows,
        summary.__dict__,
        report_xlsx,
        source_filename=excel_path.name,
        email=email,
        dry_run=args.dry_run,
    )
    print(f"\nExcel report written: {report_xlsx}")

    print_summary(summary)

    # Compact per-paper report for test runs.
    print()
    print("Per-paper results:")
    for paper, result in results_for_report:
        print(
            f"  #{paper.paper_number:03d} | {result.status:16} | "
            f"{result.destination_folder or '-':10} | "
            f"{result.oa_source or '-':10} | {result.filename or '-'}"
        )
        if result.error and result.status not in SUCCESS_STATUSES:
            print(f"           error: {result.error}")
        if result.pdf_url and result.status == "downloaded":
            print(f"           pdf:   {result.pdf_url}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
