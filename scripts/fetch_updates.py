"""
Fetch and parse What's New pages from Microsoft Learn for:
  - Windows 365 Enterprise & Frontline
  - Windows 365 Business
  - Windows 365 Link
  - Windows App

Outputs a structured data.json consumed by the aggregator page.
"""

import json
import re
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup


def normalize_text(text: str) -> str:
    """Replace problematic Unicode characters with plain ASCII equivalents."""
    replacements = {
        "\u2011": "-",   # non-breaking hyphen
        "\u2010": "-",   # hyphen
        "\u2012": "-",   # figure dash
        "\u2013": "-",   # en dash
        "\u2014": "-",   # em dash
        "\u2015": "-",   # horizontal bar
        "\u2018": "'",   # left single quote
        "\u2019": "'",   # right single quote
        "\u201c": '"',   # left double quote
        "\u201d": '"',   # right double quote
        "\u2026": "...", # ellipsis
        "\u00a0": " ",   # non-breaking space
        "\u200b": "",    # zero-width space
        "\u200e": "",    # left-to-right mark
        "\u200f": "",    # right-to-left mark
        "\ufeff": "",    # BOM / zero-width no-break space
        "\u202f": " ",   # narrow no-break space
    }
    for char, repl in replacements.items():
        text = text.replace(char, repl)
    return text

SOURCES = {
    "enterprise": {
        "url": "https://learn.microsoft.com/en-us/windows-365/enterprise/whats-new",
        "label": "Enterprise & Frontline",
    },
    "business": {
        "url": "https://learn.microsoft.com/en-us/windows-365/business/whats-new",
        "label": "Business",
    },
    "link": {
        "url": "https://learn.microsoft.com/en-us/windows-365/link/whats-new",
        "label": "Windows 365 Link",
    },
    "windowsapp": {
        "url": "https://learn.microsoft.com/en-us/windows-app/whats-new",
        "label": "Windows App",
    },
}

WEEK_RE = re.compile(
    r"Week\s+of\s+(\w+\s+\d{1,2},?\s+\d{4})", re.IGNORECASE
)

DATE_PUBLISHED_RE = re.compile(
    r"(?:Date published|Public availability date|Insider availability date):\s*(\w+\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})",
    re.IGNORECASE,
)

# Map month names to integers for robust date parsing
MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}


def parse_week_date(text: str) -> str | None:
    """Extract a YYYY-MM-DD string from a 'Week of ...' heading."""
    m = WEEK_RE.search(text)
    if not m:
        return None
    raw = m.group(1).replace(",", "")
    parts = raw.split()
    if len(parts) != 3:
        return None
    month_name, day_str, year_str = parts
    month = MONTHS.get(month_name.lower())
    if not month:
        return None
    try:
        return f"{int(year_str):04d}-{month:02d}-{int(day_str):02d}"
    except ValueError:
        return None


def parse_published_date(text: str) -> str | None:
    """Extract a YYYY-MM-DD from a 'Date published: ...' or 'Public availability date: ...' string."""
    m = DATE_PUBLISHED_RE.search(text)
    if not m:
        return None
    raw = m.group(1).replace(",", "")
    # Strip ordinal suffixes (1st, 2nd, 3rd, 10th, etc.)
    raw = re.sub(r"(\d+)(?:st|nd|rd|th)", r"\1", raw)
    parts = raw.split()
    if len(parts) != 3:
        return None
    month_name, day_str, year_str = parts
    month = MONTHS.get(month_name.lower())
    if not month:
        return None
    try:
        return f"{int(year_str):04d}-{month:02d}-{int(day_str):02d}"
    except ValueError:
        return None


def detect_tags(title: str, desc: str) -> list[str]:
    """Heuristic tagging based on title/description keywords."""
    combined = (title + " " + desc).lower()
    tags = []
    if "generally available" in combined or "moved out of preview" in combined or "is now ga" in combined:
        tags.append("ga")
    if "preview" in combined and "ga" not in tags:
        tags.append("preview")
    if "frontline" in combined:
        tags.append("frontline")
    if "government" in combined or "gcc" in combined:
        tags.append("government")
    return tags


def detect_category(title: str, desc: str) -> str:
    """Heuristic category detection from text."""
    combined = (title + " " + desc).lower()
    keywords = {
        "Provisioning": ["provisioning", "provision", "gallery image", "custom image"],
        "Device Security": ["security", "lockbox", "customer key", "capture protection", "compliance", "conditional access"],
        "Networking": ["rdp shortpath", "network connection", "udp", "stun", "turn", "captive portal", "shortpath"],
        "Monitor & Troubleshoot": ["report", "alert", "troubleshoot", "utilization", "copilot in intune", "diagnostics", "health check"],
        "Apps": ["teams", "cloud apps", "app ", "mmr", "multimedia"],
        "End User Experience": ["user experience", "windows app", "switch", "boot", "restore experience", "connection center", "split screen"],
        "Regions": ["region", "geography", "now available in"],
        "Authentication": ["sign-in", "sso", "single sign-on", "wam", "authentication"],
        "Build Update": ["quality update", "build number", "bug fixes and"],
        "Documentation": ["documentation", "new article", "new help"],
        "Accessibility": ["voice access", "accessibility"],
    }
    for cat, kws in keywords.items():
        if any(kw in combined for kw in kws):
            return cat
    return "Device Management"


def collect_description_paragraphs(element) -> str:
    """Walk siblings after a heading and collect paragraph text until the next heading."""
    parts = []
    for sib in element.next_siblings:
        if sib.name and sib.name.startswith("h"):
            break
        if sib.name == "p":
            parts.append(sib.get_text(strip=True))
        if sib.name in ("ul", "ol"):
            for li in sib.find_all("li", recursive=False):
                parts.append(li.get_text(strip=True))
    return " ".join(parts)


def fetch_and_parse(source_key: str, url: str) -> list[dict]:
    """Fetch one What's New page and return a list of announcement dicts."""
    print(f"  Fetching {source_key}: {url}")
    resp = requests.get(url, timeout=30, headers={"User-Agent": "W365-WhatsNew-Aggregator/1.0"})
    resp.raise_for_status()
    resp.encoding = "utf-8"  # Force UTF-8 — Learn pages are UTF-8 but headers may claim otherwise
    soup = BeautifulSoup(resp.text, "html.parser")

    # Find the main content area
    content = soup.select_one("#main-column") or soup.select_one("main") or soup
    headings = content.find_all(re.compile(r"^h[2-4]$"))

    entries = []
    current_date = None

    for h in headings:
        text = h.get_text(strip=True)

        # Check for week heading
        week_date = parse_week_date(text)
        if week_date:
            current_date = week_date
            continue

        level = int(h.name[1])

        # For Windows App, the structure is different — version headings are h3
        # with dates in the following paragraph instead of "Week of" headings
        if source_key == "windowsapp":
            if level == 3:
                # Extract date from the first following <p>
                pub_date = None
                desc_parts = []
                for sib in h.next_siblings:
                    if sib.name and sib.name.startswith("h"):
                        break
                    if sib.name == "p":
                        p_text = sib.get_text(strip=True)
                        d = parse_published_date(p_text)
                        if d and not pub_date:
                            pub_date = d
                            continue  # skip the date paragraph from description
                        desc_parts.append(p_text)
                    if sib.name in ("ul", "ol"):
                        for li in sib.find_all("li", recursive=False):
                            desc_parts.append(li.get_text(strip=True))
                # Also try parsing the heading itself as a date (e.g. "January 21, 2026")
                if not pub_date:
                    pub_date = parse_week_date("Week of " + text) or parse_published_date("Date published: " + text)
                if not pub_date:
                    continue
                desc = " ".join(desc_parts)
                if not desc:
                    continue
                if len(desc) > 400:
                    desc = desc[:397] + "..."
                entries.append({
                    "date": pub_date,
                    "source": source_key,
                    "title": normalize_text(text),
                    "desc": normalize_text(desc),
                    "category": detect_category(text, desc),
                    "tags": detect_tags(text, desc),
                })
            continue

        if not current_date:
            continue

        # For other sources, feature announcements are typically h3/h4
        # Skip very generic sub-section titles
        generic_sections = {
            "device management", "device provisioning", "device security",
            "provisioning", "apps", "documentation", "miscellaneous",
            "monitor and troubleshoot", "end user experience",
            "end-user experience", "role-based access control",
            "windows 365 app", "windows 365 frontline", "partners",
            "windows 365 government", "government community cloud",
            "windows app", "windows 365 boot updates",
            "copilot in intune for windows 365", "device security",
        }
        if text.lower().strip() in generic_sections:
            continue

        desc = collect_description_paragraphs(h)
        if not desc and level < 4:
            # Might be a section heading; look for child headings instead
            continue
        if not desc:
            desc = text  # fallback
            continue

        if len(desc) > 400:
            desc = desc[:397] + "..."

        entries.append({
            "date": current_date,
            "source": source_key,
            "title": normalize_text(text),
            "desc": normalize_text(desc),
            "category": detect_category(text, desc),
            "tags": detect_tags(text, desc),
        })

    return entries


def main():
    all_entries = []
    for key, info in SOURCES.items():
        try:
            entries = fetch_and_parse(key, info["url"])
            print(f"    -> {len(entries)} announcements parsed")
            all_entries.extend(entries)
        except Exception as e:
            print(f"  ERROR fetching {key}: {e}", file=sys.stderr)

    # Deduplicate by (date, source, title)
    seen = set()
    unique = []
    for e in all_entries:
        key = (e["date"], e["source"], e["title"])
        if key not in seen:
            seen.add(key)
            unique.append(e)

    # Sort by date descending
    unique.sort(key=lambda x: x["date"], reverse=True)

    out_path = Path(__file__).resolve().parent.parent / "data.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "lastUpdated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "announcements": unique,
            },
            f,
            indent=2,
            ensure_ascii=False,
        )
    print(f"\nWrote {len(unique)} announcements to {out_path}")


if __name__ == "__main__":
    main()
