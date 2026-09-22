"""
Shared fuzzy company/title dedup match, used by every scraper's
already_applied() (job_scraper.py, builtin_scraper.py, dynamite_scraper.py,
indeed_scraper.py). Pulled out because this exact logic was duplicated
byte-for-byte across all four files — a real fix from one copy never made it
to the others.

Each scraper still does its own company/title normalization before calling
this (job_scraper.py strips LinkedIn DOM artifacts and collapses whitespace;
the others just .lower().strip()) — those differences are real and
file-specific, so only the matching itself is shared here.

WHY THE MATCH IS SHAPED THIS WAY (2026-09-20): the original rule was "same
company + any 2 shared title words". For a Product Manager search every
title shares "senior product manager", so ANY second PM opening at a company
already in the sheet was silently treated as already-applied — checked live:
Upstart "Senior Product Manager, HELOC Decisioning" was dropped because the
sheet had "Senior Product Manager, ASPL" (a different role), and Salesforce/
BD/GitLab lost roles the same way on a single page. Generic role words carry
no identity; only the qualifier does ("HELOC Decisioning" vs "ASPL"). So the
match now compares the DISTINCTIVE words only.
"""

import re
from functools import lru_cache

# Words that describe the role family / seniority rather than identify a
# specific opening. Stripped before comparing titles.
_GENERIC = {
    "senior", "sr", "staff", "lead", "principal", "group", "junior", "jr",
    "associate", "director", "head", "vp", "vice", "president", "chief",
    "product", "products", "manager", "management", "mgr", "pm", "owner",
    "of", "the", "a", "an", "and", "for", "to", "in", "at", "on", "with",
    "remote", "us", "usa", "united", "states",
}

# Spelled-out forms so "Sr. PM, Growth" and "Senior Product Manager - Growth"
# normalize to the same words.
_SYNONYMS = {
    "sr": "senior", "jr": "junior", "mgr": "manager", "pm": "product manager",
    "&": "and", "+": "and",
}


@lru_cache(maxsize=8192)
def _tokens(title: str) -> tuple:
    t = re.sub(r"[^\w&+ ]+", " ", title.lower())
    words = []
    for w in t.split():
        words.extend(_SYNONYMS.get(w, w).split())
    return tuple(words)


def _distinctive(words) -> frozenset:
    return frozenset(w for w in words if w not in _GENERIC)


def pairs_match(company: str, title: str, applied_pairs: set) -> bool:
    """True if (company, title) — already normalized by the caller — matches
    an applied/skipped job.

    Match rules, in order, against every sheet entry at the same company
    (substring match either way, e.g. "upstart" ~ "upstart network"):
      1. exact (company, title) pair;
      2. titles identical once punctuation/synonyms are normalized
         ("Sr. PM - Growth" == "Senior Product Manager, Growth");
      3. the distinctive words of the shorter title are all contained in the
         longer one ("Growth" ⊆ "Growth and Retention") — a reworded repost
         keeps its distinctive core; a different role changes it.
    Two titles with NO distinctive words on either side (e.g. "Product
    Manager" vs "Senior Product Manager") are treated as different roles —
    the URL checks in each scraper still catch a true repost of those.
    """
    if (company, title) in applied_pairs:
        return True

    t_words = _tokens(title)
    t_dist  = _distinctive(t_words)

    for (ac, at) in applied_pairs:
        if not (ac and company and (ac in company or company in ac)):
            continue
        a_words = _tokens(at)
        if a_words == t_words:
            return True
        a_dist = _distinctive(a_words)
        if not t_dist or not a_dist:
            continue
        small, big = (t_dist, a_dist) if len(t_dist) <= len(a_dist) else (a_dist, t_dist)
        if small <= big:
            return True

    return False
