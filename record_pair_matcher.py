"""
compare_hotels_v3.py — offline match (v2 + postal-in-address rule) + ONLINE
verification via DuckDuckGo (DDGS), the same API used in the booking-URL script.

Per row (859):
  OFFLINE: v2 scoring + NEW rule: if either side's postal code appears inside
           the other side's address text, it counts as a postal match
           (fixes the Hotel Condado case: Zip=15893 vs address containing 32915).
  ONLINE:  2 DDGS queries (one per record, top 8 results). Signals:
             - same Booking.com hotel page in both result sets  -> Same
             - same official-site domain + name cross-match     -> Same
             - name cross-match + postal cross-match            -> Same
             - partial cross-match                             -> Review
             - nothing found                                   -> Unclear (offline stands)
  FINAL:   online "Same" wins; a partial online link upgrades Bad Match -> Review;
           otherwise the offline verdict stands.

Output: compare_hotels_online_result.xlsx with offline_*, online_* and final_* columns.
Resume-safe: rows already having final_verdict are skipped on re-run.
"""

import os
import re
import time
import random
import threading
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd

try:
    from rapidfuzz import fuzz

    def sim(a, b):
        if not a and not b:
            return 1.0
        if not a or not b:
            return 0.0
        return fuzz.token_sort_ratio(a, b) / 100.0

except ImportError:
    from difflib import SequenceMatcher

    def sim(a, b):
        if not a and not b:
            return 1.0
        if not a or not b:
            return 0.0
        a_sorted = " ".join(sorted(a.split()))
        b_sorted = " ".join(sorted(b.split()))
        return SequenceMatcher(None, a_sorted, b_sorted).ratio()

from ddgs import DDGS

# ============ CONFIG ============
INPUT_FILE = "compare.xlsx"
OUTPUT_FILE = "compare_verdict.xlsx"

EXPECTED_COLUMNS = [
    "name", "address1", "address2", "city", "country_a", "postal_code",
    "hotel_name_b", "address_b", "city_b", "state_b", "zip_b", "country_b",
]

SAME_THRESHOLD = 70
DIFFERENT_THRESHOLD = 35
DIFFERENT_LABEL = "Bad Match"
SAME_LABEL = "Same"
REVIEW_LABEL = "Review"

MAX_WORKERS = 5
RESULTS_PER_QUERY = 8
MIN_DELAY, MAX_DELAY = 0.4, 1.0
CHECKPOINT_EVERY = 25
# =================================

STREET_ABBR = {
    "street": "st", "road": "rd", "highway": "hwy", "avenue": "av",
    "boulevard": "blvd", "drive": "dr", "lane": "ln", "place": "pl",
    "square": "sq", "terrace": "tce", "close": "cl", "parkway": "pkwy",
    "circle": "cir", "court": "ct",
    "strasse": "str",
    "rue": "r", "impasse": "imp", "chemin": "ch",
    "calle": "c", "avenida": "av", "travessa": "tv",
    "via": "v", "viale": "vle", "corso": "cso", "piazza": "pz", "largo": "lgo",
    "gil": "g", "ro": "r", "dong": "d",
}

LEAD_WORDS_RE = re.compile(r"^(no|number|num|unit|flat|room|blk|block|apt|suite)\.?\s*")

GENERIC_HOTEL_WORDS = {
    "hotel", "hotels", "international", "inn", "resort", "resorts", "suite",
    "suites", "plaza", "garden", "gardens", "grand", "grande", "the", "and",
    "of", "spa", "spas", "lodge", "motel", "villa", "villas", "residence",
    "residences", "beach", "bay", "club", "palace", "tower", "towers",
    "city", "center", "centre", "park", "view", "boutique", "express",
    "holiday", "house", "guest", "guesthouse", "hostel", "pension",
    "ryokan", "bnb", "apartment", "apartments", "residency",
}

COUNTRY_ALIASES = {
    "usa": "US", "unitedstates": "US", "unitedstatesofamerica": "US",
    "america": "US", "uk": "GB", "unitedkingdom": "GB", "greatbritain": "GB",
    "england": "GB", "scotland": "GB", "wales": "GB",
    "uae": "AE", "unitedarabemirates": "AE",
    "southkorea": "KR", "korea": "KR", "republicofkorea": "KR",
    "deutschland": "DE", "germany": "DE",
    "france": "FR", "espana": "ES", "spain": "ES",
    "italia": "IT", "italy": "IT", "brasil": "BR", "brazil": "BR",
    "nederland": "NL", "netherlands": "NL", "holland": "NL",
    "russia": "RU", "nihon": "JP", "nippon": "JP", "japan": "JP",
    "zhongguo": "CN", "china": "CN", "turkiye": "TR", "turkey": "TR",
    "hellas": "GR", "greece": "GR", "polska": "PL", "poland": "PL",
    "cesko": "CZ", "czechia": "CZ", "czechrepublic": "CZ",
    "oesterreich": "AT", "austria": "AT", "schweiz": "CH",
    "switzerland": "CH", "suisse": "CH", "belgie": "BE", "belgique": "BE",
    "belgium": "BE", "sverige": "SE", "sweden": "SE", "norge": "NO",
    "norway": "NO", "danmark": "DK", "denmark": "DK", "suomi": "FI",
    "finland": "FI", "portugal": "PT", "ireland": "IE", "mexico": "MX",
    "canada": "CA", "australia": "AU", "newzealand": "NZ",
    "india": "IN", "bharat": "IN", "thailand": "TH", "vietnam": "VN",
    "malaysia": "MY", "indonesia": "ID", "philippines": "PH",
    "singapore": "SG", "hongkong": "HK", "taiwan": "TW",
    "saudiarabia": "SA", "egypt": "EG", "southafrica": "ZA",
    "argentina": "AR", "chile": "CL", "colombia": "CO", "peru": "PE",
}

BOOKING_HOTEL_RE = re.compile(
    r"https?://(?:www\.)?booking\.com/hotel/[a-z]{2,3}/[^\"'\s]+\.html", re.I
)
GENERIC_DOMAINS = {
    "google.com", "google.co.uk", "facebook.com", "instagram.com",
    "youtube.com", "wikipedia.org", "twitter.com", "x.com", "tiktok.com",
    "linkedin.com", "booking.com", "expedia.com", "hotels.com", "agoda.com",
    "kayak.com", "trivago.com", "skyscanner.com",
}

_thread_local = threading.local()
_print_lock = threading.Lock()


# ---------------- offline helpers (v2 + postal-in-address) ----------------

def normalize_text(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip().lower()
    if s in ("nan", "none", "0", ""):
        return ""
    s = s.replace("ß", "ss")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^\w\s-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_country(s):
    t = normalize_text(s)
    if not t:
        return ""
    key = re.sub(r"[^a-z]", "", t)
    if key in COUNTRY_ALIASES:
        return COUNTRY_ALIASES[key]
    if len(key) == 2:
        return key.upper()
    return t


def normalize_street(s):
    s = normalize_text(s)
    return " ".join(STREET_ABBR.get(w, w) for w in s.split())


def strip_house_number(s):
    s = normalize_street(s)
    s = LEAD_WORDS_RE.sub("", s)
    s = re.sub(r"^[\d\s\-/.,]+", "", s).strip()
    s = re.sub(r"[\d\s\-/.,]+$", "", s).strip()
    return s


def _expand(nums):
    nums = sorted(set(nums))
    if len(nums) == 2 and 0 < nums[1] - nums[0] <= 50:
        return set(range(nums[0], nums[1] + 1))
    return set(nums)


def extract_house_numbers(s):
    s = normalize_text(s)
    s = LEAD_WORDS_RE.sub("", s)
    nums = set()
    m = re.match(r"([\d][\d\s\-/.,]*)", s)
    if m:
        nums |= _expand(int(x) for x in re.findall(r"\d+", m.group(1)))
    else:
        m2 = re.search(r"([\d][\d\s\-/.,]*)\s*$", s)
        if m2:
            nums |= _expand(int(x) for x in re.findall(r"\d+", m2.group(1)))
    return nums


def house_number_overlap(addr_a, addr_b):
    nums_a = extract_house_numbers(addr_a)
    nums_b = extract_house_numbers(addr_b)
    if not nums_a or not nums_b:
        return None
    return bool(nums_a & nums_b)


def address_contained(addr_a, addr_b):
    a, b = normalize_street(addr_a), normalize_street(addr_b)
    if len(a) >= 10 and len(b) >= 10 and (a in b or b in a):
        return True
    sa, sb = strip_house_number(addr_a), strip_house_number(addr_b)
    if len(sa) >= 8 and len(sb) >= 8 and (sa in sb or sb in sa):
        return True
    return False


def distinctive_tokens(name, city):
    words = set(normalize_text(name).split())
    words -= GENERIC_HOTEL_WORDS
    if city:
        words -= set(normalize_text(city).split())
    return words


def jaccard(a, b):
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _alnum(s):
    return re.sub(r"[^a-z0-9]", "", normalize_text(s))


def postal_match(p1, p2, addr_a="", addr_b=""):
    """True / False / None. NEW: a postal code embedded in the other side's
    address text counts as a match (Hotel Condado rule)."""
    d1, d2 = _alnum(p1), _alnum(p2)
    if not d1 or not d2:
        return None
    if d1 == d2:
        return True
    ta, tb = _alnum(addr_a), _alnum(addr_b)
    if len(d1) >= 4 and d1 in tb:
        return True
    if len(d2) >= 4 and d2 in ta:
        return True
    lo, hi = (d1, d2) if len(d1) <= len(d2) else (d2, d1)
    if len(lo) >= 4 and hi.startswith(lo):
        return True
    return False


def postal_match_via_address(p1, p2, addr_a, addr_b):
    """Did the match come from the address-embedding rule (not the Zip field)?"""
    d1, d2 = _alnum(p1), _alnum(p2)
    if not d1 or not d2 or d1 == d2:
        return False
    ta, tb = _alnum(addr_a), _alnum(addr_b)
    return (len(d1) >= 4 and d1 in tb) or (len(d2) >= 4 and d2 in ta)


def _bad(notes):
    return {
        "verdict": DIFFERENT_LABEL, "score": 0,
        "country_match": False, "postal_match": None,
        "house_number_match": None, "address_contained": False,
        "street_similarity": None, "city_similarity": None,
        "name_similarity": None, "brand_similarity": None,
        "notes": notes,
    }


def compare_offline(row):
    name_a = normalize_text(row["name"])
    addr_a_full = " ".join(
        p for p in [normalize_text(row["address1"]), normalize_text(row["address2"])] if p
    )
    city_a = normalize_text(row["city"])
    country_a = normalize_country(row["country_a"])
    postal_a = row["postal_code"]

    name_b = normalize_text(row["hotel_name_b"])
    addr_b_full = normalize_text(row["address_b"])
    city_b = normalize_text(row["city_b"])
    country_b = normalize_country(row["country_b"])
    postal_b = row["zip_b"]

    notes = []

    if country_a and country_b and country_a != country_b:
        return _bad(f"Country mismatch: '{country_a}' vs '{country_b}'")

    p_match = postal_match(postal_a, postal_b, addr_a_full, addr_b_full)
    h_match = house_number_overlap(addr_a_full, addr_b_full)
    street_sim = sim(strip_house_number(addr_a_full), strip_house_number(addr_b_full))
    city_sim = sim(city_a, city_b)
    name_sim = sim(name_a, name_b)
    contained = address_contained(addr_a_full, addr_b_full)

    dist_a = distinctive_tokens(name_a, city_a)
    dist_b = distinctive_tokens(name_b, city_b)
    brand_sim = jaccard(dist_a, dist_b) if (dist_a and dist_b) else name_sim

    score = 50.0
    if p_match is True:
        score += 30
        if postal_match_via_address(postal_a, postal_b, addr_a_full, addr_b_full):
            notes.append("postal code found inside the other address")
        else:
            notes.append("postal codes match")
    elif p_match is False:
        score -= 30
        notes.append("postal codes differ")
    else:
        notes.append("postal code unknown on one/both sides")

    if h_match is True:
        score += 20
        notes.append("house numbers overlap")
    elif h_match is False:
        score -= 20
        notes.append("house numbers differ, no overlap")
    else:
        notes.append("house number unknown on one/both sides")

    if contained:
        score += 15
        notes.append("one address contains the other")
    if name_a and name_a == name_b:
        score += 10
        notes.append("names identical after normalization")

    score += (street_sim - 0.5) * 30
    score += (city_sim - 0.5) * 10
    score += (brand_sim - 0.5) * 10

    if dist_a and dist_b and brand_sim <= 0.34 and name_sim > 0.6:
        notes.append(
            "same chain/brand name but distinct branch identifiers "
            f"({sorted(dist_a)} vs {sorted(dist_b)}) — likely a different location"
        )

    score = max(0, min(100, round(score, 1)))
    verdict = SAME_LABEL if score >= SAME_THRESHOLD else (
        DIFFERENT_LABEL if score <= DIFFERENT_THRESHOLD else REVIEW_LABEL)

    return {
        "verdict": verdict, "score": score,
        "country_match": (country_a == country_b) if (country_a and country_b) else None,
        "postal_match": p_match, "house_number_match": h_match,
        "address_contained": contained,
        "street_similarity": round(street_sim, 2),
        "city_similarity": round(city_sim, 2),
        "name_similarity": round(name_sim, 2),
        "brand_similarity": round(brand_sim, 2),
        "notes": "; ".join(notes),
    }

# ---------------- online verification (DDGS) ----------------

def get_ddgs():
    if not hasattr(_thread_local, "ddgs"):
        _thread_local.ddgs = DDGS()
    return _thread_local.ddgs


def safe_print(*args, **kwargs):
    with _print_lock:
        print(*args, **kwargs)


def qpart(v):
    s = "" if v is None else str(v).strip()
    return "" if s.lower() in ("", "nan", "none", "0") else s


def ddg_text(ddgs, query):
    for _ in range(2):
        try:
            return list(ddgs.text(query, max_results=RESULTS_PER_QUERY))
        except Exception as e:
            msg = str(e).lower()
            time.sleep(random.uniform(4, 8) if ("rate" in msg or "429" in msg or "202" in msg)
                       else random.uniform(1, 2))
    return []


def clean_booking_url(url):
    if not url:
        return ""
    m = BOOKING_HOTEL_RE.search(url)
    return m.group(0).lower() if m else ""


def domain_of(url):
    m = re.search(r"https?://([^/]+)", url or "")
    if not m:
        return ""
    parts = m.group(1).lower().split(":")[0].split(".")
    return ".".join(parts[-2:]) if len(parts) >= 2 else ""


def extract_info(results):
    booking, domains, texts = set(), [], []
    for r in results:
        url = r.get("href") or r.get("url") or ""
        title = r.get("title") or ""
        body = r.get("body") or r.get("snippet") or ""
        cu = clean_booking_url(url)
        if cu:
            booking.add(cu)
        domains.append(domain_of(url))
        texts.append(f"{title} {body}")
    blob = " ".join(texts).lower()
    return {
        "booking": booking,
        "top_domains": domains[:3],
        "text": blob,
        "alnum": re.sub(r"[^a-z0-9]", "", blob),
    }


def token_hit(tokens, text):
    tokens = [t for t in tokens if len(t) >= 4]
    if not tokens:
        return 0.0
    return sum(1 for t in tokens if t in text) / len(tokens)


def online_check(row):
    ddgs = get_ddgs()
    nA, cA, coA = qpart(row["name"]), qpart(row["city"]), qpart(row["country_a"])
    nB, cB, coB = qpart(row["hotel_name_b"]), qpart(row["city_b"]), qpart(row["country_b"])
    pA, pB = _alnum(row["postal_code"]), _alnum(row["zip_b"])

    qA = " ".join(p for p in [nA, cA, coA] if p) + " hotel"
    qB = " ".join(p for p in [nB, cB, coB] if p) + " hotel"

    infoA = extract_info(ddg_text(ddgs, qA))
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))
    infoB = extract_info(ddg_text(ddgs, qB))

    shared_booking = infoA["booking"] & infoB["booking"]
    if shared_booking:
        return {"online_verdict": "Same",
                "online_evidence": f"both resolve to {sorted(shared_booking)[0]}"}

    domA = {d for d in infoA["top_domains"] if d and d not in GENERIC_DOMAINS}
    domB = {d for d in infoB["top_domains"] if d and d not in GENERIC_DOMAINS}
    shared_dom = domA & domB

    tA = distinctive_tokens(normalize_text(row["name"]), normalize_text(row["city"]))
    tB = distinctive_tokens(normalize_text(row["hotel_name_b"]), normalize_text(row["city_b"]))
    crossA = token_hit(tA, infoB["text"])  # A's name tokens in B's results
    crossB = token_hit(tB, infoA["text"])  # B's name tokens in A's results
    cross_postal = ((len(pA) >= 4 and pA in infoB["alnum"]) or
                    (len(pB) >= 4 and pB in infoA["alnum"]))

    if shared_dom and (crossA >= 0.5 or crossB >= 0.5):
        return {"online_verdict": "Same",
                "online_evidence": f"same site {sorted(shared_dom)[0]} + name cross-match"}
    if (crossA >= 0.5 or crossB >= 0.5) and cross_postal:
        return {"online_verdict": "Same",
                "online_evidence": "name + postal cross-match in search results"}
    if crossA >= 0.5 or crossB >= 0.5 or cross_postal:
        bits = []
        if crossA >= 0.5 or crossB >= 0.5:
            bits.append("name cross-match")
        if cross_postal:
            bits.append("postal cross-match")
        return {"online_verdict": "Review",
                "online_evidence": "partial online link: " + ", ".join(bits)}
    return {"online_verdict": "Unclear",
            "online_evidence": "no clear online link between the two records"}


def combine(off, on):
    ov, fv = on["online_verdict"], off["verdict"]
    if ov == "Same":
        return "Same", "confirmed online: " + on["online_evidence"]
    if ov == "Review":
        if fv == "Bad Match":
            return "Review", "offline said Bad Match but partial online link: " + on["online_evidence"]
        return fv, off["notes"] + " | online: partial link (" + on["online_evidence"] + ")"
    return fv, off["notes"] + " | online: no clear link found"


def process_one(idx, row_dict, pos, total):
    off = compare_offline(row_dict)
    on = online_check(row_dict)
    final_v, final_n = combine(off, on)
    safe_print(f"[{pos}/{total}] {str(row_dict.get('name'))[:40]!r} -> {final_v} (off:{off['verdict']}, on:{on['online_verdict']})")
    return idx, {
        "offline_verdict": off["verdict"], "offline_score": off["score"],
        "country_match": off["country_match"], "postal_match": off["postal_match"],
        "house_number_match": off["house_number_match"],
        "address_contained": off["address_contained"],
        "street_similarity": off["street_similarity"],
        "city_similarity": off["city_similarity"],
        "name_similarity": off["name_similarity"],
        "brand_similarity": off["brand_similarity"],
        "offline_notes": off["notes"],
        "online_verdict": on["online_verdict"],
        "online_evidence": on["online_evidence"],
        "final_verdict": final_v, "final_notes": final_n,
    }


def build_output(df_in, results):
    rows = []
    for i in range(len(df_in)):
        base = df_in.iloc[i].to_dict()
        base["row_id"] = i
        base.update(results.get(i, {}))
        rows.append(base)
    return pd.DataFrame(rows)


def main():
    df_in = pd.read_excel(INPUT_FILE)
    if len(df_in.columns) < 12:
        raise ValueError(f"Expected 12 columns (A-L), found {len(df_in.columns)}.")
    df_in.columns = EXPECTED_COLUMNS + list(df_in.columns[len(EXPECTED_COLUMNS):])

    results = {}
    if os.path.exists(OUTPUT_FILE):
        prev = pd.read_excel(OUTPUT_FILE)
        for _, r in prev.iterrows():
            if pd.notna(r.get("final_verdict")):
                d = r.to_dict()
                d.pop("row_id", None)
                # drop original input cols (re-added from df_in on rebuild)
                for c in list(d.keys()):
                    if c in EXPECTED_COLUMNS:
                        d.pop(c, None)
                results[int(r["row_id"])] = d
        safe_print(f"Resuming: {len(results)} rows already done.")

    todo = [i for i in range(len(df_in)) if i not in results]
    safe_print(f"{len(df_in)} rows total, {len(todo)} to process "
               f"(workers={MAX_WORKERS}, {RESULTS_PER_QUERY} results/query).\n")

    def checkpoint():
        build_output(df_in, results).to_excel(OUTPUT_FILE, index=False)

    completed = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(process_one, i, df_in.iloc[i].to_dict(), pos, len(todo)): i
                   for pos, i in enumerate(todo, start=1)}
        for fut in as_completed(futures):
            idx, res = fut.result()
            results[idx] = res
            completed += 1
            if completed % CHECKPOINT_EVERY == 0:
                checkpoint()
                safe_print(f"  (checkpoint: {completed}/{len(todo)})")

    out = build_output(df_in, results)
    out.to_excel(OUTPUT_FILE, index=False)
    print(f"\nDone. {len(out)} rows -> {OUTPUT_FILE}\n")
    print("Final verdicts:")
    print(out["final_verdict"].value_counts().to_string())
    print("\nOffline verdicts:")
    print(out["offline_verdict"].value_counts().to_string())
    print("\nOnline verdicts:")
    print(out["online_verdict"].value_counts().to_string())


if __name__ == "__main__":
    main()
