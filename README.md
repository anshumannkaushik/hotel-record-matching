# Hotel Record Matching

Verdict engine for deduplicating hotel records — given spreadsheet pairs of hotel records (name, address, city, country, phone, postal code), decide whether each pair is the **Same** property, a **Bad Match**, or needs **Review**.

## How it works

Two-stage scoring, built over thousands of real record pairs:

1. **Offline scoring** (`record_pair_matcher.py`, `offline_verdict_check.py`)
   - Normalises names and addresses (unicode folding, punctuation, chain/brand aliases, country-code mapping for 140+ country spellings)
   - Scores name similarity, address tokens, city/country agreement and postal codes — including the rule that a postal code appearing *inside* the other record's address text counts as a postal match
   - Applies an address-first rule: when the address matches, a name difference (rebrand, misspelling, different writing style) still counts as Same; Bad Match requires name **and** address to differ

2. **Online verification** — for pairs the offline stage can't settle, both records are searched (DuckDuckGo) and cross-matched: the same property page in both result sets, or name + postal cross-matches, upgrade the verdict; partial evidence routes the pair to Review instead of guessing.

Resume-safe: rows that already carry a final verdict are skipped on re-run, so multi-thousand-row batches can be processed in stages.

## Stack

Python · pandas · DuckDuckGo search (ddgs) · fuzzy string scoring
