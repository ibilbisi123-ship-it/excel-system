import os
import re
import sqlite3
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Tuple

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()


def _resolve_db_path() -> str:
    override = os.environ.get("PRICE_DB")
    if override:
        return override
    return os.path.join(BASE_DIR, "my2.db")


DB_PATH = _resolve_db_path()
TABLE_NAME = os.environ.get("PRICE_TABLE", "gpt_purchase_data_(1)")
DESC_COLUMN = os.environ.get("PRICE_DESC_COL", "Description")
PRICE_COLUMN = os.environ.get("PRICE_VALUE_COL", "price")
MAX_CANDIDATES = int(os.environ.get("PRICE_MAX_CANDIDATES", "18"))
MIN_CANDIDATE_SCORE = float(os.environ.get("PRICE_MIN_SCORE", "0.18"))
MIN_RESULTS_BEFORE_THRESHOLD = int(os.environ.get("PRICE_MIN_RESULTS_BEFORE_THRESHOLD", "5"))

TOKEN_PATTERN = re.compile(r"[a-z0-9]+(?:[/-][a-z0-9]+)*")
NUMBER_PATTERN = re.compile(r"\d+(?:\.\d+)?")
STOPWORDS = {
    "and",
    "or",
    "the",
    "to",
    "with",
    "without",
    "for",
    "in",
    "of",
    "a",
    "an",
    "per",
    "each",
    "pcs",
    "pc",
    "qty",
    "quantity",
    "unit",
}


def normalize_description(text: Any) -> str:
    if text is None:
        return ""
    return " ".join(str(text).lower().split())


def extract_tokens(text: Any) -> List[str]:
    if not text:
        return []
    raw = str(text).lower()
    base_tokens: List[str] = []
    for match in TOKEN_PATTERN.finditer(raw):
        token = match.group()
        if token in STOPWORDS:
            continue
        base_tokens.append(token)
        if "-" in token or "/" in token:
            for part in re.split(r"[-/]", token):
                if part and part not in STOPWORDS:
                    base_tokens.append(part)
    enriched: List[str] = []
    for token in base_tokens:
        if token in STOPWORDS:
            continue
        enriched.append(token)
        alnum = re.match(r"^(\d+)([a-z]+)$", token)
        if alnum:
            enriched.append(alnum.group(1))
            suffix = alnum.group(2)
            if suffix and suffix not in STOPWORDS:
                enriched.append(suffix)
        else:
            rev = re.match(r"^([a-z]+)(\d+)$", token)
            if rev:
                prefix = rev.group(1)
                suffix = rev.group(2)
                if prefix and prefix not in STOPWORDS:
                    enriched.append(prefix)
                enriched.append(suffix)
    for idx in range(len(base_tokens) - 1):
        first = base_tokens[idx]
        second = base_tokens[idx + 1]
        if first.isdigit() and second.isalpha() and second not in STOPWORDS:
            enriched.append(f"{first}{second}")
        if second.isdigit() and first.isalpha() and first not in STOPWORDS:
            enriched.append(f"{first}{second}")
    unique: List[str] = []
    for token in enriched:
        if token and token not in unique and token not in STOPWORDS:
            unique.append(token)
    return unique


def extract_numbers(text: Any) -> List[str]:
    if not text:
        return []
    numbers = NUMBER_PATTERN.findall(str(text))
    unique: List[str] = []
    for num in numbers:
        if num not in unique:
            unique.append(num)
    return unique


def compute_token_metrics(query_tokens: set, candidate_tokens: set) -> Tuple[float, float]:
    if not query_tokens or not candidate_tokens:
        return 0.0, 0.0
    intersection = query_tokens & candidate_tokens
    union = query_tokens | candidate_tokens
    jaccard = len(intersection) / len(union)
    coverage = len(intersection) / len(query_tokens)
    return jaccard, coverage


def compute_number_overlap(query_numbers: set, candidate_numbers: set) -> float:
    if not query_numbers or not candidate_numbers:
        return 0.0
    intersection = query_numbers & candidate_numbers
    return len(intersection) / len(query_numbers)


def compute_similarity_score(query_info: Dict[str, Any], row_info: Dict[str, Any]) -> Tuple[float, Dict[str, float]]:
    seq_score = SequenceMatcher(None, query_info["normalized"], row_info["normalized"]).ratio()
    token_jaccard, token_coverage = compute_token_metrics(query_info["token_set"], row_info["token_set"])
    number_overlap = compute_number_overlap(query_info["number_set"], row_info["number_set"])
    bonus = 0.0
    if query_info["normalized"] and query_info["normalized"] in row_info["normalized"]:
        bonus += 0.05
    if token_coverage >= 0.75:
        bonus += 0.05
    if number_overlap >= 0.5:
        bonus += 0.05
    if seq_score >= 0.85:
        bonus += 0.05
    base = (seq_score * 0.4) + (token_jaccard * 0.3) + (token_coverage * 0.2) + (number_overlap * 0.1)
    score = min(1.0, base + bonus)
    components = {
        "text": seq_score,
        "token_jaccard": token_jaccard,
        "token_coverage": token_coverage,
        "number": number_overlap,
        "bonus": bonus,
    }
    return score, components


def quote_identifier(identifier: str) -> str:
    escaped = identifier.replace('"', '""')
    return f'"{escaped}"'


def _to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    cleaned = text.replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def format_price(value: Any) -> str:
    numeric = _to_float(value)
    if numeric is None:
        return ""
    return f"{numeric:.2f}"


def load_database_rows() -> List[Dict[str, Any]]:
    if not os.path.exists(DB_PATH):
        return []
    conn: Optional[sqlite3.Connection] = None
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        table_expr = quote_identifier(TABLE_NAME)
        desc_expr = quote_identifier(DESC_COLUMN)
        price_expr = quote_identifier(PRICE_COLUMN)
        query = (
            f"SELECT {desc_expr} AS description, {price_expr} AS price "
            f"FROM {table_expr} WHERE {desc_expr} IS NOT NULL"
        )
        rows = conn.execute(query).fetchall()
        data: List[Dict[str, Any]] = []
        for row in rows:
            description = row["description"]
            normalized = normalize_description(description)
            tokens = extract_tokens(description)
            numbers = extract_numbers(description)
            price_value = _to_float(row["price"])
            data.append(
                {
                    "description": description,
                    "price_raw": row["price"],
                    "price_value": price_value,
                    "price_formatted": format_price(row["price"]),
                    "normalized": normalized,
                    "tokens": tokens,
                    "token_set": set(tokens),
                    "numbers": numbers,
                    "number_set": set(numbers),
                }
            )
        return data
    except sqlite3.Error:
        return []
    finally:
        if conn:
            conn.close()


DATABASE_ROWS: List[Dict[str, Any]] = load_database_rows()


def reload_price_cache() -> int:
    global DATABASE_ROWS
    DATABASE_ROWS = load_database_rows()
    return len(DATABASE_ROWS)


def get_candidate_rows(description: str, limit: int = MAX_CANDIDATES) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    query_str = "" if description is None else str(description)
    query_info: Dict[str, Any] = {
        "original": query_str,
        "normalized": normalize_description(query_str),
    }
    tokens = extract_tokens(query_str)
    if not tokens and query_info["normalized"]:
        tokens = query_info["normalized"].split()
    numbers = extract_numbers(query_str)
    query_info["tokens"] = tokens
    query_info["token_set"] = set(tokens)
    query_info["numbers"] = numbers
    query_info["number_set"] = set(numbers)

    if not query_str.strip() or not DATABASE_ROWS:
        return [], query_info

    scored: List[Tuple[float, Dict[str, Any], Dict[str, float], List[str], List[str]]] = []
    for row in DATABASE_ROWS:
        score, components = compute_similarity_score(query_info, row)
        shared_tokens = [tok for tok in row["tokens"] if tok in query_info["token_set"]][:20]
        shared_numbers = [num for num in row["numbers"] if num in query_info["number_set"]][:20]
        scored.append((score, row, components, shared_tokens, shared_numbers))

    scored.sort(key=lambda item: item[0], reverse=True)

    candidates: List[Dict[str, Any]] = []
    for score, row, components, shared_tokens, shared_numbers in scored:
        if score < MIN_CANDIDATE_SCORE and len(candidates) >= MIN_RESULTS_BEFORE_THRESHOLD:
            break
        candidates.append(
            {
                "description": row["description"],
                "price_formatted": row["price_formatted"],
                "price_value": row["price_value"],
                "normalized": row["normalized"],
                "tokens": row["tokens"],
                "numbers": row["numbers"],
                "shared_tokens": shared_tokens,
                "shared_numbers": shared_numbers,
                "score": round(score, 3),
                "score_details": (
                    f"text={components['text']:.3f}, token_jaccard={components['token_jaccard']:.3f}, "
                    f"coverage={components['token_coverage']:.3f}, numbers={components['number']:.3f}, bonus={components['bonus']:.3f}"
                ),
            }
        )
        if len(candidates) >= limit:
            break

    if not candidates and scored:
        score, row, components, shared_tokens, shared_numbers = scored[0]
        candidates.append(
            {
                "description": row["description"],
                "price_formatted": row["price_formatted"],
                "price_value": row["price_value"],
                "normalized": row["normalized"],
                "tokens": row["tokens"],
                "numbers": row["numbers"],
                "shared_tokens": shared_tokens,
                "shared_numbers": shared_numbers,
                "score": round(score, 3),
                "score_details": (
                    f"text={components['text']:.3f}, token_jaccard={components['token_jaccard']:.3f}, "
                    f"coverage={components['token_coverage']:.3f}, numbers={components['number']:.3f}, bonus={components['bonus']:.3f}"
                ),
            }
        )

    return candidates[:limit], query_info


def _best_price_from_candidates(candidates: List[Dict[str, Any]]) -> Any:
    for cand in candidates:
        val = cand.get("price_value")
        if isinstance(val, (int, float)):
            return round(float(val), 2)
    for cand in candidates:
        price = cand.get("price_formatted")
        if price:
            return price
    return ""


def get_model_price(description: str) -> Any:
    if description is None:
        return ""
    candidates, _ = get_candidate_rows(description)
    if not candidates:
        return ""
    best = _best_price_from_candidates(candidates)
    if best not in ("", None):
        return best
    lowered = description.strip().lower()
    for row in DATABASE_ROWS:
        row_desc = str(row.get("description", "")).strip().lower()
        if row_desc == lowered:
            value = row.get("price_value")
            if isinstance(value, (int, float)):
                return round(float(value), 2)
            return row.get("price_formatted", "")
    return ""
