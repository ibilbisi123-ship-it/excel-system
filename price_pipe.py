import sys
import json

try:
    from pricesqlsystem import get_candidate_rows
except Exception as exc:
    sys.stdout.write(json.dumps({"fatal": True, "error": f"ImportError: {exc}"}) + "\n")
    sys.stdout.flush()
    sys.exit(1)


def format_val(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return f"{v:.2f}"
    return str(v)

def main() -> None:
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            payload = json.loads(line)
        except Exception:
            continue
        req_id = payload.get("id")
        description = payload.get("description", "")
        
        candidates_list = []
        try:
            limit = int(payload.get("limit", 5))
            cands, _ = get_candidate_rows(str(description), limit=limit)
            for c in cands:
                # c has 'description', 'price_value', 'price_formatted'
                val = c.get('price_value')
                if val is None:
                    txt = c.get('price_formatted') or ""
                else:
                    txt = format_val(val)
                
                candidates_list.append({
                    "description": c.get("description", ""),
                    "value": txt
                })
        except Exception:
            candidates_list = []

        sys.stdout.write(json.dumps({"id": req_id, "candidates": candidates_list}) + "\n")
        sys.stdout.flush()


if __name__ == "__main__":
    main()
