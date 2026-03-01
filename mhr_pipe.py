import sys
import json

# Import the logic
try:
    from mhrsqlsystem import get_candidate_rows, learn_mhr
except Exception as e:
    sys.stdout.write(json.dumps({"fatal": True, "error": f"ImportError: {e}"}) + "\n")
    sys.stdout.flush()
    sys.exit(1)


def format_val(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return f"{v:.3f}"
    return str(v)

def main():
    """
    Simple JSONL pipe:
    - Input line: {"id": 1, "description": "..."}
    - Output line: {"id": 1, "candidates": [{"value": "...", "description": "..."}, ...]}
    """
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            payload = json.loads(line)
            req_id = payload.get("id")
            desc = payload.get("description", "")
            
            mode = payload.get("mode", "query")
            
            if mode == "learn":
                value = payload.get("value")
                try:
                    success = learn_mhr(desc, value)
                    out = {"id": req_id, "success": success}
                except Exception as e:
                    out = {"id": req_id, "success": False, "error": str(e)}
            else:
                candidates_list = []
                try:
                    # Get top N candidates (default 5)
                    limit = int(payload.get("limit", 5))
                    cands, _ = get_candidate_rows(str(desc), limit=limit)
                    
                    # Simple heuristic: if 'term' in description, prefer termination mhr
                    desc_lower = str(desc).lower()
                    is_term = "term" in desc_lower
                    
                    for c in cands:
                        # c has 'cable_laying', 'cable_termination', 'description'
                        val_laying = c.get('cable_laying')
                        val_term = c.get('cable_termination')
                        
                        # Logic: use termination if 'term' in desc AND val_term exists
                        # else use laying
                        # (Fallback to the other if one is missing)
                        
                        # If is_term is true, try term first
                        if is_term:
                            final_val = val_term if val_term not in (None, "") else val_laying
                        else:
                            final_val = val_laying if val_laying not in (None, "") else val_term
                            
                        candidates_list.append({
                            "description": c.get("description", ""),
                            "value": format_val(final_val)
                        })

                except Exception as e:
                    # On error, return empty list
                    candidates_list = []

                out = {"id": req_id, "candidates": candidates_list}
            sys.stdout.write(json.dumps(out) + "\n")
            sys.stdout.flush()
        except Exception:
            sys.stdout.write(json.dumps({"id": None, "candidates": []}) + "\n")
            sys.stdout.flush()


if __name__ == "__main__":
    main()
