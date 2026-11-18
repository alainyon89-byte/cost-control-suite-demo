# (c) 2026 Alain Yon
"""Run history: one JSON line per run in history.jsonl (spec section 8)."""
import json
import os
import threading

_lock = threading.Lock()


class History:
    def __init__(self, data_root):
        self.path = os.path.join(data_root, "history.jsonl")

    def add(self, entry):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with _lock, open(self.path, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")

    def all(self):
        out = []
        if not os.path.exists(self.path):
            return out
        with _lock, open(self.path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    out.append(json.loads(line))
                except ValueError:
                    continue            # a damaged line never hides the rest
        return out

    def query(self, tool=None, month=None, limit=200):
        """month is 'YYYY-MM'. Newest first."""
        rows = [r for r in self.all()
                if (not tool or r.get("tool") == tool) and (not month or str(r.get("time", "")).startswith(month))]
        rows.reverse()
        return rows[:limit]

    def last_run(self, tool):
        for r in reversed(self.all()):
            if r.get("tool") == tool:
                return r
        return None

    def finished_this_month(self, month):
        return {r.get("tool") for r in self.all()
                if r.get("status") == "finished" and str(r.get("time", "")).startswith(month)}
