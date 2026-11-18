# (c) 2026 Alain Yon
"""settings.json in the data folder. Unknown keys are kept; missing keys get defaults."""
import json
import os

DEFAULTS = {
    "sap_age_warning_days": 7,
    "pr_tool_copy_mm60": False,
    "pr_tool_data_folder": "",
    "data_folders": {},                 # tool folder name -> another folder (e.g. a network share)
    "month_end_checklist": ["oe_monthly", "oe_approvals", "cellar_consumption",
                            "top_bottom_sellers", "food_sales_mix", "theo_cos"],
    "window": {"width": 1280, "height": 860},
}


class Settings:
    def __init__(self, data_root):
        self.path = os.path.join(data_root, "settings.json")
        self.data = {}
        self.load()

    def load(self):
        data = {}
        if os.path.exists(self.path):
            try:
                with open(self.path, encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
        for k, v in DEFAULTS.items():
            if k not in data:
                data[k] = json.loads(json.dumps(v))
        self.data = data
        return data

    def save(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2)
        os.replace(tmp, self.path)

    def get(self, key):
        return self.data.get(key, DEFAULTS.get(key))

    def update(self, changes):
        for k, v in (changes or {}).items():
            if k == "sap_age_warning_days":
                try:
                    v = max(1, min(365, int(v)))
                except (TypeError, ValueError):
                    continue
            self.data[k] = v
        self.save()
        return self.data
