# (c) 2026 Alain Yon
"""Loads and validates tools.json, the single list of tools (spec 4.5)."""
import copy
import datetime
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
REQUIRED = ("id", "name", "group", "what")
VARIANT_KEYS = ("script", "args", "preflight", "data_files", "open_files", "outputs", "auto_answers", "prompts",
                "options_extra", "help")


class ManifestError(Exception):
    pass


class Manifest:
    def __init__(self, path=None):
        self.path = path or os.path.join(HERE, "tools.json")
        with open(self.path, encoding="utf-8") as f:
            raw = json.load(f)
        self.groups = raw.get("groups", [])
        self.tools = raw.get("tools", [])
        self._by_id = {}
        self.validate()

    def validate(self):
        group_ids = {g["id"] for g in self.groups}
        for t in self.tools:
            for k in REQUIRED:
                if not t.get(k):
                    raise ManifestError(f"tools.json: a tool is missing '{k}': {t.get('id', t)}")
            if t["id"] in self._by_id:
                raise ManifestError(f"tools.json: duplicate tool id {t['id']}")
            if t["group"] not in group_ids:
                raise ManifestError(f"tools.json: {t['id']} has unknown group {t['group']}")
            if not t.get("script") and not t.get("variants"):
                raise ManifestError(f"tools.json: {t['id']} needs 'script' or 'variants'")
            for section in ("prompts", "auto_answers", "summary_from_log", "progress"):
                for rule in t.get(section, []):
                    rx = rule.get("match") or rule.get("regex")
                    try:
                        re.compile(rx)
                        for extra in ("title_regex", "options_regex"):
                            if rule.get(extra):
                                re.compile(rule[extra])
                    except re.error as e:
                        raise ManifestError(f"tools.json: {t['id']} {section}: bad regex {rx!r}: {e}")
            t.setdefault("folder", t["id"])
            t.setdefault("inputs", [])
            t.setdefault("options", [])
            t.setdefault("args", [])
            t.setdefault("exit_codes", {"0": "finished"})
            self._by_id[t["id"]] = t

    def ids(self):
        return [t["id"] for t in self.tools]

    def get(self, tool_id):
        if tool_id not in self._by_id:
            raise ManifestError(f"Unknown tool: {tool_id}")
        return self._by_id[tool_id]

    def resolve(self, tool_id, options=None):
        """The tool definition with its chosen variant (outlet / action) merged in."""
        t = copy.deepcopy(self.get(tool_id))
        variants = t.get("variants")
        if variants:
            key = (options or {}).get(t["variant_option"]) or next(iter(variants))
            if key not in variants:
                raise ManifestError(f"{t['name']}: unknown choice {key!r}")
            for k, v in variants[key].items():
                if k in VARIANT_KEYS:
                    t[k] = v
            t["variant"] = key
            t["variant_name"] = variants[key].get("label", key)
        return t


def fill_tokens(text):
    """{this_year} in defaults and file names."""
    if not isinstance(text, str):
        return text
    return text.replace("{this_year}", str(datetime.date.today().year))
