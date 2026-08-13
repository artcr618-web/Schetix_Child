# -*- coding: utf-8 -*-
import json
from pathlib import Path

_data = json.loads(Path(__file__).with_name("blocks_data.json").read_text(encoding="utf-8"))
TILES = _data["tiles"]
BLOCKS = _data["blocks"]
CHART_GROUPS = _data.get("chart_groups") or []
PARTS = _data.get("parts") or []


def row_year(r):
    price = float(r.get("price") or 0)
    pcs = float(r.get("pcs") or 1) or 1
    years = float(r.get("years") or 0)
    split = float(r.get("split") or 1) or 1
    if years:
        return price * pcs / years / split
    period = r.get("period") or ""
    if period == "неделя":
        return price * pcs * 52 / split
    if period == "месяц":
        return price * pcs * 12 / split
    if period == "год":
        return price * pcs / split
    return 0
