# (c) 2026 Alain Yon
"""Generate the demo dataset in sample_data/.

    python tools/generate_dummy_data.py              rebuild sample_data/ (wipes it first)
    python tools/generate_dummy_data.py --out DIR    write somewhere else

Everything here is invented. Item names are generic, material numbers start at
90000001, the plants are D100 (the hotel) and D200 (its conference centre), and
storage locations are Lxxx. A fixed random seed means every run writes the same
data (file timestamps aside).

Layout, per tool:
    sample_data/<section>/<tool>/tool_files/   the tool's working files (item maps,
                                               masters, configs): the installer and the
                                               tests copy them into the tool folder
    sample_data/<section>/<tool>/inputs/       the files you drop on the tool
    sample_data/<section>/<tool>/answers_*.txt typed answers used by the regression tests
    sample_data/<section>/<tool>/expected/     golden outputs (written by the tests)

Built-in anomalies, so every tool has something to show:
    Recount Selection  26 bar lines that sales and cocktails cannot explain (over the
                       20-line limit, so the approver prompt appears), sales lines that
                       match no stock line or only loosely (Sales Item Review)
    Theoretical CoS    lines flagged for a CoS that is too high / too low, a map row
                       pointing at a material not in MM60, and (old item map) new POS
                       items to answer on screen
    Cellar             bottles missing against what was charged, Micros names that
                       differ from SAP (asked on screen with the old item map), items
                       sold with no row on the sheet yet
    Food Sales Mix     an alias, a fuzzy auto-match, a flagged near-match, an excluded
                       item, a pending new dish and items with no match
    Loss trackers      coffee ingredients used beyond what the drinks sold explain,
                       breakfast and canteen items with split percentages, canteen over
                       budget
    OE                 stock under par (HIGH urgency), over-par stock in a sister
                       location (transfer options), breakages (MB51 551), a reversal
                       (552), ignored movement types, two-step linen approvals, and a
                       signed approval file with Y, N, blank and an invalid answer
    Cost Comparator    price rises and drops, a few large jumps, new and dropped items
    Stock Age          stock last bought 0-1, 1-2, 2-3, 3-5 and 5+ years ago (or never)
"""
import argparse
import csv
import datetime as dt
import json
import math
import os
import random
import shutil
import subprocess
import sys
import tempfile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SEED = 2026
HOTEL, CC = "D100", "D200"
PROPERTY = "Demo Hotel"
CC_NAME = "Conference Centre"
BAR_LOC, KITCHEN_LOC = "L123", "L073"
YEAR = 2026
SECTIONS = {
    "theo_cos": "beverage_cellar", "cellar_consumption": "beverage_cellar", "recount_selection": "beverage_cellar",
    "top_bottom_sellers": "reporting", "food_sales_mix": "reporting", "cost_comparator": "reporting",
    "stock_age": "reporting", "coffee_loss": "loss_trackers", "breakfast_tracker": "loss_trackers",
    "canteen_loss": "loss_trackers", "oe_monthly": "operating_equipment",
}
BOLD = Font(bold=True)
GREEN = PatternFill("solid", start_color="FFE2EFDA", end_color="FFE2EFDA")
OUT = None
R = random.Random(SEED)


# ---------------------------------------------------------------------------- helpers
def path(tool, *parts):
    return os.path.join(OUT, SECTIONS[tool], tool, *parts)


def save(wb, p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    props = wb.properties
    props.creator = "Demo"
    props.lastModifiedBy = "Demo"
    props.title = None
    props.created = props.modified = dt.datetime(2026, 11, 1, 8, 0, 0)
    wb.save(p)


def r2(x):
    return round(x + 1e-9, 2)


def money_round(x):
    return int(round(x)) if abs(x - round(x)) < 1e-9 else r2(x)


def write_json(p, data):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")


def write_csv(p, rows):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8", newline="") as f:
        csv.writer(f).writerows(rows)


def write_lines(p, lines):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8", newline="\n") as f:
        f.write("".join(ln + "\n" for ln in lines))


# ---------------------------------------------------------------------------- catalogue
MATS = []
BY_DESC = {}
_next = {"bev": 90000001, "food": 90001001, "oe": 90002001, "gen": 90003001}
VAL_CLASS = {"bev": "3031", "food": "3030", "oe": "3100", "gen": "3200"}


def add(block, desc, group, uom, price, **kw):
    m = dict(mat=_next[block], desc=desc, group=group, uom=uom, price=r2(price), block=block, **kw)
    _next[block] += 1
    MATS.append(m)
    BY_DESC[desc] = m
    return m


def M(desc):
    return BY_DESC[desc]


# Spirits: desc, group, bottle cost, Micros family, price per 25ml tot (incl. VAT)
SPIRITS = [
    ("House Vodka 750ml", "BEV001900", 165, "White Spirits", 28),
    ("Premium Vodka 750ml", "BEV001900", 320, "White Spirits", 42),
    ("London Dry Gin 750ml", "BEV001900", 210, "White Spirits", 32),
    ("Botanical Gin 750ml", "BEV001900", 390, "White Spirits", 48),
    ("Silver Tequila 750ml", "BEV001900", 420, "White Spirits", 46),
    ("Blended Scotch Whisky 750ml", "BEV001800", 260, "Whiskey", 34),
    ("Single Malt 12yr 750ml", "BEV001800", 640, "Whiskey", 68),
    ("Irish Whiskey 750ml", "BEV001800", 330, "Whiskey", 40),
    ("Bourbon Whiskey 750ml", "BEV001800", 360, "Whiskey", 42),
    ("White Rum 750ml", "BEV001500", 190, "Rum", 28),
    ("Spiced Rum 750ml", "BEV001500", 240, "Rum", 32),
    ("Cape Brandy 750ml", "BEV000300", 185, "Brandy", 26),
    ("Aged Brandy 10yr 750ml", "BEV000300", 380, "Brandy", 44),
    ("VSOP Cognac 750ml", "BEV000500", 890, "Cognac", 85),
    ("Coffee Liqueur 750ml", "BEV000700", 230, "Liqueur", 30),
    ("Cream Liqueur 750ml", "BEV000700", 205, "Liqueur", 28),
    ("Orange Liqueur 750ml", "BEV000700", 310, "Liqueur", 36),
    ("Herbal Aperitif 750ml", "BEV000100", 280, "Aperitif", 34),
]
# Wine: desc, group, bottle cost, family, bottle price, glass price (None = bottle only)
WINES = [
    ("Estate Cabernet Sauvignon 750ml", "BEV001300", 95, "Red Wine", 320, 85),
    ("House Merlot 750ml", "BEV001300", 62, "Red Wine", 220, 60),
    ("Reserve Shiraz 750ml", "BEV001300", 140, "Red Wine", 420, None),
    ("Classic Pinotage 750ml", "BEV001300", 78, "Red Wine", 260, 70),
    ("Estate Chenin Blanc 750ml", "BEV002000", 58, "White Wine", 210, 58),
    ("House Sauvignon Blanc 750ml", "BEV002000", 66, "White Wine", 230, 62),
    ("Reserve Chardonnay 750ml", "BEV002000", 125, "White Wine", 380, None),
    ("Dry Rose 750ml", "BEV001400", 60, "Rose Wine", 220, 60),
    ("Brut Cap Classique 750ml", "BEV001700", 150, "Sparkling Wine", 450, 85),
    ("Sparkling Brut 750ml", "BEV001700", 72, "Sparkling Wine", 260, 55),
]
PORT = ("Cape Ruby Port 750ml", "BEV001200", 85, "Port", 32)          # sold by the 50ml glass
# Whole units: desc, group, cost, family, selling price
UNITS = [
    ("House Lager 340ml", "BEV000200", 14.50, "Beer", 35),
    ("Premium Lager 330ml", "BEV000200", 17.80, "Beer", 38),
    ("Craft Pale Ale 440ml", "BEV000200", 28.00, "Beer", 62),
    ("Light Lager 330ml", "BEV000200", 15.20, "Beer", 36),
    ("Dry Cider 330ml", "BEV000400", 16.90, "Ciders & Coolers", 38),
    ("Apple Cooler 275ml", "BEV000400", 15.40, "Ciders & Coolers", 36),
    ("Cola 300ml", "BEV000800", 7.20, "Minerals Softdrink", 22),
    ("Cola Zero 300ml", "BEV000800", 7.20, "Minerals Softdrink", 22),
    ("Lemon Soda 300ml", "BEV000800", 6.90, "Minerals Softdrink", 22),
    ("Tonic Water 200ml", "BEV000800", 6.40, "Minerals Softdrink", 18),
    ("Ginger Ale 200ml", "BEV000800", 6.40, "Minerals Softdrink", 18),
    ("Soda Water 200ml", "BEV000800", 5.80, "Minerals Softdrink", 16),
    ("Energy Drink 250ml", "BEV000800", 14.90, "Minerals Softdrink", 38),
    ("Still Water 500ml", "BEV000900", 5.10, "Minerals Water", 18),
    ("Sparkling Water 500ml", "BEV000900", 5.40, "Minerals Water", 18),
    ("Still Water 750ml", "BEV000900", 7.60, "Minerals Water", 28),
    ("Sparkling Water 750ml", "BEV000900", 7.90, "Minerals Water", 28),
]
KEG = ("Draught Lager Keg 30lt", "BEV000200", 1450.00, "Beer", 42)      # sold as a 500ml draught
JUICES = [("Orange Juice 1lt", "BEV001000", 21.0, 26), ("Cranberry Juice 1lt", "BEV001000", 26.0, 28),
          ("Pineapple Juice 1lt", "BEV001000", 22.0, 26)]               # sold by the 250ml glass
CORDIALS = [("Lime Cordial 750ml", "BEV001100", 38.0, 12), ("Sugar Syrup 750ml", "BEV001100", 29.0, 10),
            ("Grenadine Syrup 750ml", "BEV001100", 45.0, 12)]           # sold as a 25ml shot

for d, g, p, fam, tot in SPIRITS:
    add("bev", d, g, "EA", p, fam=fam, tot=tot)
for d, g, p, fam, bottle, glass in WINES:
    add("bev", d, g, "EA", p, fam=fam, bottle=bottle, glass=glass)
add("bev", PORT[0], PORT[1], "EA", PORT[2], fam=PORT[3], glass=PORT[4])
for d, g, p, fam, sell in UNITS:
    add("bev", d, g, "EA", p, fam=fam, sell=sell)
add("bev", KEG[0], KEG[1], "EA", KEG[2], fam=KEG[3], sell=KEG[4])
for d, g, p, glass in JUICES:
    add("bev", d, g, "EA", p, fam="Minerals Juice", glass=glass)
for d, g, p, shot in CORDIALS:
    add("bev", d, g, "EA", p, fam="Minerals Cordial", shot=shot)
add("bev", "Do Not Use Old Lager 340ml", "BEV000200", "EA", 0)          # SAP keeps dead codes; tools skip them

GARNISH = [("Fresh Mint Bunch", "FOO002500", "EA", 14.0), ("Limes Loose", "FOO001500", "KG", 45.0)]
BREAKFAST = [
    ("Free Range Eggs 30s", "FOO000400", "EA", 78.0), ("Streaky Bacon 1kg", "FOO002300", "KG", 165.0),
    ("Pork Sausages 1kg", "FOO002300", "KG", 98.0), ("Beef Sausages 1kg", "FOO002300", "KG", 105.0),
    ("Baked Beans 410g", "FOO001100", "EA", 14.5), ("Button Mushrooms 1kg", "FOO002500", "KG", 78.0),
    ("Tomatoes 1kg", "FOO002500", "KG", 28.0), ("Hash Browns Frozen 1kg", "FOO001200", "KG", 52.0),
    ("Croissants Frozen 50s", "FOO001400", "EA", 420.0), ("Danish Pastries Frozen 40s", "FOO001400", "EA", 460.0),
    ("White Bread Loaf 700g", "FOO000100", "EA", 19.0), ("Brown Bread Loaf 700g", "FOO000100", "EA", 20.0),
    ("Butter Portions 100s", "FOO000200", "EA", 125.0), ("Jam Portions 100s", "FOO000600", "EA", 98.0),
    ("Honey Portions 100s", "FOO000600", "EA", 135.0), ("Plain Yoghurt 1kg", "FOO000400", "KG", 42.0),
    ("Fruit Yoghurt 1kg", "FOO000400", "KG", 48.0), ("Toasted Muesli 1kg", "FOO000700", "KG", 88.0),
    ("Corn Flakes 1kg", "FOO000700", "KG", 62.0), ("Rolled Oats 1kg", "FOO000700", "KG", 34.0),
    ("Orange Juice 5lt", "FOO001800", "EA", 115.0), ("Apple Juice 5lt", "FOO001800", "EA", 110.0),
    ("Full Cream Milk 1lt", "FOO000400", "L", 19.5), ("Smoked Salmon 1kg", "FOO001300", "KG", 420.0),
    ("Cheddar Cheese 1kg", "FOO000400", "KG", 135.0), ("Sliced Ham 1kg", "FOO002300", "KG", 145.0),
    ("Pancake Mix 1kg", "FOO000500", "KG", 56.0), ("Maple Syrup 1lt", "FOO000600", "L", 160.0),
    ("Fresh Fruit Mix 1kg", "FOO001500", "KG", 45.0), ("Baby Spinach 1kg", "FOO002500", "KG", 55.0),
]
CANTEEN = [
    ("Rice 10kg", "FOO000700", "EA", 175.0), ("Maize Meal 10kg", "FOO000700", "EA", 98.0),
    ("Chicken Pieces 2kg", "FOO002100", "EA", 118.0), ("Beef Mince 1kg", "FOO001900", "KG", 115.0),
    ("Beef Stew Meat 1kg", "FOO001900", "KG", 125.0), ("Potatoes 10kg", "FOO002500", "EA", 95.0),
    ("Onions 10kg", "FOO002500", "EA", 88.0), ("Cabbage Each", "FOO002500", "EA", 18.0),
    ("Carrots 10kg", "FOO002500", "EA", 85.0), ("Cooking Oil 5lt", "FOO000700", "EA", 155.0),
    ("Stock Powder 1kg", "FOO000800", "KG", 92.0), ("Macaroni 3kg", "FOO000900", "EA", 72.0),
    ("Tinned Tomatoes 410g", "FOO001100", "EA", 12.0), ("Bread Rolls 6s", "FOO000100", "EA", 18.0),
    ("Margarine 1kg", "FOO000200", "KG", 48.0), ("White Sugar 2.5kg", "FOO000500", "EA", 52.0),
    ("Tea Bags 100s", "FOO002400", "EA", 45.0), ("Instant Coffee 750g", "FOO001700", "EA", 165.0),
    ("Frozen Mixed Veg 2kg", "FOO001200", "EA", 64.0), ("Hake Fillets Frozen 2kg", "FOO001200", "EA", 189.0),
    ("Samp 5kg", "FOO000700", "EA", 62.0), ("Sugar Beans 2kg", "FOO000700", "EA", 58.0),
    ("Peanut Butter 1kg", "FOO000600", "KG", 68.0), ("Polony 1kg", "FOO002300", "KG", 55.0),
    ("Chicken Livers 1kg", "FOO002100", "KG", 59.0), ("Pumpkin Each", "FOO002500", "EA", 32.0),
]
COFFEE = [("Espresso Beans 1kg", "FOO001700", "KG", 310.0), ("Decaf Espresso Beans 1kg", "FOO001700", "KG", 360.0),
          ("Barista Milk 1lt", "FOO000400", "L", 21.5), ("Hot Chocolate Powder 1kg", "FOO001700", "KG", 140.0),
          ("Chocolate Sauce 1lt", "FOO000600", "L", 95.0)]
KITCHEN_OTHER = [
    ("Cake Flour 10kg", "FOO000500", "EA", 145.0), ("Castor Sugar 1kg", "FOO000500", "KG", 32.0),
    ("Fresh Cream 1lt", "FOO000400", "L", 62.0), ("Parmesan Cheese 1kg", "FOO000400", "KG", 390.0),
    ("Mozzarella 2kg", "FOO000400", "EA", 210.0), ("Vanilla Ice Cream 5lt", "FOO000400", "EA", 165.0),
    ("Chicken Breast 2kg", "FOO002100", "EA", 185.0), ("Beef Fillet 1kg", "FOO001900", "KG", 365.0),
    ("Beef Rump 1kg", "FOO001900", "KG", 225.0), ("Beef Sirloin 1kg", "FOO001900", "KG", 245.0),
    ("Lamb Shank Each", "FOO001900", "EA", 95.0), ("Pork Ribs 1kg", "FOO002300", "KG", 148.0),
    ("Salmon Fillet 1kg", "FOO001300", "KG", 395.0), ("Prawns Frozen 1kg", "FOO001200", "KG", 320.0),
    ("Calamari Tubes 1kg", "FOO001200", "KG", 155.0), ("Linefish Fillet 1kg", "FOO001300", "KG", 260.0),
    ("Arborio Rice 1kg", "FOO000700", "KG", 78.0), ("Spaghetti 3kg", "FOO000900", "EA", 82.0),
    ("Lasagne Sheets 1kg", "FOO000900", "KG", 64.0), ("Basmati Rice 2kg", "FOO000700", "EA", 69.0),
    ("Olive Oil 5lt", "FOO000700", "EA", 495.0), ("Balsamic Vinegar 1lt", "FOO000600", "L", 88.0),
    ("Mayonnaise 3kg", "FOO000600", "EA", 115.0), ("Tomato Sauce 5lt", "FOO000600", "EA", 92.0),
    ("Curry Powder 1kg", "FOO000800", "KG", 125.0), ("Black Pepper 1kg", "FOO000800", "KG", 210.0),
    ("Garlic Peeled 1kg", "FOO002500", "KG", 98.0), ("Lettuce Each", "FOO002500", "EA", 16.0),
    ("Lemons 1kg", "FOO001500", "KG", 32.0), ("Avocados Each", "FOO001500", "EA", 14.0),
    ("Burger Buns 6s", "FOO000100", "EA", 26.0), ("Puff Pastry 1kg", "FOO001400", "KG", 74.0),
    ("Dark Chocolate 1kg", "FOO000500", "KG", 210.0), ("Frozen Chips 2kg", "FOO001200", "EA", 58.0),
]
for block in (GARNISH, BREAKFAST, CANTEEN, COFFEE, KITCHEN_OTHER):
    for d, g, u, p in block:
        add("food", d, g, u, p)

# Operating equipment: desc, group, uom, price, [(location, par)]
OE = [
    ("Dinner Plate 27cm", "CRO000100", 58, [("L095", 240), ("L228", 160), ("L112", 400)]),
    ("Side Plate 16cm", "CRO000100", 32, [("L095", 240), ("L228", 160), ("L112", 400)]),
    ("Soup Bowl 450ml", "CRO000100", 41, [("L095", 120), ("L228", 80)]),
    ("Cappuccino Cup 220ml", "CRO000100", 38, [("L095", 96), ("L228", 60), ("L108", 72)]),
    ("Espresso Cup 90ml", "CRO000100", 29, [("L095", 48), ("L108", 48)]),
    ("Saucer 15cm", "CRO000100", 21, [("L095", 144), ("L228", 96), ("L108", 120)]),
    ("Pasta Bowl 30cm", "CRO000100", 76, [("L095", 80), ("L228", 48)]),
    ("Oval Platter 35cm", "CRO000100", 95, [("L095", 24), ("L112", 60)]),
    ("Dinner Knife", "CUT000100", 24, [("L096", 300), ("L229", 200), ("L112", 500)]),
    ("Dinner Fork", "CUT000100", 22, [("L096", 300), ("L229", 200), ("L112", 500)]),
    ("Dessert Spoon", "CUT000100", 19, [("L096", 240), ("L229", 160)]),
    ("Teaspoon", "CUT000100", 12, [("L096", 360), ("L229", 240), ("L108", 120)]),
    ("Steak Knife", "CUT000100", 36, [("L096", 120), ("L229", 80)]),
    ("Fish Knife", "CUT000100", 28, [("L096", 96)]),
    ("Wine Glass 350ml", "GLA000100", 45, [("L097", 240), ("L230", 160), ("L112", 400)]),
    ("Highball Glass 300ml", "GLA000100", 28, [("L097", 180), ("L230", 120)]),
    ("Champagne Flute 180ml", "GLA000100", 39, [("L097", 120), ("L230", 80), ("L112", 300)]),
    ("Water Tumbler 250ml", "GLA000100", 22, [("L097", 300), ("L230", 200), ("L112", 500)]),
    ("Whisky Tumbler 300ml", "GLA000100", 34, [("L097", 96), ("L230", 60)]),
    ("Cocktail Glass 240ml", "GLA000100", 48, [("L097", 96)]),
    ("Beer Glass 400ml", "GLA000100", 31, [("L097", 144), ("L230", 96)]),
    ("Chafing Dish Full Size", "HOL000100", 1250, [("L098", 12), ("L112", 30)]),
    ("Milk Jug 300ml", "HOL000100", 85, [("L098", 48), ("L231", 30)]),
    ("Tea Pot 600ml", "HOL000100", 145, [("L098", 36), ("L231", 24)]),
    ("Bread Basket", "HOL000100", 65, [("L098", 60), ("L231", 40)]),
    ("Ice Bucket", "HOL000100", 320, [("L098", 20), ("L231", 12)]),
    ("Serving Tongs 24cm", "UTE000100", 42, [("L099", 60), ("L232", 40)]),
    ("Soup Ladle 120ml", "UTE000100", 68, [("L099", 30), ("L232", 20)]),
    ("Chef Knife 25cm", "UTE000100", 420, [("L099", 12), ("L232", 8)]),
    ("Cutting Board Red", "UTE000100", 185, [("L099", 10), ("L232", 6)]),
    ("Balloon Whisk 30cm", "UTE000100", 74, [("L099", 12)]),
    ("Napkin White 50cm", "FBL000100", 28, [("L100", 600), ("L233", 400)]),
    ("Tablecloth Round 2.3m", "FBL000100", 310, [("L100", 120), ("L233", 60)]),
    ("Tablecloth Square 1.4m", "FBL000100", 185, [("L100", 80), ("L233", 60)]),
    ("Bar Cloth", "FBL000100", 18, [("L100", 120), ("L233", 80)]),
    ("Bath Towel White", "RLI000100", 145, [("L101", 600), ("L234", 400)]),
    ("Hand Towel White", "RLI000100", 62, [("L101", 600), ("L234", 400)]),
    ("Face Cloth White", "RLI000100", 22, [("L101", 600), ("L234", 400)]),
    ("Bath Mat", "RLI000100", 75, [("L101", 300), ("L234", 200)]),
    ("Bed Sheet King", "RLI000100", 260, [("L101", 400), ("L234", 260)]),
    ("Bed Sheet Queen", "RLI000100", 230, [("L101", 300), ("L234", 200)]),
    ("Pillow Case", "RLI000100", 48, [("L101", 900), ("L234", 600)]),
    ("Duvet Cover Queen", "RLI000100", 520, [("L101", 200), ("L234", 140)]),
    ("Bathrobe White", "ROO000100", 480, [("L102", 120), ("L235", 80)]),
    ("Slippers Pair", "ROO000100", 18, [("L102", 600), ("L235", 400)]),
    ("Laundry Bag", "ROO000100", 6, [("L102", 800), ("L235", 500)]),
    ("Coat Hanger Wooden", "ROO000100", 24, [("L102", 200), ("L235", 120)]),
    ("Hair Dryer", "ROO000100", 395, [("L102", 20), ("L235", 12)]),
    ("Ballpoint Pen Blue", "STA000100", 3.5, [("L081", 500)]),
    ("A4 Copy Paper Ream", "STA000100", 72, [("L081", 60)]),
    ("Notepad A5", "STA000100", 14, [("L081", 200)]),
    ("Shampoo 30ml", "GUE000100", 4.2, [("L083", 2000)]),
    ("Shower Gel 30ml", "GUE000100", 4.2, [("L083", 2000)]),
    ("Soap Bar 25g", "GUE000100", 2.9, [("L083", 2500)]),
    ("Dental Kit", "GUE000100", 5.5, [("L083", 600)]),
    ("Multi-surface Cleaner 5lt", "CLE000100", 145, [("L080", 24)]),
    ("Dishwash Liquid 5lt", "CLE000100", 165, [("L080", 24)]),
    ("Floor Polish 5lt", "CLE000100", 210, [("L080", 12)]),
]
for d, g, p, locs in OE:
    add("oe", d, g, "EA", p, oe_locs=locs)
GENERAL = [("Toilet Paper 2ply 48s", "PAP000100", "EA", 290.0), ("Paper Towel Roll 6s", "PAP000100", "EA", 145.0),
           ("Takeaway Cup 250ml 50s", "PAP000100", "EA", 95.0), ("Cling Film 450mm", "PAP000100", "EA", 135.0),
           ("Foil Roll 450mm", "PAP000100", "EA", 188.0), ("Bin Liners Black 20s", "PAP000100", "EA", 48.0),
           ("Light Bulb LED 9W", "MAI000800", "EA", 39.0), ("Batteries AA 4s", "MAI000800", "EA", 64.0)]
for d, g, u, p in GENERAL:
    add("gen", d, g, u, p)

LOC_NAMES = {
    "L010": "MAIN STORES", "L011": "BEVERAGE STORES", "L073": "KITCHEN", "L075": "BAR - CELLAR",
    "L080": "KITCHEN CLEANING", "L081": "STATIONERY", "L083": "GUEST SUPPLIES", "L095": "CROCKERY",
    "L096": "CUTLERY", "L097": "GLASSWARE", "L098": "HOLLOWARE", "L099": "UTENSILS", "L100": "F&B LINEN",
    "L101": "ROOMS LINEN", "L102": "ROOMS OTHER", "L108": "CAFE", "L112": "BANQUETING", "L123": "POOL BAR",
    "L228": "CROCKERY TW", "L229": "CUTLERY TW", "L230": "GLASSWARE TW", "L231": "HOLLOWARE TW",
    "L232": "UTENSILS TW", "L233": "F&B LINEN TW", "L234": "ROOMS LINEN TW", "L235": "ROOMS OTHER TW",
}


# ---------------------------------------------------------------------------- MM60
MM60_HEAD = ["Plant", "Material", "Material Description", "Material Group", "Base Unit of Measure", "Price",
             "Currency", "MRP Type", "Valuation Class"]


def mm60_prices(plant, month):
    """{mat: price} for a plant and month ('sep' / 'oct'). Seeded per item, so stable."""
    out = {}
    for m in MATS:
        rr = random.Random(f"{SEED}-{m['mat']}")
        base = m["price"]
        if plant == CC:
            base = r2(base * (1 + rr.choice([0, 0, 0, 0.02, -0.02, 0.03])))
        if month == "sep":
            rr2 = random.Random(f"{SEED}-sep-{m['mat']}")
            roll = rr2.random()
            if roll < 0.62:
                factor = 1.0
            elif roll < 0.86:
                factor = 1 / (1 + rr2.uniform(0.02, 0.12))        # Oct is 2-12% dearer
            elif roll < 0.95:
                factor = 1 + rr2.uniform(0.02, 0.08)              # Oct is cheaper
            else:
                factor = 1 / (1 + rr2.uniform(0.25, 0.45))        # big jump in Oct
            base = r2(base * factor)
        out[m["mat"]] = base
    return out


ONLY_SEP = {"Batteries AA 4s", "Light Bulb LED 9W"}           # dropped from the October MM60
ONLY_OCT = {"Energy Drink 250ml", "Pumpkin Each", "Dental Kit"}           # new in October


def mm60_rows(plant, month="oct"):
    prices = mm60_prices(plant, month)
    rows = []
    for m in MATS:
        if month == "oct" and m["desc"] in ONLY_SEP:
            continue
        if month == "sep" and m["desc"] in ONLY_OCT:
            continue
        rows.append([plant, m["mat"], m["desc"], m["group"], m["uom"], prices[m["mat"]], "ZAR", "ND",
                     VAL_CLASS[m["block"]]])
    return rows


def write_mm60_sheet(ws, plant, month="oct"):
    ws.append(MM60_HEAD)
    for c in ws[1]:
        c.font = BOLD
    for r in mm60_rows(plant, month):
        ws.append(r)
    ws.freeze_panes = "A2"
    for col, w in zip("ABCDEFGHI", (7, 11, 38, 12, 8, 11, 8, 8, 9)):
        ws.column_dimensions[col].width = w


def mm60_file(p, plant, month="oct"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    write_mm60_sheet(ws, plant, month)
    save(wb, p)


PRICE = {HOTEL: mm60_prices(HOTEL, "oct"), CC: mm60_prices(CC, "oct")}


def price(desc, plant=HOTEL):
    return PRICE[plant][M(desc)["mat"]]


# ---------------------------------------------------------------------------- Micros Sales Mix (old layout)
SALES_HEAD = ["Item", "Family Group", "Major Group", "Gross Sales", "Item Discounts", "Sales Less Item Disc",
              "% Sales", "Qty Sold", "% Qty Sold", "Average Price"]


def sales_mix_sheet(ws, business_dates, location, revenue_centre, menu_items, lines, pos_version=True):
    """lines: (item, family, major, qty, avg price incl. VAT)."""
    ws.title = "Report"
    ws.append(["Sales Mix Items Summary"])
    ws.append(["Business Dates", business_dates])
    ws.append(["Locations", location])
    ws.append(["Revenue Centers", revenue_centre])
    ws.append(["Menu Items", menu_items])
    if pos_version:
        ws.append(["POS Version", "POS export 19"])
    ws.append(SALES_HEAD)
    ws["A1"].font = Font(bold=True, size=12)
    total_sales = sum(q * p for _, _, _, q, p in lines) or 1
    total_qty = sum(q for _, _, _, q, _ in lines) or 1
    for item, fam, major, qty, avg in lines:
        gross = r2(qty * avg)
        ws.append([item, fam, major, gross, 0, gross, round(gross / total_sales, 4), qty,
                   round(qty / total_qty, 4), r2(avg)])
    ws.append(["Total Item Sales:", "", "", r2(total_sales), 0, r2(total_sales), 1, total_qty, 1,
               r2(total_sales / total_qty)])
    for col, w in zip("ABCDEFGHIJ", (34, 24, 20, 12, 12, 14, 8, 9, 9, 11)):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------- bar sales
def bar_lines(rr, scale, cocktails, extra=(), skip=()):
    """Sales lines for a bar: tots, bottles, glasses, units, cocktails. Returns list of
    (item, family, major, qty, avg price)."""
    out = []
    fam = lambda f: "Ent - " + f
    major = "Beverage Enterprise"
    for d, g, p, f, tot in SPIRITS:
        base = d.replace(" 750ml", "")
        if base in skip:
            continue
        out.append((base, fam(f), major, max(1, int(rr.uniform(20, 140) * scale)), tot))
        if d in ("Premium Vodka 750ml", "Single Malt 12yr 750ml", "VSOP Cognac 750ml") and d not in skip:
            out.append((d, fam(f), major, max(1, int(rr.uniform(1, 4) * scale)), r2(p * 2.9)))
    for d, g, p, f, bottle, glass in WINES:
        if d in skip:
            continue
        out.append((d, fam(f), major, max(1, int(rr.uniform(4, 30) * scale)), bottle))
        if glass:
            out.append((d.replace(" 750ml", " GLS"), fam(f), major, max(1, int(rr.uniform(15, 90) * scale)), glass))
    out.append(("Cape Ruby Port", fam("Port"), major, max(1, int(rr.uniform(10, 40) * scale)), PORT[4]))
    for d, g, p, f, sell in UNITS:
        if d in skip:
            continue
        out.append((d, fam(f), major, max(1, int(rr.uniform(30, 260) * scale)), sell))
    out.append(("Draught Lager", fam("Beer"), major, max(1, int(rr.uniform(200, 600) * scale)), KEG[4]))
    for d, g, p, glass in JUICES:
        out.append((d.replace(" 1lt", " GLS"), fam("Minerals Juice"), major, max(1, int(rr.uniform(20, 90) * scale)), glass))
    for d, g, p, shot in CORDIALS[:1]:
        out.append(("Lime Cordial Shot", fam("Minerals Cordial"), major, max(1, int(rr.uniform(10, 40) * scale)), shot))
    for name, sell, _ in cocktails:
        out.append((name, fam("Cocktails"), major, max(1, int(rr.uniform(15, 80) * scale)), sell))
    out.extend(extra)
    return out


# Cocktail recipes: name, selling price, [(SAP description, qty in SAP units per drink)]
ml = lambda x: round(x / 750.0, 4)
COCKTAILS = [
    ("Classic Mojito", 85, [("White Rum 750ml", ml(50)), ("Lime Cordial 750ml", ml(25)), ("Sugar Syrup 750ml", ml(15)),
                            ("Soda Water 200ml", 0.5), ("Fresh Mint Bunch", 0.1), ("Limes Loose", 0.05)]),
    ("Gin and Tonic", 72, [("London Dry Gin 750ml", ml(50)), ("Tonic Water 200ml", 1), ("Limes Loose", 0.02)]),
    ("Cosmopolitan", 88, [("House Vodka 750ml", ml(50)), ("Orange Liqueur 750ml", ml(25)),
                          ("Cranberry Juice 1lt", 0.06), ("Lime Cordial 750ml", ml(15))]),
    ("Espresso Martini", 95, [("Premium Vodka 750ml", ml(50)), ("Coffee Liqueur 750ml", ml(25)),
                              ("Sugar Syrup 750ml", ml(10)), ("Espresso Beans 1kg", 0.008)]),
    ("Margarita", 92, [("Silver Tequila 750ml", ml(50)), ("Orange Liqueur 750ml", ml(25)), ("Lime Cordial 750ml", ml(25))]),
    ("Whisky Sour", 85, [("Bourbon Whiskey 750ml", ml(50)), ("Sugar Syrup 750ml", ml(20)), ("Limes Loose", 0.04)]),
    ("Rum Punch", 80, [("Spiced Rum 750ml", ml(50)), ("Pineapple Juice 1lt", 0.1), ("Orange Juice 1lt", 0.1),
                       ("Grenadine Syrup 750ml", ml(10))]),
    ("Virgin Mojito", 55, [("Lime Cordial 750ml", ml(25)), ("Sugar Syrup 750ml", ml(15)), ("Soda Water 200ml", 1),
                           ("Fresh Mint Bunch", 0.1)]),
]


def costing_workbook(p, plant, cocktails, mm60_title="MM60", summary_title="Summary All Menus", extra_col=True):
    """Cocktail costing / tracker workbook: a summary tab, the MM60, an example sheet and one 'Menu N' per recipe."""
    wb = Workbook()
    ws = wb.active
    ws.title = summary_title
    ws["A1"] = "Cocktail costing summary"
    ws["A1"].font = Font(bold=True, size=12)
    head = ["", "Menu", "Recipe", "", "Menu item cost", "Total excl. Vat", "Vat", "Total incl. Vat", "",
            "Selling Price", "", "COS %"]
    ws.append(head)
    for c in ws[2]:
        c.font = BOLD
    mm = wb.create_sheet(mm60_title)
    write_mm60_sheet(mm, plant)
    ex = wb.create_sheet("Examp Menu")
    ex["C2"] = "Example recipe"
    menu_head = ["", "SAP Nr", "ITEM DESCRIPTION (SAP)", "QTY - (kg, ea, lt)", "UOM", "SAP VALUE", "TOTAL VALUE"]
    if extra_col:
        menu_head.append("Measure / note")
    for c, h in enumerate(menu_head, 1):
        ex.cell(4, c, h or None).font = BOLD
    for i, (name, sell, ings) in enumerate(cocktails, 1):
        ms = wb.create_sheet(f"Menu {i}")
        ms["B2"] = "Recipe:"
        ms["C2"] = name
        ms["C2"].font = Font(bold=True, size=12)
        for c, h in enumerate(menu_head, 1):
            if h:
                ms.cell(4, c, h).font = BOLD
        subtotal = 0.0
        r = 5
        for desc, qty in ings:
            m = M(desc)
            pr = PRICE[plant][m["mat"]]
            cost = r2(qty * pr)
            subtotal += cost
            ms.cell(r, 2, m["mat"])
            ms.cell(r, 3, m["desc"])
            ms.cell(r, 4, qty)
            ms.cell(r, 5, m["uom"])
            ms.cell(r, 6, pr)
            ms.cell(r, 7, cost)
            if extra_col and m["uom"] == "EA" and m["desc"].endswith("750ml"):
                ms.cell(r, 8, f"{round(qty * 750)}ml")
            r += 1
        misc = r2(subtotal * 0.03)
        total = r2(subtotal + misc)
        ex_vat = r2(sell / 1.15)
        for row, label, val in ((27, "Miscellaneous 3%", misc), (28, "Menu item cost", total),
                                (30, "Selling price excl VAT", ex_vat), (31, "VAT", r2(sell - ex_vat)),
                                (33, "Selling price incl VAT", sell)):
            ms.cell(row, 2, label).font = BOLD
            ms.cell(row, 7, val)
        ms.column_dimensions["C"].width = 34
        ws.append(["", f"Menu {i}", name, "", total, r2(total), r2(total * 0.15), r2(total * 1.15), "", sell, "",
                   round(total / ex_vat, 4)])
    for col, w in zip("ABCDEFGHIJKL", (3, 9, 26, 3, 14, 14, 10, 14, 3, 12, 3, 8)):
        ws.column_dimensions[col].width = w
    save(wb, p)


# ---------------------------------------------------------------------------- Theoretical CoS
THEO_REMOVED = ["Botanical Gin", "Reserve Shiraz 750ml", "Dry Rose GLS", "Energy Drink 250ml", "Margarita",
                "Apple Cooler 275ml"]


def theo_cos():
    rr = random.Random(f"{SEED}-theo")
    cc_cocktails = [c for c in COCKTAILS if c[0] != "Virgin Mojito"]
    extra = [("Table Water Still", "Ent - Minerals Water", "Beverage Enterprise", 120, 15),
             ("Ice Bag 2kg", "Ent - Minerals Water", "Beverage Enterprise", 14, 25),
             ("Corkage", "Ent - Red Wine", "Beverage Enterprise", 6, 60)]
    lines = bar_lines(rr, 1.4, cc_cocktails, extra)
    wb = Workbook()
    sales_mix_sheet(wb.active, "10/1/2026 - 10/31/2026", CC_NAME, "All", "Beverage", lines)
    mm = wb.create_sheet("MM60")
    write_mm60_sheet(mm, CC)
    save(wb, path("theo_cos", "inputs", "CC_Bar_Sales_Oct_2026.xlsx"))

    costing_workbook(path("theo_cos", "tool_files", "Cocktail_Costing.xlsx"), CC, cc_cocktails)
    recipe_sheet = {name: f"Menu {i}" for i, (name, _, _) in enumerate(cc_cocktails, 1)}

    house = [("Sales are VAT-inclusive", "Yes"), ("VAT rate", 0.15), ("Tot (ml)", 25), ("Wine glass (ml)", 187.5),
             ("Bubbly glass (ml)", 125), ("Draught (ml)", 500), ("Cordial (ml)", 25), ("Juice glass (ml)", 250),
             ("Port glass (ml)", 50), ("Dessert glass (ml)", 50), ("Cocktail rule", "Exclude"),
             ("Cocktail standard CoS %", 0.22), ("Flag CoS below", 0.05), ("Flag CoS above", 0.8),
             ("Table water cost per bottle (R)", 6.5), ("Cocktail costing file", "Cocktail_Costing.xlsx")]
    rows = []
    for item, family, _, _, avg in lines:
        f = family.replace("Ent - ", "")
        row = dict(item=item, min_avg=0, treatment="SAP item", material="", desc="", pack=None, serving="Whole",
                   rule="", pour=None, divisor=None, cost=None, fam=f, notes="", recipe="", skip="")
        if item in recipe_sheet:
            row.update(treatment="Recipe", recipe=recipe_sheet[item], desc=item)
        elif item == "Table Water Still":
            row.update(treatment="Table water rule")
        elif item == "Ice Bag 2kg":
            row.update(treatment="Manual cost", cost=18.0, pack=1, notes="Not a stock item: bought for cash")
        elif item == "Corkage":
            row.update(treatment="Exclude", notes="A charge, not a product")
        else:
            sap = next((m for m in MATS if m["block"] == "bev" and m["desc"].startswith(
                item.replace(" GLS", "").replace("Draught Lager", "Draught Lager Keg").replace(" Shot", ""))), None)
            if item.endswith(" GLS") and sap and sap["desc"].endswith("1lt"):
                row.update(serving="House", rule="Juice glass")
            elif item.endswith(" GLS"):
                row.update(serving="House", rule="Bubbly glass" if "Sparkling" in f else "Wine glass")
            elif item == "Draught Lager":
                row.update(serving="House", rule="Draught")
            elif item == "Lime Cordial Shot":
                row.update(serving="House", rule="Cordial")
            elif item == "Cape Ruby Port":
                row.update(serving="House", rule="Port glass")
            elif f in ("White Spirits", "Whiskey", "Rum", "Brandy", "Cognac", "Liqueur", "Aperitif") and not item.endswith("ml"):
                row.update(serving="House", rule="Tot")
            if sap:
                row.update(material=sap["mat"], desc=sap["desc"])
            if item.endswith("750ml") and f in ("White Spirits", "Whiskey", "Cognac"):
                row["min_avg"] = 300            # the bottle row of a spirit also sold by the tot
        # anomalies
        if item == "Cream Liqueur":
            row.update(serving="Whole", rule="", notes="")           # wrong: sold by the tot -> CoS far too high
        if item == "Premium Lager 330ml":
            row.update(serving="Divisor", divisor=24, notes="")      # wrong: costed per case -> CoS far too low
        if item == "Craft Pale Ale 440ml":
            row.update(material=90009999, desc="Craft Pale Ale 440ml (old code)")   # not in MM60 -> Review
        if item == "Estate Chenin Blanc 750ml":
            row.update(skip="Yes", notes="Promotional price agreed, CoS checked")
        rows.append(row)

    def item_map(p, drop=()):
        wb = Workbook()
        rm = wb.active
        rm.title = "Read Me"
        for i, t in enumerate([
                "Beverage item map (demo data)",
                "One row per Micros POS item: which SAP material it is sold from and how (whole, house pour, ml, divisor).",
                "Prices are not kept here: every run reads them from the MM60 by material number.",
                "House Rules: pour sizes, VAT, the cocktail rule and the costing workbook name.",
                "Treatment 'Recipe' costs a drink from its sheet in the cocktail costing workbook.",
                "Skip CoS Flag = Yes stops a line you have checked from being flagged again."], 1):
            rm.cell(i, 1, t)
        rm["A1"].font = Font(bold=True, size=12)
        hs = wb.create_sheet("House Rules")
        hs.append(["Setting", "Value", "Notes"])
        for k, v in house:
            hs.append([k, v, None])
        ms = wb.create_sheet("Item Map")
        ms.append(["POS Item", "Min Avg Price (R)", "Treatment", "SAP Material", "SAP Description", "Pack (ml)",
                   "Serving", "House Rule", "Pour (ml)", "Divisor", "Manual Cost (R)", "POS Family", "Notes", "Added",
                   "Skip CoS Flag", "Recipe Sheet"])
        for c in list(hs[1]) + list(ms[1]):
            c.font = BOLD
        for x in rows:
            if x["item"] in drop:
                continue
            ms.append([x["item"], x["min_avg"], x["treatment"], x["material"] or None, x["desc"] or None, x["pack"],
                       x["serving"], x["rule"] or None, x["pour"], x["divisor"], x["cost"], x["fam"],
                       x["notes"] or None, "Seeded 1 Oct 2026", x["skip"] or None, x["recipe"] or None])
        ms.freeze_panes = "B2"
        ms.column_dimensions["A"].width = 32
        ms.column_dimensions["E"].width = 34
        save(wb, p)

    item_map(path("theo_cos", "tool_files", "beverage_item_map.xlsx"))
    item_map(path("theo_cos", "inputs", "beverage_item_map_OLD.xlsx"), drop=set(THEO_REMOVED))

    # answers for the old-map run, in the order the script asks
    answers = []
    for item, family, _, qty, _ in lines:
        if item not in THEO_REMOVED or qty == 0:
            continue
        if item == "Botanical Gin":
            answers += ["N", str(M("Botanical Gin 750ml")["mat"]), "2"]          # house tot
        elif item == "Reserve Shiraz 750ml":
            answers += ["N", str(M(item)["mat"]), "1"]                          # whole bottle
        elif item == "Dry Rose GLS":
            answers += ["N", str(M("Dry Rose 750ml")["mat"]), "3"]              # house wine glass
        elif item == "Energy Drink 250ml":
            answers += ["N", str(M(item)["mat"]), "1"]
        elif item == "Margarita":
            answers += ["R1"]                                                   # its recipe sheet
        elif item == "Apple Cooler 275ml":
            answers += ["S"]                                                    # skip: stays on Review
    write_lines(path("theo_cos", "answers_theo_cos_new_items.txt"), answers)
    return lines


# ---------------------------------------------------------------------------- Recount Selection
def recount():
    rr = random.Random(f"{SEED}-recount")
    extra = [("House Sauv Blanc GLS", "Ent - White Wine", "Beverage Enterprise", 34, 62),       # loose match
             ("Bar Snack Platter", "Ent - Bar Snacks", "Beverage Enterprise", 41, 95),           # no stock line
             ("Ice Bucket Hire", "Ent - Sundries", "Beverage Enterprise", 12, 40),               # no stock line
             ("Pool Bar Lager 340ml", "Ent - Beer", "Beverage Enterprise", 64, 35),              # alias in csv
             ("Frozen Strawberry Daiquiri", "Ent - Cocktails", "Beverage Enterprise", 22, 85)]   # no recipe
    lines = bar_lines(rr, 1.0, COCKTAILS, extra, skip={"Cape Brandy", "Herbal Aperitif"})
    wb = Workbook()
    sales_mix_sheet(wb.active, "10/1/2026 - 10/31/2026", PROPERTY, "Pool Bar", "Beverage", lines)
    save(wb, path("recount_selection", "inputs", "Pool_Bar_Sales_Oct_2026.xlsx"))

    costing_workbook(path("recount_selection", "inputs", "Cocktail_Tracker.xlsx"), HOTEL, COCKTAILS,
                     mm60_title="Updated Oct 26", extra_col=False)

    # theoretical usage per stock line, the same conversions the script uses
    usage = {}

    def use(desc, q):
        usage[desc] = usage.get(desc, 0.0) + q
    for item, family, _, qty, avg in lines:
        f = family.replace("Ent - ", "")
        if f == "Cocktails":
            rec = next((c for c in COCKTAILS if c[0] == item), None)
            for desc, q in (rec[2] if rec else []):
                use(desc, q * qty)
            continue
        if item == "Pool Bar Lager 340ml":
            use("House Lager 340ml", qty)
        elif item == "House Sauv Blanc GLS":
            use("House Sauvignon Blanc 750ml", qty / 4)
        elif item == "Draught Lager":
            use(KEG[0], qty / 60)
        elif item == "Lime Cordial Shot":
            use("Lime Cordial 750ml", qty / 30)
        elif item == "Cape Ruby Port":
            use(PORT[0], qty / 15)
        elif item.endswith(" GLS"):
            base = item[:-4]
            j = next((d for d, *_ in JUICES if d.startswith(base)), None)
            use(j or base + " 750ml", qty / 4)
        elif f in ("White Spirits", "Whiskey", "Rum", "Brandy", "Cognac", "Liqueur", "Aperitif") and not item.endswith("ml"):
            use(item + " 750ml", qty / 30)
        elif item in BY_DESC:
            use(item, qty)

    stock = [m for m in MATS if m["block"] == "bev" and m["price"] > 0] + [M("Fresh Mint Bunch"), M("Limes Loose")]
    unexplained = set(rr.sample([m["desc"] for m in stock if usage.get(m["desc"], 0) > 0.5], 23))
    counted_more = set(rr.sample([m["desc"] for m in stock if m["desc"] not in unexplained], 3))
    quiet = {"Cape Brandy 750ml", "Herbal Aperitif 750ml", "Sparkling Water 750ml"}           # no movement
    rows = []
    for i, m in enumerate(stock, 1):
        pr = price(m["desc"])
        t = usage.get(m["desc"], 0.0)
        book = round(max(t * 1.6, rr.uniform(6, 40)), 0 if m["uom"] == "EA" and pr < 100 else 1)
        if m["desc"] in quiet:
            diff = 0.0
        elif m["desc"] in counted_more:
            diff = round(max(300 / pr, 1) + rr.uniform(1, 4), 1)
        elif m["desc"] in unexplained:
            gap = max(420 / pr, 0.18 * t) + rr.uniform(0.5, 3)
            diff = -round(t + gap, 1)
        else:
            noise = min(0.05 * t, 120 / pr) * rr.uniform(-1, 1)
            diff = -round(t + noise, 1)
        diff = 0.0 if abs(diff) < 0.05 else diff
        if m["uom"] == "EA" and pr < 60:
            diff = float(round(diff))
        counted = round(book + diff, 1)
        val = r2(diff * pr)
        rows.append([f"2000{BAR_LOC[1:]}", str(i), str(m["mat"]), m["desc"], HOTEL, BAR_LOC, book, counted, diff,
                     val, val if val > 0 else 0, val if val < 0 else 0])
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["PhysInvDoc", "Item", "Material", "Material Description", "Plnt", "Location", "Book Qty",
               "Qty Counted", "Difference qty", "Difference", "Pos. diff. value", "neg. diff. value"])
    for c in ws[1]:
        c.font = BOLD
    for r in rows:
        ws.append(r)
    ws.column_dimensions["D"].width = 34
    save(wb, path("recount_selection", "inputs", f"{BAR_LOC}_Variance_Oct_2026.xlsx"))

    tots = 30.0
    write_csv(path("recount_selection", "tool_files", "recount_config.csv"), [
        ["family_contains", "default_unit", "servings_per_stock_unit", "by_keg_volume", "note"],
        ["whiskey", "portion", tots, "N", "bare line = tot"], ["whisky", "portion", tots, "N", "bare line = tot"],
        ["cognac", "portion", tots, "N", "bare line = tot"], ["brandy", "portion", tots, "N", "bare line = tot"],
        ["white spirits", "portion", tots, "N", "bare line = tot"], ["liqueur", "portion", tots, "N", "bare line = tot"],
        ["rum", "portion", tots, "N", "bare line = tot"], ["shooters", "portion", tots, "N", "bare line = tot"],
        ["aperitif", "portion", tots, "N", "bare line = tot"], ["port", "portion", 15.0, "N", "50ml measure"],
        ["sherry", "portion", 15.0, "N", "50ml measure"],
        ["red wine", "whole", 4.0, "N", "bare line = bottle, GLS = glass"],
        ["white wine", "whole", 4.0, "N", "bare line = bottle, GLS = glass"],
        ["rose wine", "whole", 4.0, "N", "bare line = bottle, GLS = glass"],
        ["sparkling wine", "whole", 6.0, "N", "bare line = bottle, GLS = flute"],
        ["beer", "whole", 0, "Y", "kegs convert by volume"], ["cooler", "whole", 1, "N", "each"],
        ["minerals cordial", "portion", tots, "N", "cordial shot from the bottle"],
        ["minerals", "whole", 1, "N", "each"], ["alcohol coffee", "portion", tots, "N", "spirit measure, no recipe held"],
        ["cocktails", "recipe", 0, "N", "exploded via Cocktail Tracker"]])
    write_csv(path("recount_selection", "tool_files", "item_overrides.csv"), [
        ["sales_item_contains", "servings_per_stock_unit", "note"],
        ["# example: Irish Whiskey,30,tot from a 750ml bottle", "", ""]])
    write_csv(path("recount_selection", "tool_files", "item_aliases.csv"), [
        ["sales_item_contains", "material_no", "note"],
        ["pool bar lager", str(M("House Lager 340ml")["mat"]), "house lager sold under the pool bar's own name"],
        ["ice bucket hire", "SKIP", "a charge, no stock line"]])
    return lines


# ---------------------------------------------------------------------------- Cellar Consumption
CELLAR_SECTIONS = [
    ("", ["Cola 300ml", "Cola Zero 300ml", "Lemon Soda 300ml", "Tonic Water 200ml", "Ginger Ale 200ml",
          "Soda Water 200ml", "Orange Juice 1lt", "Cranberry Juice 1lt"], 2, False),
    ("Ent - Minerals Water", ["Still Water 500ml", "Sparkling Water 500ml", "Still Water 750ml",
                              "Sparkling Water 750ml"], 2, False),
    ("Ent - Beer", ["House Lager 340ml", "Premium Lager 330ml", "Light Lager 330ml", "Dry Cider 330ml",
                    "Apple Cooler 275ml"], 3, False),
    ("Ent - Whiskey", ["Blended Scotch Whisky 750ml", "Single Malt 12yr 750ml", "Irish Whiskey 750ml",
                       "Bourbon Whiskey 750ml"], 2, True),
    ("Ent - White Wine", ["Estate Chenin Blanc 750ml", "House Sauvignon Blanc 750ml", "Reserve Chardonnay 750ml",
                          "Brut Cap Classique 750ml"], 2, False),
    ("Ent - Red Wine", ["Estate Cabernet Sauvignon 750ml", "House Merlot 750ml", "Classic Pinotage 750ml",
                        "Reserve Shiraz 750ml"], 2, False),
    ("Ent - Brandy", ["Cape Brandy 750ml", "Aged Brandy 10yr 750ml"], 2, True),
    ("Ent - Rum", ["White Rum 750ml", "Spiced Rum 750ml"], 2, True),
    ("Ent - White Spirits", ["House Vodka 750ml", "London Dry Gin 750ml", "Coffee Liqueur 750ml"], 3, True),
]
# Micros name -> SAP description, where they differ (the item map holds these)
CELLAR_RENAMES = {
    "Cab Sauv Estate 750ml": "Estate Cabernet Sauvignon 750ml",
    "Hse Merlot GLS": "House Merlot 750ml",
    "Sauv Blanc House 750ml": "House Sauvignon Blanc 750ml",
    "Chenin Estate GLS": "Estate Chenin Blanc 750ml",
    "Scotch Blended": "Blended Scotch Whisky 750ml",
    "Lager Premium 330ml": "Premium Lager 330ml",
    "Water Still 500ml": "Still Water 500ml",
    "Tonic 200ml": "Tonic Water 200ml",
}
CELLAR_OLD_MAP_KEEPS = {"Cab Sauv Estate 750ml", "Scotch Blended"}
CELLAR_MAP_ALWAYS = {"Orange Juice GLS": "Orange Juice 1lt", "Cranberry Juice GLS": "Cranberry Juice 1lt"}


def cellar():
    rr = random.Random(f"{SEED}-cellar")
    day = dt.date(2026, 11, 10)
    sheet_name = day.strftime("%d.%m.%Y")
    # what the function sold (Micros), by SAP description: (micros name, qty, avg price)
    sold = []
    for label, items, spare, serves in CELLAR_SECTIONS:
        for d in items:
            m = M(d)
            name = next((k for k, v in CELLAR_RENAMES.items() if v == d), d)
            if serves:
                name = name if name != d else d.replace(" 750ml", "")
                qty = rr.randint(18, 90)
                avg = m["tot"]
            elif name.endswith("GLS"):
                qty = rr.randint(20, 70)
                avg = m["glass"]
            elif m.get("bottle"):
                qty = rr.randint(4, 24)
                avg = m["bottle"]
            elif m.get("glass"):                    # juice by the glass
                name = d.replace(" 1lt", " GLS")
                qty = rr.randint(10, 40)
                avg = m["glass"]
            else:
                qty = rr.randint(20, 120)
                avg = m["sell"]
            if d == "Reserve Shiraz 750ml":
                qty = 0                             # on the sheet, not sold tonight
            sold.append((name, d, qty, avg))
    # sold, but no row on the sheet yet: the script adds them to a spare row
    sold.append(("Craft Pale Ale 440ml", "Craft Pale Ale 440ml", 36, M("Craft Pale Ale 440ml")["sell"]))
    sold.append(("Silver Tequila", "Silver Tequila 750ml", 24, M("Silver Tequila 750ml")["tot"]))
    sold.append(("Corkage", None, 4, 60))           # not a stock item (item map: IGNORE)
    food = [("Canape Platter", 30, 180), ("Gala Dinner Plated", 300, 420), ("Dessert Station", 1, 9500)]

    # ---- the upgraded Micros report
    families = {}
    for name, d, qty, avg in sold:
        sec = next((lbl for lbl, items, _, _ in CELLAR_SECTIONS if d in items), None)
        if sec is not None:
            fam = sec or "Ent - Minerals Softdrink"          # the sheet's top section has no label
        else:
            fam = {"Craft Pale Ale 440ml": "Ent - Beer", "Silver Tequila 750ml": "Ent - White Spirits"}.get(d, "Ent - Sundries")
        families.setdefault(fam, []).append((name, qty, avg))
    groups = [("Beverage Enterprise", families), ("Food", {"Ent - Food": food})]
    tot_q = sum(q for _, fams in groups for items in fams.values() for _, q, _ in items)
    tot_s = sum(q * a for _, fams in groups for items in fams.values() for _, q, a in items)
    wb = Workbook()
    ws = wb.active
    ws.title = "Reports"
    ws["A3"], ws["B3"] = "Business Dates", day.strftime("%m/%d/%Y")
    ws["A4"], ws["B4"] = "Locations", CC_NAME
    for r, (k, v) in enumerate((("Revenue Centers", "All"), ("Order Types", "All"), ("Menu Items", "All"),
                                ("Day Parts", "All")), 5):
        ws.cell(r, 1, k)
        ws.cell(r, 2, v)
    ws["A10"] = "Major Group"
    for i, (maj, fams) in enumerate(groups):
        ws.cell(11 + i, 1, maj)
        ws.cell(11 + i, 2, r2(sum(q * a for items in fams.values() for _, q, a in items)))
    ws["A18"] = "Run Top Menu Items Summary Report"
    ws["A19"] = "Menu Item"
    top = sorted(((n, q * a) for items in families.values() for n, q, a in items), key=lambda t: -t[1])[:8]
    for i, (n, s) in enumerate(top):
        ws.cell(20 + i, 1, n)
        ws.cell(20 + i, 2, r2(s))
    head = ["Name", "Gross Sales Before Discounts Total", "Discounts Total", "Gross Sales After Discounts Total",
            "Sales Net VAT", "Percent of Total Sales Net VAT", "Quantity Sold", "Percent of Total Quantity Sold",
            "Percent of Major Group Sales", "Percent of Family Group Sales", "Average Price"]
    hr = 41
    for c, h in enumerate(head, 1):
        ws.cell(hr, c, h).font = BOLD

    def row(r, name, sales, qty, pct_major, pct_family):
        vals = [name, r2(sales), 0, r2(sales), r2(sales / 1.15), round(sales / tot_s, 4), qty,
                round(qty / tot_q, 4), pct_major, pct_family, r2(sales / qty) if qty else 0]
        for c, v in enumerate(vals, 1):
            ws.cell(r, c, v)
    r = hr + 1
    row(r, "Total", tot_s, tot_q, "-", "-")
    for maj, fams in groups:
        ms_ = sum(q * a for items in fams.values() for _, q, a in items)
        mq = sum(q for items in fams.values() for _, q, _ in items)
        r += 1
        row(r, maj, ms_, mq, "-", "-")
        for fam, items in fams.items():
            fs = sum(q * a for _, q, a in items)
            fq = sum(q for _, q, _ in items)
            r += 1
            row(r, fam, fs, fq, round(fs / ms_, 4) if ms_ else 0, "-")
            for n, q, a in items:
                r += 1
                row(r, n, q * a, q, round(q * a / ms_, 4) if ms_ else 0, round(q * a / fs, 4) if fs else 0)
    ws.column_dimensions["A"].width = 34
    save(wb, path("cellar_consumption", "inputs", "CC_Function_Sales_10_Nov_2026.xlsx"))

    # ---- the Cellar Consumption workbook: MM60, Master and the function's dated sheet
    def cellar_sheet(ws, counts):
        ws["B1"] = "CELLAR CONSUMPTION SHEET"
        ws.merge_cells("B1:K1")
        ws["B1"].font = Font(bold=True, size=14)
        ws["B2"] = "Function Name: DEMO GALA DINNER" if counts else "Function Name:"
        ws.merge_cells("C2:F2")
        ws["B3"] = "Date: 10 NOVEMBER 2026" if counts else "Date:"
        ws.merge_cells("C3:F3")
        ws["B4"] = "Pax  300" if counts else "Pax"
        heads = ["Sap Code", "Item Description", "Cost", "Open", "Add", "Close", "Usage", "Charged on Doc",
                 "Variance ", "Selling Price", "Total Revenue", "Total Cost", None, "Variance @ Cost"]
        for c, h in enumerate(heads, 1):
            if h:
                ws.cell(6, c, h).font = BOLD
        money = '\\R\\ #,##0.00;[Red]\\-\\R\\ #,##0.00;\\R\\ \\-'
        r = 7
        subtotal_rows = []
        for label, items, spare, serves in CELLAR_SECTIONS:
            sub = r
            subtotal_rows.append(sub)
            body_start = r + 1
            body_end = body_start + len(items) + spare - 1
            if label:
                ws.cell(sub, 2, label).font = BOLD
                ws.cell(sub, 9, f"=H{sub}-G{sub}")
            ws.cell(sub, 11, f"=SUM(K{body_start}:K{body_end})")
            ws.cell(sub, 12, f"=SUM(L{body_start}:L{body_end})")
            ws.cell(sub, 13, f"=+L{sub}/K{sub}")
            if label:
                ws.cell(sub, 14, f"=+I{sub}*C{sub}")
            for k, rr_ in enumerate(range(body_start, body_end + 1)):
                d = items[k] if k < len(items) else None
                if d:
                    ws.cell(rr_, 1, M(d)["mat"])
                ws.cell(rr_, 2, f"=VLOOKUP(A{rr_},'MM60'!B:I,2,FALSE)")
                ws.cell(rr_, 3, f"=VLOOKUP(A{rr_},'MM60'!B:I,5,FALSE)" + ("/30" if serves else ""))
                if d and counts:
                    o, a, c_ = counts[d]
                    ws.cell(rr_, 4, o)
                    if a:
                        ws.cell(rr_, 5, a)
                    ws.cell(rr_, 6, c_)
                ws.cell(rr_, 7, f"=D{rr_}+E{rr_}-F{rr_}")
                ws.cell(rr_, 9, f"=H{rr_}-G{rr_}")
                ws.cell(rr_, 11, f"=+J{rr_}*H{rr_}")
                ws.cell(rr_, 12, f"=+H{rr_}*C{rr_}")
                ws.cell(rr_, 14, f"=+I{rr_}*C{rr_}")
                for c in (3, 11, 12, 14):
                    ws.cell(rr_, c).number_format = money
            r = body_end + 1
        ws.cell(r, 2, "TOTAL REVENUE ").font = BOLD
        ws.cell(r, 11, "=+" + "+".join(f"K{s}" for s in reversed(subtotal_rows)))
        ws.cell(r, 12, "=+" + "+".join(f"L{s}" for s in reversed(subtotal_rows)))
        ws.cell(r, 13, f"=+L{r}/K{r}")
        ws.cell(r, 14, f"=+I{r}*C{r}")
        ws.cell(r + 1, 14, f"=SUM(N8:N{r})")
        for col, w in zip("ABCDEFGHIJKLMN", (9, 40, 10, 9, 9, 9, 10, 15, 12, 12, 14, 13, 3, 18)):
            ws.column_dimensions[col].width = w

    counts = {}
    for name, d, qty, avg in sold:
        if not d or not any(d in items for _, items, _, _ in CELLAR_SECTIONS):
            continue
        serves = next(s for _, items, _, s in CELLAR_SECTIONS if d in items)
        if serves:
            bottles = qty / 30.0
        elif name.endswith("GLS"):
            bottles = qty / 4.0
        else:
            bottles = qty
        used = math.ceil(bottles)
        missing = rr.choice([0, 0, 0, 1, 2]) if not serves else rr.choice([0, 0, 1])
        if d in ("House Merlot 750ml", "Single Malt 12yr 750ml", "Premium Lager 330ml"):
            missing += 3                             # clear losses for the demo
        used += missing
        opening = used + rr.randint(2, 12)
        add_ = rr.choice([0, 0, 6, 12, 24]) if not serves else rr.choice([0, 0, 2])
        counts[d] = (opening, add_, opening + add_ - used)
    wb = Workbook()
    mm = wb.active
    mm.title = "MM60"
    write_mm60_sheet(mm, CC)
    cellar_sheet(wb.create_sheet("Master"), None)
    cellar_sheet(wb.create_sheet(sheet_name), counts)
    wb.active = 2
    save(wb, path("cellar_consumption", "inputs", "Cellar_Consumption_10_November_2026.xlsx"))

    # ---- item map (current) and an older one that lacks six of the renamed items
    def item_map(p, keep_all):
        wb = Workbook()
        hu = wb.active
        hu.title = "How to use"
        for i, t in enumerate([
                "Cellar item map (demo data)",
                "The Cellar Consumption updater reads this file every run. Keep it next to update_cellar_consumption.py.",
                None, "Item Map",
                "One row per Micros item name that does not match the SAP name on its own.",
                "New items are asked about on screen with numbered choices; your answer is added here automatically "
                "(a dated backup of this file is taken first).",
                "SAP Code: type IGNORE for a POS item that is not a stock item, so it is never asked about again.",
                "Serves per Unit: normally blank. Fill it in only to force a conversion (30 = 25ml tots from 750ml).",
                None, "House Rules",
                "Pour size per SAP material group. 'Serves' = the cellar sheet counts these rows in serves.",
                "'Bottles' = the sheet counts whole bottles; glasses sold are converted to bottle fractions."], 1):
            hu.cell(i, 1, t)
        hu["A1"].font = Font(bold=True, size=12)
        im = wb.create_sheet("Item Map")
        im.append(["Micros Item", "SAP Code", "SAP Description", "Serves per Unit", "Source", "Notes"])
        for name, d in CELLAR_RENAMES.items():
            if keep_all or name in CELLAR_OLD_MAP_KEEPS:
                im.append([name, M(d)["mat"], d, None, "Seeded 14 Oct 2026", None])
        for name, d in CELLAR_MAP_ALWAYS.items():
            im.append([name, M(d)["mat"], d, 4, "Seeded 14 Oct 2026", "Juice by the 250ml glass: 4 per litre"])
        im.append(["Corkage", "IGNORE", "(not a stock item)", None, "Seeded 14 Oct 2026", "A charge, not a product"])
        hs = wb.create_sheet("House Rules")
        hs.append(["Material Group", "Category", "Pour ml", "Sheet counts in", "Notes"])
        for g, cat, pour, unit in [("BEV000100", "Aperitifs & bitters", 25, "Serves"), ("BEV000300", "Brandy", 25, "Serves"),
                                   ("BEV000500", "Cognac", 25, "Serves"), ("BEV000700", "Liqueurs", 25, "Serves"),
                                   ("BEV001100", "Cordials & syrups", 25, "Serves"), ("BEV001500", "Rum", 25, "Serves"),
                                   ("BEV001800", "Whisky", 25, "Serves"),
                                   ("BEV001900", "Gin, vodka & white spirits", 25, "Serves"),
                                   ("BEV001200", "Port & dessert wine", 50, "Bottles"),
                                   ("BEV001300", "Red wine", 187.5, "Bottles"), ("BEV001400", "Rose wine", 187.5, "Bottles"),
                                   ("BEV001600", "Sherry & fortified", 50, "Bottles"),
                                   ("BEV001700", "Sparkling", 125, "Bottles"), ("BEV002000", "White wine", 187.5, "Bottles")]:
            hs.append([g, cat, pour, unit, None])
        for c in list(im[1]) + list(hs[1]):
            c.font = BOLD
        im.column_dimensions["A"].width = 30
        im.column_dimensions["C"].width = 34
        save(wb, p)
    item_map(path("cellar_consumption", "tool_files", "cellar_item_map.xlsx"), True)
    item_map(path("cellar_consumption", "inputs", "cellar_item_map_OLD.xlsx"), False)

    # answers for the old-map run: one per item asked, in report order (plus one rejected answer)
    asked = [n for fam in families.values() for n, q, _ in fam
             if n in CELLAR_RENAMES and n not in CELLAR_OLD_MAP_KEEPS and q > 0]
    answers = []
    for i, n in enumerate(asked):
        if i == 2:
            answers.append("S")                     # skip this time
        elif i == 3:
            answers += ["99", "1"]                  # not a listed number: asked again
        else:
            answers.append("1")
    write_lines(path("cellar_consumption", "answers_cellar_prompts.txt"), answers)


# ---------------------------------------------------------------------------- kitchen MI20 and the loss trackers
BREAKFAST_SPLIT = {"Free Range Eggs 30s": 0.7, "White Bread Loaf 700g": 0.5, "Full Cream Milk 1lt": 0.4,
                   "Tomatoes 1kg": 0.6, "Baby Spinach 1kg": 0.0}
CANTEEN_SPLIT = {"Free Range Eggs 30s": 0.3, "White Bread Loaf 700g": 0.5, "Full Cream Milk 1lt": 0.3,
                 "Instant Coffee 750g": 0.0}
COFFEE_DRINKS = [  # name (exactly as the script lists it), selling price, [(desc, qty per serve)]
    ("Americano", 30, [("Espresso Beans 1kg", 0.016)]),
    ("Single Esspresso", 24, [("Espresso Beans 1kg", 0.008)]),
    ("Double Esspresso", 32, [("Espresso Beans 1kg", 0.016)]),
    ("Cappuccino", 36, [("Espresso Beans 1kg", 0.008), ("Barista Milk 1lt", 0.15)]),
    ("Café Latte", 38, [("Espresso Beans 1kg", 0.008), ("Barista Milk 1lt", 0.25)]),
    ("Decaf Americano", 32, [("Decaf Espresso Beans 1kg", 0.016)]),
    ("Decaf Cappuccino", 38, [("Decaf Espresso Beans 1kg", 0.008), ("Barista Milk 1lt", 0.15)]),
    ("Decaf Café Latte ", 40, [("Decaf Espresso Beans 1kg", 0.008), ("Barista Milk 1lt", 0.25)]),
    ("Café Mocha", 44, [("Espresso Beans 1kg", 0.008), ("Barista Milk 1lt", 0.2), ("Hot Chocolate Powder 1kg", 0.02),
                        ("Chocolate Sauce 1lt", 0.015)]),
    ("Cortado", 32, [("Espresso Beans 1kg", 0.008), ("Barista Milk 1lt", 0.06)]),
]
COFFEE_SOLD = [446, 12, 30, 500, 200, 15, 10, 5, 40, 25]       # the quantities the tests type in


def kitchen():
    rr = random.Random(f"{SEED}-kitchen")
    theo = {}
    for (name, _, ings), sold in zip(COFFEE_DRINKS, COFFEE_SOLD):
        for d, q in ings:
            theo[d] = theo.get(d, 0) + q * sold
    rows = []
    items = [d for d, *_ in BREAKFAST] + [d for d, *_ in CANTEEN] + [d for d, *_ in COFFEE] + \
        [d for d, *_ in KITCHEN_OTHER] + [d for d, *_ in GARNISH]
    for i, d in enumerate(items, 1):
        m = M(d)
        pr = price(d)
        book = round(rr.uniform(8, 60), 1 if m["uom"] in ("KG", "L") else 0)
        if d in theo:
            # coffee lines: recorded as usage, a little above what the drinks sold explain
            diff = round(theo[d] * rr.uniform(1.12, 1.3), 2)
        elif any(d == x[0] for x in CANTEEN):
            diff = -round(rr.uniform(8000, 16000) / pr, 1)
        elif any(d == x[0] for x in BREAKFAST):
            diff = -round(rr.uniform(500, 2400) / pr, 1)
        else:
            roll = rr.random()
            diff = 0.0 if roll < 0.15 else (round(rr.uniform(1, 6), 1) if roll < 0.25 else -round(rr.uniform(300, 4500) / pr, 1))
        if m["uom"] == "EA":
            diff = float(round(diff))
        counted = round(book + diff, 2)
        val = r2(diff * pr)
        rows.append([f"3000{KITCHEN_LOC[1:]}", str(i), str(m["mat"]), d, HOTEL, KITCHEN_LOC, book, counted, diff, val,
                     val if val > 0 else 0, val if val < 0 else 0])

    def mi20(p):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["PhysInvDoc", "Item", "Material", "Material Description", "Plnt", "Location", "Book Qty",
                   "Qty Counted", "Difference qty", "Difference", "Pos. diff. value", "neg. diff. value"])
        for c in ws[1]:
            c.font = BOLD
        for r in rows:
            ws.append(r)
        ws.column_dimensions["D"].width = 30
        save(wb, p)
    name = "MI20_Kitchen_Oct_2026.xlsx"
    for tool in ("coffee_loss", "breakfast_tracker", "canteen_loss"):
        mi20(path(tool, "inputs", name))

    # masters
    def master(p, sheet, title, heads, codes, split, note_col):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=12)
        ws["A2"] = "Item list for the month-end rebuild. A split % below 100% shares an item with another outlet."
        ws.append(heads)
        for c in ws[3]:
            c.font = BOLD
        for d in codes:
            pct = split.get(d, 1.0)
            note = "" if pct == 1 else ("Not this outlet's cost" if pct == 0 else f"Shared item: {int(pct * 100)}% here")
            ws.append([M(d)["mat"], d, r2(rr.uniform(150, 2400)) if note_col == "value" else "Chef confirmed", pct, note or None])
        ws.column_dimensions["B"].width = 32
        save(wb, p)
    master(path("breakfast_tracker", "tool_files", "Breakfast_Item_Master.xlsx"), "Breakfast Items",
           "BREAKFAST ITEM MASTER", ["Material Code", "Material Description", "Diff Value (Mar)", "Breakfast %", "Notes"],
           [d for d, *_ in BREAKFAST], BREAKFAST_SPLIT, "value")
    master(path("canteen_loss", "tool_files", "Canteen_Item_Master.xlsx"), "Canteen Items",
           "STAFF CANTEEN ITEM MASTER", ["Material Code", "Material Description", "Chef Notes (from SAP Dec)", "Canteen %",
                                         "Split Reason / Notes"],
           [d for d, *_ in CANTEEN] + ["Free Range Eggs 30s", "White Bread Loaf 700g", "Full Cream Milk 1lt"],
           CANTEEN_SPLIT, "chef")

    # last month's workbooks, so the Trend tab carries July-September forward
    def trend_book(p, title, labels, values):
        wb = Workbook()
        ws = wb.active
        ws.title = "Trend"
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=12)
        ws["A2"] = "Demo history: July to September 2026."
        for c, mname in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 2):
            ws.cell(4, c, mname).font = BOLD
        for row, label in labels.items():
            ws.cell(row, 1, label).font = BOLD
        for col, vals in values.items():
            for row, v in vals.items():
                ws.cell(row, col, v)
        ws.column_dimensions["A"].width = 28
        save(wb, p)
    trend_book(path("breakfast_tracker", "tool_files", "Breakfast_MonthEnd_Tool.xlsx"), "BREAKFAST  |  Month-on-Month Trend",
               {5: "Total food cost (Mi20)", 6: "Breakfast cost", 7: "Breakfast % of food", 8: "Breakfast revenue",
                10: "Breakfast covers"},
               {8: {5: 402150.0, 6: 41320.5, 7: 0.1027, 8: 168400.0, 10: 3920},
                9: {5: 389940.0, 6: 39870.2, 7: 0.1022, 8: 159250.0, 10: 3710},
                10: {5: 415300.0, 6: 44110.8, 7: 0.1062, 8: 171900.0, 10: 4010}})
    trend_book(path("canteen_loss", "tool_files", "Canteen_MonthEnd_Tool.xlsx"), "STAFF CANTEEN  |  Month-on-Month Trend",
               {5: "Canteen cost", 6: "Budget", 9: "Staff meals served", 10: "Cost per meal"},
               {8: {5: 236400.0, 6: 250000, 9: 15800, 10: 14.96},
                9: {5: 244950.0, 6: 250000, 9: 16120, 10: 15.2},
                10: {5: 251300.0, 6: 250000, 9: 16440, 10: 15.29}})

    # coffee costing workbook (recipes read live by the tracker)
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["B2"], ws["C2"] = "Menu", "Recipe"
    mm = wb.create_sheet("MM60")
    write_mm60_sheet(mm, HOTEL)
    for i, (name, sell, ings) in enumerate(COFFEE_DRINKS, 1):
        s = wb.create_sheet(f"Menu {i}")
        s["B2"], s["C2"] = "Recipe:", name
        s["C2"].font = Font(bold=True, size=12)
        for c, h in enumerate(["SAP Nr", "ITEM DESCRIPTION (SAP)", "QTY - (kg, ea, lt)", "UOM", "SAP VALUE", "TOTAL VALUE"], 2):
            s.cell(4, c, h).font = BOLD
        sub = 0.0
        for r, (d, q) in enumerate(ings, 5):
            pr = price(d)
            cost = r2(q * pr)
            sub += cost
            for c, v in enumerate([M(d)["mat"], d, q, M(d)["uom"], pr, cost], 2):
                s.cell(r, c, v)
        misc = r2(sub * 0.03)
        for row, label, val in ((27, "Miscellaneous 3%", misc), (28, "Menu item cost", r2(sub + misc)),
                                (33, "Selling price incl VAT", sell)):
            s.cell(row, 2, label).font = BOLD
            s.cell(row, 7, val)
        ws.cell(2 + i, 2, f"Menu {i}")
        ws.cell(2 + i, 3, name)
        s.column_dimensions["C"].width = 30
    save(wb, path("coffee_loss", "tool_files", "Coffee_Costing_Updated.xlsx"))


# ---------------------------------------------------------------------------- Food Sales Mix
FOOD_MENU = {
    "main": {
        "Starters": ["Soup of the Day", "Caesar Salad", "Peri-Peri Chicken Livers", "Crumbed Calamari", "Tomato Bruschetta"],
        "Mains": ["Chicken Curry", "Beef Lasagne", "Vegetable Stir Fry", "Grilled Linefish", "Braised Lamb Shank",
                  "Mushroom Risotto"],
        "Grills": ["Beef Fillet 250g", "Rump Steak 300g", "Sirloin 300g", "Pork Ribs 600g", "Half Roast Chicken"],
        "Burgers & Sandwiches": ["Classic Beef Burger", "Chicken Burger", "Club Sandwich", "Toasted Cheese & Tomato",
                                 "Vegetable Wrap"],
        "Desserts": ["Malva Pudding", "Chocolate Brownie", "Baked Cheesecake", "Ice Cream Trio", "Fresh Fruit Salad"],
        "Kids": ["Kids Chicken Strips", "Kids Spaghetti Bolognaise", "Kids Mini Burger"],
    },
    "tower": {
        "Light Meals": ["Soup and Bread", "Greek Salad", "Chicken Wrap", "Quiche of the Day", "Smoked Salmon Bagel"],
        "Mains": ["Butter Chicken", "Spaghetti Bolognaise", "Fish and Chips", "Beef Burger"],
        "Platters": ["Cheese Board", "Charcuterie Board", "Mezze Platter"],
        "Desserts": ["Lemon Tart", "Chocolate Mousse", "Carrot Cake Slice", "Ice Cream Sundae"],
    },
}
# POS names that differ from the workbook, and what each one demonstrates
FOOD_POS_EXTRAS = {
    "main": [("Line Fish of the Day", "Mains", "alias"), ("Chocolate Browny", "Desserts", "auto"),
             ("Beef Burger Deluxe", "Burgers & Sandwiches", "flag"), ("Soup of the Week", "Starters", "excluded"),
             ("Prawn Linguine", "Mains", "new"), ("Extra Chips", "Sides", "skip"), ("Side Salad", "Sides", "skip")],
    "tower": [("Fish & Chips", "Mains", "alias"), ("Carrot Cake Slices", "Desserts", "auto"),
              ("Mezze Plate for Two", "Platters", "flag"), ("Gourmet Toastie", "Light Meals", "new"),
              ("Bread Roll Extra", "Sides", "skip")],
}
FOOD_CONFIG = {
    "main": {"aliases": {"Line Fish of the Day": "Grilled Linefish"}, "exclusions": ["Soup of the Week"],
             "additions": [{"pos_name": "Prawn Linguine", "mix_name": "Prawn Linguine", "category": "Mains"}]},
    "tower": {"aliases": {"Fish & Chips": "Fish and Chips"}, "exclusions": [],
              "additions": [{"pos_name": "Gourmet Toastie", "mix_name": "Gourmet Toastie", "category": "Light Meals"}]},
}
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]


def food_sales_mix():
    for unit, label, wb_name, suffix in (("main", "Main Hotel", "Food_Sales_Mix_Main.xlsx", ""),
                                         ("tower", "Tower Wing", "Food_Sales_Mix_Tower.xlsx", "_tower")):
        rr = random.Random(f"{SEED}-food-{unit}")
        menu = FOOD_MENU[unit]
        wb = Workbook()
        wb.remove(wb.active)
        for year, filled in ((2025, 12), (2026, 9)):
            ws = wb.create_sheet(str(year))
            ws.cell(2, 1, "Sales mix for Month of ").font = BOLD
            for i, mname in enumerate(MONTHS):
                ws.cell(1, 2 + i, year)
                ws.cell(2, 2 + i, "Febuary" if mname == "February" and year == 2025 else mname).font = BOLD
            r = 4
            for cat, items in menu.items():
                c = ws.cell(r, 1, cat)
                c.font = BOLD
                c.fill = GREEN
                for col in range(2, 14):
                    ws.cell(r, col).fill = GREEN
                r += 1
                for it in items:
                    ws.cell(r, 1, it)
                    for i in range(filled):
                        cell = ws.cell(r, 2 + i, rr.randint(20, 260))
                        cell.number_format = "#,##0"
                    for i in range(filled, 12):
                        ws.cell(r, 2 + i).number_format = "#,##0"
                    r += 1
            ws.column_dimensions["A"].width = 30
            ws.freeze_panes = "B3"
        save(wb, path("food_sales_mix", "tool_files", wb_name))
        cfg = FOOD_CONFIG[unit]
        write_json(path("food_sales_mix", "tool_files", f"item_aliases{suffix}.json"),
                   dict({"_comment": "POS name -> Sales Mix row label to write into"}, **cfg["aliases"]))
        write_json(path("food_sales_mix", "tool_files", f"item_exclusions{suffix}.json"),
                   {"_comment": "POS names confirmed NOT the same as any similar mix row", "items": cfg["exclusions"]})
        write_json(path("food_sales_mix", "tool_files", f"item_additions{suffix}.json"),
                   {"_comment": "Confirmed new dishes: a row is added under the category on the next run",
                    "pending": cfg["additions"]})
        lines = []
        for cat, items in menu.items():
            for it in items:
                if unit == "main" and it in ("Grilled Linefish", "Chocolate Brownie"):
                    continue                   # sold under another POS name this month (see extras)
                if unit == "tower" and it in ("Fish and Chips", "Carrot Cake Slice", "Mezze Platter"):
                    continue
                lines.append((it, cat, "Food", rr.randint(15, 240), r2(rr.uniform(65, 245))))
        for it, cat, _ in FOOD_POS_EXTRAS[unit]:
            lines.append((it, cat, "Food", rr.randint(8, 120), r2(rr.uniform(35, 230))))
        wb = Workbook()
        sales_mix_sheet(wb.active, "10/1/2026 - 10/31/2026", PROPERTY, f"{label} Restaurant", "Food", lines)
        save(wb, path("food_sales_mix", "inputs", f"{label.replace(' ', '_')}_Food_Oct_2026.xlsx"))


# ---------------------------------------------------------------------------- Top & Bottom Sellers
def top_bottom():
    food = [it for unit in FOOD_MENU.values() for items in unit.values() for it in items]
    bev = [d.replace(" 750ml", "") for d, *_ in SPIRITS] + [d for d, *_ in WINES] + [d for d, *_ in UNITS] + \
        [c[0] for c in COCKTAILS]
    write_json(path("top_bottom_sellers", "tool_files", "seed_items.json"),
               {"food": sorted(set(food)), "bev": sorted(set(bev))})


# ---------------------------------------------------------------------------- MM60 Cost Comparator
def comparator():
    mm60_file(path("cost_comparator", "inputs", "MM60_D100_Sep_2026.xlsx"), HOTEL, "sep")
    mm60_file(path("cost_comparator", "inputs", "MM60_D100_Oct_2026.xlsx"), HOTEL, "oct")
    mm60_file(path("cost_comparator", "inputs", "MM60_D200_Oct_2026.xlsx"), CC, "oct")


# ---------------------------------------------------------------------------- Stock Age
MB52_HEAD = ["Material", "Material Description", "Storage Location", "Plant", "Material Group", "Base Unit of Measure",
             "Unrestricted", "Value Unrestricted", "DF stor. loc. level"]
MB51_HEAD = ["Plant", "Storage Location", "Movement Type", "Material Group", "Cost Center", "Purchase order",
             "Material Doc.Item", "Material Document", "Material", "Material Description", "Posting Date",
             "Qty in unit of entry", "Amt.in Loc.Cur.", "Order Price Unit", "Unit of Entry"]


def stock_age():
    rr = random.Random(f"{SEED}-age")
    asof = dt.date(2026, 10, 31)
    pool = [m for m in MATS if m["price"] > 0 and m["block"] in ("bev", "food", "gen")] + \
        [m for m in MATS if m["block"] == "oe"][:20]
    stock_rows, purchases = [], []
    doc = 5000100000
    for m in pool:
        loc = "L011" if m["block"] == "bev" else "L010"
        qty = rr.randint(2, 80)
        stock_rows.append([m["mat"], m["desc"], loc, HOTEL, m["group"], m["uom"], qty, r2(qty * m["price"]), None])
        roll = rr.random()
        if roll < 0.55:
            last_days = rr.randint(5, 360)
        elif roll < 0.70:
            last_days = rr.randint(370, 720)
        elif roll < 0.80:
            last_days = rr.randint(735, 1090)
        elif roll < 0.93:
            last_days = rr.randint(1100, 1820)
        elif roll < 0.97:
            last_days = None                       # never bought in 5 years
        else:
            last_days = rr.randint(1830, 2400)     # older than the 5-year export
        if last_days is None or last_days > 1825:
            continue
        n = rr.randint(1, 6)
        days = sorted({last_days} | {rr.randint(last_days, 1825) for _ in range(n - 1)})
        for d in days:
            doc += 1
            q = rr.randint(6, 60)
            purchases.append([HOTEL, loc, 101, m["group"], None, f"45000{rr.randint(10000, 99999)}", 1, str(doc),
                              m["mat"], m["desc"], dt.datetime.combine(asof - dt.timedelta(days=d), dt.time()),
                              q, r2(q * m["price"]), m["uom"], m["uom"]])
    purchases.sort(key=lambda r: r[10])
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock on hand"
    ws.append(MB52_HEAD)
    for r in stock_rows:
        ws.append(r)
    for title, rows in (("Past year purchase", [p for p in purchases if (asof - p[10].date()).days <= 365]),
                        ("Past 5 year purchase", purchases)):
        s = wb.create_sheet(title)
        s.append(MB51_HEAD)
        for r in rows:
            s.append(r)
            s.cell(s.max_row, 11).number_format = "yyyy-mm-dd"
    for s in wb.worksheets:
        for c in s[1]:
            c.font = BOLD
        s.column_dimensions["B" if s.title == "Stock on hand" else "J"].width = 32
    save(wb, path("stock_age", "inputs", "Stock_Age_Source.xlsx"))


# ---------------------------------------------------------------------------- OE
OE_LOCATIONS = [  # code, on approval, notes, approval file, approver 1, approver 2
    ("L080", "N", "Ordered by the kitchen", "", "", ""),
    ("L081", "N", "Ordered by Admin", "", "", ""),
    ("L083", "N", "Ordered by Housekeeping", "", "", ""),
    ("L095", "Y", "", "", "", ""), ("L096", "Y", "", "", "", ""), ("L097", "Y", "", "", "", ""),
    ("L098", "Y", "", "", "", ""), ("L099", "Y", "", "", "", ""), ("L100", "Y", "", "", "", ""),
    ("L101", "Y", "", "Linen Main", "Executive Housekeeper (Main Hotel)", "Rooms Division Manager"),
    ("L102", "Y", "", "Linen Main", "Executive Housekeeper (Main Hotel)", "Rooms Division Manager"),
    ("L108", "Y", "", "", "", ""), ("L112", "Y", "Banqueting stock: checked after each function", "", "", ""),
    ("L228", "Y", "", "", "", ""), ("L229", "Y", "", "", "", ""), ("L230", "Y", "", "", "", ""),
    ("L231", "Y", "", "", "", ""), ("L232", "Y", "", "", "", ""), ("L233", "Y", "", "", "", ""),
    ("L234", "Y", "", "Linen Tower", "Executive Housekeeper (Tower Wing)", "Rooms Division Manager"),
    ("L235", "Y", "", "Linen Tower", "Executive Housekeeper (Tower Wing)", "Rooms Division Manager"),
]


def oe(python):
    rr = random.Random(f"{SEED}-oe")
    pr = PRICE[HOTEL]
    stock, pars, notes, first_seen = {}, {}, {}, {}
    for d, g, p, locs in OE:
        m = M(d)
        for i, (loc, par) in enumerate(locs):
            roll = rr.random()
            if roll < 0.28:
                ratio = rr.uniform(0.15, 0.48)          # well under par: HIGH
            elif roll < 0.55:
                ratio = rr.uniform(0.55, 0.95)          # under par
            elif roll < 0.85:
                ratio = rr.uniform(0.95, 1.15)
            else:
                ratio = rr.uniform(1.4, 2.2)            # over par: a transfer option for sister locations
            stock[(loc, m["mat"])] = max(0, int(round(par * ratio)))
            pars[(loc, m["mat"])] = par
            first_seen[(loc, m["mat"])] = rr.choice(["Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026"])
    notes[("L112", M("Chafing Dish Full Size")["mat"])] = "Two out for repair"
    notes[("L097", M("Cocktail Glass 240ml")["mat"])] = "Moving to the new cocktail glass in December"
    new_item = ("L108", M("Espresso Cup 90ml")["mat"])           # in MB52 this month, no par yet
    pars.pop(new_item, None)
    first_seen.pop(new_item, None)
    gone = ("L099", M("Balloon Whisk 30cm")["mat"])              # par 0, no stock left: dropped with a check
    pars[gone] = 0
    stock.pop(gone, None)

    # MB52 (October), MB51 scrapping (October), MM60 (October)
    rows, total = [], 0.0
    for (loc, mat), q in sorted(stock.items()):
        m = next(x for x in MATS if x["mat"] == mat)
        val = r2(q * pr[mat])
        total += val
        rows.append([mat, m["desc"], loc, HOTEL, m["group"], m["uom"], q, val, None])
    rows.append([None, None, None, None, None, None, None, r2(total), None])
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(MB52_HEAD)
    for r in rows:
        ws.append(r)
    ws.column_dimensions["B"].width = 30
    save(wb, path("oe_monthly", "inputs", "MB52_OE_Oct_2026.xlsx"))

    mb51, amt_total, doc = [], 0.0, 4900021000
    breakable = [(loc, M(d)["mat"], d) for d, g, p, locs in OE if g in ("CRO000100", "GLA000100", "CUT000100")
                 for loc, _ in locs]
    scrapped = rr.sample(breakable, 24)
    for loc, mat, d in scrapped:
        q = rr.randint(2, 18)
        doc += 1
        amt = -r2(q * pr[mat])
        amt_total += amt
        mb51.append([HOTEL, loc, 551, M(d)["group"], "910010", None, 1, str(doc), mat, d,
                     dt.datetime(2026, 10, rr.randint(1, 28)), -q, amt, "EA", "EA"])
    loc, mat, d = scrapped[0]                                  # one breakage partly reversed (552)
    doc += 1
    mb51.append([HOTEL, loc, 552, M(d)["group"], "910010", None, 1, str(doc), mat, d, dt.datetime(2026, 10, 29), 1,
                 r2(pr[mat]), "EA", "EA"])
    amt_total += r2(pr[mat])
    for k in range(2):                                         # not scrapping: ignored by the tracker
        loc, mat, d = breakable[5 + k]
        doc += 1
        mb51.append([HOTEL, loc, 201, M(d)["group"], "910010", None, 1, str(doc), mat, d, dt.datetime(2026, 10, 15),
                     -3, -r2(3 * pr[mat]), "EA", "EA"])
    mb51.sort(key=lambda r: r[10])
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(MB51_HEAD)
    for r in mb51:
        ws.append(r)
        ws.cell(ws.max_row, 11).number_format = "yyyy-mm-dd"
    ws.append([None] * 12 + [r2(amt_total), None, None])
    ws.column_dimensions["J"].width = 30
    save(wb, path("oe_monthly", "inputs", "MB51_Oct_2026.xlsx"))
    mm60_file(path("oe_monthly", "inputs", "MM60_D100_Oct_2026.xlsx"), HOTEL, "oct")

    # last month's tracking workbook: LOCATIONS settings and par levels per location tab
    wb = Workbook()
    ws = wb.active
    ws.title = "LOCATIONS"
    ws["A1"] = "LOCATIONS  |  names and approval settings"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Demo data. Approver names are role titles only."
    head = ["Code", "Name", "On Approval (Y/N)", "Notes", "Approval File", "Approver 1", "Approver 2"]
    for c, h in enumerate(head, 1):
        ws.cell(4, c, h).font = BOLD
    for r, (code, appr, note, file_, a1, a2) in enumerate(OE_LOCATIONS, 5):
        for c, v in enumerate([code, LOC_NAMES[code], appr, note or None, file_ or None, a1 or None, a2 or None], 1):
            ws.cell(r, c, v)
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["F"].width = 36
    cols = ["Material", "Description", "Unit Price (MM60)", "Stock (MB52)", "Hold Value (MB52)", "Scrap Qty (MB51)",
            "Scrap Value (MB51)", "Par Level (EA)", "Order Qty", "Order Value", "Urgency", "Flag", "Notes",
            "First Seen", "Par Group (order code)"]
    by_loc = {}
    for (loc, mat), par in pars.items():
        by_loc.setdefault(loc, []).append((mat, par))
    sep_prices = mm60_prices(HOTEL, "sep")
    for loc in sorted(by_loc):
        s = wb.create_sheet(loc)
        s["A1"] = f"{loc}  {LOC_NAMES[loc]}"
        s["A1"].font = Font(bold=True, size=14)
        for c, h in enumerate(cols, 1):
            s.cell(4, c, h).font = BOLD
        for r, (mat, par) in enumerate(sorted(by_loc[loc]), 5):
            m = next(x for x in MATS if x["mat"] == mat)
            s.cell(r, 1, mat)
            s.cell(r, 2, m["desc"])
            s.cell(r, 3, sep_prices[mat])
            s.cell(r, 8, par)
            if (loc, mat) in notes:
                s.cell(r, 13, notes[(loc, mat)])
            s.cell(r, 14, first_seen.get((loc, mat), "Sep 2026"))
        s.column_dimensions["B"].width = 30
    save(wb, path("oe_monthly", "tool_files", "OE_Tracking_Workbook.xlsx"))

    # a signed approval file: run the tracker once on the October files and answer it
    tmp = tempfile.mkdtemp(prefix="ccs_demo_oe_")
    try:
        for f in ("oe_tracker.py",):
            shutil.copy2(os.path.join(ROOT, "source_tools", "oe_monthly", f), tmp)
        shutil.copy2(path("oe_monthly", "tool_files", "OE_Tracking_Workbook.xlsx"), tmp)
        ins = [path("oe_monthly", "inputs", n) for n in ("MB52_OE_Oct_2026.xlsx", "MM60_D100_Oct_2026.xlsx",
                                                         "MB51_Oct_2026.xlsx")]
        env = dict(os.environ, PYTHONHASHSEED="0", PYTHONIOENCODING="utf-8", CCS_PROPERTY_NAME=PROPERTY)
        p = subprocess.run([python, "-X", "utf8", os.path.join(tmp, "oe_tracker.py"), *ins, "--period", "Oct 2026"],
                           cwd=tmp, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
        if p.returncode != 0:
            raise SystemExit("oe_tracker.py failed while making the signed approval file:\n" + p.stdout[-2000:] + p.stderr[-2000:])
        src = os.path.join(tmp, "OE_Approval_Oct_2026.xlsx")
        wb = load_workbook(src)
        ws = wb["APPROVAL"]
        hdr = next(r for r in range(1, 21) if any(str(ws.cell(r, c).value or "").strip().lower() == "approve (y/n)"
                                                    for c in range(1, ws.max_column + 1)))
        col = {str(ws.cell(hdr, c).value or "").strip().lower(): c for c in range(1, ws.max_column + 1)}
        c_dec, c_qty, c_sugg = col["approve (y/n)"], col["approved qty"], col["suggested qty"]
        c_notes = col.get("notes")
        rr2 = random.Random(f"{SEED}-approve")
        r = hdr + 1
        n = 0
        while ws.cell(r, 1).value and str(ws.cell(r, 1).value).strip().upper().startswith("L"):
            sugg = ws.cell(r, c_sugg).value or 0
            roll = rr2.random()
            if n == 3:
                ws.cell(r, c_dec, "Maybe")                    # not Y or N: left pending, with a note
            elif n == 5:
                ws.cell(r, c_dec, "Y")
                ws.cell(r, c_qty, int(sugg) + 6)              # more than suggested: flagged
            elif n == 8:
                ws.cell(r, c_dec, "Y")
                ws.cell(r, c_qty, 0)                          # Y with qty 0: treated as not approved
            elif roll < 0.68:
                ws.cell(r, c_dec, "Y")
                ws.cell(r, c_qty, int(sugg) if rr2.random() < 0.8 else max(1, int(sugg * 0.5)))
            elif roll < 0.86:
                ws.cell(r, c_dec, "N")
                if c_notes:
                    ws.cell(r, c_notes, rr2.choice(["Transfer from the Tower Wing instead", "Enough for this month",
                                                    "Wait for the new range"]))
            n += 1
            r += 1
        if "_INFO" in wb.sheetnames:                          # pin the build time so every run writes the same file
            for row in wb["_INFO"].iter_rows(min_col=2, max_col=2):
                for c in row:
                    if isinstance(c.value, str) and len(c.value) == 19 and c.value[4] == "-" and c.value[13] == ":":
                        c.value = "2026-11-01 08:00:00"
        save(wb, path("oe_monthly", "inputs", "OE_Approval_Oct_2026.xlsx"))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------- main
def main(argv):
    global OUT
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--out", default=os.path.join(ROOT, "sample_data"))
    ap.add_argument("--python", default=sys.executable, help="Python used to run oe_tracker.py once")
    a = ap.parse_args(argv)
    OUT = os.path.abspath(a.out)
    if os.path.isdir(OUT):
        # keep expected/ (goldens written by the tests) but rebuild everything else
        for section in os.listdir(OUT):
            sp = os.path.join(OUT, section)
            if not os.path.isdir(sp):
                continue
            for tool in os.listdir(sp):
                for sub in os.listdir(os.path.join(sp, tool)):
                    if sub != "expected":
                        q = os.path.join(sp, tool, sub)
                        shutil.rmtree(q) if os.path.isdir(q) else os.remove(q)
    os.makedirs(OUT, exist_ok=True)
    theo_cos()
    recount()
    cellar()
    kitchen()
    food_sales_mix()
    top_bottom()
    comparator()
    stock_age()
    oe(a.python)
    with open(os.path.join(OUT, "README.md"), "w", encoding="utf-8", newline="\n") as f:
        f.write("# Sample data\n\n(c) 2026 Alain Yon. Generated by `tools/generate_dummy_data.py` (seed "
                f"{SEED}). Every name, number and price here is invented; nothing comes from a real business.\n\n"
                "Per tool: `tool_files/` are the tool's working files (copied into its folder on install), "
                "`inputs/` are the files to drop on it, `expected/` are the outputs the regression tests compare "
                "against.\n")
    n = sum(len(fs) for _, _, fs in os.walk(OUT))
    print(f"sample data written to {OUT}: {n} files, {len(MATS)} materials")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
