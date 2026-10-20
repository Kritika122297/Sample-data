"""
Build BriefForge test data from Open Food Facts.

Usage:
  pip install requests openpyxl
  python make_test_data.py --query granola --count 12
  python make_test_data.py --offline            # no internet, uses a built-in sample

Output (in ./test_data):
  clean/<barcode>.xlsx     content sheets with no planted errors
  seeded/<barcode>.xlsx    same products with 1-3 planted errors
  ground_truth.json        which errors were planted in which file (for your scoreboard)
  images/<barcode>_*.jpg   packaging photos (only with --images)
"""
import argparse, copy, json, os, random
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

API = "https://world.openfoodfacts.org/api/v2/search"
HEADERS = {"User-Agent": "BriefForgeHackathon/0.1 (team test data script)"}
FIELDS = ("code,product_name,brands,quantity,product_quantity,serving_size,serving_quantity,"
          "ingredients_text,allergens_tags,labels_tags,nutriments,image_front_url,image_ingredients_url")

DEMO = {
    "code": "0000000000001", "product_name": "Demo Granola Berry", "brands": "Demo Brand",
    "quantity": "480 g", "product_quantity": "480", "serving_size": "60 g", "serving_quantity": 60,
    "ingredients_text": "Oats, almonds, pecans, sunflower seeds, dried cranberries, coconut oil, salt",
    "allergens_tags": ["en:nuts"], "labels_tags": ["en:vegan"],
    "nutriments": {"energy-kcal_100g": 450, "fat_100g": 20, "saturated-fat_100g": 5,
                   "trans-fat_100g": 0, "cholesterol_100g": 0, "sodium_100g": 0.05,
                   "carbohydrates_100g": 58, "fiber_100g": 8, "sugars_100g": 14, "proteins_100g": 10},
}

def fetch(query, count):
    params = {"categories_tags_en": query, "countries_tags_en": "united-states",
              "page_size": count * 3, "fields": FIELDS}
    r = requests.get(API, params=params, headers=HEADERS, timeout=30)
    r.raise_for_status()
    out = []
    for p in r.json().get("products", []):
        n = p.get("nutriments", {})
        if p.get("ingredients_text") and p.get("product_quantity") and "carbohydrates_100g" in n:
            out.append(p)
        if len(out) >= count:
            break
    return out

def to_record(p):
    n = p.get("nutriments", {})
    g = lambda k: round(float(n.get(k, 0) or 0), 2)
    net = float(p.get("product_quantity") or 0)
    serving = float(p.get("serving_quantity") or 0) or 40.0
    return {
        "barcode": p["code"], "name": p.get("product_name", ""), "brand": p.get("brands", ""),
        "net_weight_g": net, "serving_size_g": serving,
        "servings_per_container": round(net / serving) if serving else 0,
        "ingredients": p.get("ingredients_text", ""),
        "allergens": [a.split(":")[-1] for a in p.get("allergens_tags", [])],
        "labels": [l.split(":")[-1] for l in p.get("labels_tags", [])],
        "claims": [],
        "nutrition_per_serving": {
            "calories": round(g("energy-kcal_100g") * serving / 100),
            "total_fat_g": round(g("fat_100g") * serving / 100, 1),
            "sat_fat_g": round(g("saturated-fat_100g") * serving / 100, 1),
            "trans_fat_g": round(g("trans-fat_100g") * serving / 100, 1),
            "cholesterol_mg": round(g("cholesterol_100g") * serving * 10, 1),
            "sodium_mg": round(g("sodium_100g") * serving * 10, 1),
            "total_carb_g": round(g("carbohydrates_100g") * serving / 100, 1),
            "fiber_g": round(g("fiber_100g") * serving / 100, 1),
            "total_sugars_g": round(g("sugars_100g") * serving / 100, 1),
            "protein_g": round(g("proteins_100g") * serving / 100, 1),
        },
    }

# ---- error injectors: each returns a description, or None if not applicable ----
def e_sugar_gt_carbs(r):
    nu = r["nutrition_per_serving"]
    nu["total_sugars_g"] = round(nu["total_carb_g"] + random.randint(2, 6), 1)
    return "total_sugars_g greater than total_carb_g"

def e_servings_mismatch(r):
    r["servings_per_container"] += random.choice([2, 3, 4])
    return "servings_per_container x serving_size does not match net_weight_g"

def e_calories_mismatch(r):
    r["nutrition_per_serving"]["calories"] = int(r["nutrition_per_serving"]["calories"] * 0.5) + 1
    return "calories do not match fat/carb/protein (9/4/4)"

def e_vegan_cholesterol(r):
    if "vegan" not in r["labels"]:
        r["labels"].append("vegan")
    r["nutrition_per_serving"]["cholesterol_mg"] = 45.0
    return "vegan product with cholesterol 45 mg"

def e_no_added_sugar_honey(r):
    r["claims"].append("No Added Sugar")
    r["ingredients"] += ", honey, coconut sugar"
    return "'No Added Sugar' claim but ingredients contain honey and coconut sugar"

def e_may_contain_listed(r):
    r["allergens"] = [f"may contain {a}" for a in (r["allergens"] or ["almonds"])]
    if "almond" not in r["ingredients"].lower():
        r["ingredients"] += ", almonds"
    return "allergen listed as 'may contain' but present in ingredients"

def e_duplicate_ingredient(r):
    parts = [x.strip() for x in r["ingredients"].split(",") if x.strip()]
    if parts:
        r["ingredients"] += ", " + random.choice(parts)
    return "ingredient listed twice"

def e_trans_gt_sat(r):
    nu = r["nutrition_per_serving"]
    nu["trans_fat_g"] = round(nu["sat_fat_g"] + 3, 1)
    return "trans fat greater than saturated fat"

INJECTORS = {
    "sugar_gt_carbs": e_sugar_gt_carbs, "servings_mismatch": e_servings_mismatch,
    "calories_mismatch": e_calories_mismatch, "vegan_cholesterol": e_vegan_cholesterol,
    "no_added_sugar_honey": e_no_added_sugar_honey, "may_contain_listed": e_may_contain_listed,
    "duplicate_ingredient": e_duplicate_ingredient, "trans_gt_sat": e_trans_gt_sat,
}

def seed_errors(rec):
    r = copy.deepcopy(rec)
    kinds = random.sample(list(INJECTORS), k=random.randint(1, 3))
    log = [{"type": k, "description": INJECTORS[k](r)} for k in kinds]
    return r, log

# ---- Excel writer (mimics a "content sheet": sections, field, value, notes) ----
def write_xlsx(r, path, with_notes=False):
    wb = Workbook(); ws = wb.active; ws.title = "Content Sheet"
    ws.append(["Section", "Field", "Value", "Notes"])
    for c in ws[1]:
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDDDDD")
    nu = r["nutrition_per_serving"]
    rows = [
        ("Label Information", "Barcode", r["barcode"], ""),
        ("Label Information", "Market", "US", ""),
        ("Product Overview", "Product name", r["name"], ""),
        ("Product Overview", "Brand", r["brand"], ""),
        ("Product Overview", "Claims", "; ".join(r["claims"]), ""),
        ("Product Overview", "Labels", "; ".join(r["labels"]), ""),
        ("Legal Name and Ingredients", "Ingredients", r["ingredients"], ""),
        ("Legal Name and Ingredients", "Allergen statement", "; ".join(r["allergens"]), ""),
        ("Quantity", "Net weight", f'{r["net_weight_g"]:g} g', "Verify unit conversion" if with_notes else ""),
        ("Quantity", "Serving size", f'{r["serving_size_g"]:g} g', ""),
        ("Quantity", "Servings per container", r["servings_per_container"], ""),
    ]
    for k, v in nu.items():
        rows.append(("Nutrition Information", k, v, ""))
    for row in rows:
        ws.append(row)
    ws.column_dimensions["A"].width = 28; ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 60; ws.column_dimensions["D"].width = 30
    wb.save(path)

def download_images(p, folder):
    for key in ("image_front_url", "image_ingredients_url"):
        url = p.get(key)
        if url:
            data = requests.get(url, headers=HEADERS, timeout=30).content
            with open(os.path.join(folder, f'{p["code"]}_{key.split("_")[1]}.jpg'), "wb") as f:
                f.write(data)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", default="granola"); ap.add_argument("--count", type=int, default=12)
    ap.add_argument("--out", default="test_data"); ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--offline", action="store_true"); ap.add_argument("--images", action="store_true")
    a = ap.parse_args()
    random.seed(a.seed)
    for sub in ("clean", "seeded", "images"):
        os.makedirs(os.path.join(a.out, sub), exist_ok=True)

    products = [DEMO] if a.offline else fetch(a.query, a.count)
    truth = {}
    for p in products:
        rec = to_record(p)
        write_xlsx(rec, os.path.join(a.out, "clean", f'{rec["barcode"]}.xlsx'))
        bad, log = seed_errors(rec)
        write_xlsx(bad, os.path.join(a.out, "seeded", f'{rec["barcode"]}.xlsx'))
        truth[f'{rec["barcode"]}.xlsx'] = log
        if a.images and not a.offline:
            download_images(p, os.path.join(a.out, "images"))
    with open(os.path.join(a.out, "ground_truth.json"), "w") as f:
        json.dump(truth, f, indent=2)
    print(f"Wrote {len(products)} clean + {len(products)} seeded sheets to {a.out}/")

if __name__ == "__main__":
    main()
