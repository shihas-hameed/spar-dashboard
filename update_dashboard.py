"""
╔══════════════════════════════════════════════════════════════╗
║        SPAR Qatar — LPO Dashboard Auto-Updater               ║
║        Double-click this file to update the live dashboard   ║
╚══════════════════════════════════════════════════════════════╝

HOW TO USE:
  1. Update your Excel file as normal
  2. Double-click this file
  3. Wait ~15 seconds — dashboard updates automatically ✅

REQUIREMENTS (install once):
  pip install pandas openpyxl gitpython
"""

import os, sys, json, traceback
from datetime import datetime

# ─────────────────────────────────────────────────────────────
#  ⚙️  CONFIGURATION — only edit these 3 lines
# ─────────────────────────────────────────────────────────────
EXCEL_FILE  = r"C:\Users\shameed\Documents\LPO_Number.xlsx"
GITHUB_REPO = r"C:\Users\shameed\Documents\spar-dashboard"
SHEET_NAME  = "LPO TRACKING"
# ─────────────────────────────────────────────────────────────

# ── Category normalisation map ────────────────────────────────
# Maps every value from your Excel Category column to a clean name.
# Add rows here if you add new categories in Excel.
CATEGORY_MAP = {
    'EQUIPMENT MAINTENANCE'               : 'Equipment Maintenance',
    'EQUIPMENT MAINT'                     : 'Equipment Maintenance',
    'EQUIPMENT MAINT.'                    : 'Equipment Maintenance',
    'FF FA FS MAINTENANCE'                : 'Fire & Safety',
    'FF FA FS'                            : 'Fire & Safety',
    'FIRE & SAFETY'                       : 'Fire & Safety',
    'FIRE AND SAFETY'                     : 'Fire & Safety',
    'FIRE AMC'                            : 'Fire AMC',
    'REFRIGERATION MAINTENANCE'           : 'Refrigeration Maintenance',
    'REFRIGERATION MAINT'                 : 'Refrigeration Maintenance',
    'REFRIGERATION MAINT.'                : 'Refrigeration Maintenance',
    'REFRIGERATION AMC'                   : 'Refrigeration AMC',
    'BUILDING MAINTENANCE CIVIL WORKS'    : 'Civil Works',
    'CIVIL WORKS'                         : 'Civil Works',
    'BUILDING MAINTENANCE MECH WORKS'     : 'Building Maint. (Mech)',
    'BUILDING MAINTENANCE'                : 'Building Maintenance',
    'FURNITURE FIXTURES ASSET PROCUREMENT': 'Asset Procurement',
    'ASSET PROCUREMENT'                   : 'Asset Procurement',
    'FURNITURE FIXTURES'                  : 'Asset Procurement',
    'FURNITURE & FIXTURES'                : 'Asset Procurement',
    'CCTV MAINTENANCE'                    : 'CCTV Maintenance',
    'CCTV MAINT'                          : 'CCTV Maintenance',
    'CCTV AMC'                            : 'CCTV AMC',
    'KITCHEN HOOD AMC'                    : 'Kitchen Hood AMC',
    'KICTCHEN HOOD AMC'                   : 'Kitchen Hood AMC',
    'KITCHEN HOOD MAINTENANCE'            : 'Kitchen Hood AMC',
    'MEP WORKS'                           : 'MEP Works',
    'HVAC MAINTENANCE'                    : 'HVAC Maintenance',
    'HAVAC MAINTENANCE'                   : 'HVAC Maintenance',
    'HVAC MAINT'                          : 'HVAC Maintenance',
    'HVAC AMC'                            : 'HVAC AMC',
    'SIGNAGE STICKER BRANDING WORKS'      : 'Signage & Branding',
    'SIGNAGE & BRANDING'                  : 'Signage & Branding',
    'SIGNAGE AND BRANDING'                : 'Signage & Branding',
    'BRANDING WORKS'                      : 'Signage & Branding',
    'GENERATOR AMC'                       : 'Generator AMC',
    'IT WORKS'                            : 'IT Works',
    'IT ASSETS'                           : 'IT Assets',
    'IT WORKS IT ASSETS'                  : 'IT Works',
}

def normalise_category(raw):
    if not raw or str(raw).strip().lower() in ['nan', 'none', '']:
        return 'Uncategorised'
    cleaned = str(raw).strip().upper()
    if cleaned in CATEGORY_MAP:
        return CATEGORY_MAP[cleaned]
    for key, val in CATEGORY_MAP.items():
        if key in cleaned or cleaned in key:
            return val
    return str(raw).strip().title()

def normalise_store(s):
    s = str(s).strip()
    if '/' in s or s in ['ALL', 'PAQQALMALJ', 'TW/03/ALM', 'PA/TW', 'ALM/PA']:
        return 'MULTI-STORE'
    mapping = {'TAWR': 'TAWAR', 'TW': 'TAWAR', 'AL MANA': 'ALM', 'HO': 'OFFICE', 'BA': 'OTHER'}
    return mapping.get(s, s)

def read_excel():
    print(f"\n📂 Reading Excel: {EXCEL_FILE}")
    try:
        import pandas as pd
    except ImportError:
        print("❌ pandas not installed. Run: pip install pandas openpyxl")
        sys.exit(1)

    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Excel file not found: {EXCEL_FILE}")
        sys.exit(1)

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]
    print(f"   Excel columns found: {list(df.columns)}")

    col_map = {}
    for col in df.columns:
        cl = col.upper().replace(' ', '').replace(':', '').replace('_', '')
        if 'LPONO'       in cl: col_map['lpo']         = col
        if 'COMPANYNAME' in cl: col_map['vendor']       = col
        if 'STORENAME'   in cl: col_map['store']        = col
        if cl in ('AMT', 'AMOUNT') or cl.startswith('AMT'):
                                col_map['amt']          = col
        if 'DESCRIPTION' in cl: col_map['description']  = col
        if 'DATEOFISSUE' in cl or ('DATE' in cl and 'UPDATE' not in cl):
                                col_map['date']         = col
        if 'LPOSTA'      in cl: col_map['lpo_status']   = col
        if 'INVOICESTA'  in cl: col_map['inv_status']   = col
        if 'PAYMENTSCHE' in cl: col_map['payment']      = col
        if 'CATEGORY'    in cl: col_map['category']     = col

    print(f"   Columns mapped: {list(col_map.keys())}")

    if 'category' in col_map:
        print(f"   ✅ Category column found: '{col_map['category']}'")
        unique_cats = df[col_map['category']].dropna().unique()
        print(f"   Unique categories in Excel ({len(unique_cats)}):")
        for c in sorted(unique_cats):
            print(f"      '{c}'  →  '{normalise_category(c)}'")
    else:
        print("   ⚠️  No Category column found — check your Excel column is named 'CATEGORY'")

    records = []
    skipped = 0

    for _, row in df.iterrows():
        try:
            vendor = str(row.get(col_map.get('vendor', ''), '') or '').strip()
            if not vendor or vendor.lower() in ['nan', 'none', '']:
                skipped += 1
                continue

            lpo_raw = row.get(col_map.get('lpo', ''), '')
            try:
                lpo = int(float(str(lpo_raw))) if lpo_raw and str(lpo_raw).lower() != 'nan' else None
            except:
                lpo = str(lpo_raw).strip()

            amt_raw = row.get(col_map.get('amt', ''), 0)
            try:
                amt = float(str(amt_raw)) if amt_raw and str(amt_raw).lower() != 'nan' else 0
            except:
                amt = 0

            store_raw = str(row.get(col_map.get('store', ''), '') or '').strip()
            store = normalise_store(store_raw)

            date_raw = row.get(col_map.get('date', ''), None)
            year = None
            if date_raw and str(date_raw).lower() not in ['nan', 'none', '']:
                try:
                    dt = pd.to_datetime(date_raw, errors='coerce', dayfirst=True)
                    year = int(dt.year) if dt and not pd.isna(dt) else None
                except:
                    year = None

            desc = str(row.get(col_map.get('description', ''), '') or '').strip()
            if desc.lower() in ['nan', 'none']:
                desc = ''

            inv_raw = str(row.get(col_map.get('inv_status', ''), '') or '').strip().upper()
            inv = 'PENDING' if 'PEND' in inv_raw else None

            payment = str(row.get(col_map.get('payment', ''), '') or '').strip()
            if payment.lower() in ['nan', 'none']:
                payment = ''

            # ── Use YOUR category column from Excel ──
            cat_raw = row.get(col_map.get('category', ''), '') if 'category' in col_map else ''
            category = normalise_category(cat_raw)

            rec = {
                'lpo'        : lpo,
                'vendor'     : vendor,
                'store'      : store,
                'year'       : year,
                'category'   : category,
                'amt'        : round(amt, 2),
                'description': desc[:60],
                'payment'    : payment[:40],
            }
            if inv:
                rec['inv'] = inv
            records.append(rec)

        except Exception as e:
            skipped += 1

    print(f"\n   ✅ {len(records)} records processed, {skipped} skipped")
    return records

def write_data_json(records):
    out_path = os.path.join(GITHUB_REPO, 'data.json')
    unique_cats = sorted(set(r['category'] for r in records))
    payload = {
        'generated'    : datetime.now().isoformat(),
        'total_records': len(records),
        'total_amount' : round(sum(r['amt'] for r in records), 2),
        'categories'   : unique_cats,
        'records'      : records,
    }
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, separators=(',', ':'))
    size_kb = os.path.getsize(out_path) / 1024
    print(f"   ✅ data.json written ({size_kb:.1f} KB) → {out_path}")
    print(f"   📂 Categories: {unique_cats}")
    return out_path

def push_to_github():
    try:
        from git import Repo
    except ImportError:
        print("❌ gitpython not installed. Run: pip install gitpython")
        sys.exit(1)
    print(f"\n🚀 Pushing to GitHub...")
    try:
        repo = Repo(GITHUB_REPO)
        repo.index.add(['data.json'])
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        repo.index.commit(f'Update LPO data — {timestamp}')
        repo.remotes.origin.push()
        print(f"   ✅ Pushed successfully!")
        print(f"   ⏱️  Dashboard will update in ~30 seconds")
    except Exception as e:
        print(f"   ❌ Git push failed: {e}")
        raise

def main():
    print("=" * 60)
    print("  SPAR Qatar — LPO Dashboard Updater")
    print(f"  {datetime.now().strftime('%A, %d %B %Y — %H:%M')}")
    print("=" * 60)
    records = read_excel()
    if not records:
        print("\n⚠️  No records found. Check your Excel file and sheet name.")
        input("\nPress Enter to close...")
        return
    print(f"\n📝 Writing data.json...")
    write_data_json(records)
    push_to_github()
    total = sum(r['amt'] for r in records)
    cats  = sorted(set(r['category'] for r in records))
    print("\n" + "=" * 60)
    print("  ✅ DONE! Dashboard updated successfully.")
    print(f"  📊 {len(records)} LPOs | QAR {total:,.0f} total")
    print(f"  🏷️  {len(cats)} categories: {', '.join(cats)}")
    print("  🌐 Check your live site in ~30 seconds")
    print("=" * 60)
    input("\nPress Enter to close...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nCancelled.")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        traceback.print_exc()
        input("\nPress Enter to close...")
