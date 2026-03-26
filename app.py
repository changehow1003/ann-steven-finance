import json
from pathlib import Path
import pandas as pd
import streamlit as st
from datetime import datetime, date
import altair as alt
from supabase import create_client
import uuid


st.set_page_config(page_title="財務系統", layout="wide")
st.title("💼 Ann&Steven美甲店財務管理系統")

DATA_FILE = Path("finance_data.json")
SUPABASE_URL = "https://ztagpvfnqnekrsjmnddj.supabase.co"
SUPABASE_KEY = "sb_publishable_ME7e31h0RikcPZwonY2NTg_h-4DBFod"

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# =========================
# JSON
# =========================
def default_data():
    return {
        "expense_fixed": [],
        "expense_other": [],
        "withdrawals": [],
        "counters": {"F": 0, "O": 0, "W": 0}
    }


def load_data():
    if DATA_FILE.exists():
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("expense_fixed", [])
        data.setdefault("expense_other", [])
        data.setdefault("withdrawals", [])
        data.setdefault("counters", {"F": 0, "O": 0, "W": 0})
        for k in ["F", "O", "W"]:
            data["counters"].setdefault(k, 0)
        return data
    return default_data()


def save_data():
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "expense_fixed": st.session_state.expense_fixed,
            "expense_other": st.session_state.expense_other,
            "withdrawals": st.session_state.withdrawals,
            "counters": st.session_state.counters
        }, f, ensure_ascii=False, indent=2)


# =========================
# init
# =========================
if "init" not in st.session_state:
    data = load_data()
    st.session_state.expense_fixed = data.get("expense_fixed", [])
    st.session_state.expense_other = data.get("expense_other", [])
    st.session_state.withdrawals = data.get("withdrawals", [])
    st.session_state.counters = data.get("counters", {"F": 0, "O": 0, "W": 0})
    st.session_state.init = True

today = datetime.today()


# =========================
# 工具
# =========================
def load_income_cloud():
    resp = supabase.table("income").select("*").order("id").execute()
    rows = resp.data or []
    out = []
    for r in rows:
        out.append({
            "id": r.get("id"),
            "年份": int(r.get("year") or 0),
            "月份": int(r.get("month") or 0),
            "帳戶": r.get("account") or "",
            "金額": float(r.get("amount") or 0),
        })
    return out


def replace_all_income_cloud(rows):
    # 先清空再重寫，這版最簡單穩定
    supabase.table("income").delete().neq("id", 0).execute()

    payload = []
    for r in rows:
        payload.append({
            "year": int(r.get("年份", 0)),
            "month": int(r.get("月份", 0)),
            "account": str(r.get("帳戶", "")).strip(),
            "amount": float(r.get("金額", 0)),
        })

    if payload:
        supabase.table("income").insert(payload).execute()


def insert_income_cloud(year, month, account, amount):
    supabase.table("income").insert({
        "year": int(year),
        "month": int(month),
        "account": str(account).strip(),
        "amount": float(amount),
    }).execute()

def load_fixed_cloud():
    resp = supabase.table("expense_fixed").select("*").order("id").execute()
    rows = resp.data or []
    out = []
    for r in rows:
        out.append({
            "id": r.get("id"),
            "uid": r.get("uid") or "",
            "年份": int(r.get("year") or 0),
            "月份": int(r.get("month") or 0),
            "項目": r.get("item") or "",
            "金額": float(r.get("amount") or 0),
            "付款人": r.get("payer") or "",
            "發票": bool(r.get("invoice") or False),
            "提領過": bool(r.get("withdrawn") or False),
        })
    return out


def replace_all_fixed_cloud(rows):
    supabase.table("expense_fixed").delete().neq("id", 0).execute()

    payload = []
    for r in rows:
        payload.append({
            "uid": r.get("uid") or next_uid("F"),
            "year": int(r.get("年份", 0)),
            "month": int(r.get("月份", 0)),
            "item": str(r.get("項目", "")).strip(),
            "amount": float(r.get("金額", 0)),
            "payer": str(r.get("付款人", "")).strip(),
            "invoice": bool(r.get("發票", False)),
            "withdrawn": bool(r.get("提領過", False)),
        })

    if payload:
        supabase.table("expense_fixed").insert(payload).execute()


def insert_fixed_cloud(row):
    payload = {
        "uid": row.get("uid") or next_uid("F"),
        "year": int(row.get("年份", 0)),
        "month": int(row.get("月份", 0)),
        "item": str(row.get("項目", "")).strip(),
        "amount": float(row.get("金額", 0)),
        "payer": str(row.get("付款人", "")).strip(),
        "invoice": bool(row.get("發票", False)),
        "withdrawn": bool(row.get("提領過", False)),
    }
    supabase.table("expense_fixed").insert(payload).execute()


def load_other_cloud():
    resp = supabase.table("expense_other").select("*").order("id").execute()
    rows = resp.data or []
    out = []
    for r in rows:
        out.append({
            "id": r.get("id"),
            "uid": r.get("uid") or "",
            "年份": int(r.get("year") or 0),
            "月份": int(r.get("month") or 0),
            "日期": r.get("expense_date") or "",
            "項目": r.get("item") or "",
            "金額": float(r.get("amount") or 0),
            "付款人": r.get("payer") or "",
            "發票": bool(r.get("invoice") or False),
            "提領過": bool(r.get("withdrawn") or False),
        })
    return out


def replace_all_other_cloud(rows):
    supabase.table("expense_other").delete().neq("id", 0).execute()

    payload = []
    for r in rows:
        payload.append({
            "uid": r.get("uid") or next_uid("O"),
            "year": int(r.get("年份", 0)),
            "month": int(r.get("月份", 0)),
            "expense_date": str(r.get("日期", "")).strip(),
            "item": str(r.get("項目", "")).strip(),
            "amount": float(r.get("金額", 0)),
            "payer": str(r.get("付款人", "")).strip(),
            "invoice": bool(r.get("發票", False)),
            "withdrawn": bool(r.get("提領過", False)),
        })

    if payload:
        supabase.table("expense_other").insert(payload).execute()


def insert_other_cloud(row):
    payload = {
        "uid": row.get("uid") or next_uid("O"),
        "year": int(row.get("年份", 0)),
        "month": int(row.get("月份", 0)),
        "expense_date": str(row.get("日期", "")).strip(),
        "item": str(row.get("項目", "")).strip(),
        "amount": float(row.get("金額", 0)),
        "payer": str(row.get("付款人", "")).strip(),
        "invoice": bool(row.get("發票", False)),
        "withdrawn": bool(row.get("提領過", False)),
    }
    supabase.table("expense_other").insert(payload).execute()


def load_advance_cloud():
    resp = supabase.table("expense_advance").select("*").order("id").execute()
    rows = resp.data or []
    out = []
    for r in rows:
        out.append({
            "id": r.get("id"),
            "uid": r.get("uid") or "",
            "年份": int(r.get("year") or 0),
            "月份": int(r.get("month") or 0),
            "日期": r.get("expense_date") or "",
            "項目": r.get("item") or "",
            "金額": float(r.get("amount") or 0),
            "付款人": r.get("payer") or "",
            "發票": bool(r.get("invoice") or False),
        })
    return out


def replace_all_advance_cloud(rows):
    supabase.table("expense_advance").delete().neq("id", 0).execute()

    payload = []
    for r in rows:
        payload.append({
            "uid": r.get("uid") or next_uid("A"),
            "year": int(r.get("年份", 0)),
            "month": int(r.get("月份", 0)),
            "expense_date": str(r.get("日期", "")).strip(),
            "item": str(r.get("項目", "")).strip(),
            "amount": float(r.get("金額", 0)),
            "payer": str(r.get("付款人", "")).strip(),
            "invoice": bool(r.get("發票", False)),
        })

    if payload:
        supabase.table("expense_advance").insert(payload).execute()


def insert_advance_cloud(row):
    payload = {
        "uid": row.get("uid") or next_uid("A"),
        "year": int(row.get("年份", 0)),
        "month": int(row.get("月份", 0)),
        "expense_date": str(row.get("日期", "")).strip(),
        "item": str(row.get("項目", "")).strip(),
        "amount": float(row.get("金額", 0)),
        "payer": str(row.get("付款人", "")).strip(),
        "invoice": bool(row.get("發票", False)),
    }
    supabase.table("expense_advance").insert(payload).execute()


def load_withdrawals_cloud():
    resp = supabase.table("withdrawals").select("*").order("withdraw_date", desc=True).execute()
    rows = resp.data or []
    out = []
    for r in rows:
        out.append({
            "id": r.get("id"),
            "日期": r.get("withdraw_date") or "",
            "金額": float(r.get("amount") or 0),
            "收款帳戶": r.get("account") or "",
            "狀態": r.get("status") or "",
            "備註": r.get("note") or "",
        })
    return out


def load_withdrawal_sources_cloud(withdrawal_id):
    resp = supabase.table("withdrawal_sources").select("*").eq("withdrawal_id", withdrawal_id).execute()
    rows = resp.data or []
    out = []
    for r in rows:
        out.append({
            "類別": r.get("source_type") or "",
            "uid": r.get("source_uid") or "",
            "金額": float(r.get("amount") or 0),
            "項目": r.get("item") or "",
        })
    return out


def upload_code_to_cloud(file_name, full_text, chunk_size=4000):
    # 先刪掉同檔名舊備份
    supabase.table("code_backup").delete().eq("file_name", file_name).execute()

    rows = []
    total_len = len(full_text)

    for i in range(0, total_len, chunk_size):
        chunk = full_text[i:i + chunk_size]
        rows.append({
            "file_name": file_name,
            "part_no": i // chunk_size + 1,
            "content": chunk
        })

    if rows:
        supabase.table("code_backup").insert(rows).execute()


def load_code_from_cloud(file_name):
    resp = (
        supabase.table("code_backup")
        .select("*")
        .eq("file_name", file_name)
        .order("part_no")
        .execute()
    )
    rows = resp.data or []
    return "".join(r["content"] for r in rows)


def parse_tsv(text, cols):
    if not text.strip():
        return pd.DataFrame(columns=cols)
    rows = [r.split("\t") for r in text.splitlines() if r.strip()]
    return pd.DataFrame(rows, columns=cols)


def clean_month(df):
    df["月份"] = df["月份"].astype(str).str.replace("月", "", regex=False)
    df["月份"] = pd.to_numeric(df["月份"], errors="coerce")
    return df


def to_bool_invoice(s):
    if s is None:
        return False
    if isinstance(s, bool):
        return s
    v = str(s).strip().lower()
    return v in ["v", "true", "1", "是", "有", "y", "yes"]


def try_parse_date_ymd_to_date(s):
    if s is None:
        return None
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M"):
        try:
            dt = datetime.strptime(str(s).strip(), fmt)
            return dt.date()
        except:
            continue
    return None


def is_number(x):
    try:
        float(str(x).replace(",", ""))
        return True
    except:
        return False


def normalize_bool(v):
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    s = str(v).strip().lower()
    return s in ["true", "1", "yes", "y", "v", "是", "有"]


def normalize_amount(v):
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return 0.0


def parse_expense_bulk(text):
    records = []
    if not text.strip():
        return records

    for line in text.splitlines():
        if not line.strip():
            continue
        parts = [p.strip() for p in line.split("\t")]
        if len(parts) == 0:
            continue

        year = None
        month = None
        item = ""
        amount = None
        payer = ""
        invoice = False

        if len(parts) >= 6:
            m_raw, date_raw = parts[0], parts[1]
            d = try_parse_date_ymd_to_date(date_raw)

            try:
                month = int(str(m_raw).replace("月", ""))
            except:
                if d:
                    month = d.month

            if d:
                year = d.year

            if is_number(parts[3]):
                item = parts[2]
                amount = parts[3]
                payer = parts[4] if len(parts) >= 5 else ""
                invoice = to_bool_invoice(parts[5]) if len(parts) >= 6 else False
            else:
                item = parts[2]
                invoice = to_bool_invoice(parts[3]) if len(parts) >= 4 else False
                amount = parts[4] if len(parts) >= 5 else None
                payer = parts[5] if len(parts) >= 6 else ""
        elif len(parts) in (5, 6):
            try:
                year = int(parts[0])
            except:
                year = today.year
            try:
                month = int(str(parts[1]).replace("月", ""))
            except:
                month = None
            item = parts[2]
            amount = parts[3]
            payer = parts[4] if len(parts) >= 5 else ""
            invoice = to_bool_invoice(parts[5]) if len(parts) == 6 else False
        else:
            continue

        try:
            amount = float(str(amount).replace(",", ""))
        except:
            continue

        payer = ("" if payer is None else str(payer).strip())
        year = today.year if year is None else year
        month = today.month if month is None else month

        records.append({
            "年份": int(year),
            "月份": int(month),
            "項目": item,
            "金額": amount,
            "付款人": payer,
            "發票": bool(invoice),
            "提領過": False,
        })

    return records


def reorder_expense_columns(df):
    preferred = ["刪除", "年份", "月份", "項目", "金額", "付款人", "發票", "提領過"]
    for col in preferred:
        if col not in df.columns:
            if col in ["刪除", "提領過", "發票"]:
                df[col] = False
            elif col == "金額":
                df[col] = 0.0
            else:
                df[col] = ""
    return df.reindex(columns=preferred)


def reorder_other_columns(df):
    preferred = ["刪除", "日期", "項目", "金額", "付款人", "發票", "提領過"]
    for col in preferred:
        if col not in df.columns:
            if col in ["刪除", "提領過", "發票"]:
                df[col] = False
            elif col == "金額":
                df[col] = 0.0
            else:
                df[col] = ""
    return df.reindex(columns=preferred)


def reorder_advance_columns(df):
    preferred = ["刪除", "日期", "項目", "金額", "付款人", "發票"]
    for col in preferred:
        if col not in df.columns:
            if col in ["刪除", "發票"]:
                df[col] = False
            elif col == "金額":
                df[col] = 0.0
            else:
                df[col] = ""
    return df.reindex(columns=preferred)


def enrich_year_month(records):
    out = []
    for r in records:
        d = r.get("日期")
        if isinstance(d, str):
            d_parsed = try_parse_date_ymd_to_date(d)
        elif isinstance(d, (datetime, date)):
            d_parsed = d if isinstance(d, date) else d.date()
        else:
            d_parsed = None
        if d_parsed is None:
            d_parsed = date.today()
        r["年份"] = d_parsed.year
        r["月份"] = d_parsed.month
        r["日期"] = d_parsed.isoformat()
        r["付款人"] = (r.get("付款人") or "").strip()
        r["發票"] = normalize_bool(r.get("發票", False))
        r["提領過"] = normalize_bool(r.get("提領過", False))
        r["金額"] = normalize_amount(r.get("金額", 0))
        out.append(r)
    return out


def parse_other_bulk(text):
    recs = []
    if not text.strip():
        return recs
    for line in text.splitlines():
        if not line.strip():
            continue
        parts = [p.strip() for p in line.split("\t")]
        if len(parts) < 2:
            continue

        record = {"項目": "", "金額": 0.0, "付款人": "", "發票": False, "提領過": False}

        if len(parts) >= 5:
            d = try_parse_date_ymd_to_date(parts[0])
            if d:
                record["日期"] = d.isoformat()
                record["項目"] = parts[1]
                record["金額"] = parts[2]
                record["付款人"] = parts[3] if len(parts) >= 4 else ""
                record["發票"] = to_bool_invoice(parts[4]) if len(parts) >= 5 else False
            else:
                d2 = try_parse_date_ymd_to_date(parts[1])
                if d2:
                    if len(parts) >= 4 and is_number(parts[3]):
                        record["日期"] = d2.isoformat()
                        record["項目"] = parts[2]
                        record["金額"] = parts[3]
                        record["付款人"] = parts[4] if len(parts) >= 5 else ""
                        record["發票"] = to_bool_invoice(parts[5]) if len(parts) >= 6 else False
                    else:
                        record["日期"] = d2.isoformat()
                        record["項目"] = parts[2]
                        record["發票"] = to_bool_invoice(parts[3]) if len(parts) >= 4 else False
                        record["金額"] = parts[4] if len(parts) >= 5 else 0
                        record["付款人"] = parts[5] if len(parts) >= 6 else ""
                else:
                    continue
        else:
            continue

        recs.append(record)

    return enrich_year_month(recs)


def next_uid(prefix):
    return f"{prefix}_{uuid.uuid4().hex[:8]}"
    

def ensure_expense_uids():
    pass


def get_expense_by_uid(uid):
    for r in load_fixed_cloud():
        if r.get("uid") == uid:
            return r, "固定"
    for r in load_other_cloud():
        if r.get("uid") == uid:
            return r, "其他"
    return None, None


def build_eligible_expenses_df(year=None, month=None):
    rows = []
    for r in load_fixed_cloud():
        has_invoice = normalize_bool(r.get("發票", False))
        withdrawn = normalize_bool(r.get("提領過", False))
        if has_invoice and not withdrawn:
            if year and int(r.get("年份", 0)) != int(year):
                continue
            if month and int(r.get("月份", 0)) != int(month):
                continue
            rows.append({
                "選取": False,
                "序號": 0,
                "類別": "固定",
                "uid": r.get("uid", ""),
                "日期/年月": f'{int(r.get("年份", 0))}/{int(r.get("月份", 0))}',
                "項目": r.get("項目", ""),
                "金額": normalize_amount(r.get("金額", 0)),
                "付款人": r.get("付款人", "")
            })
    for r in load_other_cloud():
        has_invoice = normalize_bool(r.get("發票", False))
        withdrawn = normalize_bool(r.get("提領過", False))
        if has_invoice and not withdrawn:
            if year and int(r.get("年份", 0)) != int(year):
                continue
            if month and int(r.get("月份", 0)) != int(month):
                continue
            rows.append({
                "選取": False,
                "序號": 0,
                "類別": "其他",
                "uid": r.get("uid", ""),
                "日期/年月": r.get("日期", ""),
                "項目": r.get("項目", ""),
                "金額": normalize_amount(r.get("金額", 0)),
                "付款人": r.get("付款人", "")
            })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["序號"] = range(1, len(df) + 1)
    return df


def mark_expenses_withdrawn(selected_uids, withdrawn=True):
    uid_set = set(selected_uids)

    fixed_rows = load_fixed_cloud()
    fixed_changed = False
    for r in fixed_rows:
        if r.get("uid") in uid_set:
            r["提領過"] = withdrawn
            fixed_changed = True
    if fixed_changed:
        replace_all_fixed_cloud(fixed_rows)

    other_rows = load_other_cloud()
    other_changed = False
    for r in other_rows:
        if r.get("uid") in uid_set:
            r["提領過"] = withdrawn
            other_changed = True
    if other_changed:
        replace_all_other_cloud(other_rows)

    save_data()


def create_withdrawal(w_date, amount, account, selected_uids):
    sources = []
    sum_amount = 0.0

    for uid in selected_uids:
        rec, cat = get_expense_by_uid(uid)

        if not rec:
            st.error(f"找不到來源 uid: {uid}")
            return False

        if not normalize_bool(rec.get("發票", False)):
            st.error(f"來源未有發票：{rec.get('項目')}")
            return False

        if normalize_bool(rec.get("提領過", False)):
            st.error(f"來源已提領：{rec.get('項目')}")
            return False

        amt = normalize_amount(rec.get("金額", 0))
        sum_amount += amt

        sources.append({
            "類別": cat,
            "uid": uid,
            "金額": amt,
            "項目": rec.get("項目", "")
        })

    if abs(sum_amount - amount) > 0.01:
        st.error("金額不一致")
        return False

    wid = next_uid("W")

    # 🔥 寫入 withdrawals
    supabase.table("withdrawals").insert({
        "id": wid,
        "withdraw_date": w_date.isoformat(),
        "amount": float(amount),
        "account": account,
        "status": "已配對",
        "note": ""
    }).execute()

    # 🔥 寫入 sources
    for s in sources:
        supabase.table("withdrawal_sources").insert({
            "withdrawal_id": wid,
            "source_uid": s["uid"],
            "source_type": s["類別"],
            "amount": s["金額"],
            "item": s["項目"]
        }).execute()

    # 🔥 更新支出 → 提領過 = True
    for uid in selected_uids:
        rec, cat = get_expense_by_uid(uid)

        if cat == "固定":
            supabase.table("expense_fixed")\
                .update({"withdrawn": True})\
                .eq("uid", uid)\
                .execute()
        else:
            supabase.table("expense_other")\
                .update({"withdrawn": True})\
                .eq("uid", uid)\
                .execute()

    return True


def delete_withdrawal(wid):
    # 先抓來源
    srcs = load_withdrawal_sources_cloud(wid)
    source_uids = [s.get("uid") for s in srcs if s.get("uid")]

    # 先把來源支出恢復 withdrawn=False
    if source_uids:
        mark_expenses_withdrawn(source_uids, False)

    # 刪 relation
    supabase.table("withdrawal_sources").delete().eq("withdrawal_id", wid).execute()

    # 刪 main record
    supabase.table("withdrawals").delete().eq("id", wid).execute()


def create_pending_withdrawal(w_date, amount, account, note=""):
    wid = next_uid("W")
    st.session_state.withdrawals.append({
        "id": wid,
        "日期": w_date.isoformat(),
        "金額": float(amount),
        "收款帳戶": account,
        "來源": [],
        "狀態": "待配對",
        "備註": note
    })
    save_data()
    return wid


def parse_withdrawals_bulk(text):
    recs = []
    if not text.strip():
        return recs
    for line in text.splitlines():
        if not line.strip():
            continue
        parts = [p.strip() for p in line.split("\t")]
        if len(parts) < 3:
            continue
        d = try_parse_date_ymd_to_date(parts[0])
        if not d:
            continue
        try:
            amt = float(str(parts[1]).replace(",", ""))
        except:
            continue
        acc = parts[2]
        note = parts[3] if len(parts) >= 4 else ""
        recs.append({"日期": d, "金額": amt, "收款帳戶": acc, "備註": note})
    return recs


def auto_match_withdrawal(w):
    if w.get("狀態") == "已配對":
        return True

    wd = try_parse_date_ymd_to_date(w.get("日期")) if isinstance(w.get("日期"), str) else w.get("日期")
    if not wd:
        wd = date.today()
    wy, wm = wd.year, wd.month
    target = float(w.get("金額", 0))

    cands = []
    for r in st.session_state.expense_fixed:
        if normalize_bool(r.get("發票")) and not normalize_bool(r.get("提領過")):
            if int(r.get("年份", 0)) == wy and int(r.get("月份", 0)) == wm:
                cands.append(("固定", r))
    for r in st.session_state.expense_other:
        if normalize_bool(r.get("發票")) and not normalize_bool(r.get("提領過")):
            ry, rm = int(r.get("年份", 0)), int(r.get("月份", 0))
            if ry == wy and rm == wm:
                cands.append(("其他", r))

    for cat, r in cands:
        if abs(normalize_amount(r.get("金額", 0)) - target) < 0.01:
            w["來源"] = [{
                "類別": cat,
                "uid": r["uid"],
                "金額": normalize_amount(r["金額"]),
                "項目": r.get("項目", "")
            }]
            w["狀態"] = "已配對"
            mark_expenses_withdrawn([r["uid"]], True)
            save_data()
            return True

    cands_sorted = sorted(cands, key=lambda x: normalize_amount(x[1].get("金額", 0)), reverse=True)[:12]
    amounts = [normalize_amount(r.get("金額", 0)) for _, r in cands_sorted]
    result_idxs = []

    def dfs(idx, curr_sum, chosen):
        if abs(curr_sum - target) < 0.01:
            result_idxs.clear()
            result_idxs.extend(chosen)
            return True
        if curr_sum > target + 0.01:
            return False
        if idx >= len(cands_sorted) or len(chosen) >= 6:
            return False
        if dfs(idx + 1, curr_sum + amounts[idx], chosen + [idx]):
            return True
        if dfs(idx + 1, curr_sum, chosen):
            return True
        return False

    if dfs(0, 0.0, []):
        sources = []
        uids = []
        for i in result_idxs:
            cat, r = cands_sorted[i]
            sources.append({
                "類別": cat,
                "uid": r["uid"],
                "金額": normalize_amount(r["金額"]),
                "項目": r.get("項目", "")
            })
            uids.append(r["uid"])
        w["來源"] = sources
        w["狀態"] = "已配對"
        mark_expenses_withdrawn(uids, True)
        save_data()
        return True

    return False


ensure_expense_uids()


# =========================
# tabs
# =========================
##tab1, tab2, tab3 = st.tabs(["🧾 收支紀錄", "🏧 提領紀錄", "📊 分析"])
tab1, tab2, tab3, tab4 = st.tabs(["🧾 收支紀錄", "🏧 提領紀錄", "📊 分析", "🛠 程式備份"])


# =========================
# TAB1 - 收支紀錄
# =========================
with tab1:
    st.subheader("📊 固定收入")

    income_rows = load_income_cloud()
    df_income = pd.DataFrame(income_rows)

    if df_income.empty:
        df_income = pd.DataFrame(columns=["id", "年份", "月份", "帳戶", "金額"])

    if "刪除" not in df_income.columns:
        df_income.insert(0, "刪除", False)
    if "序號" not in df_income.columns:
        df_income.insert(1, "序號", range(1, len(df_income) + 1))
    else:
        df_income["序號"] = range(1, len(df_income) + 1)

    edited_income = st.data_editor(
        df_income,
        use_container_width=False,
        num_rows="dynamic",
        hide_index=True,
        key="income_editor",
        column_config={
            "刪除": st.column_config.CheckboxColumn("刪除", width="small"),
            "序號": st.column_config.NumberColumn("序號", disabled=True, width="small"),
            "id": None,
            "年份": st.column_config.NumberColumn("年份", width="small"),
            "月份": st.column_config.NumberColumn("月份", width="small"),
            "帳戶": st.column_config.TextColumn("帳戶", width="medium"),
            "金額": st.column_config.NumberColumn("金額", width="medium"),
        },
        disabled=["序號", "id"]
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 儲存收入"):
            rows = edited_income.copy()
            rows = rows[rows["刪除"] != True]
            rows = rows.drop(columns=["刪除", "序號", "id"], errors="ignore").to_dict("records")

            normalized = []
            for r in rows:
                y = pd.to_numeric(r.get("年份"), errors="coerce")
                m = pd.to_numeric(r.get("月份"), errors="coerce")
                amt = normalize_amount(r.get("金額", 0))
                acc = str(r.get("帳戶", "")).strip()

                if pd.isna(y) or pd.isna(m) or not acc:
                    continue

                normalized.append({
                    "年份": int(y),
                    "月份": int(m),
                    "帳戶": acc,
                    "金額": amt
                })

            replace_all_income_cloud(normalized)
            st.rerun()

    with c2:
        if st.button("🗑️ 刪除已勾選收入"):
            delete_df = edited_income[edited_income["刪除"] == True]
            if delete_df.empty:
                st.warning("請先勾選要刪除的收入")
            else:
                rows = edited_income[edited_income["刪除"] != True]
                rows = rows.drop(columns=["刪除", "序號", "id"], errors="ignore").to_dict("records")

                normalized = []
                for r in rows:
                    y = pd.to_numeric(r.get("年份"), errors="coerce")
                    m = pd.to_numeric(r.get("月份"), errors="coerce")
                    amt = normalize_amount(r.get("金額", 0))
                    acc = str(r.get("帳戶", "")).strip()

                    if pd.isna(y) or pd.isna(m) or not acc:
                        continue

                    normalized.append({
                        "年份": int(y),
                        "月份": int(m),
                        "帳戶": acc,
                        "金額": amt
                    })

                replace_all_income_cloud(normalized)
                st.rerun()

    st.subheader("➕ 單筆新增收入")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        year_income = st.selectbox(
            "年",
            list(range(2021, 2051)),
            index=list(range(2021, 2051)).index(today.year),
            key="income_add_year"
        )
    with c2:
        month_income = st.selectbox(
            "月",
            list(range(1, 13)),
            index=today.month - 1,
            key="income_add_month"
        )
    with c3:
        account = st.selectbox(
            "帳戶",
            ["收錢吧", "美團", "抖音", "其他", "收錢吧-私"],
            key="income_add_account"
        )
    with c4:
        amount_income = st.number_input("金額", min_value=0.0, key="income_add_amount")
    with c5:
        st.write("")
        st.write("")
        if st.button("新增收入"):
            insert_income_cloud(
                year=year_income,
                month=month_income,
                account=account,
                amount=amount_income
            )
            st.rerun()

    st.markdown("### 📥 Excel貼上（收入）")
    text_income = st.text_area("格式：年份\t月份\t帳戶\t金額", key="ta_income")

    if st.button("匯入收入"):
        df = parse_tsv(text_income, ["年份", "月份", "帳戶", "金額"])
        df = clean_month(df)
        df["金額"] = pd.to_numeric(df["金額"], errors="coerce")
        df = df.dropna()

        records = []
        for _, r in df.iterrows():
            records.append({
                "年份": int(r["年份"]),
                "月份": int(r["月份"]),
                "帳戶": str(r["帳戶"]).strip(),
                "金額": float(r["金額"]),
            })

        current = load_income_cloud()
        merged = current + records
        replace_all_income_cloud(merged)
        st.rerun()


    st.divider()

    st.subheader("📊 固定支出")
    st.caption("已提領過的支出，其『發票』欄位會鎖定；即使畫面上被改動，儲存時也會自動維持原值。")

    df_fixed = pd.DataFrame(load_fixed_cloud())
    if df_fixed.empty:
        df_fixed = pd.DataFrame(columns=["id", "uid", "年份", "月份", "項目", "金額", "付款人", "發票", "提領過"])

    uids_fixed = df_fixed.get("uid", pd.Series(dtype=str))
    ids_fixed = df_fixed.get("id", pd.Series(dtype=float))

    df_fixed = reorder_expense_columns(df_fixed.drop(columns=["id", "uid"], errors="ignore"))
    df_fixed["uid"] = uids_fixed if not uids_fixed.empty else ["" for _ in range(len(df_fixed))]
    df_fixed["id"] = ids_fixed if not ids_fixed.empty else [None for _ in range(len(df_fixed))]
    df_fixed.insert(1, "序號", range(1, len(df_fixed) + 1))
    df_fixed["發票"] = df_fixed["發票"].apply(normalize_bool)
    df_fixed["提領過"] = df_fixed["提領過"].apply(normalize_bool)

    edited_fixed = st.data_editor(
        df_fixed,
        use_container_width=False,
        num_rows="dynamic",
        hide_index=True,
        key="fixed_editor",
        column_config={
            "刪除": st.column_config.CheckboxColumn("刪除", width="small"),
            "序號": st.column_config.NumberColumn("序號", disabled=True, width="small"),
            "年份": st.column_config.NumberColumn("年份", width="small"),
            "月份": st.column_config.NumberColumn("月份", width="small"),
            "項目": st.column_config.TextColumn("項目", width="medium"),
            "金額": st.column_config.NumberColumn("金額", width="medium"),
            "付款人": st.column_config.TextColumn("付款人", width="small"),
            "發票": st.column_config.CheckboxColumn("發票", width="small"),
            "提領過": st.column_config.CheckboxColumn("提領過", width="small", disabled=True),
            "uid": None,
            "id": None
        },
        disabled=["序號", "提領過", "uid", "id"]
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 儲存固定支出"):
            rows = edited_fixed[edited_fixed["刪除"] != True].drop(columns=["刪除", "序號", "id"], errors="ignore").to_dict("records")
            old_map = {r.get("uid"): r for r in load_fixed_cloud()}

            deleting_uids = set(edited_fixed[edited_fixed["刪除"] == True]["uid"].tolist())
            locked_delete = [
                r for r in load_fixed_cloud()
                if r.get("uid") in deleting_uids and normalize_bool(r.get("提領過", False))
            ]
            if locked_delete:
                st.error("有已提領的固定支出被勾選刪除，請先刪除對應提領紀錄後再刪除支出。")
            else:
                normalized_rows = []
                for r in rows:
                    old = old_map.get(r.get("uid"), {})
                    old_withdrawn = normalize_bool(old.get("提領過", False))

                    y = pd.to_numeric(r.get("年份"), errors="coerce")
                    m = pd.to_numeric(r.get("月份"), errors="coerce")
                    if pd.isna(y) or pd.isna(m):
                        continue

                    row = {
                        "年份": int(y),
                        "月份": int(m),
                        "項目": str(r.get("項目", "")).strip(),
                        "金額": normalize_amount(r.get("金額", 0)),
                        "付款人": str(r.get("付款人", "")).strip(),
                        "uid": r.get("uid") or next_uid("F")
                    }
                    row["提領過"] = old_withdrawn
                    row["發票"] = normalize_bool(old.get("發票", False)) if old_withdrawn else normalize_bool(r.get("發票", False))
                    normalized_rows.append(row)

                replace_all_fixed_cloud(normalized_rows)
                st.rerun()

    with c2:
        if st.button("🗑️ 刪除已勾選固定支出"):
            delete_rows = edited_fixed[edited_fixed["刪除"] == True]
            if delete_rows.empty:
                st.warning("請先勾選要刪除的固定支出")
            else:
                delete_uids = set(delete_rows["uid"].tolist())
                deleting_records = [r for r in load_fixed_cloud() if r.get("uid") in delete_uids]
                locked = [r for r in deleting_records if normalize_bool(r.get("提領過", False))]
                if locked:
                    st.error("有已提領的固定支出被勾選，請先刪除對應提領紀錄後再刪除支出。")
                else:
                    remain = [r for r in load_fixed_cloud() if r.get("uid") not in delete_uids]
                    replace_all_fixed_cloud(remain)
                    st.rerun()

    st.subheader("➕ 單筆新增固定支出")
    f1, f2, f3, f4, f5, f6 = st.columns(6)
    with f1:
        year_fixed = st.selectbox("年", list(range(2021, 2051)), index=list(range(2021, 2051)).index(today.year), key="fx1")
    with f2:
        month_fixed = st.selectbox("月", list(range(1, 13)), index=today.month - 1, key="fx2")
    with f3:
        item_fixed = st.text_input("項目", key="fx3")
    with f4:
        amount_fixed = st.number_input("金額", key="fx4")
    with f5:
        payer_fixed = st.text_input("付款人", key="fx5")
    with f6:
        invoice_fixed = st.checkbox("發票", key="fx7")

    if st.button("新增固定支出"):
        insert_fixed_cloud({
            "年份": year_fixed,
            "月份": month_fixed,
            "項目": item_fixed,
            "金額": amount_fixed,
            "付款人": (payer_fixed or "").strip(),
            "發票": invoice_fixed,
            "提領過": False,
            "uid": next_uid("F")
        })
        st.rerun()

    st.markdown("### 📥 Excel貼上（固定支出）")
    st.markdown(
        "支援格式（以 Tab 分隔）：\n"
        "- 月份\t日期\t項目\t金額\t付款人\t發票\n"
        "- 年份\t月份\t項目\t金額\t付款人\t[發票]\n"
        "- 相容舊順序：月份\t日期\t項目\t發票\t金額\t付款人"
    )
    text_fixed_bulk = st.text_area("在此貼上（固定支出）", key="ta_fixed")
    if st.button("匯入固定支出"):
        recs = parse_expense_bulk(text_fixed_bulk)
        for r in recs:
            r["uid"] = next_uid("F")
            r["提領過"] = False

        current = load_fixed_cloud()
        merged = current + recs
        replace_all_fixed_cloud(merged)
        st.rerun()


    st.divider()
    st.subheader("📊 其他支出")
    st.caption("已提領過的支出，其『發票』欄位會鎖定；即使畫面上被改動，儲存時也會自動維持原值。")

    df_other = pd.DataFrame(load_other_cloud())
    if df_other.empty:
        df_other = pd.DataFrame(columns=["id", "uid", "日期", "年份", "月份", "項目", "金額", "付款人", "發票", "提領過"])

    uids_other = df_other.get("uid", pd.Series(dtype=str))
    ids_other = df_other.get("id", pd.Series(dtype=float))
    dates_other = df_other.get("日期", pd.Series(dtype=str))

    df_other = reorder_other_columns(df_other.drop(columns=["id", "uid"], errors="ignore"))
    df_other["uid"] = uids_other if not uids_other.empty else ["" for _ in range(len(df_other))]
    df_other["id"] = ids_other if not ids_other.empty else [None for _ in range(len(df_other))]
    df_other["日期"] = dates_other if not dates_other.empty else ["" for _ in range(len(df_other))]
    df_other["日期"] = pd.to_datetime(df_other["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    df_other = df_other.sort_values("日期", ascending=False).reset_index(drop=True)
    df_other.insert(1, "序號", range(1, len(df_other) + 1))
    df_other["發票"] = df_other["發票"].apply(normalize_bool)
    df_other["提領過"] = df_other["提領過"].apply(normalize_bool)

    edited_other = st.data_editor(
        df_other,
        use_container_width=False,
        num_rows="dynamic",
        hide_index=True,
        key="other_editor",
        column_config={
            "刪除": st.column_config.CheckboxColumn("刪除", width="small"),
            "序號": st.column_config.NumberColumn("序號", disabled=True, width="small"),
            "日期": st.column_config.TextColumn("日期", width="medium"),
            "項目": st.column_config.TextColumn("項目", width="medium"),
            "金額": st.column_config.NumberColumn("金額", width="medium"),
            "付款人": st.column_config.TextColumn("付款人", width="small"),
            "發票": st.column_config.CheckboxColumn("發票", width="small"),
            "提領過": st.column_config.CheckboxColumn("提領過", width="small", disabled=True),
            "uid": None,
            "id": None
        },
        disabled=["序號", "提領過", "uid", "id"]
    )

    o1, o2 = st.columns(2)
    with o1:
        if st.button("💾 儲存其他支出"):
            rows = edited_other[edited_other["刪除"] != True].drop(columns=["刪除", "序號", "id"], errors="ignore").to_dict("records")
            rows = enrich_year_month(rows)

            delete_uids = set(edited_other[edited_other["刪除"] == True]["uid"].tolist())
            deleting_records = [r for r in load_other_cloud() if r.get("uid") in delete_uids]
            locked = [r for r in deleting_records if normalize_bool(r.get("提領過", False))]
            if locked:
                st.error("有已提領的其他支出被勾選刪除，請先刪除對應提領紀錄後再刪除支出。")
            else:
                old_map = {r.get("uid"): r for r in load_other_cloud()}
                normalized_rows = []
                for r in rows:
                    old = old_map.get(r.get("uid"), {})
                    old_withdrawn = normalize_bool(old.get("提領過", False))
                    r["uid"] = r.get("uid") or next_uid("O")
                    r["項目"] = str(r.get("項目", "")).strip()
                    r["付款人"] = str(r.get("付款人", "")).strip()
                    r["金額"] = normalize_amount(r.get("金額", 0))
                    r["提領過"] = old_withdrawn
                    r["發票"] = normalize_bool(old.get("發票", False)) if old_withdrawn else normalize_bool(r.get("發票", False))
                    normalized_rows.append(r)

                replace_all_other_cloud(normalized_rows)
                st.rerun()

    with o2:
        if st.button("🗑️ 刪除已勾選其他支出"):
            delete_rows = edited_other[edited_other["刪除"] == True]
            if delete_rows.empty:
                st.warning("請先勾選要刪除的其他支出")
            else:
                delete_uids = set(delete_rows["uid"].tolist())
                deleting_records = [r for r in load_other_cloud() if r.get("uid") in delete_uids]
                locked = [r for r in deleting_records if normalize_bool(r.get("提領過", False))]
                if locked:
                    st.error("有已提領的其他支出被勾選，請先刪除對應提領紀錄後再刪除支出。")
                else:
                    remain = [r for r in load_other_cloud() if r.get("uid") not in delete_uids]
                    replace_all_other_cloud(remain)
                    st.rerun()

    st.subheader("➕ 單筆新增其他支出")
    oc1, oc2, oc3, oc4, oc5 = st.columns(5)
    with oc1:
        other_date = st.date_input("日期", value=date.today(), key="ot_date")
    with oc2:
        item_other = st.text_input("項目", key="ot_item")
    with oc3:
        amount_other = st.number_input("金額", min_value=0.0, key="ot_amt")
    with oc4:
        payer_other = st.text_input("付款人", key="ot_payer")
    with oc5:
        invoice_other = st.checkbox("發票", key="ot_invoice")

    if st.button("新增其他支出"):
        rec = {
            "日期": other_date.isoformat(),
            "項目": item_other,
            "金額": amount_other,
            "付款人": (payer_other or "").strip(),
            "發票": invoice_other,
            "提領過": False,
            "uid": next_uid("O")
        }
        rec = enrich_year_month([rec])[0]
        insert_other_cloud(rec)
        st.rerun()

    st.markdown("### 📥 Excel貼上（其他支出）")
    st.markdown(
        "推薦格式（Tab 分隔）：\n"
        "- 日期\t項目\t金額\t付款人\t發票\n"
        "也支援：\n"
        "- 月份\t日期\t項目\t金額\t付款人\t發票\n"
        "- 月份\t日期\t項目\t發票\t金額\t付款人"
    )
    text_other_bulk = st.text_area("在此貼上（其他支出）", key="ta_other")
    if st.button("匯入其他支出"):
        recs = parse_other_bulk(text_other_bulk)
        for r in recs:
            r["uid"] = next_uid("O")
            r["提領過"] = False

        current = load_other_cloud()
        merged = current + recs
        replace_all_other_cloud(merged)
        st.rerun()

    st.divider()
    st.subheader("📊 代墊支出")

    df_adv = pd.DataFrame(load_advance_cloud())
    if df_adv.empty:
        df_adv = pd.DataFrame(columns=["id", "uid", "日期", "年份", "月份", "項目", "金額", "付款人", "發票"])

    uids_adv = df_adv.get("uid", pd.Series(dtype=str))
    ids_adv = df_adv.get("id", pd.Series(dtype=float))
    dates_adv = df_adv.get("日期", pd.Series(dtype=str))

    df_adv = reorder_advance_columns(df_adv.drop(columns=["id", "uid"], errors="ignore"))
    df_adv["uid"] = uids_adv if not uids_adv.empty else ["" for _ in range(len(df_adv))]
    df_adv["id"] = ids_adv if not ids_adv.empty else [None for _ in range(len(df_adv))]
    df_adv["日期"] = dates_adv if not dates_adv.empty else ["" for _ in range(len(df_adv))]
    df_adv["日期"] = pd.to_datetime(df_adv["日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    df_adv = df_adv.sort_values("日期", ascending=False).reset_index(drop=True)
    df_adv.insert(1, "序號", range(1, len(df_adv) + 1))
    df_adv["發票"] = df_adv["發票"].apply(normalize_bool)

    edited_adv = st.data_editor(
        df_adv,
        use_container_width=False,
        num_rows="dynamic",
        hide_index=True,
        key="advance_editor",
        column_config={
            "刪除": st.column_config.CheckboxColumn("刪除", width="small"),
            "序號": st.column_config.NumberColumn("序號", disabled=True, width="small"),
            "日期": st.column_config.TextColumn("日期", width="medium"),
            "項目": st.column_config.TextColumn("項目", width="medium"),
            "金額": st.column_config.NumberColumn("金額", width="medium"),
            "付款人": st.column_config.TextColumn("付款人", width="small"),
            "發票": st.column_config.CheckboxColumn("發票", width="small"),
            "uid": None,
            "id": None
        },
        disabled=["序號", "uid", "id"]
    )

    a1, a2 = st.columns(2)
    with a1:
        if st.button("💾 儲存代墊支出", key="save_advance"):
            rows = edited_adv[edited_adv["刪除"] != True].drop(columns=["刪除", "序號", "id"], errors="ignore").to_dict("records")
            rows = enrich_year_month(rows)

            normalized_rows = []
            for r in rows:
                r["uid"] = r.get("uid") or next_uid("A")
                r["項目"] = str(r.get("項目", "")).strip()
                r["付款人"] = str(r.get("付款人", "")).strip()
                r["金額"] = normalize_amount(r.get("金額", 0))
                r["發票"] = normalize_bool(r.get("發票", False))
                normalized_rows.append(r)

            replace_all_advance_cloud(normalized_rows)
            st.rerun()

    with a2:
        if st.button("🗑️ 刪除已勾選代墊支出", key="delete_advance"):
            delete_rows = edited_adv[edited_adv["刪除"] == True]
            if delete_rows.empty:
                st.warning("請先勾選要刪除的代墊支出")
            else:
                delete_uids = set(delete_rows["uid"].tolist())
                remain = [r for r in load_advance_cloud() if r.get("uid") not in delete_uids]
                replace_all_advance_cloud(remain)
                st.rerun()

    st.subheader("➕ 單筆新增代墊支出")
    ac1, ac2, ac3, ac4, ac5 = st.columns(5)
    with ac1:
        advance_date = st.date_input("日期", value=date.today(), key="adv_date")
    with ac2:
        item_advance = st.text_input("項目", key="adv_item")
    with ac3:
        amount_advance = st.number_input("金額", min_value=0.0, key="adv_amt")
    with ac4:
        payer_advance = st.text_input("付款人", key="adv_payer")
    with ac5:
        invoice_advance = st.checkbox("發票", key="adv_invoice")

    if st.button("新增代墊支出", key="add_advance"):
        rec = {
            "日期": advance_date.isoformat(),
            "項目": item_advance,
            "金額": amount_advance,
            "付款人": (payer_advance or "").strip(),
            "發票": invoice_advance,
            "uid": next_uid("A")
        }
        rec = enrich_year_month([rec])[0]
        insert_advance_cloud(rec)
        st.rerun()

# =========================
# TAB2 - 提領紀錄
# =========================
with tab2:
    if st.session_state.get("reset_withdrawal_form", False):
        for k in ("w_amount", "w_account"):
            if k in st.session_state:
                del st.session_state[k]
        del st.session_state["reset_withdrawal_form"]

    if "flash_msg" in st.session_state:
        st.success(st.session_state.pop("flash_msg"))

    st.subheader("提領紀錄（全部）")

    withdrawals_cloud = load_withdrawals_cloud()

    if not withdrawals_cloud:
        st.info("尚無提領紀錄")
    else:
        rows = []

        for w in withdrawals_cloud:
            srcs = load_withdrawal_sources_cloud(w["id"])

            if not srcs:
                rows.append({
                    "刪除": False,
                    "序號": 0,
                    "提領ID": w.get("id"),
                    "日期": w.get("日期"),
                    "收款帳戶": w.get("收款帳戶", ""),
                    "提領金額": float(w.get("金額", 0)),
                    "來源類別": "",
                    "來源ID": "",
                    "來源項目": "",
                    "來源金額": 0.0,
                    "來源付款人": ""
                })
            else:
                for s in srcs:
                    rec, cat = get_expense_by_uid(s["uid"])
                    rows.append({
                        "刪除": False,
                        "序號": 0,
                        "提領ID": w.get("id"),
                        "日期": w.get("日期"),
                        "收款帳戶": w.get("收款帳戶", ""),
                        "提領金額": float(w.get("金額", 0)),
                        "來源類別": cat or s.get("類別", ""),
                        "來源ID": s.get("uid", ""),
                        "來源項目": (rec or {}).get("項目", s.get("項目", "")),
                        "來源金額": float(s.get("金額", 0)),
                        "來源付款人": (rec or {}).get("付款人", "")
                    })


        df_w_all = pd.DataFrame(rows)
        if not df_w_all.empty:
            df_w_all["序號"] = range(1, len(df_w_all) + 1)

        edited_w_all = st.data_editor(
            df_w_all,
            use_container_width=False,
            hide_index=True,
            key="withdrawal_all_editor",
            column_config={
                "刪除": st.column_config.CheckboxColumn("刪除", width="small"),
                "序號": st.column_config.NumberColumn("序號", disabled=True, width="small"),
                "提領ID": st.column_config.TextColumn("提領ID", disabled=True),
                "日期": st.column_config.TextColumn("日期", disabled=True),
                "收款帳戶": st.column_config.TextColumn("收款帳戶", disabled=True),
                "提領金額": st.column_config.NumberColumn("提領金額", disabled=True),
                "來源類別": st.column_config.TextColumn("來源類別", disabled=True),
                "來源ID": st.column_config.TextColumn("來源ID", disabled=True),
                "來源項目": st.column_config.TextColumn("來源項目", disabled=True),
                "來源金額": st.column_config.NumberColumn("來源金額", disabled=True),
                "來源付款人": st.column_config.TextColumn("來源付款人", disabled=True),
            },
            disabled=["序號", "提領ID", "日期", "收款帳戶", "提領金額", "來源類別", "來源ID", "來源項目", "來源金額", "來源付款人"]
        )
        if st.button("🗑️ 刪除已勾選提領紀錄"):
            selected_rows = edited_w_all[edited_w_all["刪除"] == True]
            if selected_rows.empty:
                st.warning("請先勾選要刪除的提領紀錄")
            else:
                delete_ids = set(selected_rows["提領ID"].dropna().tolist())
                for wid in delete_ids:
                    delete_withdrawal(wid)
                st.success(f"已刪除 {len(delete_ids)} 筆提領紀錄，並復原來源支出『提領過』。")
                st.rerun()


    st.divider()
    st.subheader("建立提領")

    c1, c2, c3 = st.columns(3)
    with c1:
        w_date = st.date_input("日期", value=date.today(), key="w_date")
    with c2:
        w_amount = st.number_input("金額", min_value=0.0, step=100.0, key="w_amount")
    with c3:
        w_account = st.text_input("收款帳戶", placeholder="如：滋滋農行 / 滋滋富邦 / 銀行匯費", key="w_account")

    f1, f2 = st.columns(2)
    with f1:
        year_filter = st.selectbox("篩選年份", options=["全部"] + list(range(2021, 2051)), index=0, key="w_year_filter")
    with f2:
        month_filter = st.selectbox("篩選月份", options=["全部"] + list(range(1, 13)), index=0, key="w_month_filter")

    y_sel = None if year_filter == "全部" else int(year_filter)
    m_sel = None if month_filter == "全部" else int(month_filter)

    df_eligible = build_eligible_expenses_df(y_sel, m_sel)
    if df_eligible.empty:
        st.info("目前沒有可提領的支出（需為發票=True 且 尚未提領）。")
    else:
        edited_eligible = st.data_editor(
            df_eligible,
            use_container_width=False,
            hide_index=True,
            key="eligible_editor",
            column_config={
                "選取": st.column_config.CheckboxColumn("選取", width="small"),
                "序號": st.column_config.NumberColumn("序號", disabled=True, width="small"),
                "類別": st.column_config.TextColumn("類別", disabled=True),
                "日期/年月": st.column_config.TextColumn("日期/年月", disabled=True),
                "項目": st.column_config.TextColumn("項目", disabled=True),
                "金額": st.column_config.NumberColumn("金額", disabled=True),
                "付款人": st.column_config.TextColumn("付款人", disabled=True),
                "uid": None
            },
            disabled=["序號", "類別", "日期/年月", "項目", "金額", "付款人", "uid"]
        )

        selected_uids = edited_eligible.loc[edited_eligible["選取"] == True, "uid"].tolist()
        selected_total = 0.0
        if selected_uids:
            selected_total = edited_eligible[edited_eligible["uid"].isin(selected_uids)]["金額"].sum()

        st.write(f"已選取 {len(selected_uids)} 筆，總額：{selected_total:,.2f}")

        if w_amount and abs(selected_total - w_amount) > 0.01:
            st.warning("選取總額與提領金額不一致，請調整。")

        if st.button("建立提領紀錄"):
            if w_amount <= 0:
                st.error("金額需大於 0")
            elif not selected_uids:
                st.error("請至少勾選一筆來源")
            elif abs(selected_total - w_amount) > 0.01:
                st.error("選取總額與提領金額不一致")
            elif not (w_account or "").strip():
                st.error("請輸入收款帳戶")
            else:
                ok = create_withdrawal(
                    w_date=w_date,
                    amount=float(w_amount),
                    account=w_account.strip(),
                    selected_uids=selected_uids
                )
                if ok:
                    st.session_state["flash_msg"] = "提領成功！"
                    st.session_state["reset_withdrawal_form"] = True
                    st.rerun()


# =========================
# TAB3 - 分析
# =========================
with tab3:
    df_income = pd.DataFrame(load_income_cloud())
    df_fixed = pd.DataFrame(load_fixed_cloud())
    df_other = pd.DataFrame(load_other_cloud())
    

    if df_fixed.empty:
        df_fixed = pd.DataFrame(columns=["年份", "月份", "項目", "金額", "付款人", "發票", "提領過"])
    if df_other.empty:
        df_other = pd.DataFrame(columns=["年份", "月份", "日期", "項目", "金額", "付款人", "發票", "提領過"])

    if "刪除" in df_fixed.columns:
        df_fixed = df_fixed.drop(columns=["刪除"], errors="ignore")
    if "刪除" in df_other.columns:
        df_other = df_other.drop(columns=["刪除"], errors="ignore")

    df_fixed = reorder_expense_columns(df_fixed).drop(columns=["刪除"], errors="ignore")
    for col in ["年份", "月份", "項目", "金額", "付款人", "發票", "提領過"]:
        if col not in df_other.columns:
            df_other[col] = 0 if col == "金額" else (False if col in ["發票", "提領過"] else "")
    df_other = df_other[["年份", "月份", "項目", "金額", "付款人", "發票", "提領過"]]

    if not df_fixed.empty:
        df_fixed["金額"] = df_fixed["金額"].apply(normalize_amount)
        df_fixed["發票"] = df_fixed["發票"].apply(normalize_bool)
        df_fixed["提領過"] = df_fixed["提領過"].apply(normalize_bool)

    if not df_other.empty:
        df_other["金額"] = df_other["金額"].apply(normalize_amount)
        df_other["發票"] = df_other["發票"].apply(normalize_bool)
        df_other["提領過"] = df_other["提領過"].apply(normalize_bool)

    df_exp_all = pd.concat([df_fixed, df_other], ignore_index=True) if not (df_fixed.empty and df_other.empty) else pd.DataFrame(columns=["年份", "月份", "項目", "金額", "付款人", "發票", "提領過"])

    income = df_income["金額"].sum() if not df_income.empty else 0.0
    expense = df_exp_all["金額"].sum() if not df_exp_all.empty else 0.0
    can_take = df_exp_all[(df_exp_all["發票"] == True) & (df_exp_all["提領過"] == False)]["金額"].sum() if not df_exp_all.empty else 0.0
    used = df_exp_all[df_exp_all["提領過"] == True]["金額"].sum() if not df_exp_all.empty else 0.0
    remain = can_take

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("收入", f"{income:,.1f}")
    c2.metric("支出（固定+其他）", f"{expense:,.1f}")
    c3.metric("可提領", f"{can_take:,.1f}")
    c4.metric("已提領", f"{used:,.1f}")
    c5.metric("剩餘", f"{remain:,.1f}")

    st.divider()
    st.subheader("收入堆疊直條圖（按月份）")

    if df_income.empty:
        st.info("尚無收入資料可分析")
    else:
        df_income_chart = df_income.copy()
        df_income_chart["年份"] = pd.to_numeric(df_income_chart["年份"], errors="coerce")
        df_income_chart["月份"] = pd.to_numeric(df_income_chart["月份"], errors="coerce")
        df_income_chart["金額"] = pd.to_numeric(df_income_chart["金額"], errors="coerce")
        df_income_chart = df_income_chart.dropna(subset=["年份", "月份", "帳戶", "金額"])

        if df_income_chart.empty:
            st.info("收入資料格式不足，無法繪圖")
        else:
            df_income_chart["月份標籤"] = (
                df_income_chart["年份"].astype(int).astype(str)
                + "-"
                + df_income_chart["月份"].astype(int).astype(str).str.zfill(2)
            )

            grouped = (
                df_income_chart.groupby(["月份標籤", "帳戶"], as_index=False)["金額"]
                .sum()
            )

            monthly_total = (
                grouped.groupby("月份標籤", as_index=False)["金額"]
                .sum()
                .rename(columns={"金額": "月總收入"})
            )

            grouped = grouped.merge(monthly_total, on="月份標籤", how="left")
            grouped["占比"] = grouped.apply(
                lambda r: 0 if r["月總收入"] == 0 else r["金額"] / r["月總收入"],
                axis=1
            )
            grouped["標籤"] = grouped.apply(
                lambda r: f'{r["金額"]:,.0f} ({r["占比"]:.0%})',
                axis=1
            )

            month_order = sorted(grouped["月份標籤"].unique().tolist())
            account_order = sorted(grouped["帳戶"].unique().tolist())

            grouped["帳戶"] = pd.Categorical(grouped["帳戶"], categories=account_order, ordered=True)
            grouped = grouped.sort_values(["月份標籤", "帳戶"]).reset_index(drop=True)

            grouped["stack_start"] = grouped.groupby("月份標籤")["金額"].cumsum() - grouped["金額"]
            grouped["stack_end"] = grouped.groupby("月份標籤")["金額"].cumsum()
            grouped["stack_mid"] = (grouped["stack_start"] + grouped["stack_end"]) / 2

            bars = alt.Chart(grouped).mark_bar().encode(
                x=alt.X("月份標籤:N", title="月份", sort=month_order),
                y=alt.Y("stack_end:Q", title="收入金額"),
                y2="stack_start:Q",
                color=alt.Color("帳戶:N", title="帳戶", sort=account_order),
                tooltip=[
                    alt.Tooltip("月份標籤:N", title="月份"),
                    alt.Tooltip("帳戶:N", title="帳戶"),
                    alt.Tooltip("金額:Q", title="金額", format=",.0f"),
                    alt.Tooltip("占比:Q", title="占比", format=".1%")
                ]
            )

            text = alt.Chart(grouped).mark_text(
                color="white",
                size=16,
                fontWeight="bold",
                align="center",
                baseline="middle"
            ).encode(
                x=alt.X("月份標籤:N", sort=month_order),
                y=alt.Y("stack_mid:Q"),
                detail="帳戶:N",
                text="標籤:N"
            )

            income_stack_chart = (bars + text).properties(height=450)
            st.altair_chart(income_stack_chart, use_container_width=True)

    st.divider()
    st.subheader("收入 / 總支出 / 該月盈餘 / 累計盈餘（按月份）")

    income_monthly = pd.DataFrame(columns=["月份標籤", "收入"])
    expense_monthly = pd.DataFrame(columns=["月份標籤", "總支出"])

    if not df_income.empty:
        tmp_income = df_income.copy()
        tmp_income["年份"] = pd.to_numeric(tmp_income["年份"], errors="coerce")
        tmp_income["月份"] = pd.to_numeric(tmp_income["月份"], errors="coerce")
        tmp_income["金額"] = pd.to_numeric(tmp_income["金額"], errors="coerce")
        tmp_income = tmp_income.dropna(subset=["年份", "月份", "金額"])

        if not tmp_income.empty:
            tmp_income["月份標籤"] = (
                tmp_income["年份"].astype(int).astype(str)
                + "-"
                + tmp_income["月份"].astype(int).astype(str).str.zfill(2)
            )
            income_monthly = (
                tmp_income.groupby("月份標籤", as_index=False)["金額"]
                .sum()
                .rename(columns={"金額": "收入"})
            )

    if not df_exp_all.empty:
        tmp_exp = df_exp_all.copy()
        tmp_exp["年份"] = pd.to_numeric(tmp_exp["年份"], errors="coerce")
        tmp_exp["月份"] = pd.to_numeric(tmp_exp["月份"], errors="coerce")
        tmp_exp["金額"] = pd.to_numeric(tmp_exp["金額"], errors="coerce")
        tmp_exp = tmp_exp.dropna(subset=["年份", "月份", "金額"])

        if not tmp_exp.empty:
            tmp_exp["月份標籤"] = (
                tmp_exp["年份"].astype(int).astype(str)
                + "-"
                + tmp_exp["月份"].astype(int).astype(str).str.zfill(2)
            )
            expense_monthly = (
                tmp_exp.groupby("月份標籤", as_index=False)["金額"]
                .sum()
                .rename(columns={"金額": "總支出"})
            )

    monthly_summary = pd.merge(
        income_monthly,
        expense_monthly,
        on="月份標籤",
        how="outer"
    ).fillna(0)

    if monthly_summary.empty:
        st.info("尚無足夠資料可分析收入與支出趨勢")
    else:
        monthly_summary = monthly_summary.sort_values("月份標籤").reset_index(drop=True)
        monthly_summary["該月盈餘"] = monthly_summary["收入"] - monthly_summary["總支出"]
        monthly_summary["累計盈餘"] = monthly_summary["該月盈餘"].cumsum()

        bars_df = monthly_summary.melt(
            id_vars=["月份標籤", "該月盈餘", "累計盈餘"],
            value_vars=["收入", "總支出"],
            var_name="類型",
            value_name="金額"
        )

        month_order = monthly_summary["月份標籤"].tolist()

        bar_chart = alt.Chart(bars_df).mark_bar().encode(
            x=alt.X("月份標籤:N", title="月份", sort=month_order),
            xOffset=alt.XOffset("類型:N"),
            y=alt.Y("金額:Q", title="金額"),
            color=alt.Color(
                "類型:N",
                title="類型",
                scale=alt.Scale(domain=["收入", "總支出"], range=["#4CAF50", "#F44336"])
            ),
            tooltip=[
                alt.Tooltip("月份標籤:N", title="月份"),
                alt.Tooltip("類型:N", title="類型"),
                alt.Tooltip("金額:Q", title="金額", format=",.0f")
            ]
        )

        line_profit = alt.Chart(monthly_summary).mark_line(point=True, strokeWidth=3, color="#2196F3").encode(
            x=alt.X("月份標籤:N", sort=month_order),
            y=alt.Y("該月盈餘:Q", title="金額"),
            tooltip=[
                alt.Tooltip("月份標籤:N", title="月份"),
                alt.Tooltip("該月盈餘:Q", title="該月盈餘", format=",.0f")
            ]
        )

        line_acc = alt.Chart(monthly_summary).mark_line(point=True, strokeWidth=3, color="#FF9800").encode(
            x=alt.X("月份標籤:N", sort=month_order),
            y=alt.Y("累計盈餘:Q", title="金額"),
            tooltip=[
                alt.Tooltip("月份標籤:N", title="月份"),
                alt.Tooltip("累計盈餘:Q", title="累計盈餘", format=",.0f")
            ]
        )

        combined_chart = (bar_chart + line_profit + line_acc).resolve_scale(
            y="shared"
        ).properties(height=500)

        st.altair_chart(combined_chart, use_container_width=True)

        st.markdown("### 月彙總資料")
        st.dataframe(
            monthly_summary.rename(columns={
                "月份標籤": "月份",
                "收入": "收入",
                "總支出": "總支出",
                "該月盈餘": "該月盈餘",
                "累計盈餘": "累計盈餘"
            }),
            use_container_width=True,
            hide_index=True
        )
