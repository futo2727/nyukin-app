# -*- coding: utf-8 -*-
# ------------------------------------------------------------
# 振込名義 → 店舗名・店舗番号 変換ツール（Streamlit / クラウド対応版）
#
# やること：
#   ① 通帳のカタカナ名義（複数行）を貼り付ける
#   ② 対応表でカタカナ → 読み換える店舗名 を調べる
#   ③ その店舗名をマスターであいまい検索して店舗番号を出す
#   ④ 見つからなかった名義は下部に一覧表示。常設フォームから対応表へ追記できる
#
# データの置き場所（自動で切り替わります）：
#   ・マスター（番号_店名）        … このフォルダの「番号_店名.csv」（読み取り専用）
#   ・対応表（振込名義）           … 設定があれば Google スプレッドシート、
#                                    無ければ「振込名義_店舗名対応表.csv」（ローカル）
#
#   ネット公開（Streamlit Cloud）では、Secrets に次を入れると
#   自動で Google スプレッドシート＋合言葉モードになります：
#     - gcp_service_account : Google のサービスアカウント鍵（表形式）
#     - taiou_sheet_url     : 対応表スプレッドシートのURL
#     - app_password        : 入口の合言葉
# ------------------------------------------------------------

import streamlit as st
import pandas as pd
import unicodedata
import re
import csv
import io
import os
import shutil
from datetime import datetime
from difflib import SequenceMatcher

# ============================================================
# 0. ファイルの場所を決める（このapp.pyと同じフォルダ）
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TAIOU_CSV = os.path.join(BASE_DIR, "振込名義_店舗名対応表.csv")   # ローカルで使う対応表
MASTER_CSV = os.path.join(BASE_DIR, "番号_店名.csv")             # 正しいマスター

# 文字コードは自動判定する。UTF-8を優先し、ダメならShift-JIS(cp932)を試す
ENCODINGS = ["utf-8-sig", "utf-8", "cp932"]


# ============================================================
# 1. Secrets（ネット公開時の秘密設定）の確認
# ============================================================
def has_secret(key):
    """Secretsに指定のキーがあるか安全に調べる（ローカルでSecretsが無くてもエラーにしない）。"""
    try:
        return key in st.secrets
    except Exception:
        return False


# 対応表をGoogleスプレッドシートで扱うかどうか（Secretsが揃っていればON）
USE_SHEETS = has_secret("gcp_service_account") and has_secret("taiou_sheet_url")


# ============================================================
# 2. 共通の道具（文字コード判定・文字ならし）
# ============================================================
def read_text_auto(path):
    """CSVを文字化けしないように文字コードを自動判定して読み込む。戻り値：(中身, 文字コード名)"""
    for enc in ENCODINGS:
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return f.read(), enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        return f.read(), "utf-8"


def normalize(s):
    """表記ゆれをそろえる（あいまい検索の土台）。
       全角/半角統一(NFKC) → 空白削除 → カッコや記号を削除。"""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"\s+", "", s)
    for ch in "（）()【】[]「」『』｛｝{}・,、。.／/":
        s = s.replace(ch, "")
    return s


def similarity(a, b):
    """2つの文字列の「似ている度」を0〜1で返す"""
    return SequenceMatcher(None, a, b).ratio()


# ============================================================
# 3. Google スプレッドシート接続（対応表の読み書き用）
# ============================================================
@st.cache_resource
def get_taiou_worksheet():
    """対応表スプレッドシートの1枚目のシートに接続して返す。"""
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(st.secrets["taiou_sheet_url"])
    return sheet.sheet1


# ============================================================
# 4. データの読み込み（結果はキャッシュして高速化）
# ============================================================
@st.cache_data
def load_master():
    """「番号_店名.csv」を読み込んで、きれいな (店舗番号, 店名) の表にする。

    このCSVは特殊な形：
      ・1行の中に「店舗番号,店名」が横に3組ならんでいる
      ・下のほうに文字化けした重複データがある（→使わない）
    """
    text, enc = read_text_auto(MASTER_CSV)
    rows = list(csv.reader(io.StringIO(text)))

    # --- 下部の文字化け重複を切り捨てる（全列が空の行が3連続したら以降を捨てる）---
    cutoff = len(rows)
    empty_run = 0
    for i, row in enumerate(rows):
        if all((c or "").strip() == "" for c in row):
            empty_run += 1
            if empty_run >= 3:
                cutoff = i - empty_run + 1
                break
        else:
            empty_run = 0
    rows = rows[:cutoff]

    # --- 横3組を縦につなぎ直す ---
    records = []
    for row in rows:
        if len(row) < 9:
            row = row + [""] * (9 - len(row))
        for ni, si in [(0, 1), (3, 4), (6, 7)]:
            num = (row[ni] or "").strip()
            name = (row[si] or "").replace("\n", "").replace("\r", "").strip()
            if num == "" and name == "":
                continue
            if num == "店舗番号" or name in ("店名", "店舗名"):
                continue
            if name == "":
                continue
            records.append((num, name))

    seen = set()
    uniq = []
    for num, name in records:
        key = (num, name)
        if key in seen:
            continue
        seen.add(key)
        uniq.append({"店舗番号": num, "店名": name, "_norm": normalize(name)})

    df = pd.DataFrame(uniq, columns=["店舗番号", "店名", "_norm"])
    return df, enc


@st.cache_data(ttl=60)
def load_taiou():
    """対応表（振込名義, 読み換える店舗名, メモ）を読み込む。
       設定があればGoogleスプレッドシート、無ければローカルCSVから。
       戻り値：(表, データの出どころの説明文字列)"""
    if USE_SHEETS:
        ws = get_taiou_worksheet()
        rows = ws.get_all_values()   # 見出し行を含む全行（CSVと同じ形）
        source = "Google スプレッドシート"
    else:
        text, enc = read_text_auto(TAIOU_CSV)
        rows = list(csv.reader(io.StringIO(text)))
        source = f"ローカルCSV（{enc}）"

    data = []
    for i, row in enumerate(rows):
        if i == 0:
            continue  # 1行目は見出し
        row = list(row)
        if len(row) < 3:
            row = row + [""] * (3 - len(row))
        meigi = (row[0] or "").strip()
        tenpo = (row[1] or "").strip()
        memo = (row[2] or "").strip()
        if meigi == "" and tenpo == "":
            continue
        data.append({
            "振込名義": meigi,
            "読み換える店舗名": tenpo,
            "メモ": memo,
            "_norm_meigi": normalize(meigi),
        })

    df = pd.DataFrame(data, columns=["振込名義", "読み換える店舗名", "メモ", "_norm_meigi"])
    return df, source


# ============================================================
# 5. 検索ロジック
# ============================================================
def split_stores(tenpo):
    """「読み換える店舗名」が「・」などで複数書かれている場合、分割する。"""
    parts = re.split(r"[・、,／/]", str(tenpo))
    return [p.strip() for p in parts if p.strip() != ""]


def lookup_meigi(meigi, taiou_df):
    """カタカナ名義で対応表を調べる。見つかれば {読み換える店舗名, メモ}、無ければ None。"""
    nq = normalize(meigi)
    if nq == "" or taiou_df.empty:
        return None
    hit = taiou_df[taiou_df["_norm_meigi"] == nq]
    if len(hit) == 0:
        return None
    r = hit.iloc[0]
    return {"読み換える店舗名": r["読み換える店舗名"], "メモ": r["メモ"]}


def find_numbers(store_name, master_df):
    """1つの店舗名について、マスターから店舗番号の候補をあいまい検索する。
       1) 完全一致 と 2) 部分一致 をまとめて全部表示 →
       3) それも無ければ「似ている度0.6以上」を上位5件。"""
    nq = normalize(store_name)
    if nq == "" or master_df.empty:
        return []

    results = []
    matched_keys = set()

    for _, r in master_df[master_df["_norm"] == nq].iterrows():
        matched_keys.add((r["店舗番号"], r["店名"]))
        results.append({"店舗番号": r["店舗番号"], "店名": r["店名"],
                        "種類": "完全一致", "スコア": 1.0})

    for _, r in master_df.iterrows():
        key = (r["店舗番号"], r["店名"])
        if key in matched_keys:
            continue
        nm = r["_norm"]
        if nm == "":
            continue
        if nq in nm or nm in nq:
            matched_keys.add(key)
            results.append({"店舗番号": r["店舗番号"], "店名": r["店名"],
                            "種類": "部分一致", "スコア": round(similarity(nq, nm), 2)})

    if not results:
        sims = []
        for _, r in master_df.iterrows():
            nm = r["_norm"]
            if nm == "":
                continue
            score = similarity(nq, nm)
            if score >= 0.6:
                sims.append((score, r))
        sims.sort(key=lambda x: x[0], reverse=True)
        for score, r in sims[:5]:
            results.append({"店舗番号": r["店舗番号"], "店名": r["店名"],
                            "種類": "似ている候補", "スコア": round(score, 2)})

    return results


# ============================================================
# 6. 対応表への追記
# ============================================================
def append_taiou_row(meigi, store, memo):
    """対応表に1行追記する。
       ・Googleスプレッドシート利用時：シートの末尾に1行追加（版の履歴に自動で残る）
       ・ローカルCSV利用時：バックアップを取ってから末尾に1行追記
       戻り値：バックアップのパス（CSV時のみ。シート時はNone）"""
    if USE_SHEETS:
        ws = get_taiou_worksheet()
        ws.append_row([meigi, store, memo], value_input_option="USER_ENTERED")
        return None

    # --- ローカルCSVの場合 ---
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BASE_DIR, f"振込名義_店舗名対応表.backup_{ts}.csv")
    shutil.copy2(TAIOU_CSV, backup_path)

    _, enc = read_text_auto(TAIOU_CSV)
    write_enc = "utf-8" if enc in ("utf-8", "utf-8-sig") else enc

    with open(TAIOU_CSV, "rb") as f:
        raw = f.read()
    need_newline = len(raw) > 0 and not raw.endswith((b"\n", b"\r"))

    with open(TAIOU_CSV, "a", encoding=write_enc, newline="") as f:
        if need_newline:
            f.write("\r\n")
        csv.writer(f).writerow([meigi, store, memo])

    return backup_path


# ============================================================
# 7. 画面（UI）
# ============================================================
st.set_page_config(page_title="振込名義→店舗名・番号 変換", layout="wide")


def check_password():
    """合言葉(app_password)が設定されていれば、入口でパスワードを求める。
       設定が無いローカルではそのまま通す。"""
    if not has_secret("app_password"):
        return
    if st.session_state.get("authenticated"):
        return
    st.title("🔒 合言葉を入力してください")
    pw = st.text_input("合言葉", type="password")
    if st.button("入る"):
        if pw == st.secrets["app_password"]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("合言葉が違います。")
    st.stop()


check_password()

st.markdown(
    "<h1 style='font-size:clamp(1.35rem, 3.4vw, 1.9rem); font-weight:700; margin:0 0 0.6rem 0;'>"
    "🏦 振込名義 → 店舗名・店舗番号 変換ツール</h1>",
    unsafe_allow_html=True,
)

# データ読み込み
master_df, master_enc = load_master()
taiou_df, taiou_source = load_taiou()

# サイドバー（状態の確認用）
with st.sidebar:
    st.header("データの状態")
    st.write(f"マスター（番号_店名）：**{len(master_df)}件**")
    st.caption(f"文字コード判定: {master_enc}")
    st.write(f"対応表（振込名義）：**{len(taiou_df)}件**")
    st.caption(f"保存先: {taiou_source}")
    st.info("「該当なし」は多く出ても正常です（対応表を育てながら使う道具です）。")

# ---------- ① 入力欄 ----------
st.subheader("① 通帳の名義（カタカナ）を貼り付け")
st.caption("改行で区切って、上から順に1件ずつ変換します。")
text = st.text_area(
    "名義（複数行OK）",
    height=200,
    placeholder="アメリカンハウス\nトータルプランニング\nアールサイエンス",
    key="input_text",
)

with st.expander("（将来用）画像アップロード欄 ※今は使いません"):
    st.file_uploader("通帳の画像", type=["png", "jpg", "jpeg"], disabled=True)
    st.caption("画像から文字を読む機能は後回しです。まずはテキスト入力で全機能が動きます。")

# ---------- 変換処理 ----------
meigi_list = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]

results = []
not_found = []

for meigi in meigi_list:
    info = lookup_meigi(meigi, taiou_df)
    if info is None:
        results.append({"meigi": meigi, "found": False})
        not_found.append(meigi)
    else:
        stores = split_stores(info["読み換える店舗名"])
        store_results = [{"入力店舗名": s, "候補": find_numbers(s, master_df)} for s in stores]
        results.append({
            "meigi": meigi,
            "found": True,
            "メモ": info["メモ"],
            "stores": store_results,
        })

# ---------- 結果表示 ----------
if meigi_list:
    st.subheader("変換結果")
    for item in results:
        st.markdown("---")
        st.markdown(f"### {item['meigi']}")

        if not item["found"]:
            st.warning("該当なし（未登録）")
            continue

        if item.get("メモ"):
            st.caption(f"📝 メモ：{item['メモ']}")

        for sr in item["stores"]:
            cands = sr["候補"]
            if not cands:
                st.write(f"・{sr['入力店舗名']} → 店舗番号が見つかりませんでした")
                continue
            for c in cands:
                tag = "" if c["種類"] == "完全一致" else f"　`{c['種類']}／似ている度 {c['スコア']}`"
                st.write(f"・**{c['店名']}**　店舗番号 **{c['店舗番号']}**{tag}")

# ============================================================
# 8. 該当なし一覧（あとで登録する候補リスト）
# ============================================================
if meigi_list:
    st.markdown("---")
    st.subheader("② 該当なし（未登録）の一覧")
    not_found_unique = list(dict.fromkeys(not_found))
    if not_found_unique:
        st.caption("あとで登録する候補です。「フォームへ」を押すと、下の登録フォームに名義が入ります。")
        for i, m in enumerate(not_found_unique):
            c1, c2 = st.columns([5, 1])
            c1.write(f"・{m}")
            if c2.button("フォームへ", key=f"tofrom_{i}"):
                st.session_state["reg_meigi"] = m
                st.rerun()
    else:
        st.success("該当なしはありませんでした。")

# ============================================================
# 9. 常設の登録フォーム（対応表を育てる）
# ============================================================
st.markdown("---")
st.subheader("③ 対応表に登録する（常設フォーム）")
st.caption("「読み換える店舗名」は、正しいマスター（番号_店名）の店名から検索して選びます（表記ゆれを防ぐため）。")

store_query = st.text_input("読み換える店舗名を検索（例：福臨閣、はしご など）", key="reg_store_query")

selected_store = None
if store_query.strip():
    nq = normalize(store_query)
    matched = master_df[master_df["_norm"].str.contains(re.escape(nq), na=False)]
    if len(matched) == 0:
        st.caption("該当する店名が見つかりません。文字を変えて試してください。")
    else:
        labels = [f'{row["店名"]}（{row["店舗番号"]}）' for _, row in matched.head(50).iterrows()]
        names = [row["店名"] for _, row in matched.head(50).iterrows()]
        choice = st.selectbox("候補から選ぶ", labels, key="reg_store_select")
        if choice:
            selected_store = names[labels.index(choice)]

with st.form("register_form", clear_on_submit=False):
    default_meigi = st.session_state.get("reg_meigi", "")
    reg_meigi = st.text_input("振込名義（カタカナ）", value=default_meigi)
    st.write(f"選択中の店舗名：**{selected_store if selected_store else '（上の検索で選んでください）'}**")
    reg_memo = st.text_input("メモ（任意）")
    submitted = st.form_submit_button("追加する", type="primary")

if submitted:
    if not reg_meigi.strip():
        st.error("振込名義を入力してください。")
    elif not selected_store:
        st.error("読み換える店舗名を、上の検索欄から選んでください。")
    else:
        backup_path = append_taiou_row(reg_meigi.strip(), selected_store, reg_memo.strip())
        load_taiou.clear()                       # 表を読み直すためキャッシュを消す
        st.session_state.pop("reg_meigi", None)  # フォームの名義をリセット
        st.success(f"追加しました：{reg_meigi.strip()} → {selected_store}")
        if backup_path:
            st.caption(f"（バックアップ：{os.path.basename(backup_path)}）")
        st.rerun()
