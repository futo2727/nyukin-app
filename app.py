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

# 「店舗コードなし（業者・その他）」の登録を表す目印。
# 対応表の「読み換える店舗名」の先頭にこれを付けて保存する（例：【店舗コードなし】トヨタ）
NOCODE_PREFIX = "【店舗コードなし】"


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


def normalize_meigi(s):
    """振込名義（カタカナ）用の、より強力な文字ならし。
       通帳のカタカナは表記ゆれが多いので、次まで吸収する：
       ・全角/半角統一(NFKC)
       ・ひらがな→カタカナ
       ・小さいカナ→大きいカナ（ャ→ヤ、ッ→ツ など）
       ・会社の略号を無視（カ) (カ) ユ) ド) など）
       ・空白・記号を削除
    """
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    # ひらがな→カタカナ
    s = "".join(chr(ord(c) + 0x60) if "ぁ" <= c <= "ゖ" else c for c in s)
    # 小さいカナ→大きいカナ
    small = "ァィゥェォッャュョヮヵヶ"
    large = "アイウエオツヤユヨワカケ"
    s = s.translate(str.maketrans(small, large))
    # 会社略号を無視： (カ) / (カ / カ)  などを削除
    s = re.sub(r"[（(][カユドメシイ][）)]?|[カユドメシイ][）)]", "", s)
    # 空白・記号を削除
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
            "_norm_meigi": normalize_meigi(meigi),
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
    """カタカナ名義で対応表を調べる。
       戻り値：
         ・ぴったり一致 → {"match":"exact", 読み換える店舗名, メモ}
         ・惜しい（表記ゆれ/1文字違い） → {"match":"near", "candidates":[(似ている度, 行), ...]}
         ・全然無い → None
    """
    nq = normalize_meigi(meigi)
    if nq == "" or taiou_df.empty:
        return None

    # まず、ぴったり一致
    hit = taiou_df[taiou_df["_norm_meigi"] == nq]
    if len(hit) > 0:
        r = hit.iloc[0]
        return {"match": "exact",
                "読み換える店舗名": r["読み換える店舗名"], "メモ": r["メモ"]}

    # 無ければ、似ている登録を探す（1文字違い・表記ゆれの救済）
    sims = []
    for _, r in taiou_df.iterrows():
        nm = r["_norm_meigi"]
        if not nm:
            continue
        sc = similarity(nq, nm)
        if sc >= 0.7:
            sims.append((sc, r))
    sims.sort(key=lambda x: x[0], reverse=True)
    if sims:
        return {"match": "near", "candidates": sims[:3]}
    return None


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


def load_taiou_editable():
    """削除UI用：対応表の各行を『削除の目印（行番号）』つきで返す。"""
    if USE_SHEETS:
        values = get_taiou_worksheet().get_all_values()
    else:
        text, _ = read_text_auto(TAIOU_CSV)
        values = list(csv.reader(io.StringIO(text)))

    entries = []
    for i, row in enumerate(values):
        if i == 0:
            continue  # 見出し
        row = list(row)
        if len(row) < 3:
            row = row + [""] * (3 - len(row))
        meigi = (row[0] or "").strip()
        tenpo = (row[1] or "").strip()
        memo = (row[2] or "").strip()
        if meigi == "" and tenpo == "":
            continue
        entries.append({"row": i + 1, "振込名義": meigi,
                        "読み換える店舗名": tenpo, "メモ": memo})
    return entries


def delete_taiou_row(row_number):
    """対応表から指定の行（1始まり）を削除する。"""
    if USE_SHEETS:
        get_taiou_worksheet().delete_rows(int(row_number))
        return None

    # ローカルCSV：バックアップしてから、その行を除いて書き直す
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BASE_DIR, f"振込名義_店舗名対応表.backup_{ts}.csv")
    shutil.copy2(TAIOU_CSV, backup_path)

    text, enc = read_text_auto(TAIOU_CSV)
    rows = list(csv.reader(io.StringIO(text)))
    write_enc = "utf-8" if enc in ("utf-8", "utf-8-sig") else enc
    target = int(row_number) - 1
    rows = [r for idx, r in enumerate(rows) if idx != target]
    with open(TAIOU_CSV, "w", encoding=write_enc, newline="") as f:
        csv.writer(f).writerows(rows)
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
st.caption("名義を貼り付けたら、下の「変換する」ボタンを押してください。（Enterキーは改行に使えます）")
st.text_area(
    "名義（複数行OK）",
    height=200,
    placeholder="アメリカンハウス\nトータルプランニング\nアールサイエンス",
    key="input_text",
)
if st.button("変換する", type="primary"):
    st.session_state["run_text"] = st.session_state.get("input_text", "")

with st.expander("（将来用）画像アップロード欄 ※今は使いません"):
    st.file_uploader("通帳の画像", type=["png", "jpg", "jpeg"], disabled=True)
    st.caption("画像から文字を読む機能は後回しです。まずはテキスト入力で全機能が動きます。")

# ---------- 変換処理（「変換する」ボタンで実行した内容を使う）----------
text_to_process = st.session_state.get("run_text", "")
meigi_list = [ln.strip() for ln in text_to_process.splitlines() if ln.strip() != ""]

results = []
not_found = []

def build_conversion(tenpo):
    """読み換える店舗名から変換結果を作る。
       ・「店舗コードなし」の登録 → {"no_code": True, "label": 表示名}
       ・通常（店舗名）        → {"no_code": False, "stores": [各店舗の番号候補]}"""
    t = str(tenpo).strip()
    if t.startswith(NOCODE_PREFIX):
        return {"no_code": True, "label": t[len(NOCODE_PREFIX):].strip()}
    return {"no_code": False,
            "stores": [{"入力店舗名": s, "候補": find_numbers(s, master_df)}
                       for s in split_stores(t)]}


for meigi in meigi_list:
    info = lookup_meigi(meigi, taiou_df)
    if info is None:
        # 全然見つからない → 未登録（②の一覧＆登録候補に入れる）
        results.append({"meigi": meigi, "found": False, "near": []})
        not_found.append(meigi)
    elif info["match"] == "exact":
        results.append({
            "meigi": meigi,
            "found": True,
            "メモ": info["メモ"],
            "conv": build_conversion(info["読み換える店舗名"]),
        })
    else:
        # 惜しい（表記ゆれ/1文字違い）→ 近い登録を「もしかして」で提示
        near_list = []
        for sc, r in info["candidates"]:
            near_list.append({
                "振込名義": r["振込名義"],
                "メモ": r["メモ"],
                "score": round(sc, 2),
                "conv": build_conversion(r["読み換える店舗名"]),
            })
        # 登録済みの可能性が高いので、②の未登録一覧には入れない
        results.append({"meigi": meigi, "found": False, "near": near_list})

# ---------- 結果表示 ----------
if meigi_list:
    st.subheader("変換結果")
    for item in results:
        st.markdown("---")
        st.markdown(f"### {item['meigi']}")

        def show_conversion(conv):
            # 店舗コードなし（業者など）の場合
            if conv.get("no_code"):
                lbl = conv.get("label")
                suffix = f"：{lbl}" if lbl else ""
                st.write(f"・🏷 **店舗コードなし（業者など）**{suffix}")
                return
            # 通常の店舗
            for sr in conv["stores"]:
                cands = sr["候補"]
                if not cands:
                    st.write(f"・{sr['入力店舗名']} → 店舗番号が見つかりませんでした")
                    continue
                for c in cands:
                    tag = "" if c["種類"] == "完全一致" else f"　`{c['種類']}／似ている度 {c['スコア']}`"
                    st.write(f"・**{c['店名']}**　店舗番号 **{c['店舗番号']}**{tag}")

        if item["found"]:
            if item.get("メモ"):
                st.caption(f"📝 メモ：{item['メモ']}")
            show_conversion(item["conv"])
            continue

        # 見つからなかった場合
        if item.get("near"):
            # 惜しい登録がある → もしかして提示
            st.info("ぴったり一致はありませんでした。近い登録を表示します（入力の表記ゆれかもしれません）：")
            for nc in item["near"]:
                st.markdown(f"**🔎 もしかして：{nc['振込名義']}**　`似ている度 {nc['score']}`")
                if nc["メモ"]:
                    st.caption(f"📝 メモ：{nc['メモ']}")
                show_conversion(nc["conv"])
        else:
            st.warning("該当なし（未登録）")

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

# 登録の種類：店舗（番号あり） or 店舗コードなし（業者など）
reg_mode = st.radio(
    "登録の種類",
    ["店舗として登録（番号あり）", "店舗コードなし（業者・その他）"],
    key="reg_mode",
    horizontal=True,
)
is_store_mode = reg_mode.startswith("店舗として登録")

selected_store = None   # 店舗モードで選ばれた店名
nocode_label = ""       # コードなしモードの表示ラベル

if is_store_mode:
    st.caption("「読み換える店舗名」は、正しいマスター（番号_店名）の店名から検索して選びます（表記ゆれを防ぐため）。")
    store_query = st.text_input("読み換える店舗名を検索（例：福臨閣、はしご など）", key="reg_store_query")
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
else:
    st.caption("トヨタなどの業者のように、店舗番号が無いものを登録します。")
    nocode_label = st.text_input("表示名（任意：例 トヨタ、〇〇商事、業者 など）", key="reg_nocode_label")

with st.form("register_form", clear_on_submit=False):
    default_meigi = st.session_state.get("reg_meigi", "")
    reg_meigi = st.text_input("振込名義（カタカナ）", value=default_meigi)
    if is_store_mode:
        st.write(f"選択中の店舗名：**{selected_store if selected_store else '（上の検索で選んでください）'}**")
    else:
        shown_label = nocode_label.strip() if nocode_label.strip() else "（表示名なし）"
        st.write(f"種類：**店舗コードなし（業者など）**／表示名：**{shown_label}**")
    reg_memo = st.text_input("メモ（任意）")
    submitted = st.form_submit_button("追加する", type="primary")

if submitted:
    if not reg_meigi.strip():
        st.error("振込名義を入力してください。")
    elif is_store_mode and not selected_store:
        st.error("読み換える店舗名を、上の検索欄から選んでください。")
    else:
        if is_store_mode:
            store_value = selected_store
            done_msg = f"追加しました：{reg_meigi.strip()} → {selected_store}"
        else:
            store_value = NOCODE_PREFIX + nocode_label.strip()
            done_msg = f"追加しました：{reg_meigi.strip()} → 店舗コードなし（業者など）"
        backup_path = append_taiou_row(reg_meigi.strip(), store_value, reg_memo.strip())
        load_taiou.clear()                       # 表を読み直すためキャッシュを消す
        st.session_state.pop("reg_meigi", None)  # フォームの名義をリセット
        st.success(done_msg)
        if backup_path:
            st.caption(f"（バックアップ：{os.path.basename(backup_path)}）")
        st.rerun()

# ============================================================
# 10. 登録の取り消し（間違えて追加したときに削除）
# ============================================================
st.markdown("---")
st.subheader("④ 登録の取り消し（間違えて追加したとき）")
st.caption("登録済みの対応を確認して、不要なものを削除できます。")

if st.checkbox("登録済みの一覧を表示する"):
    entries = load_taiou_editable()
    st.caption(f"現在 {len(entries)} 件 登録されています。")

    del_query = st.text_input("絞り込み（振込名義や店舗名の一部を入力）", key="del_query")
    if del_query.strip():
        nq = normalize(del_query)
        shown = [e for e in entries
                 if nq in normalize(e["振込名義"]) or nq in normalize(e["読み換える店舗名"])]
    else:
        shown = entries

    if not shown:
        st.write("該当する登録がありません。")
    else:
        for e in shown:
            c1, c2 = st.columns([6, 1])
            memo_txt = f"　📝{e['メモ']}" if e["メモ"] else ""
            c1.write(f"**{e['振込名義']}** → {e['読み換える店舗名']}{memo_txt}")
            if c2.button("削除", key=f"del_{e['row']}"):
                st.session_state["confirm_delete"] = e
                st.rerun()

    # 削除の確認（間違って消さないように一度確認する）
    cd = st.session_state.get("confirm_delete")
    if cd:
        st.warning(f"「{cd['振込名義']} → {cd['読み換える店舗名']}」を削除します。よろしいですか？")
        b1, b2 = st.columns(2)
        if b1.button("はい、削除する", type="primary"):
            delete_taiou_row(cd["row"])
            load_taiou.clear()
            st.session_state.pop("confirm_delete", None)
            st.success("削除しました。")
            st.rerun()
        if b2.button("キャンセル"):
            st.session_state.pop("confirm_delete", None)
            st.rerun()
