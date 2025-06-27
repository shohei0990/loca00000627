# app.py

import streamlit as st
import pandas as pd
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image
import datetime
from uuid import uuid4
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor


# --- 定義しておく ---
subcats = {
    "ハウススタジオ": ["和風","洋風","一軒家","マンション","アパート"],
    "オフィス":      ["執務室","会議室","ロビー"],
    "商業施設":      ["ショッピングモール","遊園地","温泉","水族館/動物園/植物園",
                     "博物館/美術館","映画館","ボーリング/ゲームセンター/ビリヤード場",
                     "商店街","ホテル"],
    "学校":          ["小学校","中学校","高校","大学/専門学校","幼稚園/保育園"],
    "病院":          ["受付","手術室"],
    "店舗":          ["コンビニ","ドラッグストア","スーパー","アパレル","ガソリンスタンド"],
    "飲食店":        ["中華料理屋","レストラン","カフェ","居酒屋","食堂","BAR"],
    "自然":          ["山","川","海","草原","森","湖/池","花畑","道"],
    "その他":        ["駐車場","屋上","神社仏閣","オープンスペース","夜景/イルミネーション",
                     "公民館","スポーツ施設"],
    "該当なし":      ["その他"]
}
detail_opts = [
    "ゼネ車使用可否","キッチン使用可否","同録の可否","養生の有無",
    "電源の有無","駐車場の有無","特機の使用可否","スモークの使用可否","火器の使用可否",
]

st.set_page_config(page_title="Location Uploader & PPTX Export", layout="wide")
st.title("📍 ロケ地情報入力＆エクセル／画像→PowerPoint")

# --- 1. 基本情報 ---
st.header("1. ロケ地基本情報")
loc_name       = st.text_input("ロケ地名")
address        = st.text_input("住所")
hp_link        = st.text_input("HPリンク（URL）")

# ロケ地種類：2カラム
c1, c2 = st.columns(2)
with c1:
    cat_main_val = st.selectbox("ロケ地種類（大分類）", list(subcats.keys()))
with c2:
    cat_sub_val  = st.selectbox("ロケ地種類（小分類）", subcats[cat_main_val])

transport_info = st.text_area("交通機関情報", height=80)

# 面積・天高：2カラム
a1, a2 = st.columns(2)
with a1:
    area_val   = st.number_input("面積 [m²]", min_value=0.0)
with a2:
    ceiling_val= st.number_input("天高 [cm]", min_value=0.0)

# 窓口連絡先：名前 + 電話3分割 + メール
st.subheader("窓口連絡先")
contact_person = st.text_input("窓口の担当者名")
ph1, ph2  = st.columns(2)
with ph1:
    phone1 = st.text_input("電話番号")
with ph2:
    contact_email = st.text_input("窓口のメールアドレス")

st.markdown("---")

# --- 2. 利用情報 --
st.header("2. 利用情報")
d1, d2 = st.columns(2)
with d1:
    price_day  = st.number_input("金額／day（¥）", min_value=0)
with d2:
    price_hour = st.number_input("金額／h（¥）",  min_value=0)

price_note = st.text_area("金額備考", height=80)

# 利用可能日時：24h + 開始・終了
st.subheader("利用可能日時")
t1, t2, t3 = st.columns([1,1,1])
with t1:
    open_24h   = st.checkbox("24時間利用可")
with t2:
    start_time = st.time_input("開始時間", value=datetime.time(9,0))
with t3:
    end_time   = st.time_input("終了時間", value=datetime.time(18,0))

st.markdown("---")

# --- 3. 詳細オプション ---
st.header("3. 詳細オプション")
detail_values = {}
for i in range(0, len(detail_opts), 3):
    cols = st.columns(3)
    for j, opt in enumerate(detail_opts[i:i+3]):
        detail_values[opt] = cols[j].selectbox(opt, ["あり","なし","不明"], index=2, key=opt)

# 支払い方法
payment    = st.selectbox("支払い方法", ["現金","カード","請求書","不明"], index=3)
pay_note   = st.text_input("支払い備考")
# 使用可能人数
st.subheader("使用可能人数")
u1, u2, u3, u4 = st.columns([1,1,1,1])
with u1:
    specify_num   = st.checkbox("人数指定")
with u2:
    max_number    = st.number_input("最大人数", min_value=0) if specify_num else None
with u3:
    unlimited     = st.checkbox("上限なし")
with u4:
    unknown_count = st.checkbox("不明")



st.markdown("---")

# --- 4. 対象作品 & 担当者 ---
st.header("4. 対象作品")
work_no     = st.text_input("作品番号")
pm, p, co   = st.columns(3)
with pm:
    pm_person   = st.text_input("担当者 PM")
with p:
    p_person    = st.text_input("担当者 P")
with co:
    coordinator = st.text_input("ロケコーディネーター")

st.markdown("---")

# --- 5. Excel 保存ボタン ---
st.header("5. 入力データをExcelで保存")
if st.button("💾 データをExcelでダウンロード"):
    # 辞書にまとめる
    data = {
        "ロケ地名":       loc_name,
        "住所":           address,
        "HPリンク":       hp_link,
        "大分類":         cat_main_val,
        "小分類":         cat_sub_val,
        "交通機関情報":   transport_info,
        "面積[m²]":       area_val,
        "天高[cm]":       ceiling_val,
        "担当者名":       contact_person,
        "電話番号":       phone1,
        "メールアドレス":  contact_email,
        "金額/day":       price_day,
        "金額/h":         price_hour,
        "金額備考":       price_note,
        "24時間可":       open_24h,
        "開始時間":       start_time.strftime("%H:%M"),
        "終了時間":       end_time.strftime("%H:%M"),
        **detail_values,
        "人数指定":       specify_num,
        "最大人数":       max_number,
        "上限なし":       unlimited,
        "不明(人数)":     unknown_count,
        "支払い方法":     payment,
        "支払い備考":     pay_note,
        "作品番号":       work_no,
        "担当者 PM":      pm_person,
        "担当者 P":       p_person,
        "コーディネーター": coordinator,
    }
    df = pd.DataFrame([data])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ロケ地情報")
    buf.seek(0)

    st.download_button(
        label="📥 Excelをダウンロード",
        data=buf,
        file_name="location_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")

# --- 6. 画像アップロード & PPTX 出力 ---
# --- Configuration ---

# Inputs for location metadata
location_name = loc_name
address = address

# --- ユーザー入力：1スライドあたりの画像枚数を選択 ---
# --- ユーザー入力：1スライドあたりの画像枚数 ---
total_per_slide = st.selectbox(
    "1スライドあたりの画像枚数",
    [6, 9],
    index=0
)
if total_per_slide == 6:
    PPT_COLS, PPT_ROWS = 3, 2
else:
    PPT_COLS, PPT_ROWS = 3, 3
PPT_PER_SLIDE = PPT_COLS * PPT_ROWS

# --- フォントサイズ定義 ---
HEADING_FONT = Pt(20)   # スライド上部の見出し
LOC_FONT     = Pt(20)   # ロケ地名
TABLE_FONT   = Pt(10)   # 表の文字

# --- プレビュー設定（変更なし）---
PREVIEW_COLS, PREVIEW_ROWS = 5, 2
PREVIEW_PER_PAGE = PREVIEW_COLS * PREVIEW_ROWS
PADDING = 5

# --- カテゴリ定義 ---
categories = [
    ("サムネイル", "thumbs", False),
    ("ロケ地写真", "photos", True),
    ("アングル写真", "angles", True),
    ("その他設備・搬入搬出経路", "others", True),
    ("平面図", "floor", True),
    ("ロケ地MAP", "map_img", True),
]

# --- セッションステート初期化 ---
for _, key, _ in categories:
    st.session_state.setdefault(f"{key}_data", {})
    st.session_state.setdefault(f"{key}_include", {})
    st.session_state.setdefault(f"{key}_page", 1)
    st.session_state.setdefault(f"{key}_ctr", 0)




# --- 画像アップロード＆プレビュー ---
for label, key, multi in categories:
    data = st.session_state[f"{key}_data"]
    ctr = st.session_state[f"{key}_ctr"]
    uploader_key = f"upl_{key}_{ctr}"
    uploaded = st.file_uploader(
        label,
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=multi,
        key=uploader_key
    )
    files = uploaded if isinstance(uploaded, list) else ([uploaded] if uploaded else [])
    if files:
        for f in files:
            if f.name not in data:
                data[f.name] = Image.open(f)
                st.session_state[f"{key}_include"][f.name] = True
        st.session_state[f"{key}_ctr"] += 1
        st.rerun()

    items = list(data.items())
    n = len(items)
    if n == 0:
        continue

    page_key = f"{key}_page"
    total_pages = (n + PREVIEW_PER_PAGE - 1) // PREVIEW_PER_PAGE
    st.session_state[page_key] = max(1, min(st.session_state[page_key], total_pages))
    page = st.session_state[page_key]

    start = (page - 1) * PREVIEW_PER_PAGE
    chunk = items[start:start + PREVIEW_PER_PAGE]

    st.markdown(
        f"<div style='border:1px solid #ddd; padding:{PADDING}px; margin-bottom:{PADDING}px; max-height:400px; overflow-y:auto;'>",
        unsafe_allow_html=True
    )
    cols_ui = st.columns(PREVIEW_COLS)
    for idx, (name, img) in enumerate(chunk):
        with cols_ui[idx % PREVIEW_COLS]:
            st.image(img, use_column_width=True)
            c1, c2 = st.columns([4, 1])
            with c1:
                inc = st.checkbox(
                    "資料出力",
                    key=f"inc_{key}_{name}",
                    value=st.session_state[f"{key}_include"][name]
                )
                st.session_state[f"{key}_include"][name] = inc
            with c2:
                if st.button("❌", key=f"del_{key}_{name}"):
                    data.pop(name)
                    st.session_state[f"{key}_include"].pop(name, None)
                    new_n = len(data)
                    new_total = max(1, (new_n + PREVIEW_PER_PAGE - 1) // PREVIEW_PER_PAGE)
                    st.session_state[page_key] = min(page, new_total)
                    st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    if total_pages > 1:
        options = [str(i) for i in range(1, total_pages + 1)]
        sel = st.radio(
            label=f"ページ ({page}/{total_pages})",
            options=options,
            index=page - 1,
            horizontal=True,
            key=f"nav_{key}"
        )
        new_page = int(sel)
        if new_page != page:
            st.session_state[page_key] = new_page
            st.rerun()

# --- PPTX 生成＆ダウンロード ---
if st.button("💾 PPTX を生成"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # --- サマリースライド ---
    summary = prs.slides.add_slide(prs.slide_layouts[5])

    # (1) ロケ地名：中央寄せテキストボックス
    tb_loc = summary.shapes.add_textbox(
        Inches(1), Inches(0.3),
        prs.slide_width - Inches(2), Inches(0.8)
    )
    tf_loc = tb_loc.text_frame
    p_loc = tf_loc.paragraphs[0]
    run_loc = p_loc.add_run()
    run_loc.text = location_name
    p_loc.alignment = PP_ALIGN.CENTER
    run_loc.font.name = "YuGothic"
    run_loc.font.size = LOC_FONT

    # (2) 住所と件数テーブル
    rows = 2 + len(categories)
    tbl = summary.shapes.add_table(
        rows, 2,
        Inches(1), Inches(1.3),  # ロケ地名下に配置
        Inches(8), Inches(0.8 * rows)
    ).table

    # ヘッダー行とデータ行のフォントサイズを後から設定
    tbl.cell(0, 0).text = "住所"
    tbl.cell(0, 1).text = address
    for i, (label, key, _) in enumerate(categories, start=1):
        tbl.cell(i, 0).text = label
        tbl.cell(i, 1).text = str(len(st.session_state[f"{key}_data"]))

    # 表フォントサイズ調整
    for row in tbl.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = TABLE_FONT
                    run.font.name = "YuGothic"

    # --- 画像スライド ---
    for label, key, _ in categories:
        selected = [
            img for name, img in st.session_state[f"{key}_data"].items()
            if st.session_state[f"{key}_include"][name]
        ]
        if not selected:
            continue

        for i in range(0, len(selected), PPT_PER_SLIDE):
            chunk = selected[i:i + PPT_PER_SLIDE]
            slide = prs.slides.add_slide(prs.slide_layouts[5])

            # (1) カテゴリ名：左寄せ・小めの見出し
            tb_cat = slide.shapes.add_textbox(
                Inches(0.5), Inches(0.3),
                Inches(3), Inches(0.5)
            )
            tf_cat = tb_cat.text_frame
            p_cat = tf_cat.paragraphs[0]
            run_cat = p_cat.add_run()
            run_cat.text = label
            p_cat.alignment = PP_ALIGN.LEFT
            run_cat.font.name = "YuGothic"
            run_cat.font.size = HEADING_FONT

            # (2) ロケ地名：中央寄せ
            tb_cat_loc = slide.shapes.add_textbox(
                Inches(3.5), Inches(0.3),
                prs.slide_width - Inches(4), Inches(0.5)
            )
            tf_cat_loc = tb_cat_loc.text_frame
            p_cat_loc = tf_cat_loc.paragraphs[0]
            run_cat_loc = p_cat_loc.add_run()
            run_cat_loc.text = location_name
            p_cat_loc.alignment = PP_ALIGN.CENTER
            run_cat_loc.font.name = "YuGothic"
            run_cat_loc.font.size = HEADING_FONT

            # (3) 画像グリッド配置（省略）
            usable_w = prs.slide_width - Inches(1)
            usable_h = prs.slide_height - Inches(1.5)
            gap_w, gap_h = Inches(0.2), Inches(0.2)
            cell_w = (usable_w - gap_w * (PPT_COLS - 1)) / PPT_COLS
            cell_h = (usable_h - gap_h * (PPT_ROWS - 1)) / PPT_ROWS
            left_m, top_m = Inches(0.5), Inches(1.5)

            for idx, img in enumerate(chunk):
                r, c = divmod(idx, PPT_COLS)
                x = left_m + c * (cell_w + gap_w)
                y = top_m + r * (cell_h + gap_h)
                ow, oh = img.size
                ratio, cell_ratio = ow / oh, cell_w / cell_h
                if ratio > cell_ratio:
                    pw, ph = cell_w, cell_w / ratio
                else:
                    ph, pw = cell_h, cell_h * ratio
                px = x + (cell_w - pw) / 2
                py = y + (cell_h - ph) / 2

                if pw < cell_w or ph < cell_h:
                    bg = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE, x, y, cell_w, cell_h
                    )
                    bg.fill.solid()
                    bg.fill.fore_color.rgb = RGBColor(0, 0, 0)
                    bg.line.fill.background()

                buf = io.BytesIO()
                img.save(buf, format=img.format or "PNG")
                buf.seek(0)
                slide.shapes.add_picture(buf, px, py, width=pw, height=ph)

    # 保存＆ダウンロード準備
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    st.session_state['pptx_bytes'] = out.getvalue()

# --- ダウンロードボタン ---
if 'pptx_bytes' in st.session_state:
    st.download_button(
        "📥 PPTXをダウンロード",
        data=st.session_state['pptx_bytes'],
        file_name="location_pictures.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )