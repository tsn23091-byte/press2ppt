import io
import re
from typing import List, Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from readability import Document
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
import streamlit as st
from openai import OpenAI

APP_VERSION = "press2ppt v1.4 - URL 取得 + コピペ&画像アップ対応"

# ========= 設定 =========
TEMPLATE_PATH = "templates/cuprum_template.pptx"
DEFAULT_FONTS = ["Meiryo", "Yu Gothic UI", "MS UI Gothic", "Calibri"]

# スタイル設定（要望通り）
TITLE_COLOR = RGBColor(255, 255, 255)   # 白
TITLE_SIZE_PT = 28
TITLE_BOLD = True

BODY_COLOR = RGBColor(0, 153, 153)      # #009999
BODY_SIZE_PT = 24
BODY_BOLD = True

# ロゴ等の除外（URLモードで使用）
IMG_EXCLUDE_RE = re.compile(
    r"(?:^|[-_/])(logo|favicon|sprite|badge|mark|header|footer|og_image|common/images/og_image\.png)\b",
    re.IGNORECASE,
)
EXACT_IMG_BLACKLIST = {
    "https://www.jx-nmm.com/common/images/og_image.png",
    "http://www.jx-nmm.com/common/images/og_image.png",
}

# ========= OpenAI =========
def get_client(api_key: Optional[str]):
    try:
        return OpenAI(api_key=api_key) if api_key else OpenAI()
    except Exception:
        return None

# ========= HTML取得（公開サイト用 / URLモード） =========
def fetch_html_public(url: str) -> str:
    r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    return r.text

def _best_from_srcset(srcset: str) -> Optional[str]:
    try:
        parts = [p.strip() for p in srcset.split(",") if p.strip()]
        pairs = []
        for p in parts:
            s = p.split()
            if len(s) == 1:
                pairs.append((s[0], 0))
            else:
                w = 0
                try:
                    if s[1].endswith("w"):
                        w = int(s[1][:-1])
                except Exception:
                    pass
                pairs.append((s[0], w))
        pairs.sort(key=lambda x: x[1], reverse=True)
        return pairs[0][0] if pairs else None
    except Exception:
        return None

def parse_page(url: str) -> dict:
    html = fetch_html_public(url)
    return _parse_common(html, base_url=url)

def _parse_common(html: str, base_url: str = "") -> dict:
    doc = Document(html)
    title = (doc.short_title() or "").strip()
    if not title:
        try:
            head = BeautifulSoup(html, "lxml").find("head")
            if head:
                t = head.find("title")
                if t and t.get_text(strip=True):
                    title = t.get_text(strip=True)
        except Exception:
            pass

    main_html = doc.summary(html_partial=True)
    soup = BeautifulSoup(main_html, "lxml")

    # 本文テキスト
    ps = [p.get_text(" ", strip=True) for p in soup.find_all("p")]
    text = " ".join(ps)

    # 画像候補（相対→絶対）
    def abs_url(u: str) -> str:
        if not base_url:
            return u
        if u and u.startswith("/"):
            return urljoin(base_url, u)
        return u

    urls: List[str] = []
    for img in soup.find_all("img"):
        cand = img.get("src") or img.get("data-src") or img.get("data-original")
        if not cand and img.get("srcset"):
            cand = _best_from_srcset(img.get("srcset"))
        if cand:
            urls.append(abs_url(cand))

    for pic in soup.find_all("picture"):
        for src in pic.find_all("source"):
            srcset = src.get("srcset")
            if srcset:
                cand = _best_from_srcset(srcset)
                if cand:
                    urls.append(abs_url(cand))

    # フォールバック: head の og/twitter image
    try:
        head = BeautifulSoup(html, "lxml").find("head")
        if head:
            for prop in ["og:image", "twitter:image"]:
                tag = head.find("meta", property=prop) or head.find("meta", attrs={"name": prop})
                if tag and tag.get("content"):
                    urls.append(abs_url(tag["content"]))
    except Exception:
        pass

    # 除外 & 重複排除
    cleaned = []
    for u in urls:
        if not u or u.startswith("data:"):
            continue
        low = u.lower()
        u_noquery = re.sub(r"[?#].*$", "", u)
        if u_noquery in EXACT_IMG_BLACKLIST:
            continue
        if IMG_EXCLUDE_RE.search(low):
            continue
        cleaned.append(u)

    uniq, seen = [], set()
    for u in cleaned:
        base = re.sub(r"[?#].*$", "", u)
        if base not in seen:
            uniq.append(u)
            seen.add(base)

    return {"title": title, "text": text, "images": uniq}

# ========= 要約 =========
SYS_TITLER = (
    "あなたは日本語のPRアシスタントです。25文字を超える場合のみ自然な見出しに短縮。"
    "句読点含め25文字以内、固有名詞は優先して保持。"
)
SYS_SUMMARY = (
    "あなたは日本語のPRアシスタントです。企業プレスリリースの要旨を指定の上限文字数以内で簡潔に要約。"
    "目的はプレスリリースの内容を社内発信することです。"
    "体言止めや重言を避け、固有名詞は維持。ですます調。"
    "JX金属株式会社が主語の場合は省略、同様に当社なども省略、それでも意味が通る内容に要約すること。"
)

def gpt_shorten_title(client: Optional[OpenAI], title: str) -> str:
    if len(title) <= 25 or not client:
        return title[:25]
    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": SYS_TITLER},
                {"role": "user", "content": f"元タイトル：{title}\n25文字以内に短く"},
            ],
            temperature=0.2,
        )
        return (resp.choices[0].message.content or "").strip()[:25]
    except Exception:
        return title[:25]

def offline_summary(text: str) -> str:
    if not text:
        return ""
    sentences = re.split(r"(?<=[。！？!?])", text)
    chunk = "".join(sentences[:3])
    chunk = re.sub(r"\s+", " ", chunk)
    return chunk

def gpt_summarize_body(client: Optional[OpenAI], text: str, max_len: int = 120) -> str:
    head = text[:4000]
    base = offline_summary(head)
    if not client:
        return base[:max_len]
    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": SYS_SUMMARY},
                {"role": "user", "content": f"{head}\n\n上限{max_len}文字で、重要点を落とさず簡潔に要約してください。"},
            ],
            temperature=0.2,
        )
        s = (resp.choices[0].message.content or "").strip()
        return s[:max_len]
    except Exception:
        return base[:max_len]

# ========= 画像DL（URLモード用） =========
def download_images(urls: List[str], limit: int = 4) -> List[Image.Image]:
    imgs: List[Image.Image] = []
    for u in urls[:limit]:
        try:
            r = requests.get(u, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            img = Image.open(io.BytesIO(r.content)).convert("RGB")
            if img.width < 300 or img.height < 180:
                continue
            imgs.append(img)
        except Exception:
            continue
    return imgs

# ========= PowerPoint生成 =========
def _first_placeholder(slide, types: tuple[int, ...]) -> Optional[object]:
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.type in types:
                return ph
        except Exception:
            continue
    return None

def _set_text(shape, text: str, size_pt: int, color: RGBColor, bold: bool):
    try:
        if not getattr(shape, "has_text_frame", False):
            return
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = text
        for fn in DEFAULT_FONTS:
            try:
                run.font.name = fn
                break
            except Exception:
                continue
        run.font.size = Pt(size_pt)
        run.font.color.rgb = color
        run.font.bold = bold
    except Exception:
        pass

def get_layout_by_name(prs, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return None

def build_pptx(template_path: str, title: str, summary: str, images: List[Image.Image]) -> bytes:
    prs = Presentation(template_path)

    n = min(len(images), 3)
    layout_name = {
        0: "Cuprum Title+Body",
        1: "Cuprum Title+Body+1Pic",
        2: "Cuprum Title+Body+2Pic",
        3: "Cuprum Title+Body+3Pic",
    }[n]
    layout = get_layout_by_name(prs, layout_name) or prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)

    # タイトル
    title_ph = _first_placeholder(slide, (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE))
    if title_ph is None:
        for ph in slide.placeholders:
            if "タイトル" in getattr(ph, "name", ""):
                title_ph = ph
                break
    if title_ph is not None:
        _set_text(title_ph, title, TITLE_SIZE_PT, TITLE_COLOR, TITLE_BOLD)

    # 本文
    body_ph = _first_placeholder(slide, (PP_PLACEHOLDER.BODY,))
    if body_ph is None:
        body_ph = _first_placeholder(slide, (PP_PLACEHOLDER.CONTENT,))
    if body_ph is None:
        for ph in slide.placeholders:
            if "テキスト" in getattr(ph, "name", ""):
                body_ph = ph
                break
    if body_ph is not None:
        _set_text(body_ph, summary, BODY_SIZE_PT, BODY_COLOR, BODY_BOLD)

    # 画像（プレースホルダー優先）
    pic_placeholders = []
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                pic_placeholders.append(ph)
        except Exception:
            continue
    pic_placeholders.sort(key=lambda sh: (sh.left, sh.top))

    for i, img in enumerate(images[:len(pic_placeholders)]):
        ph = pic_placeholders[i]
        buf = io.BytesIO()
        img.save(buf, format="JPEG")
        buf.seek(0)
        try:
            ph.insert_picture(buf)
        except Exception:
            slide.shapes.add_picture(buf, ph.left, ph.top, width=ph.width)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ========= UI =========
st.set_page_config(page_title="プレス/コピペ → Cuprum PPT", page_icon="🧩", layout="wide")
st.title(f"プレスURL or コピペ＋画像 → Cuprumテンプレ自動作成｜{APP_VERSION}")

with st.sidebar:
    st.header("設定")
    template_file = st.file_uploader("テンプレ（.pptx）を差し替え可", type=["pptx"])
    api_key = st.text_input("OpenAI API Key（未入力/失敗時はローカル要約）", type="password")
    max_images = st.slider("最大画像数（先頭から使用、上限3枚）", 0, 6, 3)
    summary_length = st.slider("要約文字数上限（目安）", 120, 400, 120, 20)
    st.caption("タイトル>25文字は短縮。本文は上限文字数で要約（コピペ版は要約オプション）。")

mode = st.radio("入力モード", ["URLモード", "コピペ＋画像アップロード"], horizontal=True)

# 共有の作業用変数
title_final = ""
summary_final = ""
images: List[Image.Image] = []
parsed_preview = None  # プレビュー用

# ============== モード1：URLモード ==============
if mode == "URLモード":
    url = st.text_input("プレスリリースのURL（社外サイト推奨）")
    if st.button("① 内容を抽出（URLから）"):
        try:
            parsed = parse_page(url)
            st.session_state["parsed_url"] = parsed
            st.success("抽出しました。下で要約・画像を確認してください。")
        except Exception as e:
            st.error(f"抽出に失敗しました: {type(e).__name__}: {e}")

    parsed = st.session_state.get("parsed_url")
    if parsed:
        st.subheader("抽出結果プレビュー（URLモード）")
        left, right = st.columns([2, 1])
        with left:
            st.write("抽出タイトル:", parsed.get("title") or "(なし)")
            raw_text = parsed.get("text") or ""
            st.write("本文（先頭プレビュー）:", raw_text[:300] + ("…" if len(raw_text) > 300 else ""))
        with right:
            st.write("候補画像URL（先頭から使用）")
            candidates = parsed.get("images", [])
            if candidates:
                for i, u in enumerate(candidates[:max_images]):
                    st.write(f"{i+1}. {u}")
            else:
                st.write("（画像候補なし）")

        # 要約 & タイトル調整
        client = get_client(api_key or None)
        title_final = gpt_shorten_title(client, parsed.get("title") or "（無題）")
        summary_final = gpt_summarize_body(client, parsed.get("text") or "", summary_length)

        # 画像DL
        sel_urls = parsed.get("images", [])[:max_images]
        images = download_images(sel_urls, limit=max_images)
        if images:
            cols = st.columns(min(len(images), 3))
            for i, img in enumerate(images):
                with cols[i % len(cols)]:
                    st.image(img, caption=f"Image {i+1}", use_container_width=True)
        else:
            st.info("表示可能な画像が見つかりませんでした。")

        parsed_preview = {"title": title_final, "summary": summary_final}

# ============== モード2：コピペ＋画像アップロード ==============
else:
    manual_title = st.text_input("記事タイトル（コピペ）")
    manual_body = st.text_area("記事本文（コピペ）", height=220)
    colA, colB = st.columns(2)
    with colA:
        do_summarize = st.checkbox("本文を要約する（上限はサイドバーの文字数）", value=True)
    with colB:
        do_shorten_title = st.checkbox("タイトルが25文字超なら短縮する", value=True)

    uploaded_files = st.file_uploader(
        "画像をアップロード（最大3枚まで）", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True
    )

    if st.button("① プレビュー生成（コピペ版）"):
        if not manual_title and not manual_body:
            st.warning("タイトルまたは本文のどちらかは入力してください。")
        else:
            # タイトル整形
            client = get_client(api_key or None)
            if do_shorten_title:
                title_final = gpt_shorten_title(client, manual_title or "（無題）")
            else:
                title_final = (manual_title or "（無題）")[:100]  # 暫定で100文字制限

            # 本文整形
            if do_summarize:
                summary_final = gpt_summarize_body(client, manual_body or "", summary_length)
            else:
                txt = manual_body or ""
                summary_final = txt[:summary_length]

            # 画像（アップロード → PIL）
            images = []
            if uploaded_files:
                for f in uploaded_files[:3]:
                    try:
                        img = Image.open(f).convert("RGB")
                        images.append(img)
                    except Exception:
                        continue

            st.session_state["manual_preview"] = {
                "title": title_final,
                "summary": summary_final,
                "images_len": len(images),
            }
            st.session_state["manual_images"] = images
            st.success("プレビューを作成しました。下で確認してください。")

    manual_prev = st.session_state.get("manual_preview")
    images = st.session_state.get("manual_images", [])
    if manual_prev:
        st.subheader("抽出結果プレビュー（コピペ版）")
        st.write("**タイトル（最終）**:", manual_prev["title"])
        st.write(f"**本文（{len(manual_prev['summary'])}文字）**:", manual_prev["summary"])
        if images:
            cols = st.columns(min(len(images), 3))
            for i, img in enumerate(images):
                with cols[i % len(cols)]:
                    st.image(img, caption=f"Uploaded {i+1}", use_container_width=True)
        else:
            st.info("画像は未アップロードです。")

        title_final = manual_prev["title"]
        summary_final = manual_prev["summary"]
        parsed_preview = manual_prev

# ============== 共通：PPT生成 ==============
st.markdown("---")
if st.button("② PPTを作成してダウンロード"):
    try:
        tpl_path = TEMPLATE_PATH
        if template_file is not None:
            tpl_path = "uploaded_template.pptx"
            with open(tpl_path, "wb") as f:
                f.write(template_file.read())

        import os
        if not os.path.exists(tpl_path):
            st.error(f"テンプレが見つかりません: {tpl_path}")
        elif not parsed_preview:
            st.error("先に①でプレビューを作成してください。")
        else:
            ppt_bytes = build_pptx(tpl_path, title_final or "（無題）", summary_final or "", images or [])
            if not isinstance(ppt_bytes, (bytes, bytearray)) or len(ppt_bytes) == 0:
                st.error("PPT生成に失敗しました（データ不正または空ファイル）。")
            else:
                st.success("PPTを生成しました。")
                st.download_button(
                    "ダウンロード",
                    data=ppt_bytes,
                    file_name="press_auto.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
    except Exception as e:
        import traceback
        st.error(f"PPT生成に失敗しました: {type(e).__name__}: {e}")
        st.code("".join(traceback.format_exc()))
else:
    st.caption("① 抽出/プレビュー → ② PPT作成 の順で操作してください。")
