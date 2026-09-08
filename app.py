"""
問診票 OCR アプリ
================
スマホで撮影した手書き問診票をOCRしてWordファイルを作成し、
Google Driveの日付フォルダに保存するStreamlitアプリ。

OCRエンジン：Google Drive 組み込みOCR（完全無料）
"""

import streamlit as st
import json
import io
import time
from datetime import datetime

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from PIL import Image

# ─────────────────────────────────────────
# ページ設定（スマホ最適化）
# ─────────────────────────────────────────
st.set_page_config(
    page_title="問診票 OCR",
    page_icon="🏥",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  /* スマホ向けボタン大型化 */
  .stButton > button {
      height: 3.2rem;
      font-size: 1.1rem;
      font-weight: bold;
  }
  /* アップロードエリア */
  .stFileUploader label { font-size: 1rem; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# Google Drive サービス取得（キャッシュ）
# ─────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_drive_service():
    """サービスアカウントでDrive APIに接続する"""
    credentials_info = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    creds = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


# ─────────────────────────────────────────
# OCR処理
# ─────────────────────────────────────────
def compress_image(image_bytes: bytes, max_width: int = 1600) -> bytes:
    """
    画像をリサイズ・圧縮してアップロードを高速化する。
    長辺が max_width を超える場合のみリサイズ。
    """
    img = Image.open(io.BytesIO(image_bytes))
    # RGBA → RGB 変換（PNGなどで必要）
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    # リサイズ
    w, h = img.size
    if w > max_width or h > max_width:
        ratio = max_width / max(w, h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=80, optimize=True)
    buf.seek(0)
    return buf.getvalue()


def run_ocr(service, image_bytes: bytes) -> str:
    """
    画像をGoogle DriveにアップロードしてOCRテキストを取得する。
    Google Driveは画像→Google Doc変換時に自動OCRを行う（無料）。
    """
    # アップロード前に圧縮（Broken pipe 対策）
    image_bytes = compress_image(image_bytes)

    # 1. 画像を「Google Doc」として保存 = OCR発動
    # ※ parents を指定することでサービスアカウント個人DriveではなくShared Folderに
    #    一時ファイルを置く → 個人Driveのストレージ超過エラー(storageQuotaExceeded)を回避
    file_metadata = {
        "name": "_ocr_temp_monshin",
        "mimeType": "application/vnd.google-apps.document",
        "parents": [st.secrets["DRIVE_FOLDER_ID"]],
    }
    media = MediaIoBaseUpload(
        io.BytesIO(image_bytes),
        mimetype="image/jpeg",
        resumable=True,   # 大きいファイルでも安定するよう resumable に変更
        chunksize=1024 * 256,
    )

    request = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
    )
    # resumable アップロードの実行
    doc_file = None
    while doc_file is None:
        _, doc_file = request.next_chunk()
    doc_id = doc_file["id"]

    try:
        # OCR処理完了を待機
        time.sleep(4)

        # 2. プレーンテキストとしてエクスポート
        request = service.files().export_media(
            fileId=doc_id,
            mimeType="text/plain",
        )
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        ocr_text = buf.getvalue().decode("utf-8").strip()
        return ocr_text if ocr_text else "（OCRでテキストを取得できませんでした）"

    finally:
        # 3. 一時ファイルを削除（Driveを汚さない）
        try:
            service.files().delete(fileId=doc_id).execute()
        except Exception:
            pass


# ─────────────────────────────────────────
# Wordドキュメント生成
# ─────────────────────────────────────────
def build_word_doc(ocr_text: str, patient_no: str, image_bytes: bytes) -> bytes:
    """
    問診票の構造に沿った見やすいWordドキュメントを生成する。
    元画像も埋め込むのでスタッフが目視確認できる。
    """
    doc = Document()

    # ── ページ余白設定 ──
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    now = datetime.now()

    # ── タイトル ──
    title = doc.add_heading("問　診　票（OCR）", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── 受診情報テーブル ──
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    cells = tbl.rows[0].cells
    cells[0].text = f"受診日：{now.strftime('%Y年%m月%d日')}"
    cells[1].text = f"患者番号：{patient_no or '─'}"
    cells[2].text = f"処理時刻：{now.strftime('%H:%M')}"
    for cell in cells:
        for para in cell.paragraphs:
            para.runs[0].font.size = Pt(10)

    doc.add_paragraph()

    # ── OCR 読み取り結果 ──
    h2 = doc.add_heading("【 OCR 読み取り結果 】", level=2)
    h2.runs[0].font.size = Pt(12)

    # 問診票の固定セクション見出し（太字で表示）
    SECTION_KEYWORDS = [
        "どのような症状",
        "その症状はいつから",
        "現在治療中",
        "過去にかかった",
        "現在服用中",
        "アレルギー",
        "お酒は",
        "タバコは",
        "女性の方",
        "妊娠",
        "授乳",
    ]

    for line in ocr_text.splitlines():
        line = line.strip()
        if not line:
            continue
        is_header = any(kw in line for kw in SECTION_KEYWORDS)
        para = doc.add_paragraph()
        run = para.add_run(line)
        run.font.size = Pt(10.5)
        if is_header:
            run.bold = True
            run.font.size = Pt(11)

    doc.add_paragraph()

    # ── スタッフ確認欄 ──
    h2b = doc.add_heading("【 スタッフ確認欄 】", level=2)
    h2b.runs[0].font.size = Pt(12)

    doc.add_paragraph("確認者：＿＿＿＿＿＿＿　　確認日時：＿＿＿＿＿＿＿")
    doc.add_paragraph()
    doc.add_paragraph("補足・修正：")
    doc.add_paragraph("　" * 80)
    doc.add_paragraph()

    # ── 元画像（目視確認用） ──
    h2c = doc.add_heading("【 問診票 原本画像 】", level=2)
    h2c.runs[0].font.size = Pt(12)

    # 画像サイズを適切にリサイズ
    img = Image.open(io.BytesIO(image_bytes))
    img_buf = io.BytesIO()
    img.save(img_buf, format="JPEG", quality=85)
    img_buf.seek(0)
    doc.add_picture(img_buf, width=Cm(14))

    # ── バイト列として返す ──
    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


# ─────────────────────────────────────────
# Google Drive への保存
# ─────────────────────────────────────────
def save_to_drive(service, doc_bytes: bytes, filename: str) -> str:
    """
    日付フォルダ（YYYY-MM-DD）を作成してWordファイルを保存する。
    フォルダが既にあればそこに追記する。
    戻り値：ファイルの共有リンク
    """
    parent_id = st.secrets["DRIVE_FOLDER_ID"]
    today     = datetime.now().strftime("%Y-%m-%d")

    # 日付フォルダを検索
    q = (
        f"name='{today}' "
        f"and '{parent_id}' in parents "
        "and mimeType='application/vnd.google-apps.folder' "
        "and trashed=false"
    )
    results = service.files().list(q=q, fields="files(id)").execute()

    if results["files"]:
        date_folder_id = results["files"][0]["id"]
    else:
        # なければ作成
        folder_meta = {
            "name": today,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        folder = service.files().create(body=folder_meta, fields="id").execute()
        date_folder_id = folder["id"]

    # Wordファイルをアップロード
    file_meta = {
        "name": filename,
        "parents": [date_folder_id],
    }
    media = MediaIoBaseUpload(
        io.BytesIO(doc_bytes),
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        ),
        resumable=False,
    )
    uploaded = service.files().create(
        body=file_meta,
        media_body=media,
        fields="id,webViewLink",
    ).execute()

    return uploaded.get("webViewLink", "")


# ─────────────────────────────────────────
# メインUI
# ─────────────────────────────────────────
def main():
    st.title("🏥 問診票 OCR")
    st.caption("写真を撮影してアップロードするとWord化してDriveに保存します")

    st.divider()

    # 患者番号（任意）
    patient_no = st.text_input(
        "患者番号（任意）",
        placeholder="例：00123",
        help="空欄の場合は処理時刻で自動命名します",
    )

    # 画像アップロード
    uploaded = st.file_uploader(
        "📷 問診票の写真を選択",
        type=["jpg", "jpeg", "png"],
        help="スマホで撮影した写真、またはスキャン画像",
    )

    if uploaded:
        # プレビュー表示
        st.image(uploaded, caption="アップロードされた問診票", use_container_width=True)

        st.divider()

        if st.button("✅ OCR処理して Drive に保存", type="primary", use_container_width=True):

            image_bytes = uploaded.read()

            # 進捗表示
            progress = st.progress(0)
            status   = st.empty()

            try:
                # Step 1: Drive 接続
                status.info("🔗 Google Drive に接続中...")
                service = get_drive_service()
                progress.progress(10)

                # Step 2: OCR
                status.info("📖 テキスト読み取り中（10〜15秒かかります）...")
                ocr_text = run_ocr(service, image_bytes)
                progress.progress(55)

                # Step 3: Word生成
                status.info("📄 Word ファイルを作成中...")
                doc_bytes = build_word_doc(ocr_text, patient_no, image_bytes)
                progress.progress(75)

                # Step 4: Drive保存
                status.info("💾 Google Drive に保存中...")
                ts       = datetime.now().strftime("%H%M%S")
                filename = (
                    f"問診票_{patient_no}_{ts}.docx"
                    if patient_no
                    else f"問診票_{ts}.docx"
                )
                link = save_to_drive(service, doc_bytes, filename)
                progress.progress(100)

                # 完了
                status.empty()
                st.success(f"✅ 保存完了：`{filename}`")
                if link:
                    st.markdown(f"[📄 Drive で開く（確認・修正）]({link})")

                # OCR結果の確認表示
                with st.expander("📝 OCR 読み取りテキストを確認する"):
                    st.text_area(
                        "読み取り結果（修正はDriveのWordファイル上で行ってください）",
                        value=ocr_text,
                        height=300,
                    )

            except Exception as e:
                progress.empty()
                status.empty()
                st.error(f"エラーが発生しました：{e}")
                st.info("設定（secrets）を確認してください。詳細はREADMEを参照。")


if __name__ == "__main__":
    main()
