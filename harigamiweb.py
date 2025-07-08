import streamlit as st
import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from datetime import datetime
import warnings
import io # インメモリファイル操作のために追加
import zipfile # zipファイル操作のために追加
import re # ファイル名サニタイズのために追加

# openpyxlの警告を非表示にする
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 固定ファイル
# Streamlit環境では、テンプレートファイルはアプリと同じディレクトリに置くか、適切にパスを指定する必要があります。
WORD_TEMPLATE = "harigami.docx" 
# 一時的な生成ファイルとZIPファイルを保存するディレクトリ
OUTPUT_DIR = "output_docs" 

# プレースホルダー定義
PLACEHOLDERS = {
    "［10月　19日（水）］": "DATE",
    "［10:00］": "START_TIME", 
    "［11:00］": "END_TIME",
    "［物件名］": "NAME",
}

def replace_placeholders_preserve_format(paragraph, replacements):
    """書式を完全に保持する置換関数"""
    full_text = paragraph.text
    should_center = False

    for ph, key in PLACEHOLDERS.items():
        if ph in full_text:
            if key in ["DATE", "START_TIME", "END_TIME"]:
                should_center = True

            for run in paragraph.runs:
                if ph in run.text:
                    original_font_size = run.font.size
                    original_bold = run.font.bold
                    original_italic = run.font.italic
                    original_underline = run.font.underline
                    original_color = run.font.color
                    
                    run.text = run.text.replace(ph, replacements[key])
                    
                    if original_font_size:
                        run.font.size = original_font_size
                    if original_bold is not None:
                        run.font.bold = original_bold
                    if original_italic is not None:
                        run.font.italic = original_italic
                    if original_underline is not None:
                        run.font.underline = original_underline
                    if original_color:
                        run.font.color.rgb = original_color.rgb
                    break
            
            current_text = paragraph.text
            if ph in current_text:
                replace_text_across_runs(paragraph, ph, replacements[key])
    
    if should_center:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def replace_text_across_runs(paragraph, search_text, replace_text):
    """複数のrunにまたがるテキストを置換（書式保持）"""
    full_text = ''.join(run.text for run in paragraph.runs)
    
    if search_text in full_text:
        new_text = full_text.replace(search_text, replace_text)
        
        if paragraph.runs:
            first_run = paragraph.runs[0]
            for run in paragraph.runs[1:]:
                run.text = ""
            first_run.text = new_text

def replace_placeholders_in_tables(doc, replacements):
    """テーブル内のプレースホルダーも置換"""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_placeholders_preserve_format(paragraph, replacements)

def replace_placeholders_comprehensive(doc, replacements):
    """文書全体の包括的な置換処理"""
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():
            replace_placeholders_preserve_format(para, replacements)
    
    replace_placeholders_in_tables(doc, replacements)
    
    for section in doc.sections:
        if section.header:
            for para in section.header.paragraphs:
                replace_placeholders_preserve_format(para, replacements)
        
        if section.footer:
            for para in section.footer.paragraphs:
                replace_placeholders_preserve_format(para, replacements)

def process_excel_and_generate_docs(excel_file_buffer):
    """
    Excelファイルを処理し、Word文書を生成するメインロジック。
    Streamlitのファイルアップローダーから受け取ったバッファを処理します。
    """
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    generated_file_paths = [] 

    try:
        df = pd.read_excel(excel_file_buffer, sheet_name="作業指示書 の一覧", engine='openpyxl')
        total_rows = len(df)
        st.info(f"📊 {total_rows}件のデータを読み込みました。文書生成を開始します...")
        
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        processed_count = 0
        for index, row in df.iterrows():
            progress_percent = (index + 1) / total_rows
            progress_bar.progress(progress_percent)
            progress_text.text(f"処理中: {index + 1} / {total_rows} 件完了")

            try:
                if pd.isna(row["物件名"]) or pd.isna(row["予定開始"]) or pd.isna(row["予定終了"]):
                    continue
                
                name = str(row["物件名"]).strip()
                
                start_dt = pd.to_datetime(row["予定開始"], errors='coerce')
                end_dt = pd.to_datetime(row["予定終了"], errors='coerce')
                
                if pd.isna(start_dt) or pd.isna(end_dt):
                    continue
                
                weekdays = {
                    'Mon': '月', 'Tue': '火', 'Wed': '水', 'Thu': '木', 
                    'Fri': '金', 'Sat': '土', 'Sun': '日'
                }
                weekday = weekdays[start_dt.strftime('%a')]
                
                date_str = f"{start_dt.month}月{start_dt.day}日（{weekday}）"
                start_str = start_dt.strftime("%H:%M")
                end_str = end_dt.strftime("%H:%M")
                
            except Exception as e:
                continue
            
            replacements = {
                "DATE": date_str,
                "START_TIME": start_str,
                "END_TIME": end_str,
                "NAME": name
            }
            
            try:
                if not os.path.exists(WORD_TEMPLATE):
                    st.error(f"❌ テンプレートファイルが見つかりません: {WORD_TEMPLATE}")
                    return [] 
                    
                doc = Document(WORD_TEMPLATE)
                
                replace_placeholders_comprehensive(doc, replacements)
                
                # --- ここからファイル名サニタイズの強化 ---
                safe_name = re.sub(r'[^\w\.\-]', '_', name)
                safe_name = re.sub(r'_{2,}', '_', safe_name)
                safe_name = safe_name.strip('_')
                if not safe_name:
                    safe_name = "untitled_document"
                # --- ファイル名サニタイズの強化ここまで ---

                output_file_name = f"{safe_name}.docx"
                output_path = os.path.join(OUTPUT_DIR, output_file_name)
                
                doc.save(output_path)
                
                # --- ファイルが実際に保存されたかを確認 ---
                if os.path.exists(output_path):
                    generated_file_paths.append(output_path)
                    processed_count += 1
                else:
                    # ファイルが保存されなかった場合は警告
                    st.warning(f"⚠️ 行 {index + 1}: Word文書 '{output_file_name}' の保存に失敗しました。このファイルはZIPに含まれません。")

            except Exception as e:
                # ここでWord文書作成時のエラーをキャッチし、ログに記録（Streamlit UIには表示しないが内部で把握）
                # st.error(f"❌ 行 {index + 1}: Word文書作成エラー - {str(e)}")
                continue
        
        progress_bar.progress(1.0)
        progress_text.text(f"処理完了: {processed_count} / {total_rows} 件完了")

        st.success(f"\n🎉 {processed_count}件の通知文書の生成が完了しました！")
        st.info(f"💡 アプリが稼働しているサーバー上では、一時的に `{OUTPUT_DIR}` フォルダにファイルが保存されています。")
        return generated_file_paths
        
    except FileNotFoundError:
        st.error(f"❌ ファイルが見つかりません。")
        return []
    except Exception as e:
        st.error(f"❌ エラーが発生しました: {str(e)}")
        return []

# --- Streamlit UI部分 ---
st.set_page_config(
    page_title="貼紙自動生成アプリ",
    page_icon="📄",
    layout="centered"
)

st.title("📄貼紙自動生成アプリ")
st.markdown("""
このアプリは、アップロードされたExcelファイル(作業指示書一覧）の情報に基づいて、
貼紙のWord文書を生成します。
""")

st.warning(f"**テンプレートの配置**: Wordテンプレートファイル (`{WORD_TEMPLATE}`) は、このスクリプトと同じディレクトリに配置してください。")
st.info(f"**日付と時刻のセンタリング**: Wordテンプレート (`{WORD_TEMPLATE}`) 内で該当する段落をあらかじめ「中央揃え」に設定してください。Pythonコードは、プレースホルダーを含む段落全体をセンタリングします。")
st.markdown("""
**保存先について**:
生成されたWord文書は、まとめてZIPファイルとして提供されます。
ダウンロードする際の保存先は、**お使いのPCのダウンロード設定** に従います（ブラウザのダウンロードダイアログで指定できます）。
アプリケーションが動作しているサーバー上で直接保存場所を選択することはできません。
""")

uploaded_file = st.file_uploader(
    "1. Excelファイルを選択してください (.xlsx, .xls)",
    type=["xlsx", "xls"],
    help="「作業指示書 の一覧」シートが含まれるExcelファイルをアップロードしてください。"
)

if uploaded_file is not None:
    st.success("Excelファイルが正常にアップロードされました！")
    
    if st.button("2. Word文書を生成する"):
        if not os.path.exists(WORD_TEMPLATE):
            st.error(f"エラー: Wordテンプレートファイルが見つかりません。`{WORD_TEMPLATE}` をこのアプリと同じディレクトリに置いてください。")
        else:
            with st.spinner("Word文書を生成中...しばらくお待ちください。"):
                excel_buffer = io.BytesIO(uploaded_file.read())
                generated_doc_paths = process_excel_and_generate_docs(excel_buffer)
            
            if generated_doc_paths:
                st.subheader("🎉 生成された文書をまとめてダウンロード")
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for doc_path in generated_doc_paths:
                        try:
                            # ZIPファイルに追加する際のファイル名を調整 (フォルダパスは含めない)
                            zf.write(doc_path, os.path.basename(doc_path))
                        except Exception as e:
                            st.warning(f"ZIPへの追加に失敗: {os.path.basename(doc_path)} - {e}")

                zip_buffer.seek(0)

                st.download_button(
                    label="全てのWord文書をZIPでダウンロード",
                    data=zip_buffer.getvalue(),
                    file_name="generated_word_documents.zip",
                    mime="application/zip"
                )
                
                for doc_path in generated_doc_paths:
                    try:
                        # ファイル削除前に再度存在チェック
                        if os.path.exists(doc_path):
                            os.remove(doc_path)
                    except Exception as e:
                        st.warning(f"一時ファイルの削除に失敗: {doc_path} - {e}")
                
            else:
                st.warning("生成された文書はありませんでした。Excelデータとエラーメッセージを確認してください。")

else:
    st.info("⬆️ 上の「参照」ボタンをクリックして、Excelファイルをアップロードしてください。")

st.markdown("---")
st.caption("Powered by Streamlit")
