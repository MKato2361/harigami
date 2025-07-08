import streamlit as st
import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from datetime import datetime
import warnings
import io # インメモリファイル操作のために追加

# openpyxlの警告を非表示にする
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 固定ファイル
# Streamlit環境では、テンプレートファイルはアプリと同じディレクトリに置くか、適切にパスを指定する必要があります。
WORD_TEMPLATE = "harigami.docx" 
OUTPUT_DIR = "output_docs" # 生成されたWordファイルを一時的に保存するディレクトリ

# プレースホルダー定義
PLACEHOLDERS = {
    "［10月　19日（水）］": "DATE",
    "［10:00］": "START_TIME", 
    "［11:00］": "END_TIME",
    "［物件名］": "NAME",
}

def replace_placeholders_preserve_format(paragraph, replacements):
    """書式を完全に保持する置換関数"""
    # 段落全体のテキストを確認
    full_text = paragraph.text
    
    # センタリングを適用するかどうかのフラグ
    should_center = False

    for ph, key in PLACEHOLDERS.items():
        if ph in full_text:
            # st.write(f"  🔄 置換中: '{ph}' → '{replacements[key]}'") # Streamlitではst.writeでデバッグ出力

            # 日付や時刻のプレースホルダーが含まれているかチェック
            if key in ["DATE", "START_TIME", "END_TIME"]:
                should_center = True

            # 各runを個別に処理
            for run in paragraph.runs:
                if ph in run.text:
                    # 書式情報を保存
                    original_font_size = run.font.size
                    original_bold = run.font.bold
                    original_italic = run.font.italic
                    original_underline = run.font.underline
                    original_color = run.font.color
                    
                    # テキストを置換
                    run.text = run.text.replace(ph, replacements[key])
                    
                    # 書式を復元
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
                    
                    # st.write(f"    ✅ 書式保持完了") # Streamlitではst.writeでデバッグ出力
                    break
            
            # 複数runにまたがる場合の処理
            current_text = paragraph.text
            if ph in current_text:
                replace_text_across_runs(paragraph, ph, replacements[key])
    
    # 日付または時刻のプレースホルダーが見つかった場合、段落をセンタリング
    if should_center:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # st.write(f"    ⭐ 段落をセンタリングしました。") # Streamlitではst.writeでデバッグ出力


def replace_text_across_runs(paragraph, search_text, replace_text):
    """複数のrunにまたがるテキストを置換（書式保持）"""
    full_text = ''.join(run.text for run in paragraph.runs)
    
    if search_text in full_text:
        new_text = full_text.replace(search_text, replace_text)
        
        if paragraph.runs:
            first_run = paragraph.runs[0]
            # original_font = first_run.font # 書式保持は上でやっているので不要であればコメントアウト
            
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
    # st.write("  📝 段落の置換処理開始...") # Streamlitではst.writeでデバッグ出力
    
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():
            # st.write(f"    段落 {i+1}: '{para.text[:50]}...'") # Streamlitではst.writeでデバッグ出力
            replace_placeholders_preserve_format(para, replacements)
    
    # st.write("  📋 テーブルの置換処理開始...") # Streamlitではst.writeでデバッグ出力
    replace_placeholders_in_tables(doc, replacements)
    
    # st.write("  📄 ヘッダー・フッターの置換処理開始...") # Streamlitではst.writeでデバッグ出力
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
    
    # 出力フォルダ作成（Streamlitの場合、コンテナ内の一時ディレクトリでも良い）
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    generated_files = [] # 生成されたファイルのパスを保持するリスト

    try:
        # Excel読込（ファイルバッファから直接読み込む）
        df = pd.read_excel(excel_file_buffer, sheet_name="作業指示書 の一覧", engine='openpyxl')
        total_rows = len(df)
        st.info(f"📊 {total_rows}件のデータを読み込みました。文書生成を開始します...")
        
        # プログレスバーのプレースホルダーを作成
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        processed_count = 0
        for index, row in df.iterrows():
            # プログレスバーを更新
            progress_percent = (index + 1) / total_rows
            progress_bar.progress(progress_percent)
            progress_text.text(f"処理中: {index + 1} / {total_rows} 件完了")

            st.markdown(f"---")
            st.subheader(f"✨ 処理中: 行 {index + 1}")
            try:
                if pd.isna(row["物件名"]) or pd.isna(row["予定開始"]) or pd.isna(row["予定終了"]):
                    st.warning(f"⚠️  行 {index + 1}: 必要なデータが不足しています（スキップ）")
                    continue
                
                name = str(row["物件名"]).strip()
                
                st.write(f"🔍 物件名: {name}")
                st.write(f"   予定開始: {row['予定開始']} (型: {type(row['予定開始'])})")
                st.write(f"   予定終了: {row['予定終了']} (型: {type(row['予定終了'])})")
                
                start_dt = pd.to_datetime(row["予定開始"], errors='coerce')
                end_dt = pd.to_datetime(row["予定終了"], errors='coerce')
                
                if pd.isna(start_dt) or pd.isna(end_dt):
                    st.error(f"❌ 行 {index + 1}: 日付・時間の変換に失敗しました（スキップ）")
                    continue
                
                weekdays = {
                    'Mon': '月', 'Tue': '火', 'Wed': '水', 'Thu': '木', 
                    'Fri': '金', 'Sat': '土', 'Sun': '日'
                }
                weekday = weekdays[start_dt.strftime('%a')]
                
                date_str = f"{start_dt.month}月{start_dt.day}日（{weekday}）"
                start_str = start_dt.strftime("%H:%M")
                end_str = end_dt.strftime("%H:%M")
                
                st.write(f"   変換結果: {date_str} {start_str}-{end_str}")
                
            except Exception as e:
                st.error(f"❌ 行 {index + 1}: データ処理エラー - {str(e)}（スキップ）")
                continue
            
            replacements = {
                "DATE": date_str,
                "START_TIME": start_str,
                "END_TIME": end_str,
                "NAME": name
            }
            
            try:
                # Word文書作成
                # テンプレートファイルの存在確認をここで行う
                if not os.path.exists(WORD_TEMPLATE):
                    st.error(f"❌ テンプレートファイルが見つかりません: {WORD_TEMPLATE}")
                    return [] # テンプレートが見つからない場合は処理を中断
                    
                doc = Document(WORD_TEMPLATE)
                
                st.write(f"  📄 Word文書の置換処理開始...")
                replace_placeholders_comprehensive(doc, replacements)
                
                safe_name = name.replace("/", "_").replace("\\", "_").replace(":", "_").replace(" ", "_")
                output_file_name = f"{safe_name}.docx"
                output_path = os.path.join(OUTPUT_DIR, output_file_name)
                
                doc.save(output_path)
                generated_files.append(output_path)
                processed_count += 1
                st.success(f"✅ 生成完了: {output_file_name}")
                
            except Exception as e:
                st.error(f"❌ 行 {index + 1}: Word文書作成エラー - {str(e)}")
                continue
        
        # 最終的なプログレスバーの状態
        progress_bar.progress(1.0)
        progress_text.text(f"処理完了: {processed_count} / {total_rows} 件完了")

        st.success(f"\n🎉 {processed_count}件の通知文書の生成が完了しました！")
        st.write(f"📚 生成されたWord文書は、以下のダウンロードリンクから取得できます。")
        st.info(f"💡 アプリが稼働しているサーバー上では、一時的に `{OUTPUT_DIR}` フォルダにファイルが保存されています。")
        return generated_files
        
    except FileNotFoundError:
        st.error(f"❌ ファイルが見つかりません。")
        return []
    except Exception as e:
        st.error(f"❌ エラーが発生しました: {str(e)}")
        return []

# --- Streamlit UI部分 ---
st.set_page_config(
    page_title="Word文書自動生成アプリ",
    page_icon="📄",
    layout="centered"
)

st.title("📄 ExcelデータからWord文書を自動生成")
st.markdown("""
このアプリは、アップロードされたExcelファイルの情報に基づいて、
指定されたWordテンプレートのプレースホルダーを自動的に置き換え、
新しいWord文書を生成します。
""")

st.warning(f"**テンプレートの配置**: Wordテンプレートファイル (`{WORD_TEMPLATE}`) は、このスクリプトと同じディレクトリに配置してください。")
st.info(f"**日付と時刻のセンタリング**: Wordテンプレート (`{WORD_TEMPLATE}`) 内で該当する段落をあらかじめ「中央揃え」に設定してください。Pythonコードは、プレースホルダーを含む段落全体をセンタリングします。")
st.markdown("""
**保存先について**:
このWebアプリケーションでは、生成されたWord文書は、お使いのブラウザを通じてダウンロードしていただきます。
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
            with st.empty(): # プログレスバーとテキストのために空のコンテナを用意
                generated_doc_paths = process_excel_and_generate_docs(io.BytesIO(uploaded_file.read()))
            
            if generated_doc_paths:
                st.subheader("🎉 生成された文書をダウンロード")
                for doc_path in generated_doc_paths:
                    try:
                        with open(doc_path, "rb") as f:
                            st.download_button(
                                label=f"ダウンロード: {os.path.basename(doc_path)}",
                                data=f.read(),
                                file_name=os.path.basename(doc_path),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        # ダウンロードリンク提供後、ファイルを削除する（必要に応じて）
                        # os.remove(doc_path)
                    except Exception as e:
                        st.error(f"ダウンロードリンクの作成に失敗しました: {os.path.basename(doc_path)} - {e}")
            else:
                st.warning("生成された文書はありませんでした。Excelデータとエラーメッセージを確認してください。")

else:
    st.info("⬆️ 上の「参照」ボタンをクリックして、Excelファイルをアップロードしてください。")

st.markdown("---")
st.caption("Powered by Streamlit")
