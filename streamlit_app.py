import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

# === UI設定 ===
st.set_page_config(page_title="多言語GLデータ翻訳支援", layout="wide")
st.title("🌐 多言語GLデータ内容解釈支援（Web版）")

# === 処理制限 ===
is_web = True  # 100件制限

# === レイアウト ===
left_col, right_col = st.columns([1, 2])

# === 入力UI ===
with left_col:
    st.header("🔐 入力")

    if "api_key" not in st.session_state:
        st.session_state.api_key = ""

    st.session_state.api_key = st.text_input("OpenAI APIキー", type="password", value=st.session_state.api_key)
    uploaded_file = st.file_uploader("Excelファイル（5列構成）", type=["xlsx"])

# === 翻訳プロンプト設定 ===
with right_col:
    st.header("📝 翻訳プロンプトの設定")

    default_context = """本データは製薬業界のGL（総勘定元帳）データであり、1行ごとに「国名」「サプライヤ名」「費目」「案件名」「摘要」から構成される構造化データである。
各項目には略語、業界用語、ベンダ名、費目コード、内部プロジェクト名などが含まれている。"""

    default_instruction = """- 原文の意味・意図を正確に逐語訳すること（省略・要約・意訳はNG）。
- 不明な企業名や略語はWeb検索を行い、注釈に企業説明を加えること。
- 数字・単位は原文を保持するが意味が伝わるよう記述。
- 「翻訳結果」と「注釈」の2段構成で出力すること。"""

    context = st.text_area("【前提】", value=default_context, height=150)
    instruction = st.text_area("【翻訳ルール】", value=default_instruction, height=250)

# === Web検索関数 ===
def search_web(query):
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-search-preview",
            web_search_options={
                "search_context_size": "medium",
                "user_location": {"type": "approximate", "approximate": {"country": "JP"}},
            },
            messages=[{"role": "user", "content": query}],
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Web検索エラー: {e}"

# === 翻訳関数 ===
def call_openai_api(text, context, instruction, supplier_query=None):
    supplier_info = ""
    if supplier_query:
        supplier_info = search_web(supplier_query)

    prompt = f"""あなたは製薬業界のGLデータに関するプロフェッショナル翻訳者である。

以下の原文は、「国名」「サプライヤ名」「費目」「案件名」「摘要」から構成される構造化データの1行である。

【原文】
{text}

【前提】
{context}

【翻訳指示】
{instruction}

{"【Web補足情報（企業名）】\n" + supplier_info if supplier_info else ""}

【出力形式】
翻訳結果: <逐語訳された日本語テキスト>
注釈: <専門用語や略語、会社名、サービス名に関する補足情報・解説>
"""

    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "あなたは正確な逐語翻訳を行うプロ翻訳者です。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )
        content = response.choices[0].message.content
        translation, note = "翻訳失敗", "注釈取得失敗"
        lines = content.splitlines()
        for line in lines:
            if "翻訳結果:" in line:
                translation = line.split("翻訳結果:")[1].strip()
            elif "注釈:" in line:
                note = line.split("注釈:")[1].strip()
                for next_line in lines[lines.index(line)+1:]:
                    if next_line.strip():
                        note += f" {next_line.strip()}"
                    else:
                        break
        return translation, note, supplier_info
    except Exception as e:
        return "エラー", f"APIエラー: {e}", ""

# === サンプル翻訳 ===
with left_col:
    st.subheader("🔍 サンプル実行")
    sample_input = st.text_input("例：Japan / GSK / Oncology Study / P1 Trial / SAP invoice")
    if st.button("サンプル翻訳を実行"):
        with st.spinner("翻訳中..."):
            tr, note, sup = call_openai_api(
                sample_input, context, instruction, supplier_query="GSK 製薬会社 概要"
            )
            st.success("✅ 翻訳完了")
            st.markdown(f"**翻訳結果：** {tr}")
            st.markdown(f"**注釈：** {note}")
            st.markdown(f"**サプライヤ情報：** {sup}")

# === メイン翻訳処理 ===
if st.session_state.api_key and uploaded_file:
    openai.api_key = st.session_state.api_key

    try:
        df = pd.read_excel(uploaded_file)
        required_cols = ["国名", "サプライヤ名", "費目", "案件名", "摘要"]
        if not all(col in df.columns for col in required_cols):
            st.error("⚠️ 必須列：国名, サプライヤ名, 費目, 案件名, 摘要")
            st.stop()
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

    st.success(f"{len(df)}件のデータを処理します。")

    if is_web and len(df) > 100:
        st.error("⚠️ Webバージョンは最大100件までです。")
        st.stop()

    if left_col.button("🚀 一括翻訳実行"):
        with st.spinner("処理中..."):
            results = {}
            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(i):
                percent = int((i + 1) / len(df) * 100)
                progress_bar.progress(percent)
                status_text.text(f"{i + 1}/{len(df)} 件処理中...")

            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = {}
                for idx, row in df.iterrows():
                    try:
                        full_text = f"{row['国名']} / {row['サプライヤ名']} / {row['費目']} / {row['案件名']} / {row['摘要']}"
                        supplier_query = f"{row['サプライヤ名']} 会社情報  国:{row['国名']} 費目:{row['費目']}"
                        futures[executor.submit(call_openai_api, full_text, context, instruction, supplier_query)] = idx
                    except Exception as e:
                        futures[executor.submit(lambda: ("エラー", f"データ処理エラー: {e}", ""))] = idx

                for i, future in enumerate(as_completed(futures)):
                    idx = futures[future]
                    results[idx] = future.result()
                    update_progress(i)

            df["翻訳結果"], df["注釈"], df["サプライヤ情報"] = zip(*[results[i] for i in sorted(results)])

            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            filename = f"翻訳結果_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

            with left_col:
                st.success("✅ 処理完了。以下からExcelをダウンロード可能です。")
                st.download_button(
                    label="📥 翻訳済みExcelをダウンロード",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
