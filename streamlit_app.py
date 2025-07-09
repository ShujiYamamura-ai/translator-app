import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

# === UI初期設定 ===
st.set_page_config(page_title="GL翻訳支援アプリ", layout="wide")
st.title("🌐 多言語GLデータ翻訳支援（Web版v202507091200）")

is_web = True  # Webバージョン制限：最大100件

# === レイアウト設定 ===
left_col, right_col = st.columns([1, 2])

# === 入力エリア ===
with left_col:
    st.header("🔐 入力")

    if "api_key" not in st.session_state:
        st.session_state.api_key = ""

    st.session_state.api_key = st.text_input("OpenAI APIキー", type="password", value=st.session_state.api_key)
    uploaded_file = st.file_uploader("Excelファイル（国名, サプライヤ名, 費目, 案件名, 摘要）", type=["xlsx"])

# === 翻訳設定エリア ===
with right_col:
    st.header("📝 翻訳プロンプトの設定")

    search_enabled = st.checkbox("🔎 不明なサプライヤに対してWeb検索を有効にする", value=True)

    default_context = """本データは製薬業界のGL（総勘定元帳）データであり、1行ごとに「国名」「サプライヤ名」「費目」「案件名」「摘要」から構成される構造化データである。
各項目には略語、業界用語、ベンダ名、費目コード、内部プロジェクト名などが含まれている。"""

    default_instruction = """- 原文の意味・意図を正確に逐語訳すること（省略・要約・意訳は不可）。
- 不明な企業名や略語は必要な場合のみWeb検索で補足し、注釈に説明を加えること。
- 出力は「翻訳結果」「注釈」の2段構成で記載。
- 数字・単位は原文を保持しつつ意味が伝わるように記述する。"""

    context = st.text_area("【前提】", value=default_context, height=150)
    instruction = st.text_area("【翻訳ルール】", value=default_instruction, height=250)

# === Web検索関数 ===
def search_web(supplier, country):
    query = f"{supplier} とは？  国:{country}"
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-search-preview",
            web_search_options={
                "search_context_size": "medium",
                "user_location": {
                    "type": "approximate",
                    "approximate": {"country": country if len(country) == 2 else "JP"},
                },
            },
            messages=[{"role": "user", "content": query}],
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Web検索エラー: {e}"

# === 翻訳関数 ===
def call_openai_api(text, context, instruction, supplier_name, country, search_enabled=True):
    prompt = f"""あなたは製薬業界のGLデータに関するプロ翻訳者です。

この原文は「国名」「サプライヤ名」「費目」「案件名」「摘要」で構成されるGLデータの1行である。

【原文】
{text}

【前提】
{context}

【翻訳指示】
{instruction}

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
        translation, note, supplier_info = "翻訳失敗", "注釈取得失敗", ""

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

        if search_enabled and (
            "不明" in note or "情報が見つかりません" in note or "補足情報なし" in note
        ):
            supplier_info = search_web(supplier_name, country)
        return translation, note, supplier_info
    except Exception as e:
        return "エラー", f"APIエラー: {e}", ""

# === サンプル実行 ===
with left_col:
    st.subheader("🔍 サンプル翻訳")
    sample_input = st.text_input("例：日本 / GSK / 臨床試験費 / 肺がんP1 / SAP請求")
    if st.button("サンプル翻訳を実行"):
        with st.spinner("翻訳中..."):
            tr, note, sup = call_openai_api(
                sample_input, context, instruction,
                supplier_name="GSK", country="JP", search_enabled=search_enabled
            )
            st.success("✅ 完了")
            st.markdown(f"**翻訳結果：** {tr}")
            st.markdown(f"**注釈：** {note}")
            st.markdown(f"**サプライヤ情報：** {sup}")

# === メイン処理 ===
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
        with st.spinner("ChatGPTによる翻訳中..."):
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
                        futures[executor.submit(
                            call_openai_api,
                            full_text, context, instruction,
                            supplier_name=row["サプライヤ名"],
                            country=row["国名"],
                            search_enabled=search_enabled
                        )] = idx
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
                st.success("✅ 翻訳完了。以下からダウンロードしてください。")
                st.download_button(
                    label="📥 翻訳済みExcelをダウンロード",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
