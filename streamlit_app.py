import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

# === ISO国コード最小変換辞書（日本語＋英語 最低限）
COUNTRY_NAME_TO_ISO = {
    "日本": "JP", "Japan": "JP",
    "中国": "CN", "China": "CN",
    "韓国": "KR", "Korea": "KR", "South Korea": "KR",
    "アメリカ": "US", "USA": "US", "United States": "US"
}

def normalize_country_code(name):
    if isinstance(name, str):
        return COUNTRY_NAME_TO_ISO.get(name.strip(), "JP")
    return "JP"

# === Streamlit UI ===
st.set_page_config(page_title="GL翻訳支援", layout="wide")
st.title("🌐 多言語GLデータ翻訳支援（Web版）｜2025-07-09 13:45 JST")

left_col, right_col = st.columns([1, 2])

with left_col:
    st.header("🔐 入力")
    if "api_key" not in st.session_state:
        st.session_state.api_key = ""
    st.session_state.api_key = st.text_input("OpenAI APIキー", type="password", value=st.session_state.api_key)
    uploaded_file = st.file_uploader("Excelファイル（国名、サプライヤ名、費目、案件名、摘要）", type=["xlsx"])

with right_col:
    st.header("📝 翻訳設定")
    search_enabled = st.checkbox("🔎 不明な企業に対してWeb検索を行う", value=True)

    default_context = """本データは製薬業界のGL（総勘定元帳）データであり、「国名」「サプライヤ名」「費目」「案件名」「摘要」から構成されている構造化データです。"""
    default_instruction = """- 各項目の内容を正確に逐語訳してください。
- 不明な企業名がある場合、必要に応じてWeb検索し補足情報を注釈に加えてください。
- 出力形式は「翻訳結果」「注釈」に分け、すべて日本語で記載すること。"""

    context = st.text_area("【前提】", value=default_context, height=150)
    instruction = st.text_area("【翻訳ルール】", value=default_instruction, height=250)

# === Web検索関数 ===
def search_web(supplier, country_name):
    iso_code = normalize_country_code(country_name)
    query = f"{supplier} の会社概要、事業内容、拠点情報"
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-search-preview",
            web_search_options={
                "search_context_size": "medium",
                "user_location": {
                    "type": "approximate",
                    "approximate": {"country": iso_code},
                },
            },
            messages=[{"role": "user", "content": query}],
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Web検索エラー: {e}"

# === 翻訳関数 ===
def call_openai_api(text, context, instruction, supplier_name, country_name, search_enabled=True):
    prompt = f"""あなたは製薬業界のGLデータに関するプロ翻訳者です。

この原文は「国名」「サプライヤ名」「費目」「案件名」「摘要」で構成された構造化データの1行です。

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
                {"role": "system", "content": "あなたは丁寧な逐語翻訳を行う日本語専門のプロ翻訳者です。すべて日本語で出力してください。"},
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

        if search_enabled and ("不明" in note or "情報が見つかりません" in note or "補足情報なし" in note):
            supplier_info = search_web(supplier_name, country_name)
        return translation, note, supplier_info
    except Exception as e:
        return "エラー", f"APIエラー: {e}", ""

# === サンプル翻訳 ===
with left_col:
    st.subheader("🔍 サンプル実行")
    sample_text = st.text_input("例：日本 / GSK / 臨床試験費 / 肺がん研究 / 請求処理")
    if st.button("サンプル翻訳を実行"):
        with st.spinner("翻訳中..."):
            tr, note, info = call_openai_api(sample_text, context, instruction, "GSK", "日本", search_enabled)
            st.success("✅ 翻訳完了")
            st.markdown(f"**翻訳結果：** {tr}")
            st.markdown(f"**注釈：** {note}")
            st.markdown(f"**サプライヤ情報：** {info}")

# === 一括処理 ===
if st.session_state.api_key and uploaded_file:
    openai.api_key = st.session_state.api_key

    try:
        df = pd.read_excel(uploaded_file)
        required_cols = ["国名", "サプライヤ名", "費目", "案件名", "摘要"]
        if not all(col in df.columns for col in required_cols):
            st.error("⚠️ 必須列が不足：国名／サプライヤ名／費目／案件名／摘要")
            st.stop()
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

    if len(df) > 100:
        st.error("⚠️ Webバージョンは最大100件までに制限されています。")
        st.stop()

    if left_col.button("🚀 一括翻訳を開始"):
        with st.spinner("翻訳中..."):
            results = {}
            progress = st.progress(0)
            status = st.empty()

            def update_progress(i):
                pct = int((i + 1) / len(df) * 100)
                progress.progress(pct)
                status.text(f"{i + 1}/{len(df)} 件処理中...")

            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = {}
                for idx, row in df.iterrows():
                    text = f"{row['国名']} / {row['サプライヤ名']} / {row['費目']} / {row['案件名']} / {row['摘要']}"
                    futures[executor.submit(
                        call_openai_api,
                        text, context, instruction,
                        supplier_name=row["サプライヤ名"],
                        country_name=row["国名"],
                        search_enabled=search_enabled
                    )] = idx

                for i, future in enumerate(as_completed(futures)):
                    idx = futures[future]
                    results[idx] = future.result()
                    update_progress(i)

            df["翻訳結果"], df["注釈"], df["サプライヤ情報"] = zip(*[results[i] for i in sorted(results)])

            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            filename = f"翻訳結果_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

            st.success("✅ 翻訳完了！ダウンロード可能です。")
            st.download_button(
                label="📥 翻訳済みExcelをダウンロード",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
