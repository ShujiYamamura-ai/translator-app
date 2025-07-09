import streamlit as st
import pandas as pd
import openai
import os
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed
import io
import time

# === コード更新日時（固定表示用）===
CODE_UPDATED_AT = "2025-07-09 23:40 JST"

# === タイトル表示 ===
st.set_page_config(page_title="GL翻訳支援", layout="wide")
st.title(f"🌐 多言語GLデータ翻訳支援（Web版｜更新: 20250709 2300")

# === レイアウト設定 ===
left_col, right_col = st.columns([1, 2])

# === ISO国コード読み込み ===
ISO_XLSX_PATH = os.path.join("data", "iso_country_codes.xlsx")

@st.cache_data
def load_country_iso_map(path):
    try:
        df = pd.read_excel(path)
        return {str(k).strip(): str(v).strip() for k, v in zip(df["国名"], df["ISOコード"])}
    except Exception as e:
        st.error(f"ISOコードファイルの読み込みに失敗しました: {e}")
        st.stop()

iso_map = load_country_iso_map(ISO_XLSX_PATH)

def normalize_country_code(name):
    if isinstance(name, str):
        return iso_map.get(name.strip(), "JP")
    return "JP"

# === 入力エリア ===
with left_col:
    st.header("🔐 入力ファイルとAPIキー")
    if "api_key" not in st.session_state:
        st.session_state.api_key = ""
    st.session_state.api_key = st.text_input("OpenAI APIキー", type="password", value=st.session_state.api_key)
    uploaded_file = st.file_uploader("Excelファイル（国名、サプライヤ名、費目、案件名、摘要）", type=["xlsx"])

# === 翻訳・検索設定エリア ===
with right_col:
    st.header("📝 翻訳ルール・Web検索設定")
    st.markdown("""
    ### 🔎 Web検索の実行方法

    Web検索は、GLデータ中のサプライヤや企業情報が曖昧な場合に、
    **事業内容やグループ関係などを補足的に取得する目的で利用**します。
      """)

    # Web検索モード選択
    web_search_mode = st.selectbox(
        "🔎 Web検索の実行方法",
        options=["不明な場合のみ実行", "すべての行に対して実行", "Web検索を使用しない"],
        index=0
    )

    # 対象企業名・業界名入力
    target_company = st.text_input("🏢 対象企業名（任意）", value="")
    target_industry = st.text_input("🏭 業界名（任意）", value="")

    # サプライヤ情報プロンプト（テンプレ入り）
    default_supplier_prompt = f"{target_company} との関係、所在地、事業概要、売上高、競合企業、企業グループ構成"
    supplier_prompt = st.text_input("📘 サプライヤ情報に含めたい項目", value=default_supplier_prompt)

    # 翻訳プロンプト設定
    default_context = f"""本データは{target_industry or "各種業界"}における会計・経理関連のGL（総勘定元帳）データであり、「国名」「サプライヤ名」「費目」「案件名」「摘要」から構成された構造化データです。"""
    default_instruction = """- 各項目の意味を正確に逐語訳すること（省略・意訳・要約は不可）。
- 不明な企業名がある場合は必要に応じてWeb検索を行い、注釈およびサプライヤ情報に記載すること。
- サプライヤ情報には次を含める：所在地、事業概要、売上高、競合企業、親会社やグループ関係。
- 注釈にすでに十分な情報がある場合、サプライヤ情報には「注釈に記載の通り」と記載してよい。
- 出力形式は「翻訳結果」「注釈」「サプライヤ情報」の3段構成、すべて日本語で記載。"""

    context = st.text_area("【前提】", value=default_context, height=150)
    instruction = st.text_area("【翻訳ルール】", value=default_instruction, height=250)

# === Web検索の目的 説明 ===
st.markdown("""
---
📌 **Web検索の目的**  
注釈に企業情報が不足している場合や、企業名が曖昧な場合に補足情報を取得するために Web検索を活用します。  
検索精度向上のため、**国名はできるだけ ISOコード（例：JP, CN, US）** で記載してください。  
※ 検索精度を重視しない場合は、Web検索を無効にしても構いません。
""")

# === Web検索条件関数 ===
def should_execute_web_search(note, mode):
    if mode == "Web検索を使用しない":
        return False
    elif mode == "すべての行に対して実行":
        return True
    elif mode == "不明な場合のみ実行":
        return ("不明" in note or "情報が見つかりません" in note or "補足情報なし" in note)
    return False

# === Web検索関数 ===
def search_web(supplier, country_name, prompt_hint, target_company, retries=2, delay=2):
    iso_code = normalize_country_code(country_name)
    query = build_supplier_search_query(supplier, target_company, prompt_hint)

    for attempt in range(retries + 1):
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
            if attempt < retries:
                time.sleep(delay)
            else:
                return f"Web検索失敗（{retries+1}回試行後）: {e}"
 
 # === 検索クエリ生成関数 ===   
def build_supplier_search_query(supplier_name, target_company, prompt_hint):
    return f"""
次の目的でWeb検索を行ってください：

1. {supplier_name} と {target_company} の企業間関係（親会社／グループ会社など）があるかを確認してください。
2. 関係がある場合は、その関係性を説明したうえで {supplier_name} の以下の情報を中心に調査してください。
3. 関係がない場合は、{target_company} については触れず、{supplier_name} のみに集中して調査を行ってください。

調査対象項目：{prompt_hint}

※ 対象企業に関する記述は最大2文以内とし、関係がない場合は一切記載しないでください。
"""

# === 翻訳関数 ===
def call_openai_api(text, context, instruction, supplier_name, country_name, prompt_hint, web_mode, target_company):
    prompt = f"""あなたはGLデータ（総勘定元帳）データに関するプロ翻訳者です。

この原文は「国名」「サプライヤ名」「費目」「案件名」「摘要」から構成された構造化データの1行です。

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

        if should_execute_web_search(note, web_mode):
            supplier_info = search_web(supplier_name, country_name, prompt_hint, target_company)
        else:
            supplier_info = "注釈に記載の通り"

        return translation, note, supplier_info
    except Exception as e:
        return "エラー", f"APIエラー: {e}", ""

# === サンプル翻訳 ===
with left_col:
    st.subheader("🔍 サンプル翻訳（構造化入力）")

    col1, col2, col3, col4, col5, col6 = st.columns(6)

    with col1:
        sample_country = st.text_input("国名", value="US", label_visibility="collapsed")
        st.caption("🌍 国名")
    with col2:
        sample_supplier = st.text_input("サプライヤ名", value="JWALK, LLC", label_visibility="collapsed")
        st.caption("🏢 サプライヤ名")
    with col3:
        sample_category = st.text_input("費目名", value="Consulting Fee", label_visibility="collapsed")
        st.caption("💼 費目名")
    with col4:
        sample_project = st.text_input("案件名", value="US Market Trend Research", label_visibility="collapsed")
        st.caption("📁 案件名")
    with col5:
        sample_summary = st.text_input("摘要", value="Local Consumer Behavior Analysis in NY", label_visibility="collapsed")
        st.caption("📝 摘要")
    with col6:
        sample_target_company = st.text_input("対象企業", value="Shiseido", label_visibility="collapsed")
        st.caption("🎯 対象企業(想定クライアント)")

    if st.button("サンプル翻訳を実行"):
        sample_text = f"{sample_country} / {sample_supplier} / {sample_category} / {sample_project} / {sample_summary}"
        with st.spinner("翻訳中..."):
            tr, note, info = call_openai_api(
                sample_text,
                context,
                instruction,
                supplier_name=sample_supplier,
                country_name=sample_country,
                prompt_hint=supplier_prompt,
                web_mode=web_search_mode,
                target_company=sample_target_company
            )
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
            st.error("⚠️ 入力ファイルに必要な列が不足しています。")
            st.stop()
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

    if len(df) > 100:
        st.error("⚠️ Web版では最大100件までです。")
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
                        prompt_hint=supplier_prompt,
                        web_mode=web_search_mode,
                        target_company=target_company
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

            st.success("✅ 翻訳完了！以下からダウンロードしてください。")
            st.download_button("📥 翻訳済みExcelをダウンロード", data=output, file_name=filename)
