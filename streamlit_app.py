import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

# === 設定 ===
st.set_page_config(page_title="多言語GLデータ内容解釈支援 - Web ver", layout="wide")
st.title("🌐 多言語GLデータ内容解釈支援（Web ver）")

# === モード判定 ===
is_web = True  # Web ver → 100件制限／ローカル版は False にすれば無制限

# === 説明 ===
st.markdown("""
**このアプリは Web バージョンです（Streamlit Cloud 上で動作）**  
- 処理できる件数は **最大100件まで** に制限されています  
- より大量データ（100件以上）を扱いたい場合は、**ローカルアプリ版をご利用ください（処理件数制限なし）**
""")

# === レイアウト ===
left_col, right_col = st.columns([1, 2])

# === 入力エリア（左カラム）===
with left_col:
    st.header("🔐 入力・操作")

    if "api_key" not in st.session_state:
        st.session_state.api_key = ""

    st.session_state.api_key = st.text_input(
        "OpenAI APIキー",
        type="password",
        value=st.session_state.api_key,
    )

    uploaded_file = st.file_uploader("Excelファイル（A列を翻訳）", type=["xlsx"])

# === プロンプト設定（右カラム）===
with right_col:
    st.header("📝 翻訳プロンプトの設定")

    default_context = """本データは、製薬業界における会計・経理関連のGL（総勘定元帳）データである。
各テキストには、費目名・プロジェクト名・業務摘要・請求項目・ベンダ（外部委託業者）名など、複数の情報が混在しており、文脈依存の要素が多い。
形式としては1つのセル内に複数情報が非構造的に記載されており、略語・記号・社内表記が含まれる可能性がある。"""

    default_instruction = """- 原文の意味・意図を正確に汲み取り、日本語に逐語的に翻訳すること。省略・要約・意訳は一切行わない。
- 不明な企業名やサービス名が含まれる場合、Web検索を行って補足情報を注釈に記載すること。
- 専門用語、略語、製品名、ベンダ名（企業名）については、訳語に加えて注釈を付記すること。
- 数字や日付、単位などは原文のフォーマットを維持しつつ、意味が伝わるように記述すること。
- 出力形式は「翻訳結果」「注釈」に分けて記載し、必要に応じて🔍Web補足情報も追加すること。"""

    context = st.text_area("【前提】", value=default_context, height=150)
    instruction = st.text_area("【翻訳指示】", value=default_instruction, height=300)

# === 翻訳関数）===
def call_openai_api(text, context, instruction):
    system_prompt = (
        "あなたは多言語のGL（総勘定元帳）テキストを翻訳し、企業・サービス・商品情報に基づいて補足注釈を付ける翻訳アシスタントです。"
        "不明な企業名やサービス名が含まれる場合は、Web検索を用いて関連性の高い企業やサービス情報を収集し、注釈の中で補足してください。"
        "検索対象とすべきキーワードを文中から自動的に抽出して構いません。"
    )

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": f"""以下のテキストを翻訳し、内容に関連する企業やサービスが不明な場合はWeb検索で補足してください。

原文:
{text}

【前提】
{context}

【指示】
{instruction}

【出力形式】
翻訳結果: <翻訳された日本語テキスト>
注釈: <訳語の補足・用語の背景、Webからの補足情報があれば「🔍 Web補足情報：...」として追記>
"""}
    ]

    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            tools=[{"type": "web-search"}],
            tool_choice="auto",
            temperature=0.3
        )

        content = response.choices[0].message.content or ""
        translation, note = "翻訳失敗", "取得できませんでした"
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

        return translation, note

    except Exception as e:
        return "エラー", f"APIエラー: {e}"

# === サンプル分析（左カラム）===
with left_col:
    st.subheader("🔍 サンプル実行（1件だけ試す）")
    sample_text = st.text_input("例：翻訳対象文をここに入力", value="SAP invoice for oncology P1 study; GSK")
    if st.button("サンプル翻訳を実行"):
        with st.spinner("翻訳中..."):
            sample_result, sample_note = call_openai_api(sample_text, context, instruction)
            st.success("✅ 翻訳完了")
            st.markdown(f"**翻訳結果：** {sample_result}")
            st.markdown(f"**注釈：** {sample_note}")

# === メイン処理（アップロードファイルがあれば）===
if st.session_state.api_key and uploaded_file:
    openai.api_key = st.session_state.api_key
    try:
        df = pd.read_excel(uploaded_file)
        first_col = df.iloc[:, 0].astype(str)
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

    st.success(f"{len(first_col)}件のテキストを翻訳します。")

    # Web版では件数制限
    if is_web and len(first_col) > 100:
        st.error("⚠️ このWebバージョンでは最大100件までしか処理できません。\nファイルを調整するか、ローカルアプリ版をご利用ください。")
        st.stop()

    if left_col.button("🚀 翻訳を開始"):
        with st.spinner("ChatGPTによる翻訳中..."):
            results = {}
            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(i):
                percent = int((i + 1) / len(first_col) * 100)
                progress_bar.progress(percent)
                status_text.text(f"{i + 1}/{len(first_col)} 件処理中...")

            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = {
                    executor.submit(call_openai_api, text, context, instruction): idx
                    for idx, text in enumerate(first_col)
                }
                for i, future in enumerate(as_completed(futures)):
                    idx = futures[future]
                    results[idx] = future.result()
                    update_progress(i)

            df["翻訳結果"], df["注釈"] = zip(*[results[i] for i in sorted(results)])

            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            filename = f"翻訳結果_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

            with left_col:
                st.success("✅ 翻訳が完了しました。以下からダウンロードしてください。")
                st.download_button(
                    label="📥 翻訳済みExcelをダウンロード",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
