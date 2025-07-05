import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

st.set_page_config(page_title="テキスト一括翻訳ツール", page_icon="🧾")

st.title("🧾 テキスト一括翻訳ツール（ChatGPT API対応）")

st.markdown("""
このアプリでは、**Excelファイルの1列目（A列）のテキスト**をChatGPT（GPT-4o）で一括翻訳します。  
翻訳プロンプトの「前提」「追加指示」はブラウザ上で自由に編集可能です。  
出力は元のデータ＋翻訳結果＋注釈付きのExcelファイルとなります。

**使い方：**

1. OpenAIのAPIキーを入力（取得：[https://platform.openai.com/account/api-keys](https://platform.openai.com/account/api-keys)）
2. Excelファイルをアップロード（翻訳対象はA列）
3. 「前提」「追加指示」を必要に応じて編集
4. 「翻訳を開始」を押す
5. 結果ファイルをダウンロード
""")

# APIキー入力欄
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

st.session_state.api_key = st.text_input(
    "🔑 OpenAI APIキーを入力",
    type="password",
    value=st.session_state.api_key,
)

# 翻訳プロンプト設定
default_context = "- 翻訳対象は製薬企業のGLデータです。\n- 費目名、案件名、摘要、サプライヤ名等がまとめて入っています"

fixed_instruction = """- 厳密に内容を日本語に翻訳してください。内容を要約せず、漏らさないようにお願いします。
- 専門用語・略語・ベンダ名の説明は注釈として加えてください。
- 略語は正式名称を付記してください。
- 出力の際は、後でエクセルに分割して貼り付けられるように、改行や順番を保ってください。
- 翻訳内容は必ず出力してください。
- 注釈内容はできるだけ出力してください。
- 外国語のまま出力されることがありますが、必ず日本語に翻訳してください。"""

st.subheader("📄 翻訳プロンプトのカスタマイズ")

context = st.text_area("【前提】", value=default_context, height=120)

st.markdown("【指示（固定）】")
st.code(fixed_instruction, language="markdown")

extra_instruction = st.text_area("【追加指示（任意）】", value="", height=100)

# ファイルアップロード
uploaded_file = st.file_uploader("📁 Excelファイルをアップロード（A列を翻訳対象とします）", type=["xlsx"])

# 条件チェック
if not st.session_state.api_key:
    st.warning("🔑 OpenAI APIキーを入力してください。")
elif not uploaded_file:
    st.warning("📄 Excelファイルをアップロードしてください。")

# API呼び出し関数
def call_openai_api(text, context, fixed_instruction, extra_instruction):
    prompt = f"""以下のテキストを翻訳してください：

原文: {text}

前提:
{context}

指示:
{fixed_instruction}
{extra_instruction}

出力形式:
翻訳結果: <翻訳内容>
注釈: <注釈>
"""

    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "あなたは優秀な翻訳専門家です。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )
        content = response.choices[0].message.content
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

# 実行ロジック
if st.session_state.api_key and uploaded_file:
    openai.api_key = st.session_state.api_key

    try:
        df = pd.read_excel(uploaded_file)
        first_col = df.iloc[:, 0].astype(str)
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

    st.success(f"{len(first_col)}件のテキストを翻訳します。")

    if st.button("🚀 翻訳を開始"):
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
                    executor.submit(call_openai_api, text, context, fixed_instruction, extra_instruction): idx
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

            st.success("✅ 翻訳完了！以下からダウンロードできます。")
            st.download_button(
                label="📥 翻訳済みExcelをダウンロード",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
