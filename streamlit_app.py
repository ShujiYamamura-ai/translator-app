import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

st.set_page_config(page_title="テキスト一括翻訳ツール", layout="wide")

st.title("🧾 テキスト一括翻訳ツール（ChatGPT API対応）")

# 説明文は横幅を活かして表示
st.markdown("""
このアプリでは、**ExcelファイルのA列（1列目）のテキスト**をChatGPT（GPT-4o）で一括翻訳します。  
翻訳プロンプトの「前提」「追加指示」は自由に編集可能です。出力は元データ＋翻訳＋注釈のExcelファイルです。
""")

# 横並びレイアウト：左 = 入力操作、右 = プロンプト設定
left_col, right_col = st.columns([1, 2])

# ----------------------------
# 左カラム：APIキー、ファイル、翻訳ボタン
# ----------------------------
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

    if not st.session_state.api_key:
        st.warning("APIキーを入力してください。")
    elif not uploaded_file:
        st.warning("Excelファイルをアップロードしてください。")

# ----------------------------
# 右カラム：翻訳プロンプト設定
# ----------------------------
with right_col:
    st.header("📝 プロンプト設定")

    default_context = "- 翻訳対象は製薬企業のGLデータです。\n- 費目名、案件名、摘要、サプライヤ名等がまとめて入っています"

    fixed_instruction = """- 厳密に内容を日本語に翻訳してください。内容を要約せず、漏らさないようにお願いします。
- 専門用語・略語・ベンダ名の説明は注釈として加えてください。
- 略語は正式名称を付記してください。
- 出力の際は、後でエクセルに分割して貼り付けられるように、改行や順番を保ってください。
- 翻訳内容は必ず出力してください。
- 注釈内容はできるだけ出力してください。
- 外国語のまま出力されることがありますが、必ず日本語に翻訳してください。"""

    context = st.text_area("【前提】", value=default_context, height=150)
    st.markdown("【指示（固定）】")
    st.code(fixed_instruction, language="markdown")
    extra_instruction = st.text_area("【追加指示（任意）】", value="", height=120)

# ----------------------------
# API処理・実行
# ----------------------------
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

# ----------------------------
# 実行トリガー（翻訳処理）
# ----------------------------
if st.session_state.api_key and uploaded_file:
    openai.api_key = st.session_state.api_key

    try:
        df = pd.read_excel(uploaded_file)
        first_col = df.iloc[:, 0].astype(str)
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

    st.success(f"{len(first_col)}件のテキストを翻訳します。")

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
