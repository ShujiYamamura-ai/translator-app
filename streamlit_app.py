import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

st.set_page_config(page_title="GL翻訳ツール", page_icon="🧾")

st.title("🧾 GL翻訳ツール（OpenAI GPT-4o）")
st.markdown("製薬企業向けGLデータを翻訳します。以下にAPIキーとファイルを入力してください。")

# APIキー入力欄（常に表示）
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

st.session_state.api_key = st.text_input(
    "OpenAI APIキー（取得: https://platform.openai.com/account/api-keys）",
    type="password",
    value=st.session_state.api_key,
)

# ファイルアップロード欄（常に表示）
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=["xlsx"])

# 条件が揃ったら処理
if uploaded_file and st.session_state.api_key:
    openai.api_key = st.session_state.api_key

    try:
        data = pd.read_excel(uploaded_file, sheet_name='Sheet1')
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        st.stop()

    if not {'ID', 'target', 'merge'}.issubset(data.columns):
        st.error("必要な列（ID, target, merge）がExcelに存在しません。")
        st.stop()

    data = data[data['target'] == 1]
    data = data.drop_duplicates(subset='ID', keep='first')

    st.success(f"{len(data)}件の対象行を処理します。")

    def classify_project(row, index, total):
        prompt = f"""
以下の情報を基に翻訳してください：

翻訳対象: {row['merge']}

前提:
- 翻訳対象は製薬企業のGLデータです。
- 費目名、案件名、摘要、サプライヤ名等がまとめて入っています

指示:
- 上記の厳密に内容を日本語に翻訳してください。内容を要約せず、漏らさないようにお願いします。
- 専門用語・略語・ベンダ名の説明は注釈として加えてください。
- 略語は正式名称を付記してください
- 上記の内容は";"で区切られた複数の列をマージしたものです。
- 出力の際は、後でエクセルに分割して貼り付けられるように、";"で区切ったまま翻訳してください
- 翻訳内容は必ず出力してください
- 注釈内容はできるだけ出力してください
- たまに外国語のまま出力されることがありますが、絶対に日本語に翻訳してください

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
            result = response.choices[0].message.content
            classification, reason = None, None
            lines = result.splitlines()
            for line in lines:
                if "翻訳結果:" in line:
                    classification = line.split("翻訳結果:")[1].strip()
                elif "注釈:" in line:
                    reason = line.split("注釈:")[1].strip()
                    for next_line in lines[lines.index(line)+1:]:
                        if next_line.strip():
                            reason += f" {next_line.strip()}"
                        else:
                            break
            if classification is None or reason is None:
                classification, reason = "翻訳失敗", result
        except Exception as e:
            classification, reason = "エラー", f"失敗: {e}"
        return classification, reason

    def classify_all(data):
        results = {}
        total = len(data)
        progress_bar = st.progress(0)
        status_text = st.empty()

        def update_progress(i):
            percent = int((i + 1) / total * 100)
            progress_bar.progress(percent)
            status_text.text(f"{i + 1}/{total} 件処理中...")

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {
                executor.submit(classify_project, row, idx, total): idx
                for idx, row in data.iterrows()
            }
            for i, future in enumerate(as_completed(futures)):
                idx = futures[future]
                results[idx] = future.result()
                update_progress(i)

        progress_bar.empty()
        status_text.empty()
        return results

    if st.button("翻訳を開始"):
        with st.spinner("ChatGPTによる翻訳中...（数分かかる場合があります）"):
            results = classify_all(data)
            sorted_results = [results[idx] for idx in sorted(results.keys())]
            data["翻訳結果"], data["注釈"] = zip(*sorted_results)

            output = io.BytesIO()
            data.to_excel(output, index=False)
            output.seek(0)
            now_str = datetime.now().strftime('%Y%m%d_%H%M')
            st.success("翻訳完了！以下からダウンロードできます。")
            st.download_button(
                label="翻訳結果をダウンロード",
                data=output,
                file_name=f"翻訳結果_{now_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    # 入力待ちの案内を常に表示
    if not st.session_state.api_key:
        st.warning("🔑 OpenAI APIキーを入力してください。")
    if not uploaded_file:
        st.warning("📄 Excelファイルをアップロードしてください。")
