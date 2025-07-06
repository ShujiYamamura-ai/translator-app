import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import io

st.set_page_config(page_title="多言語GLデータ内容解釈支援", layout="wide")

st.title("🧾 多言語GLデータ内容解釈支援（ChatGPT API対応）")

st.markdown("""
このアプリでは、**ExcelファイルのA列（1列目）のテキスト**をChatGPT（GPT-4o）で一括翻訳します。  
翻訳プロンプトの「前提」「翻訳指示」は自由に編集可能です。  
出力は元データ＋翻訳結果＋注釈のExcelファイルとなります。
""")

# レイアウト：左 = 操作、右 = プロンプト
left_col, right_col = st.columns([1, 2])

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

with right_col:
    st.header("📝 翻訳プロンプトの設定")

    default_context = """本データは、製薬業界における会計・経理関連のGL（総勘定元帳）データである。
各テキストには、費目名・プロジェクト名・業務摘要・請求項目・ベンダ（外部委託業者）名など、複数の情報が混在しており、文脈依存の要素が多い。
形式としては1つのセル内に複数情報が非構造的に記載されており、略語・記号・社内表記が含まれる可能性がある。"""

    default_instruction = """- 原文の意味・意図を正確に汲み取り、日本語に逐語的に翻訳すること。省略・要約・意訳は一切行わない。
- 原文内の文法ミスや略記がある場合も、意味を正確に汲み取って正しい日本語に置き換えること。
- 専門用語、略語、製品名、ベンダ名（企業名）については、訳語に加えて注釈を付記すること。
    - 例：GSK → GSK（グラクソ・スミスクライン、英国の製薬会社）
    - 例：IQVIA → IQVIA（医療データ解析およびCRO事業を展開するグローバル企業）
- 英語以外（例：ドイツ語、フランス語など）の語句が含まれる場合、すべての単語について注釈を付与すること。単語単位で区切って解釈する。
- 略語は正式名称とセットで訳出すること（例：SAP → SAP（Systems, Applications and Products））。
- 数字や日付、単位などは原文のフォーマットを維持した上で、意味を正確に反映した訳語を記述する。
- 出力形式は「翻訳結果」「注釈」に明確に分けること。読みやすいように各セクションは改行で区切ること。
- 翻訳内容はExcelで後から貼り付け・加工できるよう、文の順序や改行は維持し、余計な記号や括弧は追加しない。
- 外国語・記号・略称が混在する場合でも、すべて日本語に翻訳・説明を付けること。未翻訳は不可。"""

    context = st.text_area("【前提】", value=default_context, height=150)
    instruction = st.text_area("【翻訳指示】", value=default_instruction, height=300)

# 翻訳API呼び出し
def call_openai_api(text, context, instruction):
    prompt = f"""以下のテキストを翻訳してください：

原文:
{text}

【前提】
{context}

【指示】
{instruction}

【出力形式】
翻訳結果: <翻訳された日本語テキスト>
注釈: <訳語の補足・用語の背景など>
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

# 実行トリガー
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

            st.success("✅ 翻訳完了！以下からダウンロードできます。")
            st.download_button(
                label="📥 翻訳済みExcelをダウンロード",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
