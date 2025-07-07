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
    st.header("📝 翻訳プロンプト（固定形式）")

    st.markdown("""
翻訳は以下の厳密な前提とルールに従って実行されます。  
内容の変更が必要な場合は「上級者モード」をONにしてください。
""")

    use_custom_prompt = st.checkbox("🔧 上級者モード（翻訳ルールを自分で書く）", value=False)

    if use_custom_prompt:
        custom_prompt = st.text_area("プロンプト全文を編集", height=400)
    else:
        custom_prompt = """あなたは製薬業界の財務データに精通したプロフェッショナル翻訳者である。

以下の原文は、製薬企業の経理GLデータに含まれる情報であり、費目名・プロジェクト名・業務概要・ベンダ名（サプライヤ）・略語・記号などが混在した、文脈依存かつ非構造なテキストである。

【翻訳対象】
<原文テキストをここに挿入>

【翻訳ルール】
- 全ての情報を逐語的に翻訳し、省略・要約・意訳は禁止する
- ベンダ（会社名）はどのような会社か調査して注釈に記載する
- 略語・記号・記述の意味は正式名称と背景を注釈に加える（例："P1" = 第1相試験）
- セミコロンやカンマなどの区切りは保持し、元の構造を再現する
- 数字・日付・通貨表記も元の形式で残しつつ意味が分かるように訳す

【出力形式】
翻訳結果: <翻訳内容>
注釈: <用語の補足／ベンダ説明／略語解説などを詳細に記載>
"""

# 翻訳関数（改修後）
def call_openai_api(text, prompt_template):
    filled_prompt = prompt_template.replace("<原文テキストをここに挿入>", text)

    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "あなたはプロフェッショナルな逐語翻訳者です。"},
                {"role": "user", "content": filled_prompt}
            ],
            temperature=0
        )
        content = response.choices[0].message.content
        translation, note = "翻訳失敗", "注釈取得不可"
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
