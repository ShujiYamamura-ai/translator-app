# GL翻訳ツール（Streamlit版）

製薬企業向けのGLデータを、ChatGPT（GPT-4o）で正確に翻訳するWebアプリです。

## 使い方

1. OpenAIのAPIキーを[こちら](https://platform.openai.com/account/api-keys)から取得
2. このアプリのURLにアクセス（例: https://your-app.streamlit.app）
3. APIキーを入力
4. Excelファイル（merge列・ID列・target列を含む）をアップロード
5. 「翻訳を開始」をクリック
6. 完了後、翻訳済みExcelをダウンロード

## 必要な列

- `ID`: 各行の一意識別子
- `target`: 1で処理対象
- `merge`: 翻訳対象のテキスト列（`;`で複数列を結合した文字列）

## 注意点

- 並列処理で高速に動作しますが、件数が多い場合は数分かかります。
- OpenAI APIキーは各自でご用意ください。
