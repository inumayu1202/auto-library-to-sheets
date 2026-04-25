# Kita Library Scraper (北区立図書館データ自動取得スクリプト)

北区立図書館のWebサイトから「新着案内（一般書）」および「利用状況（読書記録）」のデータを自動で取得し、指定したGoogleスプレッドシートへ書き込むPythonスクリプトです。

## 機能 (Features)

- Playwrightを使用したヘッドレスブラウザによる自動ログインと画面操作
- 「新着案内」の一般書（前日分）のデータを取得し、スプレッドシートに出力
- 「利用状況一覧」から読書記録の全ページデータを取得し、別のスプレッドシートに出力

## 動作環境 (Prerequisites)

- Python 3.8 以上
- Google Cloud Platform (GCP) サービスアカウント（Google Sheets API, Google Drive API へのアクセス権限）

## セットアップ (Setup)

1. **リポジトリのクローン**
   ```bash
   git clone https://github.com/your-username/your-repo-name.git
   cd your-repo-name
   ```

2. **必要なライブラリのインストール**
   ```bash
   pip install -r requirements.txt
   ```
   Playwrightのブラウザをインストールする必要があります：
   ```bash
   playwright install chromium
   ```

3. **環境変数の設定**
   `.env.example` をコピーして `.env` ファイルを作成し、ご自身の環境に合わせて設定を入力してください。
   ```bash
   cp .env.example .env
   ```
   *注意: `.env` ファイルにはIDやパスワードが含まれるため、絶対にGitHub等に公開しないでください。*

4. **Googleサービスアカウントの認証情報**
   Google Cloud Console から取得したJSON形式の認証鍵ファイルを `credentials.json` という名前でプロジェクトのルートディレクトリに配置してください。
   *注意: `credentials.json` も公開しないでください。*

5. **スプレッドシートの共有設定**
   出力先となるGoogleスプレッドシートの共有設定から、`credentials.json` 内に記載されている `client_email` のアドレスに対して「編集者」権限を付与してください。

## 使用方法 (Usage)

以下のコマンドでスクリプトを実行します。

```bash
python main.py
```

実行が完了すると、コンソールに進行状況と取得結果が表示され、対象のスプレッドシートが自動的に更新されます。

## ライセンス (License)

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
