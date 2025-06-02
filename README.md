# md_test_case_to_excel

マークダウン形式で書いたテスト仕様書をExcel形式に変換するためのツールです。マークダウンの編集機能とGitHubでの差分管理を活用しながら、必要に応じてExcel形式で共有できます。

![](attachments/excel-image.png)

## 環境要件

- Python 3.11以上
- 以下のPythonパッケージ:
  - pandas 2.2.2
  - openpyxl 3.1.5
  - pydantic 2.9.1
  - pyyaml 6.0.2

## インストール手順

### 1. Pythonのインストール

まだPythonがインストールされていない場合は、[Python公式サイト](https://www.python.org/downloads/)からインストールしてください。

### 2. md_test_case_to_excelのインストール

#### 方法1: pipを使用してインストール

```bash
pip install md-test-case-to-excel
```

#### 方法2: ソースからインストール

```bash
# リポジトリをクローン
git clone https://github.com/username/md_test_case_to_excel.git
cd md_test_case_to_excel

# 依存関係をインストール
pip install -r requirements.txt

# パッケージをインストール
pip install -e .
```

### 3. アンインストール方法

#### pipでインストールした場合

```bash
pip uninstall md-test-case-to-excel
```

#### ソースからインストールした場合

```bash
# インストールディレクトリに移動
cd path/to/md_test_case_to_excel

# アンインストール
pip uninstall md-test-case-to-excel

# もしくは以下のコマンドでも可能
python setup.py develop --uninstall
```

## 使い方

### テスト仕様書を作成する

以下の形式に従ってテスト仕様書を作成します:

```markdown
# テスト仕様書

## 大項目
### 中項目
#### [正常|異常|準正常] [OK|NG|未実施|--] テストケース名
1. 確認手順1
2. 確認手順2
* [ ] 想定動作1
* [ ] 想定動作2
- 備考内容
```

### Excelに変換する

#### コマンドラインから変換する場合

```bash
# 基本的な使い方
md-test-case-to-excel -f path/to/your/testspec.md --template

# または直接Pythonモジュールを実行する場合
python -m md_test_case_to_excel.converter -f path/to/your/testspec.md --template
```

## コマンドラインオプション

|オプション名|説明|
|:---|:---|
|-h, --help| 引数のヘルプ表示|
|-f, --file| 入力ファイルパス|
|--template| テンプレートExcelファイルを使用する場合に指定|
|--test-type| テストの種別（test:テスト仕様書、ut:単体試験、it:結合試験）|
|--ut| 単体試験シートに出力する（--test-type utのショートカット）|
|--it| 結合試験シートに出力する（--test-type itのショートカット）|
|--no-auto-width| 列幅の自動調整を無効にする場合に指定|

## 応用例

### 効率的なワークフロー

1. **新規テスト仕様書の作成**:
   - テンプレートファイルから新規テスト仕様書を作成
   - マークダウン形式で作成・編集

2. **バージョン管理**:
   - Gitを使用して変更履歴を管理
   - マークダウン形式のため、差分確認が容易

3. **Excel出力と共有**:
   - レビューや提出が必要な場合はExcel形式に変換
   - コマンドを実行するだけで最新内容をExcelに反映

### 既存Excelファイルの更新

既存のExcelファイルがある場合、そのファイルに追記する形で更新できます。
Excelファイル内のJ列以降のコメントや試験結果などのデータは自動的に保持されます。

```bash
# Markdownファイルを更新後、既存のExcelファイルに追記する
md-test-case-to-excel -f example/updated_sample.md
```

### シート選択機能

```bash
# 単体試験シートに書き込む
md-test-case-to-excel -f example/testcases.md --ut

# 結合試験シートに書き込む
md-test-case-to-excel -f example/testcases.md --it
```

## カスタマイズ

設定ファイル`config.yaml`を編集することで、様々なカスタマイズが可能です:

- フォント名や各シート名の変更
- 列幅や列のフォーマットの調整
- マークダウンの解析パターンの変更

```yaml
excel_settings:
  font_name: Meiryo UI
  sheet_name:
    summary: サマリー
    test: テスト仕様書
    ut: 単体試験
    it: 結合試験

columns:
  number:
    name: 'NO'
    length: 6
    horizontal: 'center'
    vertical: 'center'
  # 他の設定は省略
```

## トラブルシューティング

### エクセルファイルが更新できない
エクセルファイルが他のアプリケーションで開かれていると更新できません。エラーが表示される場合は、エクセルファイルを閉じてから再実行してください。

### テンプレートが見つからない
`--template`オプション指定時、デフォルトでは`assets/ARMDXP_単体・結合試験_DAS-M_テンプレート_md.xlsx`を使用します。このファイルが存在しない場合は、新規にファイルを作成します。

## リリースノート

### v0.3.0 (2025/5/29)
- 単体試験と結合試験のシート選択機能を追加
- コマンドラインオプションを拡充

### v0.2.0 (2024/9/16)
- ソースコード全体をリファクタリング
- 設定ファイルの構造を見直し
- 最新バージョンのライブラリに対応

### v0.1.0 (2020/12/08)
- 初回リリース

## 謝辞

マークダウン形式は[ryuta46/eval-spec-maker](https://github.com/ryuta46/eval-spec-maker)を参考にしています。Pythonコードは[torisawa/convert.py](https://gist.github.com/toriwasa/37c690862ddf67d43cfd3e1af4e40649)を参考にしています。

## 制限事項

- Excelヘッダーは日本語のみ対応しています
