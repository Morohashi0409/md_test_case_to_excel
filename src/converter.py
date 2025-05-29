"""
Markdownで書かれたテスト仕様書をエクセルファイルに変換します。

Usage:
    md2excel -h
    md2excel [-f] <file> [--template] [--no-auto-width] [--test-type <type>]
    md2excel [-f] <file> [--ut|--it]  # 単体試験・結合試験の略称
"""

import argparse
import os
import sys
from pathlib import Path
import shutil

from src.config_loader import load_config
from src.excel import ExcelWriter
from src.markdown import MarkdownTestParser, read_markdown_file

def find_package_root():
    """
    パッケージのルートディレクトリを探します。
    インストール済みの場合はインストール先ディレクトリ、
    開発中の場合はプロジェクトのルートディレクトリを返します。
    """
    # まずは実行ファイルの場所を取得
    if getattr(sys, 'frozen', False):
        # PyInstallerなどで実行ファイルになっている場合
        base_path = Path(sys.executable).parent
    else:
        # 通常のPythonスクリプトとして実行されている場合
        # モジュールの場所を基準にする
        base_path = Path(__file__).parent.parent
        
    # config.yaml の場所を確認
    config_path = base_path / "config.yaml"
    if config_path.exists():
        return base_path
    
    # インストールされたパッケージの場合、siteパッケージ内を探す
    try:
        import pkg_resources
        dist = pkg_resources.get_distribution("md_test_case_to_excel")
        return Path(dist.location) / "md_test_case_to_excel"
    except:
        pass
    
    # 最後の手段として環境変数をチェック
    if "MD_TEST_CASE_TO_EXCEL_ROOT" in os.environ:
        return Path(os.environ["MD_TEST_CASE_TO_EXCEL_ROOT"])
    
    # それでも見つからない場合はカレントディレクトリを使用
    return Path.cwd()

def convert_md_to_excel(file_path, template=False, no_auto_width=False, test_type="test"):
    """
    Markdownファイルをエクセルファイルに変換する関数
    
    Args:
        file_path (str): 入力ファイルパス
        template (bool): テンプレートを使用するかどうか
        no_auto_width (bool): 列幅の自動調整を無効にするかどうか
        test_type (str): テストの種別（test, ut, it）
        
    Returns:
        Path: 出力されたファイルのパス
    """
    package_root = find_package_root()
    
    # 設定ファイルの読み込み
    config = load_config(package_root / "config.yaml")
    
    # Markdownファイルの読み込みと解析
    markdown_content = read_markdown_file(Path(file_path))
    parser = MarkdownTestParser(markdown_content, config)
    df = parser.parse()
    print(f"-------\n{df}\n-------")

    writer = ExcelWriter(df, config)
    
    # テンプレートパスの設定
    template_path = None
    if template:
        template_path = package_root / "assets" / "ARMDXP_単体・結合試験_DAS-M_テンプレート_md.xlsx"
        if not template_path.exists():
            print(f"警告: テンプレートファイル {template_path} が見つかりません。新規ファイルを作成します。")
            template_path = None
        else:
            print(f"テンプレートファイル {template_path} を使用します。")
    
    # 既存のExcelファイルが存在し、--templateオプションが指定されていない場合に既存ファイルをテンプレートとして使用
    output_path = Path(file_path).parent / f"{Path(file_path).stem}.xlsx"
    if output_path.exists() and not template:
        print(f"既存のExcelファイル {output_path} をテンプレートとして使用します。")
        template_path = output_path
    
    # 出力先のパスを決定
    if template_path and template_path.exists():
        # テンプレートファイルを使用する場合
        if template_path != output_path:  # テンプレートと出力先が異なる場合はコピー
            # テンプレートファイルをそのまま残して、コピーをして使う
            output_path = writer(output_path, 
                                merge_cells=True, 
                                template_path=template_path,
                                auto_adjust_width=not no_auto_width,
                                auto_adjust_height=True,
                                preserve_additional_columns=True,
                                test_type=test_type)
        else:  # 既存ファイルの上書き更新の場合
            output_path = writer(output_path, 
                                merge_cells=True, 
                                template_path=template_path,
                                auto_adjust_width=not no_auto_width,
                                auto_adjust_height=True,
                                preserve_additional_columns=True,
                                test_type=test_type)
    else:
        # 従来通りの処理 (新規ファイル作成)
        output_path = writer(output_path, 
                            merge_cells=True,
                            auto_adjust_width=not no_auto_width,
                            auto_adjust_height=True,
                            test_type=test_type)
    
    # 出力したシート名を表示する
    sheet_name = ""
    if test_type == "ut":
        sheet_name = config.excel_settings.sheet_name.ut
    elif test_type == "it":
        sheet_name = config.excel_settings.sheet_name.it
    else:
        sheet_name = config.excel_settings.sheet_name.test
        
    print(f"\nDone! The file is saved at `{output_path}` (シート: {sheet_name}).")
    
    return output_path

def main():
    """
    コマンドラインツールのエントリーポイント
    """
    parser = argparse.ArgumentParser(description="Markdownで書かれたテスト仕様書をエクセルファイルに変換します。")
    parser.add_argument("-f", "--file", type=str, required=True, help="入力ファイルパス")
    parser.add_argument("--template", action="store_true", help="テンプレートExcelファイルを使用する場合に指定")
    parser.add_argument("--no-auto-width", action="store_true", help="列幅の自動調整を無効にする場合に指定")
    
    # テスト種別の指定方法（ショートカットと詳細オプションのグループ化）
    test_type_group = parser.add_mutually_exclusive_group()
    test_type_group.add_argument("--test-type", type=str, choices=["test", "ut", "it"], 
                               default="test",
                               help="テストの種別（test:テスト仕様書、ut:単体試験、it:結合試験）")
    test_type_group.add_argument("--ut", action="store_const", const="ut", dest="test_type",
                               help="単体試験シートに出力する（--test-type utのショートカット）")
    test_type_group.add_argument("--it", action="store_const", const="it", dest="test_type",
                               help="結合試験シートに出力する（--test-type itのショートカット）")
    
    args = parser.parse_args()
    
    convert_md_to_excel(
        args.file,
        template=args.template,
        no_auto_width=args.no_auto_width,
        test_type=args.test_type
    )

if __name__ == "__main__":
    main()