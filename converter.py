"""
Markdownで書かれたテスト仕様書をエクセルファイルに変換します。

Usage:
    python converter.py -h
    python converter.py [-f] <file> [--no-merge] [--template] [--no-auto-width] [--no-auto-height] [--preserve-columns] [--test-type <type>]
    python converter.py [-f] <file> [--ut|--it]  # 単体試験・結合試験の略称
"""

import argparse
from pathlib import Path
import shutil

from src.config_loader import load_config
from src.excel import ExcelWriter
from src.markdown import MarkdownTestParser, read_markdown_file

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Markdownで書かれたテスト仕様書をエクセルファイルに変換します。")
    parser.add_argument("-f", "--file", type=str, required=True, help="入力ファイルパス")
    parser.add_argument("--no-merge", action="store_true", help="セルをマージしない場合に指定")
    parser.add_argument("--template", action="store_true", help="テンプレートExcelファイルを使用する場合に指定")
    parser.add_argument("--no-auto-width", action="store_true", help="列幅の自動調整を無効にする場合に指定")
    parser.add_argument("--no-auto-height", action="store_true", help="行高の自動調整を無効にする場合に指定")
    parser.add_argument("--preserve-columns", action="store_true", help="既存ファイルのJ列以降の内容を保持する場合に指定")
    
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

    config = load_config(Path(__file__).parent.joinpath("config.yaml"))

    markdown_content_example = read_markdown_file(Path(args.file))
    parser = MarkdownTestParser(markdown_content_example, config)
    df = parser.parse()
    print(f"-------\n{df}\n-------")

    writer = ExcelWriter(df, config)
    
    # テンプレートパスの設定
    template_path = None
    if args.template:
        template_path = Path(__file__).parent.joinpath("assets", "ARMDXP_単体・結合試験_DAS-M_テンプレート_md.xlsx")
        if not template_path.exists():
            print(f"警告: テンプレートファイル {template_path} が見つかりません。新規ファイルを作成します。")
            template_path = None
        else:
            print(f"テンプレートファイル {template_path} を使用します。")
    
    # 既存のExcelファイルが存在し、--templateオプションが指定されていない場合に既存ファイルをテンプレートとして使用
    output_path = Path(args.file).parent.joinpath(f"{Path(args.file).stem}.xlsx")
    if output_path.exists() and not args.template:
        print(f"既存のExcelファイル {output_path} をテンプレートとして使用します。")
        template_path = output_path
    
    # 出力先のパスを決定
    if template_path and template_path.exists():
        # テンプレートファイルを使用する場合
        if template_path != output_path:  # テンプレートと出力先が異なる場合はコピー
            # テンプレートファイルをそのまま残して、コピーをして使う
            output_path = writer(output_path, 
                                merge_cells=not args.no_merge, 
                                template_path=template_path,
                                auto_adjust_width=not args.no_auto_width,
                                auto_adjust_height=not args.no_auto_height,
                                preserve_additional_columns=args.preserve_columns,
                                test_type=args.test_type)
        else:  # 既存ファイルの上書き更新の場合
            output_path = writer(output_path, 
                                merge_cells=not args.no_merge, 
                                template_path=template_path,
                                auto_adjust_width=not args.no_auto_width,
                                auto_adjust_height=not args.no_auto_height,
                                preserve_additional_columns=args.preserve_columns,
                                test_type=args.test_type)
    else:
        # 従来通りの処理 (新規ファイル作成)
        output_path = writer(output_path, 
                            merge_cells=not args.no_merge,
                            auto_adjust_width=not args.no_auto_width,
                            auto_adjust_height=not args.no_auto_height,
                            test_type=args.test_type)
    
    # 出力したシート名を表示する
    sheet_name = ""
    if args.test_type == "ut":
        sheet_name = config.excel_settings.sheet_name.ut
    elif args.test_type == "it":
        sheet_name = config.excel_settings.sheet_name.it
    else:
        sheet_name = config.excel_settings.sheet_name.test
        
    print(f"\nDone! The file is saved at `{output_path}` (シート: {sheet_name}).")
