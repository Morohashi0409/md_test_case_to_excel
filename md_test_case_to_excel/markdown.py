import re
from pathlib import Path

import pandas as pd

from md_test_case_to_excel.config_loader import Config


class MarkdownTestParser:

    def __init__(self, markdown_content: str, config: Config):
        """Markdownテスト仕様書を解析し、データフレームに変換するクラス

        Args:
            markdown_content (str):  Markdown形式のテスト仕様書
            config (Config):         設定情報


        """

        self.markdown_content = markdown_content
        self.config = config

        # 新しいカラム順序: ["NO", "大分類", "中分類", "小分類", "試験内容", "確認事項"]
        self.columns = self.config.columns.model_fields.keys()
        self.data = []

        self.pattern_section = re.compile(self.config.columns.section.md_pattern, re.MULTILINE)
        self.pattern_subsection = re.compile(self.config.columns.subsection.md_pattern, re.MULTILINE)
        self.pattern_testcase = re.compile(self.config.columns.testcase.md_pattern, re.MULTILINE)
        self.pattern_step = re.compile(self.config.columns.step.md_pattern, re.MULTILINE)
        self.pattern_expectation = re.compile(self.config.columns.expectation.md_pattern, re.MULTILINE)

        # 階層構造のNO管理用の変数
        self.section_count = 0
        self.section_map = {}  # 大分類名をキーとし、その番号を値とするディクショナリ
        self.subsection_count = 0
        self.subsection_map = {}  # '大分類名:中分類名'をキーとし、その番号を値とするディクショナリ
        self.testcase_count = 0

    def parse(self) -> pd.DataFrame:
        """Markdownファイルを解析し、データフレーム用のデータを作成します。

        Returns:
            pd.DataFrame: 解析結果のデータフレーム

        Notes:
            - テスト仕様書のデフォルト形式は以下の通り
                # 大項目
                ## 中項目
                ### 小項目
                #### テストケース名
                1. 確認手順
                2. 確認手順
                * [ ] 期待値
        """
        current_section = None
        current_subsection = None
        last_section = None
        last_subsection = None

        lines = self.markdown_content.split('\n')

        for i, line in enumerate(lines):
            section_match = self.pattern_section.match(line)
            subsection_match = self.pattern_subsection.match(line)
            testcase_match = self.pattern_testcase.match(line)

            if section_match:
                current_section = section_match.group(1)
                current_subsection = None  # Reset subsection when a new section is found
                
                # 大分類が変わった場合、カウンターを増やす
                if current_section != last_section:
                    if current_section not in self.section_map:
                        self.section_count += 1
                        self.section_map[current_section] = self.section_count
                    last_section = current_section
                    self.subsection_count = 0 # 中分類のカウンターをリセット
                    self.subsection_map = {}  # 中分類のマップもリセット
            
            elif subsection_match:
                current_subsection = subsection_match.group(1)
                
                # 中分類が変わった場合、カウンターを増やす
                if current_subsection != last_subsection or current_section != last_section:
                    subsection_key = f"{current_section}:{current_subsection}"
                    if subsection_key not in self.subsection_map:
                        self.subsection_count += 1
                        self.subsection_map[subsection_key] = self.subsection_count
                    last_subsection = current_subsection
                    self.testcase_count = 0  # 小分類（テストケース）のカウンターをリセット
            
            elif testcase_match:
                # テストケースごとにカウンターを増やす
                self.testcase_count += 1
                
                # 新しいパターン - 直接テストケース名を取得
                test_case_name = testcase_match.group(1)
                steps, expectations = [], []

                # 直前の行からの継続かどうかを判断するフラグ
                continuing_expectation = False
                current_expectation = ""

                for subline in lines[i + 1:]:
                    step_match = self.pattern_step.match(subline)
                    expectation_match = self.pattern_expectation.match(subline)

                    if step_match:
                        steps.append(step_match.group(1))
                        continuing_expectation = False  # 確認事項の継続をリセット
                    elif expectation_match:
                        # 新しい確認事項が始まる場合は、前の確認事項を追加
                        if continuing_expectation and current_expectation:
                            expectations.append(current_expectation)
                        # 新しい確認事項を開始
                        current_expectation = expectation_match.group(1)
                        continuing_expectation = True
                    elif subline.strip() and continuing_expectation:
                        # 空行でなく、かつ確認事項の継続中なら、その行を現在の確認事項に追加
                        current_expectation += "\n" + subline.strip()
                    elif subline.startswith('####') or subline.startswith('###') or subline.startswith('##'):
                        # 次のセクションが始まったら処理終了
                        break
                    elif not subline.strip():
                        # 空行の場合、確認事項の継続が終了
                        if continuing_expectation and current_expectation:
                            expectations.append(current_expectation)
                            continuing_expectation = False
                            current_expectation = ""

                # 最後の確認事項が残っていれば追加
                if continuing_expectation and current_expectation:
                    expectations.append(current_expectation)

                # 階層構造のNO値を設定
                section_num = self.section_map.get(current_section, 0)
                subsection_num = self.subsection_map.get(f"{current_section}:{current_subsection}", 0) if current_subsection else 0
                hierarchical_no = f"{section_num}-{subsection_num}-{self.testcase_count}"
                
                # データフレーム用の行データ作成（新しい順序）
                self.data.append([
                    hierarchical_no,     # 階層化されたNO
                    current_section,     # 大分類
                    current_subsection,  # 中分類  
                    test_case_name,      # 小分類
                    '\n'.join([f"{i + 1}. {step}" for i, step in enumerate(steps)]),  # 試験内容
                    '\n'.join([f"・{exp}" for exp in expectations])   # 確認事項
                ])

        return pd.DataFrame(self.data, columns=self.columns)


def read_markdown_file(file_path: Path) -> str:
    if not file_path.exists():
        raise FileNotFoundError(f"Markdownファイルが見つかりません: {file_path}")
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()
