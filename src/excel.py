from itertools import product
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from src.config_loader import Config, load_column_names


def apply_cell_style(cell, font, fill=None, alignment=None, border=None):
    cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border


class ExcelWriter:

    def __init__(self, df: pd.DataFrame, config_excel: Config):
        self.df = df.copy()
        self.config = config_excel

        self.columns = load_column_names(self.config)
        self.col_names = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[:len(self.columns)]
        if len(self.col_names) < len(self.columns):
            raise IndexError("列は26列以下にしてください")

    def __write_test_specification_sheet(self,
                                         writer: pd.ExcelWriter,
                                         merge_cells: bool = False,
                                         ):
        """テスト仕様書をエクセルシートに書き込みます。

        Args:
            writer (pd.ExcelWriter):    ExcelWriterオブジェクト
            merge_cells (bool):         セルをマージするかどうか
        """
        df_excel = self.df.copy()
        df_excel.columns = self.columns

        # マージできるようにマルチインデックス化
        multi_idx_cols = []
        if merge_cells:
            multi_idx_cols = [v["name"] for k, v in self.config.columns.model_dump().items() if v.get("multi_idx")]
            if multi_idx_cols:
                # マージセルを使うが、NOをA列に確実に表示するため別の方法でマージする
                pass
            else:
                merge_cells = False  # マージ対象の列がなければマージを無効化

        # NOをA列に配置するため、常にindex=Falseでデータを書き出す
        df_excel.to_excel(
            writer, 
            sheet_name=self.config.excel_settings.sheet_name.test, 
            merge_cells=False,  # 常にFalseにして自前でマージ処理を行う 
            index=False,  # インデックス列を非表示にする
        )

        # ワークシートの取得
        worksheet = writer.sheets[self.config.excel_settings.sheet_name.test]

        # ヘッダーのスタイル設定
        for col, length in zip(self.col_names, [col["length"] for col in self.config.columns.model_dump().values()]):
            apply_cell_style(
                worksheet[f"{col}1"],
                font=Font(name=self.config.excel_settings.font_name, bold=True, color="ffffff"),
                fill=PatternFill(patternType="solid", fgColor="4f81bd"),
                alignment=Alignment(vertical="center", horizontal="center", wrap_text=True),
            )

            # 列幅調整
            worksheet.column_dimensions[col].width = length

        # データセルのスタイル調整
        for i in range(len(df_excel)):
            for col, (horizontal, vertical) in (
                    zip(self.col_names, zip([x["horizontal"] for x in self.config.columns.model_dump().values()],
                                            [x["vertical"] for x in self.config.columns.model_dump().values()])
                        )):
                apply_cell_style(
                    worksheet[f"{col}{i + 2}"],
                    font=Font(name=self.config.excel_settings.font_name),
                    alignment=Alignment(horizontal=horizontal, vertical=vertical, wrap_text=True),
                    border=Border(left=Side(style="thin"), right=Side(style="thin"),
                                  top=Side(style="thin"), bottom=Side(style="thin"))
                )
                
        # マージセルの処理
        if merge_cells and multi_idx_cols:
            # マージ対象の列のインデックスを取得
            merge_col_indices = [self.columns.index(col) for col in multi_idx_cols]
            
            # 各マージ対象列に対して処理
            for col_idx in merge_col_indices:
                col_letter = self.col_names[col_idx]
                
                # マージするセルの範囲を決定
                current_value = None
                start_row = 2  # データは2行目から始まる
                
                for i in range(len(df_excel)):
                    row = i + 2  # Excel行番号（2行目からデータ開始）
                    cell_value = worksheet[f"{col_letter}{row}"].value
                    
                    # 値が変わったか、最後の行の場合
                    if cell_value != current_value or i == len(df_excel) - 1:
                        # 前のグループがあればマージ
                        if current_value is not None and row - start_row > 1:
                            # 最後の行の場合は現在の行も含める
                            end_row = row if cell_value != current_value else row + 1
                            worksheet.merge_cells(f"{col_letter}{start_row}:{col_letter}{end_row-1}")
                        
                        # 新しいグループの開始
                        current_value = cell_value
                        start_row = row

    def __call__(self, output_path: Path, merge_cells: bool = True):
        """
        convert_md_to_df()により生成されたデータフレームをエクセルファイルに変換します。

        Args:
            df (pd.DataFrame):    convert_md_to_df()により生成されたデータフレーム
        """
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                self.__write_test_specification_sheet(writer, merge_cells)
        except PermissionError:
            raise PermissionError(f"出力先のファイルを開いている可能性があります。エクセルファイルを閉じてください。")
        except Exception as e:
            raise ValueError(f"エクセルファイル出力中に不明なエラーが発生しました：\n{e}")

        return output_path
