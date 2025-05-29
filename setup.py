from setuptools import setup, find_packages
import os
from glob import glob

# アセットディレクトリが存在するか確認し、なければ作成
assets_dir = os.path.join('md_test_case_to_excel', 'assets')
if not os.path.exists(assets_dir):
    os.makedirs(assets_dir, exist_ok=True)

# テンプレートファイルの存在を確認
template_file = os.path.join(assets_dir, 'ARMDXP_単体・結合試験_DAS-M_テンプレート_md.xlsx')
if not os.path.exists(template_file):
    # ルートのassetsからコピー
    root_template = os.path.join('assets', 'ARMDXP_単体・結合試験_DAS-M_テンプレート_md.xlsx')
    if os.path.exists(root_template):
        import shutil
        shutil.copy2(root_template, template_file)
        print(f"テンプレートファイルをコピーしました: {root_template} -> {template_file}")
    else:
        print(f"警告: テンプレートファイルが見つかりません: {root_template}")
        # 空のテンプレートファイルを作成（最低限のビルドは成功させるため）
        with open(template_file, 'wb') as f:
            f.write(b'dummy')
        print(f"空のテンプレートファイルを作成しました: {template_file}")

setup(
    name="md_test_case_to_excel",
    version="0.3.0",
    description="Markdownで書かれたテスト仕様書をExcel形式に変換するツール",
    author="k-watanb",
    author_email="",
    packages=find_packages(),
    include_package_data=True,
    package_data={
        "md_test_case_to_excel": ["*.yaml", "assets/*.xlsx"],
    },
    data_files=[
        ('md_test_case_to_excel', ['md_test_case_to_excel/config.yaml']),
    ],
    install_requires=[
        "pandas>=2.2.0",
        "openpyxl>=3.1.0",
        "pydantic>=2.9.0", 
        "pyyaml>=6.0.0"
    ],
    entry_points={
        'console_scripts': [
            'md2excel=md_test_case_to_excel.converter:main',
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
)