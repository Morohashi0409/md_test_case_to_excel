from setuptools import setup, find_packages

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
        ('md_test_case_to_excel/assets', ['md_test_case_to_excel/assets/ARMDXP_単体・結合試験_DAS-M_テンプレート_md.xlsx']),
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