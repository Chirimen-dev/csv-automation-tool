# CSV自動整理・統合ツール

複数のCSVファイルを統合し、空白の整理、欠損セルの集計、重複行の削除を行うPythonツールです。

## 使い方

1. CSVファイルを input/ に入れる
2. 仮想環境を有効化する: source .venv/bin/activate
3. 実行する: python csv_tool.py

## 出力

- output/merged.csv: 統合後のCSV
- output/report.txt: 処理結果のレポート

## 必須列

- 氏名
- メールアドレス
- 電話番号
