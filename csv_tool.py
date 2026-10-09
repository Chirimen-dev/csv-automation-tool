"""
csv_tool.py - CSV自動整理・統合ツール
 
input/ フォルダ内の複数のCSVファイルを読み込み、表記を整え、
1つのCSV (output/merged.csv) に統合します。
処理結果は output/report.txt に出力します。
 
【実行方法】
    python csv_tool.py
 
【フォルダ構成】(csv_tool.py があるフォルダを基準にします)
    csv_tool.py
    input/        ← ここに統合したいCSVを入れる
    output/       ← 結果が出力される(なければ自動作成)
 
動作環境: Python 3.12 / pandas
"""
 
import sys
from dataclasses import dataclass
from pathlib import Path
 
import pandas as pd
 
# ---------------------------------------------------------------------------
# 設定(変更したい場合はここだけ直せばよい部分)
# ---------------------------------------------------------------------------
 
# csv_tool.py があるフォルダ。実行した場所(カレントディレクトリ)には影響されない。
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
MERGED_CSV_PATH = OUTPUT_DIR / "merged.csv"
REPORT_PATH = OUTPUT_DIR / "report.txt"
 
# 必須列。これらが1つでも無いCSVはエラーになる。
REQUIRED_COLUMNS = ["氏名", "メールアドレス", "電話番号"]
 
# 入力CSVの文字コード。左から順に試す。
# utf-8-sig は「BOMあり/なし」どちらのUTF-8も読める。
# cp932 は Excel で保存した日本語CSV(Shift_JIS系)用。
INPUT_ENCODINGS = ["utf-8-sig", "cp932"]
 
# 出力の文字コード。BOM付きUTF-8にすると Excel で開いても文字化けしにくい。
OUTPUT_ENCODING = "utf-8-sig"
 
 
# ---------------------------------------------------------------------------
# データの入れ物とエラー用の型
# ---------------------------------------------------------------------------
 
class CsvToolError(Exception):
    """このツール独自のエラー。メッセージはそのまま利用者に表示する。"""
 
 
@dataclass
class LoadedCsv:
    """読み込み・整形が終わった1つのCSVの情報をまとめて持つ入れ物。"""
 
    path: Path            # ファイルの場所
    df: pd.DataFrame      # 整形済みのデータ本体
    encoding: str         # 実際に読めた文字コード
    missing_cells: int    # このCSVの欠損セル数(統合前)
 
    @property
    def name(self) -> str:
        """ファイル名(例: sales.csv)。エラーメッセージ用。"""
        return self.path.name
 
 
# ---------------------------------------------------------------------------
# 1. 入力ファイルの探索
# ---------------------------------------------------------------------------
 
def find_csv_files() -> list[Path]:
    """input/ 直下のCSVファイルを、ファイル名順で返す。(サブフォルダは対象外)"""
    if not INPUT_DIR.exists():
        raise CsvToolError(
            f"inputフォルダが見つかりません: {INPUT_DIR}\n"
            f"    → csv_tool.py と同じ場所に「input」フォルダを作り、CSVを入れてください。"
        )
    if not INPUT_DIR.is_dir():
        raise CsvToolError(
            f"「input」はフォルダではなくファイルです: {INPUT_DIR}\n"
            f"    → そのファイルの名前を変えて、「input」フォルダを作ってください。"
        )
 
    try:
        # 拡張子は大文字小文字を区別しない(.csv も .CSV も対象)
        csv_files = [
            p for p in INPUT_DIR.iterdir()
            if p.is_file() and p.suffix.lower() == ".csv"
        ]
    except OSError as e:
        raise CsvToolError(f"inputフォルダの中身を取得できませんでした: {INPUT_DIR} ({e})") from e
 
    if not csv_files:
        raise CsvToolError(
            f"CSVファイルが1つもありません: {INPUT_DIR}\n"
            f"    → 拡張子が .csv のファイルを input フォルダの直下に入れてください。"
        )
 
    # ファイル名順(大文字小文字は区別しない)。同じ順序で毎回処理されるようにする。
    return sorted(csv_files, key=lambda p: (p.name.lower(), p.name))
 
 
# ---------------------------------------------------------------------------
# 2. CSVの読み込みと整形
# ---------------------------------------------------------------------------
 
def read_csv_with_encoding_fallback(path: Path) -> tuple[pd.DataFrame, str]:
    """UTF-8(BOM可) → cp932 の順に文字コードを試してCSVを読み込む。"""
    for encoding in INPUT_ENCODINGS:
        try:
            df = pd.read_csv(
                path,
                dtype=str,              # 全列を文字列として読む(電話番号の先頭の0を守る)
                keep_default_na=False,  # "NA" や "null" という文字を勝手に欠損扱いにしない
                encoding=encoding,
            )
            return df, encoding
        except UnicodeDecodeError:
            continue  # この文字コードでは読めなかったので、次を試す
 
    raise CsvToolError(
        f"{path.name}: 文字コードを判別できませんでした"
        f"(試した文字コード: {', '.join(INPUT_ENCODINGS)})。\n"
        f"    → ファイルが壊れているか、CSV以外のファイルの可能性があります。"
    )
 
 
def load_one_csv(path: Path) -> LoadedCsv:
    """1つのCSVを読み込み、列名と文字列の前後の空白を削除し、欠損セル数を数える。"""
    # --- 読み込み(失敗したら原因が分かるメッセージにして CsvToolError にする) ---
    try:
        df, encoding = read_csv_with_encoding_fallback(path)
    except CsvToolError:
        raise
    except pd.errors.EmptyDataError:
        raise CsvToolError(
            f"{path.name}: ファイルが空です(列名の行もありません)。"
        ) from None
    except pd.errors.ParserError as e:
        raise CsvToolError(
            f"{path.name}: CSVの形式が正しくありません。({e})\n"
            f"    → カンマの数が行によって違う、引用符(\")が閉じていない、などが原因です。"
        ) from e
    except OSError as e:
        raise CsvToolError(
            f"{path.name}: ファイルを開けませんでした。({e})\n"
            f"    → Excelなどで開いたままになっていないか確認してください。"
        ) from e
    except Exception as e:  # 想定外のエラーも、ファイル名付きで表示する
        raise CsvToolError(
            f"{path.name}: 読み込み中に想定外のエラーが発生しました。({type(e).__name__}: {e})"
        ) from e
 
    # データ行が見出し行より1項目多い場合、pandasが先頭列を行番号として扱ってしまう。
    # 列がずれたまま処理が進むと危険なので、ここでエラーにする。
    if not isinstance(df.index, pd.RangeIndex):
        raise CsvToolError(
            f"{path.name}: データ行の項目数が、見出し行(列名の行)より多くなっています。\n"
            f"    → 見出し行の列数と、データ行の列数が一致しているか確認してください。"
        )
 
    # 項目が足りない行のセルは欠損値(NaN)になることがあるので、空文字に統一する。
    df = df.fillna("")
 
    # --- 列名の前後の空白を削除 ---
    df.columns = [str(col).strip() for col in df.columns]
 
    duplicated = df.columns[df.columns.duplicated()].unique().tolist()
    if duplicated:
        raise CsvToolError(
            f"{path.name}: 列名が重複しています: {', '.join(duplicated)}\n"
            f"    → 前後の空白を除くと同じ名前になる列がないか確認してください。"
        )
 
    # --- 文字列データの前後の空白を削除(全角スペースも対象) ---
    for col in df.columns:
        df[col] = df[col].str.strip()
 
    # --- 欠損セルを数える(空白削除後に空文字になっているセル) ---
    missing_cells = int((df == "").sum().sum())
 
    return LoadedCsv(path=path, df=df, encoding=encoding, missing_cells=missing_cells)
 
 
# ---------------------------------------------------------------------------
# 3. 列名のチェック
# ---------------------------------------------------------------------------
 
def check_columns(loaded_files: list[LoadedCsv]) -> list[str]:
    """必須列の不足と、CSV間の列構成の違いを調べ、エラーメッセージのリストを返す。
 
    列の「順番」は問わず、列名の「集合」が同じかどうかだけを見る。
    """
    errors: list[str] = []
    valid_files: list[LoadedCsv] = []  # 必須列がそろっているCSV
 
    # (a) 必須列のチェック
    for item in loaded_files:
        missing = [col for col in REQUIRED_COLUMNS if col not in item.df.columns]
        if missing:
            errors.append(
                f"{item.name}: 必須列が不足しています。不足している列: {', '.join(missing)}"
            )
        else:
            valid_files.append(item)
 
    # (b) 列構成の一致チェック
    # 必須列がそろっているCSVのうち、ファイル名順で最初のものを基準にする。
    # (エラーがなければ、基準は「ファイル名順で最初のCSV」と一致する)
    if valid_files:
        reference = valid_files[0]
        reference_columns = list(reference.df.columns)
 
        for item in valid_files[1:]:
            if set(item.df.columns) == set(reference_columns):
                continue
            lacking = [c for c in reference_columns if c not in item.df.columns]
            extra = [c for c in item.df.columns if c not in reference_columns]
            detail = []
            if lacking:
                detail.append(f"基準にあるが、このファイルにない列: {', '.join(lacking)}")
            if extra:
                detail.append(f"このファイルにだけある列: {', '.join(extra)}")
            errors.append(
                f"{item.name}: 列構成が {reference.name} と異なります。\n"
                + "\n".join(f"    - {line}" for line in detail)
            )
 
    return errors
 
 
# ---------------------------------------------------------------------------
# 4. 統合と重複削除
# ---------------------------------------------------------------------------
 
def merge_and_deduplicate(loaded_files: list[LoadedCsv]) -> tuple[pd.DataFrame, int]:
    """CSVを1つに統合し、完全に同一の行を削除する。
 
    戻り値: (処理後のデータ, 重複削除数)
    """
    # 統合後の列順は、最初のCSV(ファイル名順)の列順に合わせる。
    base_columns = list(loaded_files[0].df.columns)
 
    # 他のCSVの列順が違っていても、列名で並べ替えてから結合する。
    frames = [item.df[base_columns] for item in loaded_files]
    merged = pd.concat(frames, ignore_index=True)
 
    rows_before_dedup = len(merged)
    # keep="first": 同じ行が複数あるとき、最初に出てきた行を残す
    merged = merged.drop_duplicates(keep="first").reset_index(drop=True)
    duplicates_removed = rows_before_dedup - len(merged)
 
    return merged, duplicates_removed
 
 
# ---------------------------------------------------------------------------
# 5. 出力
# ---------------------------------------------------------------------------
 
def build_report(
    file_count: int,
    rows_before: int,
    duplicates_removed: int,
    missing_cells: int,
    rows_after: int,
) -> str:
    """report.txt に書く文章を作る。"""
    lines = [
        f"処理したCSVファイル数: {file_count}",
        f"処理前の総行数: {rows_before}",
        f"重複削除数: {duplicates_removed}",
        f"欠損セル数: {missing_cells}",
        f"処理後の総行数: {rows_after}",
    ]
    return "\n".join(lines) + "\n"
 
 
def write_outputs(merged: pd.DataFrame, report_text: str) -> None:
    """output/merged.csv と output/report.txt を保存する(既存ファイルは上書き)。"""
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)  # なければ作る
    except OSError as e:
        raise CsvToolError(f"outputフォルダを作成できませんでした: {OUTPUT_DIR} ({e})") from e
 
    try:
        merged.to_csv(MERGED_CSV_PATH, index=False, encoding=OUTPUT_ENCODING)
    except PermissionError as e:
        raise CsvToolError(
            f"{MERGED_CSV_PATH.name} を保存できませんでした。({e})\n"
            f"    → Excelなどでこのファイルを開いたままになっていないか確認してください。"
        ) from e
    except OSError as e:
        raise CsvToolError(f"{MERGED_CSV_PATH.name} を保存できませんでした。({e})") from e
 
    try:
        REPORT_PATH.write_text(report_text, encoding=OUTPUT_ENCODING)
    except PermissionError as e:
        raise CsvToolError(
            f"{REPORT_PATH.name} を保存できませんでした。({e})\n"
            f"    → メモ帳などでこのファイルを開いたままになっていないか確認してください。"
        ) from e
    except OSError as e:
        raise CsvToolError(f"{REPORT_PATH.name} を保存できませんでした。({e})") from e
 
 
# ---------------------------------------------------------------------------
# 6. エラー表示とメイン処理
# ---------------------------------------------------------------------------
 
def print_errors(errors: list[str]) -> None:
    """エラーを番号付きで一覧表示する。"""
    print("", file=sys.stderr)
    print(f"【エラー】{len(errors)}件の問題が見つかったため、処理を中止しました。", file=sys.stderr)
    print("出力ファイル(merged.csv / report.txt)は作成・更新していません。", file=sys.stderr)
    print("", file=sys.stderr)
    for number, message in enumerate(errors, start=1):
        print(f"{number}. {message}", file=sys.stderr)
    print("", file=sys.stderr)
    print("上記を修正してから、もう一度実行してください。", file=sys.stderr)
 
 
def main() -> int:
    """全体の流れ。正常終了なら 0、エラーで中止したら 1 を返す。"""
    print("CSV自動整理・統合ツールを開始します。")
 
    # --- 1. 入力ファイルを探す ---
    try:
        csv_files = find_csv_files()
    except CsvToolError as e:
        print_errors([str(e)])
        return 1
    print(f"対象のCSVファイル: {len(csv_files)}件")
 
    # --- 2. 1ファイルずつ読み込む(エラーはためておき、最後にまとめて表示) ---
    errors: list[str] = []
    loaded_files: list[LoadedCsv] = []
    for path in csv_files:
        try:
            item = load_one_csv(path)
        except CsvToolError as e:
            errors.append(str(e))
            continue
        loaded_files.append(item)
        print(
            f"  読み込み完了: {item.name} "
            f"(データ{len(item.df)}行 / 欠損{item.missing_cells}セル / 文字コード {item.encoding})"
        )
 
    # --- 3. 列名をチェックする ---
    errors.extend(check_columns(loaded_files))
 
    # エラーが1つでもあれば、ここで中止する
    if errors:
        print_errors(errors)
        return 1
 
    # --- 4. 統合して重複を削除する ---
    rows_before = sum(len(item.df) for item in loaded_files)
    missing_cells = sum(item.missing_cells for item in loaded_files)
    merged, duplicates_removed = merge_and_deduplicate(loaded_files)
    rows_after = len(merged)
 
    # --- 5. 保存する ---
    report_text = build_report(
        file_count=len(loaded_files),
        rows_before=rows_before,
        duplicates_removed=duplicates_removed,
        missing_cells=missing_cells,
        rows_after=rows_after,
    )
    try:
        write_outputs(merged, report_text)
    except CsvToolError as e:
        print_errors([str(e)])
        return 1
 
    print("")
    print("処理が完了しました。")
    print(report_text, end="")
    print(f"出力先: {MERGED_CSV_PATH}")
    print(f"        {REPORT_PATH}")
    return 0
 
 
if __name__ == "__main__":
    sys.exit(main())