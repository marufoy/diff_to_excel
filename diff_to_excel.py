import os
import difflib
import imgkit
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image
from PIL import Image as PILImage

# 設定エリア
#比較元のベースとなるルートディレクトリ（環境ごとの大元）
DIR1_BASE = os.path.normpath(r"C:\Users\kkarai\Desktop\compare")
DIR2_BASE = os.path.normpath(r"C:\Users\kkarai\Desktop\compare2")

#比較したい「フォルダ」「ファイル」をリストで指定
#フォルダも単一ファイルもごちゃ混ぜ指定可能
#大元を丸ごと再起チェックしたいとき、リストを空にする

TARGET_SUBDIRS = [
    #rsrc/folder1,
    # rsrc/folder2,
    # rsrc/file2.txt
    ]

OUTPUT_EXCEL = "diff_report.xlsx" #出力するExcelファイル名

#wkhtmktoimageのパスを設定（wkhtmltoimage.exeのパス）
WKHTMLTOIMAGE_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe"

def create_diff_html(file_path, file2_path, rel_path):
    """
    2つファイルの差分をHTMLで出力する関数
    """
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f1, \
         open(file2_path, "r", encoding="utf-8", errors="ignore") as f2:
        lines1 = f1.readlines()
        lines2 = f2.readlines()

    if lines1 == lines2:
        return None

    differ = difflib.HtmlDiff()
    html_content = differ.make_file(
        lines1, lines2,
        fromdesc=f"Dir1:{rel_path}",
        todesc=f"Dir2:{rel_path}",
    )

    # 1. difflib固有の折り返しを削除
    html_content = html_content.replace('nowrap', '')

    # 2. 独自に折り返し
    custom_style = """
    <style>
    body {
        margin: 15px;
        background-color: #ffffff;
        font-family: 'Courier New', Courier, monospace;
        }
        table.diff {
        width: 100% !important;
        }
        td {
        word-break: break-all !important;
        white-space: pre-wrap !important;
        }
        .diff_header {
        min-width: 40px !important;
        text-align: right !important;
        padding-right: 5px !important;

    }
    </style>
    """

    if "<body>" in html_content:
        html_content = html_content.replace("<body>", f"{custom_style}\n<body>")
    else:
        html_content += custom_style

    return html_content

def process_single_file(file1_path, filename, DIR1_BASE, DIR2_BASE, temp_dir, config, options, wb, searched_directories, diff_files, sheet_mapping):
    """見つかった1つのファイルを処理し、差分があればExcelに追加する関数"""
    #dir1_baseからの相対パスを計算
    rel_path = os.path.relpath(file1_path, DIR1_BASE)
    file2_path = os.path.join(DIR2_BASE, rel_path)

    path_parts = rel_path.split(os.sep)
    if path_parts:
        searched_directories.add(path_parts[0])

    if not os.path.exists(file2_path):
        print(f"【スキップ】{rel_path}はディレクトリ2側に存在しません。")
        return

    html_diff = create_diff_html(file1_path, file2_path, rel_path)
    if html_diff is None:
        print(f" 【一致】{rel_path}は一致しているためスルーします")
        return

    print(f"【差分あり】{rel_path}のスクショを生成中・・・")
    diff_files.append(rel_path)

    safe_temp_name = rel_path.replace(os.sep, "_")
    html_path = os.path.join(temp_dir, f"{safe_temp_name}.html")
    img_path = os.path.join(temp_dir, f"{safe_temp_name}.jpg")

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_diff)

    try:
        imgkit.from_file(html_path, img_path, config=config, options=options)

        if os.path.exists(img_path):
            with PILImage.open(img_path) as pil_img:
                new_size = (int(pil_img.width * 0.75), int(pil_img.height * 0.75))
                resized_img = pil_img.resize(new_size, PILImage.Resampling.LANCZOS)
                resized_img.save(img_path, "JPEG", quality=85, optimize=True)

    except Exception as e:
        if not os.path.exists(img_path):
            print(f" 【エラー】{rel_path}のスクショ生成に失敗しました。: {e}")
            if rel_path in diff_files:
                diff_files.remove(rel_path)
            return

    clean_name = filename
    for char in ["\\", "/", ":", "*", "?", "<", ">", "|", '"', "'", "[", "]"]:
        clean_name = clean_name.replace(char, "")

    sheet_title = clean_name[:31]
    counter = 1
    base_title = clean_name[:28]
    while sheet_title in wb.sheetnames:
        suffix = f"_{counter}" if counter > 1 else ""
        sheet_title = f"{base_title}{suffix}"
        counter += 1

    #目次用のマッピングデーターを記憶
    sheet_mapping.append({"rel_path": rel_path, "sheet_title": sheet_title})

    #差分シートの作成
    ws = wb.create_sheet(title=sheet_title)

    #グリッド線を非表示に
    ws.sheet_view.showGridLines = False

    ws.row_dimensions[1].height = 25
    ws["B2"] = f"対象ファイル: {rel_path}"
    ws["B2"].font = Font(color="FFFFFF", bold=True, size=11)

    ws["B2"].fill = PatternFill(start_color="2c3e50", end_color="2c3e50", fill_type="solid")
    ws["B2"].border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    ws["B2"].alignment = Alignment(vertical="center")

    ws.column_dimensions["B"].width = 120

    img = Image(img_path)
    img.width = int(img.width * 1.1)
    img.height = int(img.height * 1.1)
    ws.add_image(img, "B4")

def main():
    config = imgkit.config(wkhtmltoimage=WKHTMLTOIMAGE_PATH)

    options = {
        'format': 'jpg',
        'quality': 85,
        'width': 2000,
        'encoding': 'UTF-8',
        'quiet': ''
    }

    wb = openpyxl.Workbook()

    #最初のシートを目次専用ページにする
    index_ws = wb.active
    index_ws.title = "目次"
    index_ws.sheet_view.showGridLines = False

    #目次ヘッダーデザイン
    index_ws["B2"] = "差分レポートの目次"
    index_ws["B2"].font = Font(color="333333", bold=True, size=16)
    index_ws["B4"] = "No"
    index_ws["C4"] = "対象ファイルパス"
    index_ws["D4"] = "シートリンク"

    #目次テーブルのヘッダー色
    header_fill = PatternFill(start_color="34495e", end_color="34495e", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True,)
    for col in ["B4", "C4", "D4"]:
        index_ws[col].fill = header_fill
        index_ws[col].font = header_font
        index_ws[col].alignment = Alignment(vertical="center")

    index_row = 5

    temp_dir = "temp_diff"
    os.makedirs(temp_dir, exist_ok=True)

    searched_directories = set()
    diff_files = []
    sheet_mapping = []

    print("【ファイル比較開始】")

    #もしリストがからなら、大元のルートディレクトリそのまま対象とする
    current_targets = TARGET_SUBDIRS if TARGET_SUBDIRS else [""]

    #再起ループ処理
    for subdir in current_targets:
        dir1_target = os.path.normpath(os.path.join(DIR1_BASE, subdir))
        dir2_target = os.path.normpath(os.path.join(DIR2_BASE, subdir))

        if subdir:
            print(f"ターゲット探索中: {subdir}")
        else:
            print(f"大元のルート丸ごとチェック中・・・")

        if not os.path.exists(dir1_target):
            print(f"【スキップ】{dir1_target}は存在しません。")
            continue

        #ファイル単体の場合の処理
        if os.path.isfile(dir1_target):
            filename = os.path.basename(dir1_target)
            process_single_file(dir1_target, filename, DIR1_BASE, DIR2_BASE, temp_dir, config, options, wb, searched_directories, diff_files, sheet_mapping)
            continue

        #フォルダの場合の処理
        for root, dirs, files in os.walk(dir1_target):
            for filename in files:
                file1_path = os.path.join(root, filename)
                process_single_file(file1_path, filename, DIR1_BASE, DIR2_BASE, temp_dir, config, options, wb, searched_directories, diff_files, sheet_mapping)

    if not sheet_mapping:
        print("【差分があるファイルはすべての環境で一つもありませんでした。】")
        index_ws["B4"] = "差分があるファイルはすべての環境で一つもありませんでした。"
    else:
        thin_border = Border(
            left=Side(style="thin", color="DDDDDD"),
            right=Side(style="thin", color="DDDDDD"),
            top=Side(style="thin", color="DDDDDD"),
            bottom=Side(style="thin", color="DDDDDD"),
        )

        for i, item in enumerate(sheet_mapping, 1):
            index_ws[f"B{index_row}"] = i
            index_ws[f"C{index_row}"] = item["rel_path"]
            index_ws[f"D{index_row}"] = "シートへ移動→"

            index_ws[f"D{index_row}"].hyperlink = f"#'{item['sheet_title']}'!A1"

            index_ws[f"B{index_row}"].alignment = Alignment(horizontal="center")
            index_ws[f"D{index_row}"].font = Font(color="1b65b9", underline="single", bold=True)
            index_ws[f"D{index_row}"].alignment = Alignment(horizontal="center")
            index_ws[f"C{index_row}"].font = Font(color="333333")

            for col in ["B", "C", "D"]:
                index_ws[f"{col}{index_row}"].border = thin_border

            index_row += 1

    #目次シートの列幅を文字数に合わせて自動調整
    index_ws.column_dimensions["B"].width = 6
    index_ws.column_dimensions["C"].width = 80
    index_ws.column_dimensions["D"].width = 18

    #不要なデフォルトの最初のシートが残っていたら削除する
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    wb.save(OUTPUT_EXCEL)
    print(f"【差分レポートを{OUTPUT_EXCEL}に保存しました。】")

    print("\n" + "="*40)
    print("実際にチェックしたディレクトリ一覧")
    print("="*40 + "\n")

    if searched_directories:
        for d in sorted(searched_directories):
            print(f"- {d}")
    else:
        print("チェックしたディレクトリはありませんでした。パスを確認してください")

    file_count = len(diff_files)
    print(f"差分があるファイルは{file_count}件です。")
    if file_count > 0:
        print(f"差分があるファイル一覧: {', '.join(diff_files)}")
    else:
        print("差分があるファイルはありませんでした。パスを確認してください")

    print("\n" + "="*40)
    print("差分レポートの詳細な内容は、Excelファイルを開いてご確認ください。")
    print("="*40 + "\n")

    if os.path.exists(temp_dir):
        for f in os.listdir(temp_dir):
            os.remove(os.path.join(temp_dir, f))
        os.rmdir(temp_dir)

    print("【ファイル比較処理が完了しました。】")

if __name__ == "__main__":
    main()
