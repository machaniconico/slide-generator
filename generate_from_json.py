"""
JSONファイルからスライド資料を生成
使い方:
  py generate_from_json.py                          # content.json + ランダムテンプレ
  py generate_from_json.py --template Bold          # テンプレ指定
  py generate_from_json.py --input research.json    # 入力JSON指定
  py generate_from_json.py --output result.pptx     # 出力ファイル指定
"""

import argparse
import json
import sys
import os

# generate_from_template の関数を再利用
from generate_from_template import pick_template, generate_presentation, list_templates, convert_to_pdf, FONT_STYLES, COLOR_PALETTES


def load_content(json_path):
    """JSONファイルを読み込み、tupleが必要な箇所を変換"""
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # JSONではtupleがlistになるので、cards/stats/timeline を tuple に変換
    for section in data.get("sections", []):
        for slide in section.get("slides", []):
            if "cards" in slide:
                slide["cards"] = [tuple(c) for c in slide["cards"]]
            if "stats" in slide:
                slide["stats"] = [tuple(s) for s in slide["stats"]]
            if "timeline" in slide:
                slide["timeline"] = [tuple(t) for t in slide["timeline"]]

    return data


def main():
    parser = argparse.ArgumentParser(description="JSONからスライド資料を生成")
    parser.add_argument("--input", "-i", default="content.json",
                        help="入力JSONファイル (default: content.json)")
    parser.add_argument("--template", "-t", default=None,
                        help="テンプレート名 (部分一致, 省略でランダム)")
    parser.add_argument("--output", "-o", default="output_presentation.pptx",
                        help="出力ファイル名 (default: output_presentation.pptx)")
    parser.add_argument("--style", "-s", default=None,
                        help="フォントスタイル名 (省略でJSONから読取 or ランダム)")
    parser.add_argument("--color", "-c", default=None,
                        help="配色パターン名 (省略でJSONから読取 or ランダム)")
    parser.add_argument("--pdf", action="store_true",
                        help="PDFも同時に出力（PowerPointデスクトップ版が必要）")
    parser.add_argument("--list", action="store_true",
                        help="利用可能なテンプレート・スタイル一覧を表示")
    args = parser.parse_args()

    if args.list:
        print("利用可能なテンプレート:")
        for t in list_templates():
            print(f"  - {t}")
        print(f"\nフォントスタイル（{len(FONT_STYLES)}種）:")
        for s in FONT_STYLES:
            print(f"  - {s['name']}")
        print(f"\n配色パターン（{len(COLOR_PALETTES)}種）:")
        for c in COLOR_PALETTES:
            print(f"  - {c['name']}")
        print(f"\n組み合わせ: {len(FONT_STYLES) * len(COLOR_PALETTES)}通り")
        return

    # JSON読み込み
    json_path = os.path.join("E:/ws/document", args.input)
    if not os.path.exists(json_path):
        # カレントディレクトリも探す
        if os.path.exists(args.input):
            json_path = args.input
        else:
            print(f"エラー: {args.input} が見つかりません")
            sys.exit(1)

    content = load_content(json_path)
    print(f"コンテンツ読み込み: {json_path}")

    # テンプレート選択
    template_path = pick_template(args.template)
    print(f"テンプレート: {os.path.basename(template_path)}")

    # スタイル・配色の決定（優先順: CLI引数 > JSONのstyle/color > ランダム）
    style_name = args.style or content.get("style")
    color_name = args.color or content.get("color")

    # 生成
    output_path = generate_presentation(template_path, content, args.output, style_name, color_name)

    # PDF出力
    if args.pdf:
        convert_to_pdf(output_path)


if __name__ == "__main__":
    main()
