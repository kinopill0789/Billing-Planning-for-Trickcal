import os
import pandas as pd

# 実際のExcelファイル名（全角スペース等も含めて正確に書いてください）
target_excel = "database.xlsx"

if not os.path.exists(target_excel):
  print(
      f"【エラー】 '{target_excel}' が見つかりません。"
      "Pythonファイルと同じフォルダにあるか確認してください。"
  )
else:
  try:
    # 1つのExcelファイル内の全シートを読み込む
    sheets_dict = pd.read_excel(target_excel, sheet_name=None)

    # シート名 ➔ JSONファイル名の対応表
    mapping = {
        "特別構成商品": "Special_Bundle.json",
        "エリーフショップ": "Elif_Shop.json",
        "月間パス": "Monthly_Pass.json",
        "エリーフ商品": "Elief_Products.json",
        "週間パック": "Weekly_Pack.json",
        "月間パック": "Monthly_Pack.json",
        "私服・探検パス": "Level_Pass.json",
    }

    for sheet_name, json_filename in mapping.items():
      if sheet_name in sheets_dict:
        # ★ここがポイント：画像のように見出しが4行にわたる場合、
        # header=[0, 1, 2, 3] を指定して正確に項目を読み込ませます。
        # （もしシートによって行数が違う場合は一番下の行を指定します）
        df = pd.read_excel(target_excel, sheet_name=sheet_name, header=3)

        # 列名（ヘッダー）にNaN（空白）が含まれている場合の対策
        df = df.dropna(how="all")

        # JSONファイルとして出力
        df.to_json(json_filename, orient="records", force_ascii=False, indent=2)
        print(f"【成功】 シート「{sheet_name}」 -> {json_filename}")
      else:
        print(
            f"【スキップ】 Excel内に「{sheet_name}」シートが見つかりませんでした"
        )

  except Exception as e:
    print(f"【エラー】 変換中に問題が発生しました: {e}")

input("\n処理が完了しました。Enterキーを押すと終了します。")
