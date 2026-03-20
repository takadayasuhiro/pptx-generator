import csv
import random
from datetime import datetime, timedelta

def generate_sales_csv(filename="test_data.csv", rows=300):
    categories = ["電子機器", "オフィス用品", "家具", "マーケティング資料"]
    channels = ["ECサイト", "直営店", "法人営業"]
    products = {
        "電子機器": [("ワイヤレスイヤホン", 8500), ("モバイルバッテリー", 3800), ("スマートウォッチ", 12000)],
        "オフィス用品": [("デスクライト", 4200), ("ノートPCスタンド", 3500), ("シュレッダー", 15000)],
        "家具": [("オフィスチェア", 18000), ("昇降式デスク", 45000), ("モニタアーム", 7500)],
        "マーケティング資料": [("販促チラシセット", 500), ("展示パネル", 5000)]
    }

    start_date = datetime(2026, 1, 1)
    
    with open(filename, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(["日付", "カテゴリ", "商品名", "単価", "数量", "売上金額", "チャネル", "在庫数"])
        
        for i in range(rows):
            date = start_date + timedelta(days=i // 4) 
            cat = random.choice(categories)
            prod, base_price = random.choice(products[cat])
            
            unit_price = int(base_price * random.uniform(0.9, 1.1))
            quantity = random.randint(1, 5) + (i // 100) 
            sales = unit_price * quantity
            channel = random.choice(channels)
            stock = random.randint(0, 100)
            
            writer.writerow([date.strftime("%Y-%m-%d"), cat, prod, unit_price, quantity, sales, channel, stock])

    print(f"成功: {filename} を作成しました（{rows}行）")

if __name__ == "__main__":
    generate_sales_csv()
