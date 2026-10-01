import pandas as pd
from datetime import datetime
from nselib import capital_market
import requests
import json # <-- NAYA: Website ke data ke liye

# === TELEGRAM BOT CREDENTIALS ===
TELEGRAM_BOT_TOKEN = "8912515907:AAHDoxsiwqDoJf6M1OZNIxPD6eUO7Ur2rsE"  
TELEGRAM_CHAT_ID = "795314530" # Apna Chat ID dalein
# ================================

def get_trading_dates():
    now = datetime.now()
    end_date = now if now.hour >= 18 else now - pd.Timedelta(days=1)
    b_days = pd.bdate_range(end=end_date, periods=5)[::-1] 
    
    valid_dates = []
    for date in b_days:
        date_str = date.strftime('%d-%m-%Y')
        try:
            df = capital_market.bhav_copy_with_delivery(date_str)
            if not df.empty:
                valid_dates.append(date_str)
            if len(valid_dates) == 2:
                break
        except Exception:
            pass
            
    if len(valid_dates) < 2:
        raise ValueError("Could not find enough trading data.")
    return valid_dates[0], valid_dates[1]

def send_telegram_message(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}
    try:
        response = requests.post(url, data=payload)
        if response.status_code == 200:
            print("✅ Telegram alert sent successfully!")
        else:
            print(f"❌ Failed to send Telegram message: {response.text}")
    except Exception as e:
        print(f"❌ Telegram Error: {e}")

def run_automated_screener_with_sectors():
    curr_date, prev_date = get_trading_dates()
    print(f"\nRunning Screener for: {curr_date} & {prev_date}\n")
    
    bhav_curr = capital_market.bhav_copy_with_delivery(curr_date)
    bhav_prev = capital_market.bhav_copy_with_delivery(prev_date)
    df_curr = pd.DataFrame(bhav_curr)
    df_prev = pd.DataFrame(bhav_prev)
    
    df_curr = df_curr[df_curr['SERIES'] == 'EQ'].copy()
    df_prev = df_prev[df_prev['SERIES'] == 'EQ'].copy()
    
    num_cols = ['OPEN_PRICE', 'HIGH_PRICE', 'LOW_PRICE', 'CLOSE_PRICE', 'TTL_TRD_QNTY', 'DELIV_QTY']
    for col in num_cols:
        df_curr[col] = pd.to_numeric(df_curr[col].astype(str).str.replace(',', ''), errors='coerce')
        df_prev[col] = pd.to_numeric(df_prev[col].astype(str).str.replace(',', ''), errors='coerce')
        
    df_merged = pd.merge(df_curr, df_prev, on='SYMBOL', suffixes=('_curr', '_prev'))
    
    print("Loading Sector Mapping...")
    try:
        df_sectors = pd.read_csv('nse_sectors.csv') 
        df_merged = pd.merge(df_merged, df_sectors[['SYMBOL', 'Sector']], on='SYMBOL', how='left')
        df_merged['Sector'] = df_merged['Sector'].fillna('Others')
    except FileNotFoundError:
        df_merged['Sector'] = 'Unknown'

    df_merged['Vol_Ratio'] = df_merged['TTL_TRD_QNTY_curr'] / df_merged['TTL_TRD_QNTY_prev']
    df_merged['Del_Ratio'] = df_merged['DELIV_QTY_curr'] / df_merged['DELIV_QTY_prev']
    
    df_merged['Vol_Breakout'] = df_merged['Vol_Ratio'] > 1.5
    df_merged['Del_Breakout'] = df_merged['Del_Ratio'] > 1.5
    
    df_merged['HH_HL'] = (df_merged['HIGH_PRICE_curr'] > df_merged['HIGH_PRICE_prev']) & \
                         (df_merged['LOW_PRICE_curr'] > df_merged['LOW_PRICE_prev'])
                         
    watchlist = df_merged[df_merged['Vol_Breakout'] & df_merged['Del_Breakout'] & df_merged['HH_HL']]
    final_cols = ['SYMBOL', 'Sector', 'CLOSE_PRICE_curr', 'TTL_TRD_QNTY_curr', 'DELIV_QTY_curr', 'Vol_Ratio', 'Del_Ratio']
    final_watchlist = watchlist[final_cols].sort_values(by='Del_Ratio', ascending=False)
    
    print("Calculating Sector-wise Breakouts...")
    sector_summary = df_merged.groupby('Sector').agg(
        Total_TRD_QNTY_curr=('TTL_TRD_QNTY_curr', 'sum'),
        Total_DELIV_QTY_curr=('DELIV_QTY_curr', 'sum'),
        Total_TRD_QNTY_prev=('TTL_TRD_QNTY_prev', 'sum'),
        Total_DELIV_QTY_prev=('DELIV_QTY_prev', 'sum')
    ).reset_index()
    
    sector_summary['Sector_Del_Ratio'] = sector_summary['Total_DELIV_QTY_curr'] / sector_summary['Total_DELIV_QTY_prev']
    sector_summary['Rank_Del'] = sector_summary['Sector_Del_Ratio'].rank(ascending=False)
    sector_summary.sort_values(by='Rank_Del', inplace=True)
    
    output_filename = f"Breakout_Analysis_{curr_date}.xlsx"
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        final_watchlist.to_excel(writer, sheet_name='WatchList', index=False)
        sector_summary.to_excel(writer, sheet_name='Sector_Break', index=False)
        
    # === NEW: WEBSITE KE LIYE JSON DATA SAVE KARNA ===
    print("Generating JSON data for Website...")
    web_data = {
        "last_updated": f"{curr_date} (End of Day)",
        "total_breakouts": len(final_watchlist),
        "top_sectors": sector_summary.head(3)[['Sector', 'Sector_Del_Ratio']].to_dict(orient='records'),
        "stocks": final_watchlist.head(20)[['SYMBOL', 'Sector', 'Vol_Ratio', 'Del_Ratio']].to_dict(orient='records')
    }
    
    with open("data.json", "w") as json_file:
        json.dump(web_data, json_file, indent=4)
    print("✅ data.json generated successfully!")
    # =================================================
    
    print("Generating Telegram Alert...")
    top_sectors = sector_summary.head(3)
    msg = f"📊 <b>MARKET BREAKOUT REPORT ({curr_date})</b>\n\n"
    msg += "🔥 <b>Top 3 Performing Sectors:</b>\n"
    for index, row in top_sectors.iterrows():
        msg += f"• {row['Sector']} (Del: {row['Sector_Del_Ratio']:.2f}x)\n"
        
    msg += "\n🚀 <b>Top 10 Breakout Stocks:</b>\n"
    msg += "<pre>SYMBOL     | SECTOR     | DEL_RATIO</pre>\n" 
    top_stocks = final_watchlist.head(10)
    for index, row in top_stocks.iterrows():
        sym = str(row['SYMBOL'])[:10].ljust(10)
        sec = str(row['Sector'])[:10].ljust(10)
        del_rt = f"{row['Del_Ratio']:.2f}x"
        msg += f"<code>{sym} | {sec} | {del_rt}</code>\n"
        
    send_telegram_message(msg)

if __name__ == "__main__":
    run_automated_screener_with_sectors()
