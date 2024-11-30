import requests
import pandas as pd
import xlwings as xw
import time

# Fetch live data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"API request failed with status code {response.status_code}")

# Analyze the data
def analyze_data(data):
    df = pd.DataFrame(data)
    df = df[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
    
    top_5 = df.nlargest(5, "market_cap")[["name", "market_cap"]]
    avg_price = df["current_price"].mean()
    highest_change = df.nlargest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
    lowest_change = df.nsmallest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
    
    return df, top_5, avg_price, highest_change, lowest_change

# Update Excel file
def update_excel(df):
    # Open Excel workbook
    wb = xw.Book("live_crypto_data.xlsx")
    sheet = wb.sheets["Crypto Data"]
    
    # Clear previous data
    sheet.clear_contents()
    
    # Write the updated DataFrame
    sheet.range("A1").value = df
    
    wb.save()
    print("Excel updated successfully.")

# Write the analysis report
def write_report(top_5, avg_price, highest_change, lowest_change):
    with open("analysis_report.txt", "w") as f:
        f.write("Crypto Analysis Report\n")
        f.write("======================\n\n")
        f.write("Top 5 Cryptocurrencies by Market Cap:\n")
        f.write(top_5.to_string(index=False))
        f.write("\n\nAverage Price of Top 50 Cryptocurrencies: ${:.2f}".format(avg_price))
        f.write("\n\nHighest 24-hour Percentage Change:\n")
        f.write(highest_change.to_string(index=False))
        f.write("\n\nLowest 24-hour Percentage Change:\n")
        f.write(lowest_change.to_string(index=False))

# Main function
if __name__ == "__main__":
    # Ensure an Excel file exists
    file_name = "live_crypto_data.xlsx"
    try:
        xw.Book(file_name)
    except:
        # Create the Excel file if it doesn't exist
        wb = xw.Book()
        wb.sheets.add("Crypto Data")
        wb.save(file_name)
        wb.close()
    
    while True:
        try:
            data = fetch_crypto_data()
            df, top_5, avg_price, highest_change, lowest_change = analyze_data(data)
            
            update_excel(df)
            write_report(top_5, avg_price, highest_change, lowest_change)
            
            print("Data updated. Waiting for next update...")
            time.sleep(300)  # 5-minute interval
        except Exception as e:
            print(f"An error occurred: {e}")
