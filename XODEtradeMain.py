import pandas as pd
from openpyxl import load_workbook
from forex_python.converter import CurrencyRates
from pycoingecko import CoinGeckoAPI
import uuid
import os
import time

EXCEL_FILE = 'TradeLedger.xlsx'

def ensure_excel_exists():
    if not os.path.exists(EXCEL_FILE):
        # Initiate Excel file and required sheets
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            users_df = pd.DataFrame(columns=['UID', 'Name', 'Email'])
            users_df.to_excel(writer, sheet_name='Users', index=False)
            balance_df = pd.DataFrame(columns=[
                'UID', 'INR', 'USD', 'BTC', 'ETH', 'NFT'])
            balance_df.to_excel(writer, sheet_name='Balances', index=False)

def get_user_df():
    return pd.read_excel(EXCEL_FILE, sheet_name='Users')

def get_balance_df():
    return pd.read_excel(EXCEL_FILE, sheet_name='Balances')

def save_user_df(df):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Users', index=False)

def save_balance_df(df):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Balances', index=False)

def create_user(name, email):
    users = get_user_df()
    user_uid = str(uuid.uuid4())
    # Add user to Users sheet
    users = users.append({'UID': user_uid, 'Name': name, 'Email': email}, ignore_index=True)
    save_user_df(users)
    # Add user with zero balances to Balances sheet
    balances = get_balance_df()
    balances = balances.append({'UID': user_uid, 'INR': 0, 'USD': 0, 'BTC': 0, 'ETH': 0, 'NFT': 0}, ignore_index=True)
    save_balance_df(balances)
    print(f"User created. UID: {user_uid}")
    return user_uid

def add_money(uid, currency, amount):
    balances = get_balance_df()
    if uid not in balances['UID'].values:
        print("UID not found.")
        return
    balances.loc[balances['UID'] == uid, currency] += amount
    save_balance_df(balances)

def withdraw_money(uid, currency, amount):
    balances = get_balance_df()
    if uid not in balances['UID'].values:
        print("UID not found.")
        return
    oldval = balances.loc[balances['UID'] == uid, currency].values[0]
    if oldval < amount:
        print("Insufficient funds.")
        return
    balances.loc[balances['UID'] == uid, currency] -= amount
    save_balance_df(balances)

def get_conversion_rates():
    # Fiat
    forex = CurrencyRates()
    inr_to_usd = forex.get_rate('INR', 'USD')
    usd_to_inr = forex.get_rate('USD', 'INR')
    # Crypto
    cg = CoinGeckoAPI()
    cg_prices = cg.get_price(ids=['bitcoin', 'ethereum'], vs_currencies=['usd', 'inr'])
    btc_usd = cg_prices['bitcoin']['usd']
    btc_inr = cg_prices['bitcoin']['inr']
    eth_usd = cg_prices['ethereum']['usd']
    eth_inr = cg_prices['ethereum']['inr']
    # NFT placeholders (since NFT values are not universal)
    nft_usd = 100  # Fake value for demo
    nft_inr = nft_usd * usd_to_inr
    return {
        'inr_to_usd': inr_to_usd,
        'usd_to_inr': usd_to_inr,
        'btc_usd': btc_usd,
        'btc_inr': btc_inr,
        'eth_usd': eth_usd,
        'eth_inr': eth_inr,
        'nft_usd': nft_usd,
        'nft_inr': nft_inr,
    }

def convert(uid, from_curr, to_curr, amount):
    rates = get_conversion_rates()
    balances = get_balance_df()
    if uid not in balances['UID'].values:
        print("UID not found.")
        return
    if balances.loc[balances['UID'] == uid, from_curr].values[0] < amount:
        print("Insufficient funds.")
        return

    # Determine conversion
    curr_map = {'INR': 'inr', 'USD': 'usd', 'BTC': 'btc', 'ETH': 'eth', 'NFT': 'nft'}
    if from_curr == to_curr:
        print("Can't convert same currency.")
        return

    # Handle INR <-> USD
    if (from_curr == 'INR' and to_curr == 'USD'):
        converted = amount * rates['inr_to_usd']
    elif (from_curr == 'USD' and to_curr == 'INR'):
        converted = amount * rates['usd_to_inr']
    # Handle INR/USD to Crypto
    elif (from_curr == 'INR' and to_curr == 'BTC'):
        converted = amount / rates['btc_inr']
    elif (from_curr == 'INR' and to_curr == 'ETH'):
        converted = amount / rates['eth_inr']
    elif (from_curr == 'USD' and to_curr == 'BTC'):
        converted = amount / rates['btc_usd']
    elif (from_curr == 'USD' and to_curr == 'ETH'):
        converted = amount / rates['eth_usd']
    # Handle Crypto to INR/USD
    elif (from_curr == 'BTC' and to_curr == 'INR'):
        converted = amount * rates['btc_inr']
    elif (from_curr == 'BTC' and to_curr == 'USD'):
        converted = amount * rates['btc_usd']
    elif (from_curr == 'ETH' and to_curr == 'INR'):
        converted = amount * rates['eth_inr']
    elif (from_curr == 'ETH' and to_curr == 'USD'):
        converted = amount * rates['eth_usd']
    # Handle to/from NFT (demo conversion, you can modify to real APIs)
    elif (to_curr == 'NFT'):
        if from_curr == 'INR':
            converted = amount / rates['nft_inr']
        elif from_curr == 'USD':
            converted = amount / rates['nft_usd']
        elif from_curr == 'BTC':
            converted = (amount * rates['btc_usd']) / rates['nft_usd']
        elif from_curr == 'ETH':
            converted = (amount * rates['eth_usd']) / rates['nft_usd']
    elif (from_curr == 'NFT'):
        if to_curr == 'INR':
            converted = amount * rates['nft_inr']
        elif to_curr == 'USD':
            converted = amount * rates['nft_usd']
        elif to_curr == 'BTC':
            converted = (amount * rates['nft_usd']) / rates['btc_usd']
        elif to_curr == 'ETH':
            converted = (amount * rates['nft_usd']) / rates['eth_usd']
    else:
        print("Unsupported conversion.")
        return

    # Deduct and add to balances
    balances.loc[balances['UID'] == uid, from_curr] -= amount
    balances.loc[balances['UID'] == uid, to_curr] += converted
    save_balance_df(balances)
    print(f"Converted {amount:.4f} {from_curr} -> {converted:.4f} {to_curr}")

def view_balance(uid):
    balances = get_balance_df()
    users = get_user_df()
    u = users[users['UID'] == uid]
    if u.empty:
        print("User not found.")
        return
    b = balances[balances['UID'] == uid]
    if b.empty:
        print("User balance not found.")
        return
    b = b.iloc[0]
    print(f"User: {u['Name'].values[0]} (UID: {uid})")
    print(f"  INR: {b['INR']:.2f}")
    print(f"  USD: {b['USD']:.2f}")
    print(f"  BTC: {b['BTC']:.6f}")
    print(f"  ETH: {b['ETH']:.6f}")
    print(f"  NFT: {b['NFT']:.2f}")
    # Show converted rates for convenience:
    rates = get_conversion_rates()
    total_usd = b['INR']/rates['inr_to_usd'] + b['USD'] + b['BTC']*rates['btc_usd'] + b['ETH']*rates['eth_usd'] + b['NFT']*rates['nft_usd']
    print(f"  Approx Total USD Value: {total_usd:.2f}")

def main_cli():
    ensure_excel_exists()
    print("Welcome to Option Index Trade Program")
    while True:
        print("""
1. New User
2. Add Money
3. Withdraw Money
4. Convert Currency
5. View Balance
6. Exit
        """)
        ch = input("Select option: ").strip()
        if ch == '1':
            name = input("Enter Name: ")
            email = input("Enter Email: ")
            create_user(name, email)
        elif ch == '2':
            uid = input("Enter UID: ")
            currency = input("Currency (INR/USD/BTC/ETH/NFT): ").upper()
            amount = float(input("Amount: "))
            add_money(uid, currency, amount)
        elif ch == '3':
            uid = input("Enter UID: ")
            currency = input("Currency (INR/USD/BTC/ETH/NFT): ").upper()
            amount = float(input("Amount: "))
            withdraw_money(uid, currency, amount)
        elif ch == '4':
            uid = input("Enter UID: ")
            from_curr = input("From Currency (INR/USD/BTC/ETH/NFT): ").upper()
            to_curr = input("To Currency (INR/USD/BTC/ETH/NFT): ").upper()
            amount = float(input("Amount: "))
            convert(uid, from_curr, to_curr, amount)
        elif ch == '5':
            uid = input("Enter UID: ")
            view_balance(uid)
        elif ch == '6':
            print("Exiting. Bye!")
            break
        else:
            print("Invalid choice.")

if __name__ == "__main__":
    main_cli()
