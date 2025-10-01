#! /usr/bin/env python3.9

import yfinance as yf
import json
import pandas as pd
import os
import time
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
from getpass import getuser
import socket
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes
from curl_cffi import requests
import subprocess
import sys
from pathlib import Path
from typing import Optional

sender = "hgcrhan@gmail.com"
receivers = ["hgcrhan@gmail.com", "sino_han@hotmail.com"]

message_template = """
<html>
<body>
<pre style="font-family:courier;font-size:100%;">
{message}
</pre>
<div style="text-align:center;">
  <img src="cid:{cid1}" alt="Chart" style="max-width:100%; height:auto; display:inline-block;">
</div>
<div style="text-align:center;">
  <img src="cid:{cid2}" alt="Chart" style="max-width:100%; height:auto; display:inline-block;">
</div>
</body>
</html>
"""

def get_share_price(symbol):
    session = requests.Session(impersonate="chrome")
    stock = yf.Ticker(symbol,session=session)
    # return stock.fast_info['last_price']
    return stock.history(period='1d')['Close'].iloc[0]

class EmailSender(object):
    def __init__(self, sender, receivers):
        self.gmail_app_password = "nuim lgxe ptgb hdsj"
        self.sender = sender
        self.receivers = receivers

    def send_email_smtp_gmail(self, message: str, new_balance_sheet_chart: str, legacy_asset_curve: str):
        msg = self.prepare_email_msg(message, new_balance_sheet_chart, legacy_asset_curve)   
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(self.sender, self.gmail_app_password)
            smtp.send_message(msg)

    def prepare_email_msg(self, message: str, new_balance_sheet_chart: str, legacy_asset_curve: str) -> EmailMessage:
        """
        Build an email with a text body and two embedded images, referenced inline via cid:.
    
        Args:
            message: The text/HTML-safe body content (plain text; will be HTML-escaped minimally by wrapping).
            new_balance_sheet_chart: Path to the first image file (e.g., PNG/JPEG).
            legacy_asset_curve: Path to the second image file.
    
        Returns:
            EmailMessage ready to send (set From/To/Subject before sending).
        """

        # Create message container
        msg = EmailMessage()
        msg['Subject'] = 'Daily Assets Report'
        msg['From'] = self.sender
        msg['To'] = self.receivers
        msg.set_content(message)
    
        # Plain text fallback (no images)
        # Keep the plain text simple; some clients prefer/only display this part.
        msg.set_content(message)
    
        # Generate content IDs for the inline images
        cid1 = make_msgid(domain="inline.local")[1:-1]  # strip < >
        cid2 = make_msgid(domain="inline.local")[1:-1]
    
        # Build HTML body that references the images by their content IDs
    
        # Add HTML alternative
        msg.add_alternative(message_template.format(message=message, cid1=cid1, cid2=cid2), subtype="html")
    
        # Helper to attach an image inline to the HTML part
        def _attach_inline_image(filename: str, cid: str):
            path = Path(filename)
            if not path.is_file():
                raise FileNotFoundError(f"Image file not found: {path}")
            ctype, encoding = mimetypes.guess_type(path.name)
            if ctype is None or encoding is not None:
                ctype = "application/octet-stream"
            maintype, subtype = ctype.split("/", 1)
            if maintype != "image":
                raise ValueError(f"Unsupported image type for inline embedding: {ctype} ({path})")
            with open(path, "rb") as f:
                # Attach to the HTML alternative (the last part we added)
                msg.get_payload()[-1].add_related(
                    f.read(),
                    maintype=maintype,
                    subtype=subtype,
                    cid=f"<{cid}>",
                    filename=path.name,
                )
    
        # Attach both images
        _attach_inline_image(new_balance_sheet_chart, cid1)
        _attach_inline_image(legacy_asset_curve, cid2)
    
        return msg

class BalanceSheetItem:
    def __init__(self, ticker_symbol, fullname, positions, share_price, currency_unit, category):
        self.ticker_symbol = ticker_symbol
        self.fullname = fullname
        self.positions = positions
        self.initial_share_price = share_price
        self.share_price = share_price
        self.currency_unit = currency_unit
        self.category = category
        self.usd_and_cny_exchange_rate = get_share_price('CNY=X')

        if self.initial_share_price == 0:
            self.share_price = get_share_price(self.ticker_symbol)
            if self.currency_unit == 'USD':
                self.share_price = self.share_price * self.usd_and_cny_exchange_rate


    def get_book_value(self):
        return self.positions * self.share_price 

    def __str__(self):
        return 'Ticker: {}, Fullname: {}, Positions: {}, Share Price: {}, Book Value: {}'.format(self.ticker_symbol, self.fullname, self.positions, self.share_price, self.get_book_value())


class AssetsManager:
    def __init__(self):
        self.balance_sheet_json_file = './balance_sheet_repo/balancesheet.json'
        self.balance_sheet_db = "./balance_sheet_repo/balancesheet.csv"
        self.balance_sheet_chart = "./balance_sheet_repo/balancesheet.png"
        self.legacy_asset_db = "./balance_sheet_repo/assets_db.csv"
        self.legacy_asset_curve = "./balance_sheet_repo/assets_curve.png"
        self.assets = []
        self.liabilities = []
        self.usd_and_cny_exchange_rate = 0
        self.book_value_per_category = {}
        self.get_usd_and_cny_exchange_rate()
        self.load_balance_sheet_json_file()
        self.get_book_values_per_category()

    def get_usd_and_cny_exchange_rate(self):
        self.usd_and_cny_exchange_rate = get_share_price('CNY=X')

    def load_balance_sheet_json_file(self):
        with open(self.balance_sheet_json_file, 'r') as file:
            self.balance_sheet = json.load(file)


        for key, value in self.balance_sheet.items():
            for item in value:
                balance_sheet_item = BalanceSheetItem(item['ticker symbol'],
                                                    item['fullname'],
                                                    item['positions'],
                                                    item['share price'],
                                                    item['currency unit'],
                                                    item['category'])
                # print (str(balance_sheet_item))
                if key == 'assets':
                    self.assets.append(balance_sheet_item)
                elif key == 'liabilities':
                    self.liabilities.append(balance_sheet_item)

    def get_book_values_per_category(self):
        for asset in self.assets:
            if asset.category in self.book_value_per_category:
                self.book_value_per_category[asset.category] += asset.get_book_value()
            else:
                self.book_value_per_category[asset.category] = asset.get_book_value()


    def get_total_book_value(self, balance_sheet_items):
        total_book_value = 0
        for item in balance_sheet_items:
            total_book_value += item.get_book_value()
        return total_book_value

    def get_investment_value(self):
        investment_value = 0
        for asset in self.assets:
            if asset.category.find("Investment") != -1 or asset.category.find("Mars") != -1:
                investment_value += asset.get_book_value()
        return investment_value

    def show_seperator(self):
        print (self.get_seperator())

    def get_seperator(self):
        return '-'*70

    def show_assets(self):
        print (self.get_assets_text_report())

    def get_assets_text_report(self):
        report_msg = ''
        for balance_sheet_category in [self.assets, self.liabilities]:
            if balance_sheet_category is self.assets:
                report_msg += 'Assets:' + '\n'
            else:
                report_msg += self.get_seperator() + '\n'
                report_msg += 'Liabilities:' + '\n'

            pd_data = []
            for balance_sheet_item in balance_sheet_category:
                pd_data.append([balance_sheet_item.ticker_symbol, balance_sheet_item.fullname, balance_sheet_item.positions, balance_sheet_item.share_price, balance_sheet_item.get_book_value()])

            pd_data.append(['Total', '', '', '', self.get_total_book_value(balance_sheet_category)])

            df = pd.DataFrame(pd_data, columns=[
                                'Ticker',
                                'Fullname',
                                'Positions',
                                'Share Price',
                                'Book Value'])
            report_msg += str(df) + '\n'


        net_value = self.get_total_book_value(self.assets) - self.get_total_book_value(self.liabilities)
        report_msg += self.get_seperator() + '\n'
        report_msg += '{:26s}: {:.4f}'.format("USD/CNY Rate",get_share_price('CNY=X')) + '\n'
        for asset_category in self.book_value_per_category:
            category_book_value = self.book_value_per_category[asset_category]
            report_msg += '{:26s}: {:<10.2f} ({:.2f}%)'.format(asset_category, category_book_value, category_book_value / net_value * 100) + '\n'
        report_msg += '{:26s}: {:<10.2f}'.format('Net Value', net_value) + '\n'

        report_msg += self.get_seperator() + '\n'

        # show the latest 10 days' assets history
        df = pd.read_csv(self.legacy_asset_db)
        report_msg += 'Last 10 days assets history:' + '\n'
        report_msg += str(df.tail(10)) + '\n'
        return report_msg

    def update_balance_sheet_db(self):
        if os.path.exists(self.balance_sheet_db):
            df = pd.read_csv(self.balance_sheet_db)
        else:
            df = pd.DataFrame(columns=['Datetime', 'Type', 'Ticker', 'Fullname', 'Positions', 'Share Price', 'Book Value', 'Category'])

        datetime_string = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for item in self.assets + self.liabilities:
            item_type = 'Assets' if item in self.assets else 'Liabilities'
            new_row = [datetime_string, item_type, item.ticker_symbol, item.fullname, item.positions, item.share_price, item.get_book_value(), item.category]
            df.loc[len(df)] = new_row

        df.to_csv(self.balance_sheet_db, index=False)


    def update_legacy_assets_db(self):
        if os.path.exists(self.legacy_asset_db):
            df = pd.read_csv(self.legacy_asset_db)
        else:
            df = pd.DataFrame(columns=['Datetime', 'Assets', 'Liabilities', 'Investment', 'Investment %', 'Net Value'])

        net_value = self.get_total_book_value(self.assets) - self.get_total_book_value(self.liabilities)
        new_row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                   self.get_total_book_value(self.assets),
                   self.get_total_book_value(self.liabilities),
                   self.get_investment_value(), self.get_investment_value() / net_value,
                   net_value]

        df.loc[len(df)] = new_row
        df.to_csv(self.legacy_asset_db, index=False)


    def generate_legacy_assets_curve(self):
        plt.figure(figsize=(24, 12))   # Set the figure size
        df = pd.read_csv(self.legacy_asset_db)
        x = [datetime.strptime(date_string, "%Y-%m-%d %H:%M:%S").date() for date_string in df['Datetime']]
        # x = [i for i in range(len(df['Datetime']))]
        y1 = df['Assets']
        y2 = df['Liabilities']
        y3 = df['Investment']
        y4 = df['Net Value']
        plt.plot(x, y4, marker='o', color='b', label='Net Value')

        # Formatting the date on x-axis
        plt.gcf().autofmt_xdate()

        plt.title('Assets Curve')
        plt.grid(True)
        plt.savefig(self.legacy_asset_curve)
        plt.close()

    def summarize_balance_sheet_db(self):
        # generate category series
        self.category_series = {}
        df = pd.read_csv(self.balance_sheet_db)
        for row in df.iterrows():
            category = row[1]['Category']
            timestamp = row[1]['Datetime']
            if category in self.category_series:
                if timestamp in self.category_series[category]:
                    self.category_series[category][timestamp] += row[1]['Book Value']
                else:
                    self.category_series[category][timestamp] = row[1]['Book Value']

            else:
                self.category_series[category] = {timestamp: row[1]['Book Value']}

        # generate equity series
        self.net_equity_series = {}
        for row in df.iterrows():
            book_value = row[1]['Book Value'] if row[1]['Type'] == 'Assets' else row[1]['Book Value'] * -1
            if row[1]['Datetime'] in self.net_equity_series:
                self.net_equity_series[row[1]['Datetime']] += book_value
            else:
                self.net_equity_series[row[1]['Datetime']] = book_value

    def generate_balance_sheet_chart(self):
        plt.figure(figsize=(24, 12))   # Set the figure size

        # Plot the category series
        for category in self.category_series:
            x = []
            y = []
            for timestamp in sorted(self.category_series[category].keys()): 
                x.append(datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S"))
                if category == 'Liabilities':
                    y.append(self.category_series[category][timestamp]*-1)
                else:
                    y.append(self.category_series[category][timestamp])
            plt.plot(x, y, marker='o', label=category)

        # Plot the net equity series
        x = []
        y = []
        for timestamp in sorted(self.net_equity_series.keys()):
            x.append(datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S"))
            y.append(self.net_equity_series[timestamp])
        plt.plot(x, y, marker='o', color='black', label="Net Equity")

        # Format the X-axis to show dates  
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))  

        # Set the density of date ticks  
        # plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=3))  # Show every 3rd day

        # Formatting the date on x-axis
        plt.gcf().autofmt_xdate()

        plt.xticks(fontsize=18)  # Set font size for X-axis ticks  
        plt.yticks(fontsize=18)  # Set font size for Y-axis ticks

        plt.title('Assets & Liabilities Chart', fontsize=18, fontweight='bold')
        plt.xlabel('Date', fontsize=18)
        plt.ylabel('Book Value', fontsize=18)
        # Add the legend outside the plot  
        plt.legend(loc='upper left', bbox_to_anchor=(1, 1), fontsize=18)
        plt.subplots_adjust(right=0.83)
        plt.grid(True)
        plt.savefig(self.balance_sheet_chart)
        plt.close()

    def revert_balance_sheet_db(self):
        df = pd.read_csv(self.balance_sheet_db)
        last_row = df.iloc[-1]
        df = df[df['Datetime'] != last_row['Datetime']]
        df.to_csv(self.balance_sheet_db, index=False)

    def revert_legacy_assets_db(self):
        df = pd.read_csv(self.legacy_asset_db)
        last_row = df.iloc[-1]
        df = df[df['Datetime'] != last_row['Datetime']]
        df.to_csv(self.legacy_asset_db, index=False)

    def send_email(self):
        email_sender = EmailSender(sender=sender, receivers=receivers)
        email_sender.send_email_smtp_gmail(self.get_assets_text_report(), self.balance_sheet_chart, self.legacy_asset_curve)

