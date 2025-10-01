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



sender="dpsim-bot@dell.com"
receiver="Geng.Han@dell.com"

message_template = """
<html>
<body>
<pre style="font-family:courier;font-size:100%;">
{message}
</pre>
<img src="cid:{image_cid}">
</body>
</html>
"""

def get_share_price(symbol):
    session = requests.Session(impersonate="chrome")
    stock = yf.Ticker(symbol,session=session)
    # return stock.fast_info['last_price']
    return stock.history(period='1d')['Close'].iloc[0]

class EmailSender(object):
    def __init__(self, sender, receiver):
        self.this_user = None
        try:
            self.sender = sender
            self.receivers = [receiver]
            self.this_user = getuser()
            self.hostname = socket.gethostname()
            for remote_hostname in ("mailserver.xiolab.lab.emc.com", "mailhub.lss.emc.com"):
                rc = os.system("ping -c 1 -w 2 {0} >/dev/null 2>&1".format(remote_hostname))
                if rc == 0:
                    self.smtpObj = smtplib.SMTP(remote_hostname, timeout=2)
                    return
            self.smtpObj = None
        except Exception:
            print ("EmailSender: failed to connect to mail server")
            pass

    def send_email(self, message, image_file_name):
        if self.smtpObj is None:
            print ("send_email: failed to connect to mail server")
            return
        try:
            with open(image_file_name, 'rb') as f:
                img_data = f.read()
    
            msg = EmailMessage()
            msg['Subject'] = 'PowerStore dpsim longevity run report'
            msg['From'] = self.sender
            msg['To'] = self.receivers

            msg.set_content(message)

            # now create a Content-ID for the image
            image_cid = make_msgid(domain='dell.com')
            # set an alternative html body
            msg.add_alternative(message_template.format(message=message, image_cid=image_cid[1:-1]), subtype='html')
            # image_cid looks like <long.random.number@xyz.com>
            # to use it as the img src, we don't need `<` or `>`
            # so we use [1:-1] to strip them off

            # now open the image and attach it to the email
            with open(image_file_name, 'rb') as img:
                # know the Content-Type of the image
                maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
            
                # attach it
                msg.get_payload()[1].add_related(img.read(), 
                                                 maintype=maintype,
                                                 subtype=subtype,
                                                 cid=image_cid)
            # the message is ready now
            # you can write it to a file
            # or send it using smtplib
            self.smtpObj.sendmail(self.sender, self.receivers, msg.as_string())
            self.smtpObj.quit()
        except Exception:
            pass

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

    def send_email(self):
        email_sender = EmailSender(sender=sender, receiver=receiver)
        email_sender.send_email(self.get_assets_text_report(), self.balance_sheet_chart)

