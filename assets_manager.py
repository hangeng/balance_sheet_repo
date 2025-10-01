#! /usr/bin/env python3.9

import yfinance as yf
import json
import pandas as pd
import os
import schedule
import time
import matplotlib.pyplot as plt
from datetime import datetime
from getpass import getuser
import socket
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes
from curl_cffi import requests

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


def get_stock_price(symbol):
    session = requests.Session(impersonate="chrome")
    stock = yf.Ticker(symbol, session=session)
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

class Asset:
    def __init__(self, symbol, name, amount, price, unit, attr):
        self.symbol = symbol
        self.name = name
        self.amount = amount
        self.initial_price = price
        self.price = price
        self.unit = unit
        self.attr = attr
        self.usd_cny_rate = get_stock_price('CNY=X')

        if self.price == 0:
            self.price = get_stock_price(self.symbol)
            if self.unit == 'USD':
                self.price = self.price * self.usd_cny_rate


    def get_value(self):
        return self.amount * self.price 

    def __str__(self):
        return 'Symbol: {}, Name: {}, Amount: {}, Price: {}, Value: {}'.format(self.symbol, self.name, self.amount, self.price, self.get_value())


class AssetsManager:
    def __init__(self):
        self.json_file = 'assets.json'
        self.assets_db = "assets_db.csv"
        self.assets_curve = "assets_curve.png"
        self.income_assets = []
        self.outcome_assets = []
        self.usd_cny_rate = 0
        self.get_usd_cny_rate()
        self.load_json()
        self.load_assets()

    def get_usd_cny_rate(self):
        self.usd_cny_rate = get_stock_price('CNY=X')

    def load_json(self):
        with open(self.json_file, 'r') as file:
            self.json = json.load(file)

    def load_assets(self):
        for asset in self.json['income']:
            self.income_assets.append(Asset(asset['symbol'], asset['name'], asset['amount'], asset['price'], asset['unit'], asset['attribute']))
        for asset in self.json['outcome']:
            self.outcome_assets.append(Asset(asset['symbol'], asset['name'], asset['amount'], asset['price'], asset['unit'], asset['attribute']))

    def get_total_value(self, assets):
        total_value = 0
        for asset in assets:
            total_value += asset.get_value()
        return total_value

    def get_investment_value(self):
        investment_value = 0
        for asset in self.income_assets:
            if asset.attr == "Investment":
                investment_value += asset.get_value()
        return investment_value

    def show_seperator(self):
        print (self.get_seperator())

    def get_seperator(self):
        return '-'*70

    def show_assets(self):
        print (self.get_assets_text_report())

    def get_assets_text_report(self):
        report_msg = ''
        for assets in [self.income_assets, self.outcome_assets]:
            if assets is self.income_assets:
                report_msg += 'Income Assets:' + '\n'
            else:
                report_msg += self.get_seperator() + '\n'
                report_msg += 'Oncome Assets:' + '\n'

            pd_data = []
            for asset in assets:
                pd_data.append([asset.symbol, asset.name, asset.amount, asset.price, asset.get_value()])

            pd_data.append(['Total', '', '', '', self.get_total_value(assets)])

            df = pd.DataFrame(pd_data, columns=[
                                'Symbol',
                                'Name',
                                'Amount',
                                'Price',
                                'Value'])
            report_msg += str(df) + '\n'


        net_value = self.get_total_value(self.income_assets) - self.get_total_value(self.outcome_assets)
        report_msg += self.get_seperator() + '\n'
        report_msg += 'USD/CNY Rate:     {:.4f}'.format(get_stock_price('CNY=X')) + '\n'
        report_msg += 'Investment Value: {:.2f} ({:.2f}%)'.format(self.get_investment_value(), self.get_investment_value() / net_value * 100) + '\n'
        report_msg += 'Net        Value: {:.2f}'.format(net_value) + '\n'


        report_msg += self.get_seperator() + '\n'

        # show the latest 10 days' assets history
        df = pd.read_csv(self.assets_db)
        report_msg += 'Last 10 days assets history:' + '\n'
        report_msg += str(df.tail(10)) + '\n'
        return report_msg

    def update_assets_db(self):
        if os.path.exists(self.assets_db):
            df = pd.read_csv(self.assets_db)
        else:
            df = pd.DataFrame(columns=['Datetime', 'Income', 'Outcome', 'Investment', 'Investment %', 'Net Value'])

        net_value = self.get_total_value(self.income_assets) - self.get_total_value(self.outcome_assets)
        new_row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                   self.get_total_value(self.income_assets), 
                   self.get_total_value(self.outcome_assets), 
                   self.get_investment_value(), self.get_investment_value() / net_value,
                   self.get_total_value(self.income_assets) - self.get_total_value(self.outcome_assets)]

        df.loc[len(df)] = new_row
        df.to_csv(self.assets_db, index=False)

    def generate_assets_curve(self):
        df = pd.read_csv(self.assets_db)
        x = [datetime.strptime(date_string, "%Y-%m-%d %H:%M:%S").date() for date_string in df['Datetime']]
        # x = [i for i in range(len(df['Datetime']))]
        y1 = df['Income']
        y2 = df['Outcome']
        y3 = df['Investment']
        y4 = df['Net Value']
        # plt.plot(x, y1, label='Income')
        # plt.plot(x, y2, label='Outcome')
        # plt.plot(x, y3, label='Investment')
        plt.plot(x, y4, marker='o', color='b', label='Net Value')

        # Formatting the date on x-axis
        plt.gcf().autofmt_xdate()

        plt.title('Assets Curve')
        plt.grid(True)
        plt.savefig(self.assets_curve)
        plt.close()

    def send_email(self):
        email_sender = EmailSender(sender=sender, receiver=receiver)
        email_sender.send_email(self.get_assets_text_report(), self.assets_curve)


def my_job():
    print("Running my job at 8am")
    am = AssetsManager()
    am.update_assets_db()
    am.generate_assets_curve()
    am.show_assets()
    am.send_email()



if __name__ == '__main__':
    # # Schedule the job to run every day at 8am
    # my_job()

    schedule.every().day.at("20:00").do(my_job)
    # Keep the script running
    while True:
        try:    
            schedule.run_pending()
            time.sleep(30)
        except Exception as e:
            print(e)



