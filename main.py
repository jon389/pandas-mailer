import pandas as pd, numpy as np
import xlwings as xw
import yfinance as yf  # https://github.com/ranaroussi/yfinance

# data = yf.download(tickers='^GSPC ES=F CL=F GBPUSD=X USDCAD=X BTC-USD', period='2d', interval='5m', prepost=True)
# xw.view(data['Close'].sort_index())

# es = yf.Ticker('ES=F')
# es.info.get('shortName')
# >>> 'E-Mini S&P 500 Mar 21'

# stock indices
# '^GSPC'   : 'S&P 500 Index'
# Nasdaq
# FTSE 100 Index
# VMID.L

# Crypto
# 'BTC-USD' : 'Bitcoin'
# 'ETH-USD' : 'Ethereum'

# FX
# 'GBPUSD=X'    : 'GBP/USD'
# 'USDCAD=X'
# 'GBPCAD=X'
# 'USDCNH=X'

# Commodity Futures
#
# 'CL=F'    : 'WTI Oil'
# 'GC=F'    : 'Comex Gold'

# '^TNX'    : '10y Treasury Yield'
# Eurodollar /SOFR futures

# pd.options.display.width = 150
# pd.set_option('display.max_columns', 10)

def get_data() -> pd.DataFrame:
    closes = []
    for t in yf.Tickers(tickers='^GSPC ES=F CL=F GBPUSD=X USDCAD=X BTC-USD').tickers:
        h = t.history(period='1d', interval='30m', prepost=True)

        closes.append(dict(
            **{x: t.info.get(x) for x in 'symbol shortName'.split()},
            price=round(h.Close.iloc[-1], t.info['priceHint']) if 'priceHint' in t.info else h.Close.iloc[-1],
            ccy=t.info.get('currency'),
            at= f'{h.index[-1]:%Y-%m-%d %H:%M:%S%z(%Z)}',
            tzone=h.index[-1].tzinfo.zone if h.index[-1].tzinfo is not None else '',

            # _gmtOffset=int(t.info.get('gmtOffSetMilliseconds', 0))/1000/60/60,

        ))

    return pd.DataFrame(closes)


from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def email_table(smtp_svr:str, smtp_user:str, smtp_pass:str, from_email:str, to_email:str):
    message = MIMEMultipart()
    message['Subject'] = f'Pandas-Mailer {pd.to_datetime("now"):%Y-%m-%d %H:%M:%S}'
    message['From'] = from_email
    message['To'] = to_email

    df = get_data()
    df_html = (df.style
                 .hide_index()
                 .set_properties(**{'font-size': '9pt',
                                    'font-family': 'Calibri'})
                 .render()
    )
    message.attach(MIMEText(df_html, "html"))
    msg_body = message.as_string()

    server = SMTP(smtp_svr, 587)
    server.starttls()
    server.login(smtp_user, smtp_pass)
    server.sendmail(message['From'], message['To'], msg_body)
    server.quit()


from argparse import ArgumentParser
if __name__ == '__main__':
    argp = ArgumentParser()
    run_func = email_table
    for a in run_func.__code__.co_varnames[:run_func.__code__.co_argcount]:
        argp.add_argument(a)
    run_func(**vars(argp.parse_args()))

# Program/script: %ComSpec%
# Add arguments:  /c start "" /min %ComSpec% /c "%LOCALAPPDATA%\anaconda3\condabin\conda.bat activate .\envs & python main.py smtp_svr... "
# Start in: <pandas-mailer directory>