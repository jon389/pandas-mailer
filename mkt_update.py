from zoneinfo import ZoneInfo
from datetime import datetime, time, timedelta
from tzlocal import get_localzone

import pandas as pd, numpy as np
from pandas.tseries.holiday import USFederalHolidayCalendar
bday_us = pd.offsets.CustomBusinessDay(calendar=USFederalHolidayCalendar())

import xlwings as xw
from log_conf import logger as log
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


def get_iday_times(now) -> dict:

    # for datetime.combine, if date=datetime, its time components and tzinfo attributes ignored
    iday_times = {
        'nyc_close': datetime.combine(now - timedelta(days=1), time(17, 0, 0, tzinfo=ZoneInfo('America/New_York'))),
        # 'asia_open': datetime.combine(now, time(8, 0, 0, tzinfo=ZoneInfo('Asia/Tokyo'))),
        'lndn_open': datetime.combine(now, time(8, 0, 0, tzinfo=ZoneInfo('Europe/London'))),
        'ldn_close': datetime.combine(now, time(16, 0, 0, tzinfo=ZoneInfo('Europe/London'))),
    }

    for ts in iday_times:
        if (iday_times[ts] - now).total_seconds() > -55 * 60:  # is within 55min or after, so shift day earlier
            iday_times[ts] -= timedelta(days=1)
        if iday_times[ts].weekday() >= 5:  # is Sat/Sun, so shift to prior Friday
            iday_times[ts] -= timedelta(days=iday_times[ts].weekday() - 4)  # if Sat, weekday=5, Sun=6

    # sort by datetime
    # noinspection PyTypeChecker
    iday_times = dict(sorted(iday_times.items(), key=lambda item: item[1]))

    # ensure all times before now
    assert all((iday_times[dt] - now).total_seconds() < 0 for dt in iday_times)

    return iday_times

def get_data() -> pd.DataFrame:
    yf_tickers = ('^GSPC ES=F ^TNX CL=F GC=F '
                  'GBPUSD=X GBPEUR=X USDCAD=X '
                  'BTC-USD ETH-USD '
                  'VUSA.L VFEM.L VUTY.L'
                  )
    yf_mkt_data = []

    now = datetime.now(tz=get_localzone())
    mnth_ago = (now - pd.offsets.DateOffset(months=1)).date() - pd.offsets.BDay(1)   # .date() chops off timepart
    iday_times = get_iday_times(now)

    log.info(f'loading yf tickers {yf_tickers}')
    for t in yf.Tickers(tickers=yf_tickers).tickers.values():
        try:
            log.info(f'loading history for {t.info["symbol"]}')
            iday = t.history(period='5d', interval='1h', prepost=True)
            mnth = t.history(start=mnth_ago, interval='1d')
            yf_mkt_data.append((t, iday, mnth,))
        except Exception as e:
            log.exception(f'error loading {t.info["symbol"]}')

    # find common month ago date
    mnth_ago = min(set.intersection(*[set(mnth.index) for t, iday, mnth in  yf_mkt_data]))

    # parse loaded data
    parsed_yf_data = []
    for t, iday, mnth in yf_mkt_data:
        try:
            def pHint_round(num):
                return round(num, t.info.get('priceHint', 2)) if pd.notna(num) else num

            iday.index = iday.index.tz_convert(now.tzinfo)
            iday['IntradayPeriod'] = pd.cut(
                iday.index,
                bins=[iday_times[dt].astimezone(now.tzinfo) for dt in iday_times] + [now + timedelta(days=1)],
                labels=[f'{iday_times[dt].astimezone(now.tzinfo):%I%p %d%b%y}'.lstrip('0') for dt in iday_times],
                right=False,
            )
            iday_periods = iday.groupby('IntradayPeriod').agg(Open=('Open', 'first'),
                                                              High_since=('High', 'max'),
                                                              Low_since=('Low', 'min'))

            iday_gaps = iday_periods.stack().reset_index().assign(
                new_head=lambda d: d.IntradayPeriod.astype(str).where(d.level_1 == 'Open',
                                                                      d.level_1.str.cat(
                                                                          d.IntradayPeriod.astype(str).str.split(
                                                                              expand=True)[0]))
            ).set_index('new_head')[0]

            parsed_yf_data.append(dict(
                #**{x: t.info.get(x) for x in 'symbol shortName'.split()},
                symbol=t.info.get('symbol'),
                shortName=(t.info.get('shortName')
                           if not t.info.get('shortName', '').startswith('VANGUARD FUNDS PLC')
                           else (t.info.get('longName', '')[:t.info.get('longName', '').index(' UCITS')]
                                    ).replace('Vanguard Funds Public Limited Company - ', '')
                           ),
                price=pHint_round(iday.Close.iloc[-1]),
                ccy=t.info.get('currency', ''),
                **{f'{now.tzname()} - {now.tzinfo.zone}': f'{iday.index[-1].astimezone(now.tzinfo):%a %d%b%y %H:%M:%S}',
                   },
                #tzone=iday.index[-1].tzinfo.zone if iday.index[-1].tzinfo is not None else '',
                # _gmtOffset=int(t.info.get('gmtOffSetMilliseconds', 0))/1000/60/60,
                #**iday_gaps.map(pHint_round).to_dict(),
                **{
                    'prevClose': pHint_round(t.info.get('previousClose')),
                    #'mth_ago': mnth_ago,  # mnth.index[0],
                    '-1mth ' f'{mnth_ago:%d%b%y}'.lstrip('0'):
                                    pHint_round(mnth.loc[mnth_ago].Close),  # mnth.iloc[0].Close,

                    'high_since'    : pHint_round(mnth.High.max()),
                    'high_when'     : f'{mnth.High.idxmax():%d%b%y}',
                    'low_since'     : pHint_round(mnth.High.min()),
                    'low_when'      : f'{mnth.High.idxmin():%d%b%y}',
                    }
            ))
        except Exception as e:
            log.exception(f'error parsing data for {t.info["symbol"]}')

    return (pd.DataFrame(parsed_yf_data)
              # .assign(at=lambda d: pd.to_datetime(d['at'], utc=True))
            )[parsed_yf_data[0].keys()]


if __name__ == '__main__':
    from mailer import email_table
    def run_func(smtp_svr: str, smtp_user: str, smtp_pass: str, from_email: str, to_email: str):
        email_table(get_data(), smtp_svr, smtp_user, smtp_pass, from_email, to_email)

    from argparse import ArgumentParser
    argp = ArgumentParser()
    for a in run_func.__code__.co_varnames[:run_func.__code__.co_argcount]:
        argp.add_argument(a)

    try:
        run_func(**vars(argp.parse_args()))
    except Exception as e:
        log.exception()

# Program/script: %ComSpec%
# Add arguments:  /c start "" /min %ComSpec% /c "%LOCALAPPDATA%\anaconda3\condabin\conda.bat activate .\envs & python mkt_update.py smtp_svr... "
# Start in: <pandas-mailer directory>