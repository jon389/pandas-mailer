# pandas-mailer

practice emailing nice pandas tables

[pretty_html_table](https://dev.to/siddheshshankar/convert-a-dataframe-into-a-pretty-html-table-and-send-it-over-email-4663)
is not necessary, given [pandas styles](https://pandas.pydata.org/pandas-docs/stable/user_guide/style.html)


[email example](https://stackoverflow.com/a/50566309)
```python
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import smtplib
import sys


recipients = ['ToEmail@domain.com'] 
emaillist = [elem.strip().split(',') for elem in recipients]
msg = MIMEMultipart()
msg['Subject'] = "Your Subject"
msg['From'] = 'from@domain.com'


html = """\
<html>
  <head></head>
  <body>
    {0}
  </body>
</html>
""".format(df_test.to_html())

part1 = MIMEText(html, 'html')
msg.attach(part1)

server = smtplib.SMTP('smtp.gmail.com', 587)
server.sendmail(msg['From'], emaillist , msg.as_string())
```


### Also practice getting intraday OHLC market data
https://gist.github.com/jon389/315b16e143fe550af10feff19aa9696f

[yfinance](https://github.com/ranaroussi/yfinance)

![image](https://user-images.githubusercontent.com/24356268/110604296-d0de6e80-817f-11eb-9be4-e8c56d682756.png)
