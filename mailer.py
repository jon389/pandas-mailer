from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from io import BytesIO

import pandas as pd
from tzlocal import get_localzone

# TODO https://stackoverflow.com/questions/32808383/formatting-numbers-so-they-align-on-decimal-point
def style_html_table(df: pd.DataFrame) -> str:
    datetime_cols = (df.select_dtypes('datetime').columns.to_list() +
                     [col for col in df.select_dtypes('object').columns
                      if df[col].map(lambda x: type(x).__name__).str.contains('datetime|Timestamp').all()]
                     )

    df_html = (df
               .style

               # set text to display
               .hide_index()
               .format({col: '{:%a %d %b %y %H:%M:%S %Z}' for col in datetime_cols})
               .format({col: lambda x: f'{x:,.8f}'.rstrip('0').rstrip('.') if pd.notna(x) else ''
                        for col in df.select_dtypes('float').columns})

               # CSS classes are attached to the generated HTML
               # Index and Column names include index_name and level<k> where k is its level in a MultiIndex
               # Index label cells include
               #  - row_heading
               #  - row<n> where n is the numeric position of the row
               #  - level<k> where k is the level in a MultiIndex
               # Column label cells include
               #  - col_heading
               #  - col<n> where n is the numeric position of the column
               #  - level<k> where k is the level in a MultiIndex
               # Blank cells include blank
               # Data cells include data

               # note that Outlook only reads the first class (currently th.col_heading and td.data)
               # also, certain properties must be applied at lowest th/td level to render in Outlook
               # https://tatham.blog/2009/07/05/whats-wrong-with-outlook/
               # https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338201(v=office.12)

               # set_table_styles: put css in a <style> tag before the generated HTML table
               .set_table_styles([
                     # non-specific table properties
                     dict(selector='',
                          props=[('mso-table-lspace', '0pt !important'),
                                 ('mso-table-rspace', '0pt !important'),
                                 ('font-family'     , 'Calibri, sans-serif'),
                                 ('font-size'       , '11pt'),
                                 ('color'           , 'black'),
                                 ('text-align'      , 'left'),
                                 ('border-collapse' , 'collapse'),
                                 ('border'          , '1px solid #8ea9db'),  # light blue
                                 ]),

                     # table row properties (includes header), probably redundant
                     dict(selector='tr',
                          props=[('height', '15px')]),

                     # header row, could also select with class=.col_header
                     dict(selector='th',
                          props=[('background', '#4472c4'),  # dark blue
                                 ('color', 'white'),         # font color
                                 ('font-weight', 'bold'),
                                 ('text-align', 'left'),
                                 ('white-space', 'nowrap'),
                                 ('border-bottom', '1px solid #8ea9db'),  # light blue
                                 ('padding', '2px 7px 2px 7px'),
                                 ('height', '15px'),
                                 ]),

                     # data rows, could also select with class=.data
                     dict(selector='td',
                          props=[('white-space', 'nowrap'),
                                 ('border-bottom', '1px solid #8ea9db'),  # light blue
                                 ('padding', '2px 7px 2px 7px'),
                                 ('height', '15px'),
                                 ]),
               ])

               # text-align: right    # for numeric values

               # add items to the opening <table> tag, legacy HTML stuff
               .set_table_attributes('border="0" cellpadding="0" cellspacing="0"')

               # .apply(), .applymap(), and .set_properties() also add css to the <style> tag
               #  before the generated HTML table (like .set_table_styles)
               #  but instead use #id selectors, so are compatible with Outlook

               # .set_properties is for non-data dependent css style properties (just calls applymap)
               #  slicer to select every other data row:
               #           subset=pd.IndexSlice[df.index[::2], :]   # subset=(df.index[::2],) also works
               .set_properties(background='#d9e1f2', subset=pd.IndexSlice[df.index[::2], :])
               #.applymap(lambda x: 'background: #d9e1f2', subset=pd.IndexSlice[df.index[::2], :])

               # right align numeric values, like Excel
               .set_properties(**{'text-align': 'right'}, subset=df.select_dtypes('number').columns)
               # there's no way to apply this to specific th.col_headers in Outlook, since it ignores 2nd+ classes
               #     to right align specific th, would have to do something like adding #id to
               #     re.sub('<th class="col_heading level0 col2">',
               #            '<th id="col_heading_level0_col2" class="col_heading level0 col2">',
               #            df_html)

               .render()
               )

    return df_html


def email_table(df: pd.DataFrame, smtp_svr: str, smtp_user: str, smtp_pass: str, from_email: str, to_email: str):
    msg = MIMEMultipart()
    msg['Subject'] = f'Pandas-Mailer {pd.Timestamp.now(get_localzone()):%d%b %H:%M:%S %Z}'
    msg['From'] = from_email
    msg['To'] = to_email

    df_html = style_html_table(df)

    # provide meta in head to suppress warning in Outlook about how message is displayed
    msg_body = ('''
            <!DOCTYPE html>
            <html>
            <head><meta name="ProgId" content="Word.Document" /></head>
            <body>
            ''' + df_html + '''
            </body>
            </html>
            ''')

    msg.attach(MIMEText(msg_body, 'html'))

    # attach dataframe as an xlsx file
    _xlsx = BytesIO()
    #df['at'] = df['at'].map(lambda x: x.replace(tzinfo=None))  # openpyxl Excel does not support timezones
    df.to_excel(_xlsx, index=False)
    _xlsx.seek(0, 0)
    attach1 = MIMEBase('application', 'octet-stream')
    attach1.set_payload(_xlsx.read())
    encoders.encode_base64(attach1)
    attach1.add_header('Content-Disposition', 'attachment',
                       filename=f'Pandas-Mailer {pd.Timestamp.now(get_localzone()):%Y-%m-%d}.xlsx')
    msg.attach(attach1)

    server = SMTP(smtp_svr, 587)
    server.starttls()
    server.login(smtp_user, smtp_pass)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
