# coding: utf8

import yagmail
import pandas as p

message="Dear sir/madam,"
html = '''
        <html>
    <body>
        <h1 style="text-align: center;">Heading Here</h1>
        <article>
            <p style="text-align: justify; font-size: 14px;font-family: 'Lucida Sans', 'Lucida Sans Regular';">
               Content Body Here
            </p>
        </article>
    </body>
</html>
'''
def SendMail(clientMail,message,html):

    yag = yagmail.SMTP('Your User Id', 'Your Genrated Password')
    to = clientMail
    subject = 'This Is Just Testing Msg'
    body = message
    img = 'img.jpg'
    html = html
    yag.send(to, subject, [body,html, img])

    print('Mail sent.')

def my_function():
    data = p.read_excel("EmailData.xlsx")
    email_col = data.get("Email")
    list_of_email = list(email_col)
    print(list_of_email)
    for mail in list_of_email:
        SendMail(mail,message,html)

my_function()
