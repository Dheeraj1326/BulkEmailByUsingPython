# coding: utf8

import yagmail
import pandas as p

message="Dear sir/madam,"
html = '''
        <html>
    <body>
        <h1 style="text-align: center;">📣 Get ready for a limited time offer that's too good to miss!</h1>
        <article>
            <p style="text-align: justify; font-size: 14px;font-family: 'Lucida Sans', 'Lucida Sans Regular';">
                😱🔥💰 Introducing the C & C++ Limited Offer for only 999/- 
                💻💲Have you always wanted to master the programming languages of C & C++?
                Now is your chance! 🚀 Whether you're a beginner or looking to level up your coding skills, 
                this offer is tailor-made for you.With C & C++ being the foundation of many software applications, 
                learning these languages opens up a world of opportunities for you in the tech industry. 
                🌐💡Imagine being able to develop your own software, create impressive applications,or even contribute to open-source projects. The possibilities are endless! 
                💪💻But wait, that's not all! 🎉 When you grab this limited offer, 
                you'll also get access to:
                <p style="font-size: 14px;font-family: 'Lucida Sans', 'Lucida Sans Regular';"> 1️⃣     Comprehensive video tutorials taught by industry experts. 🎥👨‍🏫<br>
                    2️⃣     Exciting coding challenges to put your skills to the test. 💪💡<br>
                    3️⃣     A supportive online community of fellow learners to connect with. 👥💬<br>
                    4️⃣     An official certification upon completion. 🎓<br>
                       ✅ Don't miss out on this incredible opportunity to boost your programming prowess and future-proof your career. <br>
                🚀🌟So what are you waiting for? </p>
                Click the link below to avail this limited offer: https://bit.ly/stechnosolutions 
                Let's dive into the world of C & C++ together and level up our coding game! 🙌💻
            </p>
        </article>
    </body>
</html>
'''
def SendMail(clientMail,message,html):

    yag = yagmail.SMTP('Dheeraj2698dm@gmail.com', 'jrtkqanodovazjvh')
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