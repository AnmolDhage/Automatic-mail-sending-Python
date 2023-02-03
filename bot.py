from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import smtplib
from email.mime.text import MIMEText
import pandas as pd
import time
import openpyxl
from email.mime.base import MIMEBase
from email import encoders
from datetime import date

today = date.today()
d1 = today.strftime("%d/%m/%Y")

def send_email2():
    data=pd.read_excel("emails.xlsx")
    length=len(data)
    df = pd.read_excel('emails.xlsx', index_col='Company').fillna(0)
  
    

    print("total {}, mail send on :- ".format(length))
    for i in range(0,length):
        message = MIMEMultipart('related')
        
        port = 465  # For SSL
        sender_email = ""
        password = ""
        msgAlternative  = MIMEMultipart('alternative')
        message.attach(msgAlternative)
        msgText = MIMEText('<html><body> Do you have trouble getting customers with your current Website?<br><br>'

'Well, that’s what we’re trying to fix at Yelloyolk Media.<br><br>'

'We help businesses gain more customers by delivering innovative Websites, UI/UX Solutions & Branding Services.<br><br>'

'The reason we think Yelloyolk Media will be a great fit for you is because we’ve studied your website and found that some improvements can be done in design and interaction of the site.<br><br>'

'If you want to learn more about Yelloyolk Media, visit us on our website www.yelloyolk.com<br><br>'

'In the last 4 years, we have successfully delivered 500+ Projects, Covered 20+ Different Domains, Served customers in 8 Countries.<br><br>'

'Get a glimpse of our work here (https://www.yelloyolk.com/work)<br><br>'

'If you want to talk more about Branding, UIUX, and Website Experiences.<br><br>'

'Get on a call @ +91-7767842722 or email us: connect@yelloyolk.com <br><br>'

'Name<br>'
'Team Name<br>'
'Company Name<br>'
'www.company.com <br><br> <img src="cid:image1" alt="Yelloyolk Media"> </body></html>', 'html', 'utf-8')
        msgAlternative.attach(msgText)

        fp = open('yellowyolk.png', 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()
        #Define the image's ID as referenced above
        msgImage.add_header('Content-ID', '<image1>')
        message.attach(msgImage)

        # open the file in bynary
        binary_pdf = open('Yelloyolk_Presentation.pdf', 'rb')
 
        payload = MIMEBase('application', 'octate-stream', Name='Yelloyolk_Presentation.pdf')
        payload.set_payload((binary_pdf).read())
 
        # enconding the binary into base64
        encoders.encode_base64(payload)
 
        # add header with pdf name
        payload.add_header('Content-Decomposition', 'attachment', filename='Yelloyolk_Presentation.pdf')
        message.attach(payload)

       
        receiver_email = data["Email ID"].loc[i]
        df.loc[data["Company"].loc[i]]['Feedback'] = 'sent mail on {}'.format(d1)
        df.to_excel('emails.xlsx')
        message['Subject'] = "Your website could be more attractive & modern looking."
        message['From'] = sender_email
        message['To'] = receiver_email
        server = smtplib.SMTP_SSL("smtp.gmail.com", port)
        server.login(sender_email, password)
        server.sendmail(sender_email, [receiver_email], message.as_string())
        server.quit()
        # prt="{}) {}".format(i+1, data["Email ID"].loc[i])
        # print(prt)

        print(df)
        print("{} left".format(length-i))
        time.sleep(10)

if __name__ == "__main__":
    send_email2()
    print("Mails Sent Successfully!!")


# rohit.dalvi@cgl.co.in
