from datetime import datetime
import xlrd 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

now = datetime.now()

month = int(str(now.month))
date = int(str(now.day))
 
loc = ("birthday.xlsx") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  

sheet.cell_value(0, 0) 

for i in range(1,sheet.nrows): 
    
    if int(sheet.cell_value(i, 2)) ==  month :
       
        if int(sheet.cell_value(i, 1)) ==  date :
            mail = sheet.cell_value(i, 3)
            name = sheet.cell_value(i, 0)
            fromaddr = "sendermail"
            toaddr = mail

            msg = MIMEMultipart()

            msg['From'] = fromaddr
            msg['To'] = toaddr
            msg['Subject'] = "Happy BirthDay "

            body = "A million wishes are flooding your timelines and private message box, but mine is merely telling you to live a much more eventful life because you have given me so many sweet memories to look back. I know how lucky I am to have you as my friend because people with good hearts as yours don’t come by all the time. Happy birthday, my friend." ;


            msg.attach(MIMEText(body, 'plain'))

            filename = "happy.jpg"
            attachment = open("happy.jpg", "rb")

            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

            msg.attach(part)

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(fromaddr, "sender_password")
            text = msg.as_string()
            server.sendmail(fromaddr, toaddr, text)
            server.quit()
            print("done")
            
        else:
            continue
    else:
        continue


