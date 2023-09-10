# Email Module

This module is located under the `mail` directory. 

## Getting Started

In this module, there's a need to create a Python script named `send_mail.py`.

### Send_mail.py

This file is crucial for sending emails. It should contain two functions which should be manually added. 
However, these functions will require some sensitive information. Please make sure this information is secured and not exposed to maintain your privacy and security.

Here are the steps to get you started:

1. Navigate to the `mail` directory.
2. Create a Python file named `send_mail.py` if it isn't present already.
3. In the `send_mail.py` file, add the necessary functions filled with your sensitive information. Please ensure you secure your code.

**Important Note:** Be wary of the potential risks of exposing sensitive information in your code. Thus, be sure to follow best practices of securing sensitive information in your scripts.

For more information on how to secure your sensitive information in Python scripts, you may refer to Python's official documentation or other trusted sources.
```python

import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# LIST OF TITLES AND E_MAILS TO SEND ADD THEM HERE
users = {'Title': ['Hello John, Check This Out', 'Hello Maria, Report is Ready'],
         'mail': ['johndoe@demo.com', 'mariadoe@demo.com'],
         }

# ADD     email_user = ''  AND email_password = '' with VALUES
def a_gmail(email_send, subj, word, path_to_file, output_file):
    email_user = ''
    email_password = ''
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subj
    body = word
    msg.attach(MIMEText(body, 'html'))
    attachment = open(path_to_file, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + output_file)
    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_user, email_password)
    server.sendmail(email_user, email_send, text)
    server.quit()
    print(" Το e-mail {} στάλθηκε επιτυχώς".format(email_send))

def send_mail(mail_lst, mail_names, word, path_to_file, output_file):
    for i in range(len(mail_lst)):
        c = 'S: {}'.format(mail_names[i])
        print('Αποστολή μηνύματος στον παραλήπτη {}'.format(mail_names[i]))
        a_gmail(mail_lst[i], c, word, path_to_file, output_file)
```