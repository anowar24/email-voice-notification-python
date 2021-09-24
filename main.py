# Email Voice Notification
# By Anowar Hussain Mondal
# Reference has been taken from GeeksforGeeks, py4u.net and stackoverflow


import imaplib
import email
from email.header import decode_header
from gtts import gTTS
import time
from datetime import datetime
from playsound import playsound

import os

# Edit this part to your data
username = "YourUserName@gmail.com"
password = "Password"

language = 'en'

def clean(text):
    return "".join(c if c.isalnum() else "_" for c in text)

# create an IMAP4 class with SSL 
imap = imaplib.IMAP4_SSL("imap.gmail.com")

# Log in
imap.login(username, password)


while(1):
    try:
        print('Last Checked at: '+str(datetime.now()))
        time.sleep(10)
        imap.select("INBOX")

        playthis = ''
        status, messages = imap.search(None, '(UNSEEN)')
        # Fetch mails
        messages = messages[0].split()
        N = len(messages)
        if not messages:
            continue

        # total number of emails
        messages = int(messages[-1])

        
        for i in range(messages, messages-N, -1):
            # fetch the email message by ID
            res, msg = imap.fetch(str(i), "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])
                    # decode the email subject
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode(encoding)
                    # decode email sender
                    From, encoding = decode_header(msg.get("From"))[0]
                    if isinstance(From, bytes):
                        From = From.decode(encoding)
                    playthis += 'New Mail '
                    playthis += "From "+ From
                    playthis += "Subject "+ subject
                    playthis += 'Content '
                    # if the email message is multipart
                    if msg.is_multipart():
                        # iterate over email parts
                        for part in msg.walk():
                            # extract content type of email
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            try:
                                # get the email body
                                body = part.get_payload(decode=True).decode()
                            except:
                                pass
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                # print text/plain emails and skip attachments
                                playthis += body
                            elif "attachment" in content_disposition:
                                # download attachment
                                filename = part.get_filename()
                                if filename:
                                    folder_name = clean(subject)
                                    if not os.path.isdir(folder_name):
                                        # make a folder for this email (named after the subject)
                                        os.mkdir(folder_name)
                                    filepath = os.path.join(folder_name, filename)
                                    # download attachment and save it
                                    open(filepath, "wb").write(part.get_payload(decode=True))
                    else:
                        # extract content type of email
                        content_type = msg.get_content_type()
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == "text/plain":
                            # print only text email parts
                            playthis += body

        myobj = gTTS(text=playthis, lang=language, slow=False)
        print(playthis)
        #save the mp3 and play mp3
        myobj.save("play.mp3")
        playsound('play.mp3', True)
        imap.store(str(messages), '+FLAGS', '\Seen')
        playthis = ''
    except KeyboardInterrupt:
        break

# close the connection and logout
print('Logging Out')
imap.close()
imap.logout()