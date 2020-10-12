import re
import smtplib
import sys
from datetime import datetime
from email.message import EmailMessage
from time import sleep

import imapclient

from NetworkManager import NetworkManager
from Settings import Passwords


class EmailManager:
    outlook = None
    _passwords = Passwords()

    def __init__(self):
        self.outlook = imapclient.IMAPClient('imap.outlook.com', ssl=True)
        self.outlook.login(self._passwords.email_login_information['email'],
                           self._passwords.email_login_information['password'])

    def retrieve_login_code(self, request_date=datetime.now()):
        print("Retrieving login code.")
        try:
            self.outlook.select_folder('INBOX')

            emails_from_joe = self.outlook.search(['FROM', 'aguilar.jo@monet.k12.ca.us', 'SUBJECT', 'FW: Login code'])
            if emails_from_joe:
                raw_message = self.outlook.fetch([emails_from_joe[-1]], ['BODY[]', 'FLAGS', 'ENVELOPE'])
                for key, value in raw_message.items():
                    timestamp = value[b'ENVELOPE'].date
                    if timestamp < request_date:
                        print(
                            f"Emails from Joe exist. However, latest email is {timestamp.strftime('%m/%d/%Y %H:%M:%S')} which is before {request_date.strftime('%m/%d/%Y %H:%M:%S')}. Waiting 30 seconds.")
                        sleep(30)
                        return self.retrieve_login_code(request_date)
                    body = str(value[b'BODY[]'])
                    codes = re.findall(r"\\n([0-9]+)<", body)

                    return codes[0]
            print("Emails from Joe do not exist. Waiting 60 seconds.")
            sleep(60)
            return self.retrieve_login_code(request_date)
        except:
            if not NetworkManager().is_connected():
                print("Network issue.")
                sleep(90)
                self.retrieve_login_code(request_date)
            else:
                print(f"Unexpected error: {sys.exc_info()[0]}")

    def clear_emails(self):
        self.outlook.select_folder('INBOX')
        all_emails = self.outlook.search()
        self.outlook.delete_messages(all_emails)
        self.outlook.expunge()

    def send_email(self, subject: str = None, body: str = None, to: str = 'aguilar.jo@monet.k12.ca.us'):
        """
        Sends an email from the assessment.evaluation account to whomever.
        :param subject: The subject of the email. Defaults to 'TOMS Settings Upload Error'
        :param body: The body of the email. Defaults to 'An error occurred in the TOMS Settings Upload.'
        :param to: The recipient of the email. Defaults to 'Aguilar.Jo@monet.k12.ca.us'
        """
        subject = subject if subject is not None else "TOMS Settings Upload Error"
        body = body if body is not None else "An error occurred in the TOMS Settings Upload."
        login_info = self._passwords.email_login_information

        smtp_srv = 'smtp.office365.com'
        smtp_server = smtplib.SMTP(smtp_srv, 587)

        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = login_info['email']
        msg['To'] = to
        msg.set_type('text/html')
        msg.set_content(body)

        smtp_server.ehlo()
        smtp_server.starttls()
        smtp_server.login(login_info['email'], login_info['password'])
        smtp_server.send_message(msg)
        smtp_server.close()

    def __del__(self):
        print("Bye :(")
        self.outlook.logout()
