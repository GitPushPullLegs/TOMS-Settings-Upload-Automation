from datetime import datetime, timedelta
from sys import exit
from time import sleep
from webbot import Browser
from EmailManager import EmailManager
from Settings import Passwords


class WebbotManager:
    web = None
    _passwords = Passwords()

    def __init__(self, download_path=None):
        self.web = Browser(showWindow=True, downloadPath=download_path)
        self.web.maximize_window()

    def __del__(self):
        self.web.quit()

    def download_seis(self):
        """
        Navigates to the seis website, logs in, and downloads the seis summative export. Not the one for TOMS,
        the first one which provides us with more data like the case manager.
        """
        self.web.go_to("https://seis.org/reports/toms")
        self.web.type(self._passwords.seis_login_information['username'], id="username")
        self.web.type(self._passwords.seis_login_information['password'], id="password")
        self.web.click("Login")
        sleep(3)
        self.web.click(xpath='//*[@id="form"]/div/div/div/div/div[1]/div/div/button')
        sleep(30)
        self.web.click("XLSX", number=1)

    def cleanup_seis(self):
        """
        Deletes the SEIS report that was just downloaded.
        """
        self.web.click(classname='seis-icon-delete')
        self.web.click(xpath="/html/body/div[3]/div/div/div[3]/button[1]")

    def sign_in_to_caaspp(self):
        """
        Navigates to the TOMS system and logs in. TOMS thinks it's a new device logging in so it asks for a login
        code which it emails to me and the login manager grabs.
        """
        self.web.go_to("https://mytoms.ets.org/")
        self.web.type(self._passwords.caaspp_login_information['username'], id="username")
        self.web.type(self._passwords.caaspp_login_information['password'], id="password")
        self.web.click("Secure Logon")
        self.web.type(EmailManager().retrieve_login_code(), id="emailcode")
        self.web.click("Submit")
        self.web.click(xpath='//*[@id="roleOrgSelect"]/optgroup/option[2]')
        self.web.click(xpath='//*[@id="okButton"]')

    def upload_test_settings(self, filepath: str):
        self.web.go_to("https://mytoms.ets.org/mt/dt/uploadaccoms.htm")
        self.web.click("Next")
        self.web.type(filepath, xpath='//*[@id="uploadfilepath"]')
        self.web.click("Next")

    def upload_is_valid(self) -> bool:
        print(
            f"File is processing. Will check at {(datetime.now() + timedelta(0, 120)).strftime('%m/%d/%Y %I:%M:%S %p')}")
        sleep(120)
        loops = 20
        while loops >= 0:
            elements = self.web.find_elements(tag='td')
            validation_data = {
                'datetime': datetime.strptime(elements[1].text, '%b %d, %Y, %I:%M %p'),
                'status': elements[3],  # .text
                'action': elements[4]
            }
            if validation_data['status'].text.startswith('Errors'):
                validation_data['action'].click()
                return False
            elif validation_data['status'].text.startswith('Validated'):
                validation_data['action'].click()
                return True
            else:
                print(
                    f"Still processing. Will re-try at {(datetime.now() + timedelta(0, 120)).strftime('%m/%d/%Y %I:%M:%S %p')}")
                loops -= 1
                sleep(120)
                # Had initially used recursion but I didn't want to risk running into the stack loop limit.

        # If managed to get out of loop, report error.
        EmailManager().send_email(subject="TOMS Settings Upload Validation Failed Error",
                                  body="Something went wrong with trying to validate the TOMS Settings Upload file in TOMS.")
        exit()
