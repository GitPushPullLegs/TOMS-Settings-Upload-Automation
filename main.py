from EmailManager import EmailManager
from OSManager import OSManager
from PandasManager import PandasManager
from Settings import Settings
from WebbotManager import WebbotManager

web = None
os_manager = OSManager()
pandas_manager = PandasManager()

seis_file_path: str
errors_file_path: str


def revalidate():
    global errors_file_path
    pandas_manager.evaluate_upload(seis_file_path, errors_file_path)
    web.upload_test_settings(seis_file_path)
    if not web.upload_is_valid():
        errors_file_path = os_manager.get_errors_file_path()  # It will download the new error file.
        revalidate()

def main():
    global web, seis_file_path, errors_file_path
    web = WebbotManager(Settings().download_path)
    web.download_seis()
    web.cleanup_seis()
    seis_file_path = os_manager.get_seis_file_path()
    pandas_manager.stash_case_manager(seis_file_path)
    web.sign_in_to_caaspp()
    web.upload_test_settings(seis_file_path)
    if not web.upload_is_valid():
        errors_file_path = os_manager.get_errors_file_path()
        revalidate()
    os_manager.cleanup()
    EmailManager().clear_emails()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()
