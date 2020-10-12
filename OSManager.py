from glob import glob
import os
from os.path import basename
from Settings import Settings
from time import sleep
from zipfile import ZipFile
from datetime import datetime


class OSManager:
    def get_seis_file_path(self) -> str:
        sleep(7)
        return self._get_filepath(ending='.xlsx')

    def get_errors_file_path(self) -> str:
        sleep(7)
        return self._get_filepath(starting='download', ending='.csv')

    def _get_filepath(self, starting: str = None, ending: str = None) -> str:
        try:
            allFiles = glob(Settings().download_path + r'\\' + f'{starting if starting is not None else ""}' + '*' + f'{ending}')
            return max(allFiles, key=os.path.getctime)
        except:
            sleep(23)
            return self._get_filepath(starting, ending)

    def file_exists(self, filepath: str) -> bool:
        return os.path.isfile(filepath)

    def cleanup(self):
        """
        Zips up the files and store them in the W drive. Delete the originals.
        """
        with ZipFile(Settings().stash_path + fr"\TOMS Settings Upload Files {datetime.now().strftime('%Y%m%d%H%M%S')}.zip", 'w') as zipFile:
            allFiles = [file for file in glob(Settings().download_path + r"\*") if not os.path.isdir(file)]
            for file in allFiles:
                zipFile.write(file, basename(file))
                os.remove(file)
