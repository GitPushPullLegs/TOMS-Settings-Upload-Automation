import re
from datetime import datetime
from typing import Dict, Any, Union

import numpy
import pandas as pd

from OSManager import OSManager
from Settings import Settings


class ErrorData:
    ssid = None
    error = None
    action = None
    _column = None

    def __init__(self, ssid, error, column):
        self.ssid = ssid
        self.error = error
        self._column = column

    # The conversion table is because the TOMS errors output uses different spacing and dashes on some columns.
    _conversion_table: Dict[Union[str, Any], Union[str, Any]] = {  # errorFile : seisFile
        'Student SSID': 'Student SSID',
        '*Alternate Response Options (Computer - E, M, S, CE, CM, CS, RO) (Paper - E, M, S)': '*Alternate Response Options (Computer – E, M, S, CE, CM, CS, RO)(Paper – E, M, S)',
        '*Color Contrast (NON-EMBEDDED) (Computer - E, M, S, CE, CM, CS, RO)': '*Color Contrast (NON-EMBEDDED) (Computer – E, M, S, CE, CM, CS, RO)',
        '*Magnification (ET) (Computer - E, M, S, CE, CM, CS, RO)': '*Magnification (ET) (Computer – E, M, S, CE, CM, CS, RO)',
        '*Noise Buffers (Computer - E, M, S, CE, CM, CS, RO) (Paper - E, M, S)': '*Noise Buffers (Computer – E, M, S, CE, CM, CS, RO)(Paper – E, M, S)',
        '*Print on Demand (Computer - E, M, S, CE, CM, CS, RO)': '*Print on Demand (Computer – E, M, S, CE, CM, CS, RO)',
        '*Read-Aloud in Spanish (Computer - M)': '*Read-Aloud in Spanish(Computer – M)(Paper – M)',
        '*Read-Aloud in Spanish (Computer - S)': '*Read-Aloud in Spanish(Computer – S)(Paper – S)',
        '*Read-Aloud Items (NON-EMBEDDED) (Computer - E, M, S, CE, CM, CS, RO) (Paper - E, M, S)': '*Read-Aloud Items (NON-EMBEDDED) (Computer – E, M, S, CE, CM, CS, RO)(Paper – E, M, S)',
        '*Read-Aloud Passages (NON-EMBEDDED) (Computer - E, CE, RO) (Paper - E)': '*Read-Aloud Passages (NON-EMBEDDED) (Computer – E, CE, RO)(Paper – E)',
        '*Scribe (Writing) (Computer - E, CE, RO) (Paper - E)': '*Scribe (Writing) (Computer – E, CE, RO)(Paper – E)',
        '*Scribe Items (Computer - E and CE - Non-Writing, M, S, CM, CS, RO) (Paper - E, M, S)': '*Scribe Items (Computer – E and CE – Non-Writing, M, S, CM, CS, RO)(Paper – E, M, S)',
        '*Simplified Test Directions (Computer - E, M, S, CE, CM, CS, RO) (Paper - E, M, S)': '*Simplified Test Directions(Computer – E, M, S, CE, CM, CS, RO)(Paper – E, M, S)',
        '*Speech-to-Text (Computer - E, M, S) (Paper -  E, M, S)': '*Speech-to-Text (Computer – E, M, S)(Paper –  E, M, S)',
        '*Text-to-Speech - Item and Stimuli (Computer - S)': '*Text-to-Speech – Item and Stimuli(Computer – S)',
        '*Text-to-Speech (EMBEDDED) (Computer - E, M)': '*Text-to-Speech (EMBEDDED) (Computer – E, M)',
        '*Text-to-Speech Items (Computer - RO)': '*Text-to-Speech Items(Computer – RO)',
        '*Text-to-Speech Passages (EMBEDDED)  (Computer - E)': '*Text-to-Speech Passages (EMBEDDED) (Computer – E)',
        '*Text-to-Speech Passages (EMBEDDED) (Computer - RO)': '*Text-to-Speech Passages (EMBEDDED)(Computer – RO)',
        '*Translation Glossaries (ET) (EMBEDDED) (Computer - M)': '*Translation Glossaries (ET) (EMBEDDED)(Computer – M)',
        '*Translation Glossaries (ET) (EMBEDDED) (Computer - S)': '*Translation Glossaries (ET) (EMBEDDED)(Computer – S)',
        '*Translation Glossaries (ET) (NON-EMBEDDED) (Paper - M)': '*Translation Glossaries (ET) (NON-EMBEDDED)(Paper – M)',
        '*Translation Glossaries (ET) (NON-EMBEDDED) (Paper - S)': '*Translation Glossaries (ET) (NON-EMBEDDED)(Paper – S)',
        '100s Number Table (Grades 4 and Up) (Computer - M, CM) (Paper - M)': '100s Number Table(Computer – M, CM)(Paper – M)',
        '100s Number Table (Computer - S, CS) (Paper - S)': '100s Number Table(Computer – S, CS)(Paper – S)',
        'Abacus (M, S, CM, and CS) (Paper - M, S)': 'Abacus (Computer – M, S, CM, CS) (Paper – M, S)',
        'Additional Instructional Supports for Alternate Assessments (NON-EMBEDDED) (Computer - CE, CM, CS)': 'Additional Instructional Supports for Alternate Assessments (NON-EMBEDDED)(Computer – CE, CM, CS)',
        'American Sign Language (Computer - E - Listening, M, S)': 'American Sign Language (Computer – E – Listening, M, S)',
        'Amplification (Computer - E, M, S, CE, CM, CS, RO)': 'Amplification(Computer – E, M, S, CE, CM, CS, RO)',
        'Audio Transcript (Computer - E, RO - Listening)': 'Audio Transcript(Computer – E, RO – Listening)',
        'Bilingual Dictionary (ET)  (Computer - E - PT Full Write) (Paper - E)': 'Bilingual Dictionary (ET) (Computer – E ‒ PT Full Write)(Paper – E)',
        'Braille (ET)  (Computer - E, M)': 'Braille (ET) (Computer – E, M)',
        'Braille (ET) (Computer - S)': 'Braille (ET)(Computer – S)',
        'Braille (NON-EMBEDDED) (Paper - E, M, S)': 'Braille (NON-EMBEDDED)(Paper – E, M, S)',
        'Braille (Computer - RO)': 'Braille(Computer – RO)',
        'Calculator (Four Function for Grade 5 and Scientific for Grades 8, 10, 11, and 12) (Computer - S) (Paper - S)': 'Calculator (Four Function for Grade 5 and Scientific for Grades 8, 10, 11, and 12)(Computer – S)(Paper – S)',
        'Calculator (Grades 6-8 and 11) (Computer - M) (Paper - M)': 'Calculator (Grades 6–8 and 11) (Computer – M)(Paper – M)',
        'Closed Captioning  (Computer - E - Listening, RO)': 'Closed Captioning (Computer – E ‒ Listening, RO)',
        'Color Contrast (EMBEDDED) (Computer - E, M, S, CE, CM, CS, RO)': 'Color Contrast (EMBEDDED) (Computer – E, M, S, CE, CM, CS, RO)',
        'Color Overlay (Computer - E, M, S, CE, CM, CS, RO)': 'Color Overlay (Computer – E, M, S, CE, CM, CS, RO)',
        'Large-print Special Form (NON-EMBEDDED) (Paper - E, M, S)': 'Large-print Special Form (NON-EMBEDDED)(Paper – E, M, S)',
        'Masking (Computer - E, M, S, CE, CM, CS, RO)': 'Masking (Computer – E, M, S, CE, CM, CS, RO)',
        'Medical Supports (Computer - E, M, S, CE, CM, CS, RO) (Paper - E, M, S)': 'Medical Supports(Computer – E, M, S, CE, CM, CS, RO)(Paper – E, M, S)',
        'Mouse Pointer (Size and Color) (Computer - E, M, S, CE, CM, CS, RO)': 'Mouse Pointer (Size and Color)(Computer – E, M, S, CE, CM, CS, RO)',
        'Multiplication Table (NON-EMBEDDED) (Grades 4-8 and 11) (Computer - M, CM) (Paper - M)': 'Multiplication Table (NON-EMBEDDED) (Computer – M, CM)(Paper – M)',
        'Multiplication Table (NON-EMBEDDED) (Computer - S, CS) (Paper - S)': 'Multiplication Table (NON-EMBEDDED)(Computer – S, CS)(Paper – S)',
        'Permissive Mode (Computer - E, M, S, CE, CM, CS, RO)': 'Permissive Mode(Computer – E, M, S, CE, CM, CS, RO)',
        'Print Size (Computer - E, M, S, CE, CM, CS, RO)': 'Print Size(Computer – E, M, S, CE, CM, CS, RO)',
        'Science Charts [State-approved] (i.e., Periodic Table of the Elements, reference sheets)  (Computer - S) (Paper - S)': 'Science Charts [State-approved] (i.e., Periodic Table of the Elements, reference sheets) (Computer – S)(Paper – S)',
        'Separate Setting (Computer - E, M, S, CE, CM, CS, RO) (Paper - E, M, S)': 'Separate Setting (Computer – E, M, S, CE, CM, CS, RO)(Paper – E, M, S) ',
        'Stacked Translations and Translated Test Directions (Spanish) (EMBEDDED) (Computer - M)': 'Stacked Translations and Translated Test Directions (Spanish) (EMBEDDED)(Computer – M)',
        'Stacked Translations and Translated Test Directions (Spanish) (EMBEDDED) (Computer - S)': 'Stacked Translations and Translated Test Directions (Spanish) (EMBEDDED)(Computer – S)',
        'Streamline (Computer - E, M, S, CE, CM, CS, RO)': 'Streamline(Computer – E, M, S, CE, CM, CS, RO)',
        'Translated Test Directions (ET) (PDF on CAASPP.org) (NON-EMBEDDED) (Computer - E, M, S) (Paper - E, M, S)': 'Translated Test Directions (ET) (PDF on CAASPP.org) (NON-EMBEDDED)(Computer – E, M, S)(Paper – E, M, S)',
        'Turn Off Any Universal Tool(s) (Computer - E, M, S, CE, CM, CS, RO)': 'Turn Off Any Universal Tool(s)(Computer – E, M, S, CE, CM, CS, RO)',
        'Word Prediction (Computer - E, M, S, CE, CM, CS)': 'Word Prediction(Computer – E, M, S, CE, CM, CS)'
    }

    @property
    def seis_column(self) -> str:
        return self._conversion_table[self._column]


class PandasManager:
    _stashed_case_manager_data = None
    _isoCalendar = datetime.today().isocalendar()

    def stash_case_manager(self, seis_file_path: str, test=False):
        """
        Captures the School and Case Manager information to later be merged with the error data so we can email the file to the correct people.
        Also, removes all columns that TOMS doesn't want.
        :param seis_file_path: The original seis file.
        :param test: For testing purposes. Uses a test file and outputs to a separate file.
        """
        settings = Settings()
        seis_file_path = seis_file_path if not test else settings.test_seis_file_path
        seis_output_file_path = seis_file_path if not test else settings.test_seis_file_path + '1.xlsx'
        seisDF = pd.read_excel(seis_file_path, header=0)
        self._stashed_case_manager_data = seisDF[['Student SSID', 'School', 'Case Manager']].copy()
        seisDF = seisDF.drop(
            ['SEIS ID', 'FirstName', 'LastName', 'Birthdate', 'District', 'School', 'School Type', 'GradeLevel',
             'Case Manager'], axis=1)
        seisDF.to_excel(seis_output_file_path, index=False)

    def evaluate_upload(self, seis_file_path: str, errors_file_path: str, test=False):
        """
        Using the error file, it clears out error cells in the seis file. If it cannot clear out the error cell, it removes the row.
        It also outputs an errors file to be emailed to the case managers.
        :param seis_file_path: The seis file.
        :param errors_file_path: The error file from TOMS.
        :param test: For testing purposes. Uses a test file and outputs to a separate file.
        """
        settings = Settings()
        seis_file_path = seis_file_path if not test else settings.test_seis_file_path
        seis_output_file_path = seis_file_path if not test else settings.test_seis_file_path + '1.xlsx'
        errors_file_path = errors_file_path if not test else settings.test_error_file_path

        print(seis_file_path)
        seis_df = pd.read_excel(seis_file_path, header=0)
        errors_df = pd.read_csv(errors_file_path, header=0)
        output_errors_df = pd.DataFrame(columns=['SSID', 'Action', 'Error'])
        error_data = []
        offset = self.calculate_offset(errors_df, seis_df)

        for (_, row) in errors_df.iterrows():
            seis_row_number = row['Row Number'] + offset
            seis_row = seis_df.loc[[seis_row_number]]
            temp_error_data = ErrorData(numpy.int64(seis_row['Student SSID'].to_string(index=False, header=False)),
                                        row['Error'], row['Column Name'])
            print(temp_error_data.ssid, temp_error_data.error)
            error_data.append(temp_error_data)
        for error in error_data:

            if (error.seis_column != 'Student SSID') and (error.seis_column in seis_df.columns):
                try:
                    seis_df.at[
                        seis_df[seis_df['Student SSID'] == error.ssid].index[0],
                        error.seis_column
                    ] = ''
                except:
                    pass
                error.action = 'Removed Value'
                print(f"Removed Value for {error.ssid}: {error.seis_column}")
            else:
                seis_df = seis_df[seis_df['Student SSID'] != error.ssid]
                error.action = 'Deleted Row'
                print(f"Deleted Row for {error.ssid}: {error.seis_column}")

            output_errors_df = output_errors_df.append({
                'SSID': error.ssid,
                'Action': error.action,
                'Error': error.error
            }, ignore_index=True)

        output_errors_df = output_errors_df.merge(self._stashed_case_manager_data, left_on='SSID',
                                                  right_on='Student SSID')
        self.store_errors(output_errors_df, test=test)
        seis_df.to_excel(seis_output_file_path, index=False)

    def store_errors(self, output_errors_df, test=False):
        settings = Settings()
        download_path = settings.weekly_email_file_path if not test else settings.test_download_path
        file_path = download_path + r'\\' + f"{self._isoCalendar[0]}{self._isoCalendar[1]}-ErrorFile.csv"
        if OSManager().file_exists(file_path):
            original_errors_df = pd.read_csv(file_path)
            original_errors_df = original_errors_df.append(output_errors_df, ignore_index=True)
            original_errors_df.to_csv(file_path)
        else:
            output_errors_df.to_csv(file_path)

    # TODO: - Setup weekly email. Right now the above function stores the error log for the week in the weekly email folder. I need another script to run Friday at 9 AM that takes that, emails it, then stores it in the W drive.

    def calculate_offset(self, error_df, seis_df) -> int:
        """
        Today, Oct 8, 2020, TOMS is having an issue where the reported error row number is off by 80 rows. We're reporting the error but we don't know when they'll fix it. This will catch the offset, if there is one, and adjust for it.
        :param error_df: the errors dataframe
        :param seis_df: the seis dataframe
        :return: Returns an integer representing the offset between the alleged error row in the error file and the actual row in the seis file.
        """
        for (_, row) in error_df.iterrows():
            if row['Error'].startswith("'The student with SSID") or row['Error'].startswith("'A student with SSID"):
                ssid = numpy.int64(re.findall(r"[0-9]+", row['Error'])[0])
                alleged_row_number = row['Row Number']
                actual_row_number = seis_df[seis_df['Student SSID'] == ssid].index[0]
                return actual_row_number - alleged_row_number

        return 0
