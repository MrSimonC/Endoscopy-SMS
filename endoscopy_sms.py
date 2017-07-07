from custom_modules.mssql import QueryDB
from custom_modules.outlook import Outlook
import csv
import os
import sys
import pyodbc
import datetime
__version__ = '3.0'


class SMS:
    """
    Patient SMS Reminder system
    Reads Lorenzo database for patients in endoscopy, finds mobile numbers and sends reminder sms via DrDr
    """
    def __init__(self):
        self.lorenzo_db = QueryDB('SERVER', 'DATABASE', 'USERNAME', 'PASSWORD')
        self.o = Outlook()

    def get_patient_data(self, sql_file, days):
        """
        Run sql_file contents in self.db, replacing $days in the file with days, process datetime found
        :param sql_file: path to sql file
        :param days: -1, 0, 1 etc of days away from now to find patients for endoscopy
        :return: List of ordered dicts of database rows
        """
        sql = open(sql_file).read().replace('$days', days)
        patients = self.lorenzo_db.exec_sql(sql)  # Returns list of ordered dicts
        for num, pt in enumerate(patients, start=1):
            # e.g. Mon 12 Jan at 21:00
            tci_str = pt['OFFERDTTM'].strftime('%a %d %b') + ' at ' + pt['OFFERDTTM'].strftime('%H:%M')
            pt['tci'] = tci_str
            pt['num'] = num
        return patients

    @staticmethod
    def print_patient_data(patients):
        for patient in patients:
            print(str(patient['num']) + ': ' + patient['Patient Name'].title() + ' at ' + patient['tci'])

    def _drdrsms_send(self, number, body):
        self.o.send(False, number + '@sms.drdoctor.co.uk', 'SMS message to ' + number, body + '--end')

    def send_all(self, patients, message, csv_out, days):
        for patient in patients:
            print('Processing: ' + str(patient['num']) + ': ' + patient['Patient Name'].title())
            try:
                send_msg = message.replace('$tci_datetime', patient['tci'])
                self._drdrsms_send(patient['mobile'], send_msg)
                csv_file = open(csv_out, 'at', newline='')
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow([patient['mobile'], patient['tci'], send_msg, datetime.date.today(), days])
                csv_file.close()
            except IOError:
                print("Can't access output file: " + csv_out)
                return
            # TODO keyerror from db data

if __name__ == '__main__':
    main_folder = r'K:\Coding\Python\supporting files\endoscopy'
    csv_results = os.path.join(main_folder, 'SentMessages.csv')
    sql_path = os.path.join(main_folder, 'endoscopy.sql')
    sms_message = 'You have an appointment at Southmead Hospital on $tci_datetime'

    print('Endoscopy SMS Reminder program v' + __version__ + ' (by Simon Crouch, IM&T Apr2016). \n'
          'Press Ctrl+c to cancel at anytime.\n')
    try:
        print('Contacting database...')
        s = SMS()
    except pyodbc.Error as e:
        print(e.args[1])
        input('Press Enter to exit.')
        sys.exit(0)
    print('Ready.')
    patient_data = None
    days_valid = False
    days_input = input('Type amount of days from now: ')
    while not days_valid:
        try:
            if int(days_input) <= 0:
                raise ValueError
            patient_data = s.get_patient_data(sql_path, str(int(days_input)))
            days_valid = True
        except ValueError:
            days_input = input('That\'s not allowed. Type amount of days from now: ')
    if patient_data:
        s.print_patient_data(patient_data)
        remove_no = input('Type number to remove_no [or blank to move on]: ')
        while remove_no:
            try:
                patient_data[:] = [x for x in patient_data if x.get('num') != int(remove_no.strip())]
                s.print_patient_data(patient_data)
            except ValueError:
                print('Please enter a number. ', end='', flush=True)
            remove_no = input('Type number to remove_no [or blank to move on]: ')
        input('Press Enter to send, or "ctrl+c" to abort.')
        s.send_all(patient_data, sms_message, csv_results, days_input)
        print('Done.')
    else:
        print('No patients found in database')
    input('Press Enter to exit.')
    sys.exit(0)
