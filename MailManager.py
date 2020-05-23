"""
UNIVERSITY OF ST. GALLEN

STUDENT PROJECT: Mail Manager

A PROGRAM BY: Claudius Dreysse, Sebastian Gerl, Moritz Köhler

LECTURER: Dr. Mario Silic
"""

#-------------------------------------------------------------------------------
#%%
"""
NOW ALL PACKAGES THAT ARE NEEDED TO RUN THE CODE ARE IMPORTED AND
(IF NEEDED) PIP INSTALLED BEFORE.
"""
#-------------------------------------------------------------------------------
import os
# Indicator that the program was run before: 'get-pip.py' file exists.
# If that is the case, the program does not need to install pip nor
# the special packages again, which saves run time.
if 'get-pip.py' not in os.listdir():
    os.system('curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py')
    os.system('python3 get-pip.py')
    os.system('pip install regex')
    os.system('pip install tqdm')
import re
import mimetypes
import smtplib
import imaplib
import email
import datetime as dt
import time
from tqdm import tqdm
from getpass import getpass
from email.message import EmailMessage

#-------------------------------------------------------------------------------
#%%
"""
THE CLASS 'MailUser' WILL BE CREATED. IT HOLDS DIFFERENT FUNCTIONS WHICH WILL
INTERACT WITH THE USER IN ORDER TO CUSTOMIZE THE PROGRAM.
"""
#-------------------------------------------------------------------------------
class MailUser():
    def __init__(self):
        """
        Initialization function that defines all attributes for the class object.
        """
        self.program_directory = os.getcwd()
        self.email_address = ''
        self.email_password = ''
        self.attachment_directory = ''
        self.providers_dict = {'gmail': 'gmail.com', 'googlemail':'gmail.com', 'yahoo': 'mail.yahoo.com'}
        self.server = ''
        self.attachment_directory = ''
        self.sorting_criterion = ''
        self.refresh_frequency = ''
        self.refresh_pause = 0
        self.report_preference = ''
        self.report_frequency = ''
        self.report_pause = 0
        # TO CHECK FUNCTIONALITY WE RECOMMEND TO MANIPULATE 'time_dictionary',
        # VALUES ARE DENOTED IN SECONDS
        self.time_dictionary = {"minute":60, "hour":3600, "day":86400, "week":604800}

    def define_credentials(self):
        """
        Function that asks the user to enter her email address and password.
        """
        # checking for the 'credentials.txt' file which holds the last entered credentials
        try:
            # suggest to use email adress from last time
            credentials_file = open(f'{self.program_directory}/credentials.txt','r').readlines()[0]
            self.email_address = credentials_file.split(",")[0]
            while True:
                confirm_email = input(f"\n Would you like to continue with '{self.email_address}'? (Yes/No): ").lower()
                if confirm_email in ['yes', 'no']:
                    break
            if confirm_email == 'yes':
                # check whether a password was stored during the last run
                try:
                    self.email_password = credentials_file.split(",")[1]
                except:
                    self.email_password = getpass("\n Enter your password: ")
                    while True:
                        save_password = input('\n Do you want to save your password? (Yes/No): ').lower()
                        if save_password in ['yes', 'no']:
                            break
                    if save_password == 'yes':
                        self.write_file('credentials.txt', [self.email_address, ",", self.email_password])

            else:
                os.remove(f'{self.program_directory}/mail_id.txt')
                raise exception

        # if no 'credentials.txt' is found, the program askes the user for credentials
        except:
            self.email_address = input('\n Enter your email address: ')
            self.email_password = getpass('\n Enter your password: ')
            while True:
                save_password = input('\n Do you want to save your password? (Yes/No): ').lower()
                if save_password in ['yes', 'no']:
                    break
            # store the account details in local file
            if save_password == 'yes':
                self.write_file('credentials.txt', [self.email_address, ",", self.email_password])
            else:
                self.write_file('credentials.txt', [self.email_address])

        # check if the program can login in with the entered account details
        self.check_credentials()


    def define_saving_location(self):
        """
        Function that asks the user where in the operating system the attachments
        should be saved.
        """
        # checking for 'attachment_directory.txt' file that would hold the directory
        # where the attachments were stored during the last run
        try:
            path_file = open(f'{self.program_directory}/attachment_directory.txt', 'r').readlines()[0]
            directory_name = path_file.split("/")[-2]
            folder_name = path_file.split("/")[-1]
            while True:
                confirm_folder = input(f"\n Your attachmets go in '{directory_name}'/'{folder_name}'. \n Choose to 'continue' or 'change': ")
                if confirm_folder in ['continue', 'change']:
                    break
            if confirm_folder == 'change':
                while True:
                    attachments_path = input("\n Choose to store attachments in 'Documents' or on 'Desktop': ").capitalize()
                    if attachments_path in ["Documents", "Desktop"]:
                        break
                folder_name = input("\n Name the attachment folder: ")
                # use regex to find the base path dependent on the individual user name
                base_path = re.search(".*/Users/\w+/", self.program_directory)[0]
                if attachments_path == 'Documents':
                    self.attachment_directory = os.path.join(base_path, 'Documents', folder_name)
                    if not os.path.exists(self.attachment_directory):
                        os.mkdir(self.attachment_directory)
                        self.write_file('attachment_directory.txt', [self.attachment_directory])
                else:
                    self.attachment_directory = os.path.join(base_path, 'Desktop', folder_name)
                    if not os.path.exists(self.attachment_directory):
                        os.mkdir(self.attachment_directory)
                        self.write_file('attachment_directory.txt', [self.attachment_directory])

            # recreate the folder in case of deletion
            else:
                if not os.path.exists(path_file):
                    os.mkdir(path_file)
                    print(f"\n We recreated your folder '{folder_name}'.")
                self.attachment_directory = path_file

        # if no directory path was saved, the program will ask the user to define
        # a storing location for the attachments
        except:
            while True:
                attachments_path = input("\n Choose to store attachments in 'Documents' or on 'Desktop': ").capitalize()
                if attachments_path in ['Documents', 'Desktop']:
                    break
            folder_name = input("\n Name the attachment folder: ")
            # use regex to find the base path dependent on the individual user name
            base_path = re.search(".*/Users/\w+/", self.program_directory)[0]
            if attachments_path == 'Documents':
                self.attachment_directory = os.path.join(base_path, 'Documents', folder_name)
                if not os.path.exists(self.attachment_directory):
                    os.makedirs(self.attachment_directory)
                self.write_file('attachment_directory.txt', [self.attachment_directory])

            else:
                self.attachment_directory = os.path.join(base_path, 'Desktop', folder_name)
                if not os.path.exists(self.attachment_directory):
                    os.makedirs(self.attachment_directory)
                self.write_file('attachment_directory.txt', [self.attachment_directory])

    def define_sort_structure(self):
        """
        Function that asks the user by what information the attachments should be
        organised within in the user defined directory for the attachments.
        """
        # checking for the 'sorting_criterion.txt' that stores the user's preference
        # by what the attachments should be sorted in the above defined directory
        try:
            sorting_file = open(f'{self.program_directory}/sorting_criterion.txt', 'r').readlines()[0]
            self.sorting_criterion = sorting_file
            while True:
                confirm_sorting = input(f"\n Your attachments are sorted by '{sorting_file}'. \n Choose to 'continue' or 'change': ")
                if confirm_sorting in['continue','change']:
                    break
            if confirm_sorting == 'change':
                while True:
                    self.sorting_criterion = input("\n By what do you want to sort the attachments? \n 'Date', 'Sender', 'Document Type', 'Subject': ").lower()
                    if self.sorting_criterion in ["date", "sender", "document type", "subject"]:
                        break
                self.write_file('sorting_criterion.txt', [self.sorting_criterion])

        # let the user define the sorting structure
        except:
            while True:
                self.sorting_criterion = input("\n By what do you want to sort the attachments? \n 'Date', 'Sender', 'Document Type', 'Subject': ").lower()
                if self.sorting_criterion in ['sender', 'document type', 'subject', 'date']:
                    break
            self.write_file('sorting_criterion.txt', [self.sorting_criterion])


    def define_refresh_frequency(self):
        """
        Function that asks the user to define how often the program should check for
        new mails with attachments.
        """
        while True:
            self.refresh_frequency = input("\n Checking for new mails every minute.\n Choose to 'Continue' or 'Change': ").lower()
            if self.refresh_frequency in ['continue', 'change']:
                break
        if self.refresh_frequency == 'change':
            while True:
                self.refresh_frequency = input("\n Change to every 'Minute', 'Hour' or every 'Day': ").lower()
                if self.refresh_frequency in ['minute', 'hour', 'day']:
                    break
        else:
            self.refresh_frequency = "minute"
        self.refresh_pause = self.time_dictionary[self.refresh_frequency]


    def define_report_frequency(self):
        """
        Function that asks the user to define if and how often the program should
        send a report of the downloaded files via email.
        """
        while True:
            self.report_preference = input("\n Do you want to receive reports on saved attachments? (Yes/No): ").lower()
            if self.report_preference in ["yes", "no"]:
                break
        if self.report_preference=="yes":
            while True:
                self.report_frequency = input("\n How often do you want to receive reports?\n Once every 'Hour', 'Day', 'Week': ").lower()
                if self.report_frequency in ['hour', 'day', 'week']:
                    break
            self.report_pause = self.time_dictionary[self.report_frequency]
            print(f'\n You will receive reports every {self.report_frequency}.')
        else:
            print("\n You will not receive reports.")


    def download_attachments(self):
        """
        Function that uses the previously defined specifications (e.g. email account or
        saving location) to search for attachments in newly received mails and to save
        them in the right place on the computer.
        """
        self.start_interval = dt.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
        report_time = dt.datetime.strptime(self.start_interval,'%Y-%m-%d %H:%M:%S') + dt.timedelta(seconds=self.report_pause)
        try:
            while True:
                print("\n Checking for new mails...")
                # connect to server each time the loop restarts in order to update
                # the mail-ids pulled from the inbox
                mail = self.connect_to_server()
                search_result, search_data = mail.uid('search', None, 'All')
                inbox_list = search_data[0].split()
                # pull id of last examined mail, if available
                try:
                    last_id = open(f'{self.program_directory}/mail_id.txt', "r").readlines()[0]
                except:
                    self.write_file('mail_id.txt', [str(inbox_list[-1], 'utf-8', 'ignore')])
                    last_id = open(f'{self.program_directory}/mail_id.txt', 'r').readlines()[0]

                i, current_id = 1, inbox_list[-1]
                # store the mail-id of the latest e-mail ('current_id') to identify new messages in the next program run
                self.write_file('mail_id.txt', [str(current_id, 'utf-8', 'ignore')])

                print(f' New mails: {int(current_id) - int(last_id)}')

                # identify new mails by comparing the id of the last examined email with
                # the latest new email in the inbox
                while int(last_id) < int(current_id):
                    # fetch the email with the respective id and convert it into
                    # an 'email.message' object
                    fetch_result, fetch_data = mail.uid('fetch', current_id, '(RFC822)')
                    raw_email = fetch_data[0][1].decode("utf-8")
                    email_message = email.message_from_string(raw_email)

                    # loop through different parts of the 'email_message'-object to
                    # identify files and save them in the previously user defined folder
                    for part in email_message.walk():
                        filename = part.get_filename()
                        # store filename in an external textfile if a name was found
                        if type(filename) == str:
                            file = open(f'{self.program_directory}/saved_attachments.txt', 'a')
                            file.writelines(f'\n    - {filename}')
                            file.close()

                            # define the folders based on user defined sorting criterion
                            if self.sorting_criterion == 'date':
                                sorted_by = '_'.join(email_message[self.sorting_criterion].split(' ')[1:4])
                            elif self.sorting_criterion == 'sender':
                                from_ = email_message['from']
                                sorted_by = from_.split('<')[1].replace('>', '')
                            elif self.sorting_criterion == 'subject':
                                sorted_by = email_message[self.sorting_criterion]
                            elif self.sorting_criterion == 'document type':
                                sorted_by = filename.split('.')[-1]

                            # save the identified files from the email
                            save_path = os.path.join(self.attachment_directory, sorted_by)
                            if not os.path.exists(save_path):
                                os.makedirs(save_path)
                            with open(os.path.join(save_path, filename), 'wb') as fp:
                                fp.write(part.get_payload(decode=True))

                    i += 1
                    current_id = inbox_list[-i]

                # send a report via email depending on user preference
                if self.report_preference == 'yes':
                    intermediate_interval = dt.datetime.strptime(dt.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'), '%Y-%m-%d %H:%M:%S')
                    if report_time < intermediate_interval:
                        self.send_report()
                        report_time += dt.timedelta(seconds=self.report_pause)


                print("\n Waiting for next run... To exit press [ctrl]+[c]")
                with tqdm(total=self.refresh_pause, bar_format=" {bar} {l_bar} Time until next run: {remaining} ") as pbar:
                    for i in range(self.refresh_pause):
                        time.sleep(1)
                        pbar.update(1)

        # terminate the program if user interrupts via 'crtl'+'c'
        except KeyboardInterrupt:
            print("\n You have terminated the program.")
            exit()

    #---------------------------------------------------------------------------
    """
    FOR PURPOSES OF CODE READABILITY THE FOLLOWING FUNCTIONS WERE DEFINED. THEY
    EITHER EXECUTE RECURRING CODE SNIPPETS OR DIVIDE LONG FUNCTIONS FROM ABOVE
    INTO SHORTER AND MORE INTUITVE PARTS.
    """
    #---------------------------------------------------------------------------

    def connect_to_server(self):
        """
        Function that connects to the respective mail server and directly accesses
        the inbox for further usage. This function is used to either test if the
        entered credentials are correct or to access email.
        """
        self.server = self.providers_dict[self.email_address.split('@')[1].split('.')[0]]
        m = imaplib.IMAP4_SSL(f'imap.{self.server}')
        m.login(self.email_address, self.email_password)
        m.select('inbox')
        return m

    def check_credentials(self):
        """
        Check if the entered email address and password are correct.
        """
        while True:
            try:
                mail = self.connect_to_server()
                break
            except:
                print("\n Your email address or password might be wrong.")
                self.email_address = input('\n Enter your email address: ')
                self.email_password = getpass('\n Enter your password: ')
                while True:
                    save_password = input('\n Do you want to save your password? (Yes/No): ').lower()
                    if save_password=="yes" or save_password=="no":
                        break
                if save_password=="yes":
                    self.write_file('credentials.txt', [self.email_address, ",", self.email_password])
                else:
                    self.write_file('credentials.txt', [self.email_address])

                try:
                    mail = self.connect_to_server()
                    break
                except:
                    continue

    def send_mail(self, msg):
        """
        Function that enables sending the passed message.
        """
        with smtplib.SMTP_SSL(f'smtp.{self.server}', 465) as smtp:
            smtp.login(self.email_address, self.email_password)
            smtp.send_message(msg)


    def send_report(self):
        """
        Function that sends a report email to the user's email adress with a list
        of the downloaded filenames.
        """
        # intermediate_interval = dt.datetime.strptime(dt.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S'), '%Y-%m-%d %H:%M:%S')
        # # report_time = dt.datetime.strptime(self.start_interval,'%Y-%m-%d %H:%M:%S') + dt.timedelta(seconds=self.report_pause)
        # if time < intermediate_interval:
        # adjust the list 'downloaded_files' so that they will be displayed below each other
        try:
            downloads_list = ''.join(open(f'{self.program_directory}/saved_attachments.txt').readlines())

            # set up the report email
            attachment_report = EmailMessage()
            attachment_report['Subject'] = 'Attachment Report'
            attachment_report['From'] = self.email_address
            attachment_report['To'] = self.email_address
            attachment_report.set_content(f'We identified new files in your inbox. The following files were downloaded and saved in {self.attachment_directory}:\n{downloads_list}.\n\n\n\nYou will receive the next report in 1 {self.report_frequency} if new attachments were identified.')

            # actual sending
            self.send_mail(attachment_report)
            print("\n Attachment report has been sent.")
            # remove 'saved_attachments.txt' so that the next update will not
            # double report older attachments
            os.remove("saved_attachments.txt")
        except:
            pass

        # report_time += dt.timedelta(seconds=self.report_pause)


    def write_file(self, name, input_list):
        """
        Function that newly writes, overwrites or reads an existing text file.
        """
        file = open(f'{self.program_directory}/{name}', "w+")
        file.writelines(input_list)
        file.close()

#-------------------------------------------------------------------------------
#%%
"""
FOR AUTOMATED EXECUTION.
"""
#-------------------------------------------------------------------------------
def run_program(user):
    """Function that runs the program in the right order"""
    user.define_credentials()
    user.define_saving_location()
    user.define_sort_structure()
    user.define_refresh_frequency()
    user.define_report_frequency()
    user.download_attachments()

run_program(MailUser())
