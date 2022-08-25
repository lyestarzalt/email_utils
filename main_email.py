# author:Lyes Tarzalt
from imap_tools import AND, MailBox, MailboxLoginError, MailboxLogoutError
from enum import Enum
import pandas as pd
import imaplib
import socket
import time
import traceback
import os
from pathlib import Path


class Gmailfolders(Enum):
    INBOX = 'INBOX'
    SENT = '[Gmail]/Sent Mail'
    TRASH = '[Gmail]/Trash'
    SPAM = '[Gmail]/Spam'
    DRAFTS = '[Gmail]/Drafts'
    ALL = "[Gmail]/All Mail"


class Email:
    """
    Args:
        host_address (str): host name of the email server eg: imap.gmail.com
        email_address (str): email address of the recipient
        email_password (str): email password, or generate one at https://support.google.com/accounts/answer/185833?hl=en
        subject_email_find (str): subject of the email, that we want to download
    """

    def __init__(self, host_address: str, email_address: str, email_password: str, subject_email_find: str) -> None:
        self.host: str = host_address
        self.address: str = email_address
        self.password: str = email_password
        self.subject: str = subject_email_find
        self.attachment_object: bytes = b''
        self.file_name: str = ''
        self.attachment_dataframe: pd.DataFrame = pd.DataFrame()
        self.file_location: str = 'no file'
        self.current_dir: Path = Path(os.getcwd())
        self.found_email: bool = False

    def search_email(self) -> bool:
        # function that get the attachment from the email based on the subject
        # and return the attachment as bytes object.
        with MailBox(self.host).login(self.address, self.password, Gmailfolders.ALL.value) as mailbox:
            emails = mailbox.fetch(AND(subject=self.subject))
            for msg in emails:
                for att in msg.attachments:
                    self.file_name = att.filename
                    self.attachment_object = att.payload
                    self.found_email = True

            if not self.found_email:
                print('no email found')
                return False

    def catch_email(self, wait_time_hours: int = 1):
        """_summary_
        function to catch an email and download the attachment
        as soon as it arrives.
        How: IMAP idle (https://www.rfc-editor.org/rfc/rfc2177)
        and Imap-tools for python.
        Args:

        wait_time_hour (float): how many hours should the func wait/run

        Returns:
            str: file full path
        """
        connection_start_time: float = time.monotonic()
        connection_current_time: float = 0.0
        done: bool = False
        try:
            with MailBox(self.host).login(self.address, self.password, Gmailfolders.ALL.value) as mailbox:
                print('@@ New connection', time.asctime())
                while connection_current_time < wait_time_hours * 60 * 60:

                    # *it will idle for 60sec and listen to any changes in the
                    # *email
                    responses = mailbox.idle.wait(timeout=60)

                    print(time.asctime(), 'IDLE responses:',
                          'No updates' if len(responses) == 0 else responses)
                    # if any changes it will return a list containing new/updated/deleted
                    if responses:
                        # !here we filter by subject and unseen.
                        # !Ideally we should filter by date as well.
                        for msg in mailbox.fetch(AND(subject=self.subject, seen=False)):
                            for att in msg.attachments:
                                # ?download the file
                                self.file_name = att.filename
                                self.attachment_object = att.payload
                                self.file_location = self.current_dir / self.file_name
                                self.found_email = True
                            print('found the email!', msg.subject)
                            done = True
                    if done:
                        break

                    connection_current_time = time.monotonic() - connection_start_time

        except (TimeoutError, ConnectionError,
                imaplib.IMAP4.abort, MailboxLoginError, MailboxLogoutError,
                socket.herror, socket.gaierror, socket.timeout) as e:
            print(
                f'## Error\n{e}\n{traceback.format_exc()}')
            time.sleep(60)

    def get_file(self) -> str:
        """convert the bytes object to excel file and return the file path
        """
        if self.found_email:
            with open(self.file_name, "wb") as fp:
                fp.write(self.attachment_object)
            print(f'file saved to {self.file_location}')
            self.file_location = f'{self.current_dir} / {self.file_name}'
        else:
            self.file_location = 'email not found'
            print('no email found to download attachment')

        return self.file_location

    def get_dataframe(self, sheet: str = None) -> pd.DataFrame:
        """convert the bytes object to pandas dataframe and return the dataframe

        args:
        sheet (string): sheet name of the excel file
        if no sheet selected it will return the first sheet
        """
        if self.found_email:
            self.attachment_dataframe = pd.read_excel(
                self.attachment_object, sheet_name=sheet)
        else:
            self.attachment_dataframe = pd.DataFrame()

        return self.attachment_dataframe


if __name__ == "__main__":
    # EXAMPLE HOW TO USE THE CLASS
    host = "imap.gmail.com"
    emailAddr = "xxxx@gmail.com "
    password = 'generate one at https://support.google.com/accounts/answer/185833?hl=en' 
 
    email = Email(host_address=host, email_address=emailAddr, email_password=password,
                  subject_email_find='sales number for 20/08/2022')
    email.search_email()
    print(email.file_name)
    print(email.get_dataframe(sheet=None))
