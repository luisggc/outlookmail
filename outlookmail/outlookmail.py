"""Main module."""
import pandas as pd
import math
import re
from validate_email import validate_email
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')

class Mail:
    def __init__(self,
                email_template_path,
                to=0,
                cc=0,
                bcc=0,
                limit_contacts_by_email=290,
                limit_contacts_type="bcc",
                send_on_behalf=0,
                mail_properties={},
                send=False,
                **options
        ):
        """
            Returns a Mail object.
            
                Call signatures:
                    om.Mail(email_template_path, to=0, cc=0, bcc=0, limit_contacts_by_email=290, limit_contacts_type="bcc", send_on_behalf=0, mail_properties={}, send=False, **options)
                
                >>> email = om.Mail(r"C:\Users\luisc\OneDrive\Projetos\outlookmail\Jupyter\Emails\Arquivo do Email.msg",
                >>>     to= "johndoe@hotmail.com; another_email@email.com",
                >>>     cc=["johndoe2@hotmail.com; another_email2@email.com"],
                >>>     bcc= r"path_to_excel_with_emails.xlsx"
                >>> )

                >>> email.display()
                >>> email.send()
            
            Parameters
            ----------
            email_template_path : path
                Path to file with ".msg" format that contains the tmplate used to send the email.
            
            to : string or list or xlsx_path, optional
                Primary receiver(s) of the email.
                Define the emails information like a string, list or path to xlsx file with the emails.
                If path is used as input, all contacts must be in the column A (fisrt column) and the first row will be ignored.
                Possible example:
                'johndoe@hotmail.com; another_email@email.com',
                ["johndoe2@hotmail.com; another_email2@email.com"],
                r"C:\path_to_excel_with_emails.xlsx"
            
            cc : string or list or xlsx_path, optional
                Secondary receiver(s) of the email.
                Define the emails information like a string, list or path to xlsx file with the emails.
                If path is used as input, all contacts must be in the column A (fisrt column) and the first row will be ignored.
                Possible example:
                'johndoe@hotmail.com; another_email@email.com',
                ["johndoe2@hotmail.com; another_email2@email.com"],
                r"C:\path_to_excel_with_emails.xlsx"

            bcc : string or list or xlsx_path, optional
                Secret receiver(s) of the email. The other contacts will not be able to see that the emial was sent to these contacts.
                Define the emails information like a string, list or path to xlsx file with the emails.
                If path is used as input, all contacts must be in the column A (fisrt column) and the first row will be ignored.
                Possible example:
                'johndoe@hotmail.com; another_email@email.com',
                ["johndoe2@hotmail.com; another_email2@email.com"],
                r"C:\path_to_excel_with_emails.xlsx"
            
            limit_contacts_by_email: number, optional
                Outlooks usually limits the number of contacts that you are sending the emails.
                This parameter will make Outlook send multiple emails to keep the maximum of 'limit_contacts_by_email' emails in each dispatch.
                This limits depends on the organization, emails length and other factors. By default It will slit in packs of 290.
                If It is not interesting this option for you, set as a huge number such as 9999999.

            limit_contacts_type: to, cc, bcc, optional
                Type of destination that the rules for 'limit_contacts_by_email' will be applicable.
                Only options: "to", "cc", "bcc"

            send_on_behalf:
                If you main email in outlook have access to other emails, you can set this other email to be the sender.

            mail_properties:
                Check all MailItem properties here https://docs.microsoft.com/en-us/office/vba/api/Outlook.MailItem#properties. Example:
                {'Subject' : 'this is my new subject'}

            send:
                Directly send the e-mail with no check.
        """
        
        self.to_input = to
        self.cc_input = cc
        self.bcc_input = bcc

        self.email_template_path = email_template_path
        self.limit_contacts_by_email = limit_contacts_by_email
        self.send_on_behalf = send_on_behalf
        self.mail_properties = mail_properties
        
        self.to_list = self.contacts_to_list(to)
        self.cc_list = self.contacts_to_list(cc)
        self.bcc_list = self.contacts_to_list(bcc)

        self.to, self.cc, self.bcc = ";".join(self.to_list), ";".join(self.cc_list), ";".join(self.bcc_list)

        contacts_limited = getattr(self, limit_contacts_type+"_list")
        self.packs = []
        if len(contacts_limited)>limit_contacts_by_email:
            n_emails_will_be_sent = math.ceil(len(contacts_limited)/limit_contacts_by_email)
            self.packs = [ "; ".join(contacts_limited[i*limit_contacts_by_email:(i+1)*limit_contacts_by_email]) for i in range(0,n_emails_will_be_sent) ]

        if send:
            self.send()

    def display(self):
        mail = self.create_mail_instance()
        mail.Display()

    def create_mail_instance(self):
        mail = outlook.CreateItemFromTemplate(self.email_template_path)
        mail.HTMLBody = mail.HTMLBody
        mail.to = self.to if self.to_input != 0 else mail.to
        mail.cc = self.cc if self.cc_input != 0 else mail.cc
        mail.bcc = self.bcc if self.bcc_input != 0 else mail.bcc
        mail.SentOnBehalfOfName = self.send_on_behalf if self.send_on_behalf != 0 else mail.SentOnBehalfOfName

        for prop, value in self.mail_properties.items():
            setattr(mail, prop, value)

        return mail

    def send(self):
        if len(self.packs) == 0:
            self.send_one()
        else:
            self.send_pack()

    def send_one(self):
        mail = self.create_mail_instance()
        mail.Send()

    def send_pack(self):
        for pack in self.packs:
            mail = self.create_mail_instance()
            setattr(mail, self.limit_contacts_type+"_list", pack)
            mail.Send()


    def contacts_to_list(self, contacts):
        if type(contacts) == str:
            if contacts[-4:] == "xlsx":
                contacts = self.read_contacts_from_excel(contacts)
            else:
                contacts = re.findall(r'[\w\.-]+@[\w\.-]+', contacts)
        elif type(contacts) == list:
            pass
        elif contacts == 0:
            #It means default value
            contacts = []
        else:
            raise ValueError("Format input from contacts information not identified")

        contacts = self.remove_invalid_email(contacts)
        return contacts


    @staticmethod
    def read_contacts_from_excel(contacts_file_path):
        emails_o = pd.read_excel(contacts_file_path)
        emails = emails_o.rename(columns={emails_o.columns[0]: "Email"}).copy()
        emails = emails["Email"].copy()
        return emails.unique()

    @staticmethod
    def remove_invalid_email(contact_list):
        emails = " ".join(contact_list).lower()
        emails = re.findall(r'[\w\.\+\-]+\@[\w]+\.[a-z]{2,3}', emails)
        emails = [ email for email in emails if validate_email(email)]
        return emails

    

