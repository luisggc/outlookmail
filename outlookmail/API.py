from outlookmail import Mail
import os
from datetime import datetime
import io


class API:
    @staticmethod
    def read_from_txt(file_path, success_message=False, do_on_send_success=lambda x:x, do_on_send_failure=lambda x:x):
        #file_data = open(file_path,'r')
        with io.open(file_path, mode="r", encoding="utf-8") as file_data:
            txt = file_data.read()
        args = { row.split("=")[0].strip(): "=".join(row.split("=")[1:]).strip() for row in txt.split("\n") }
        args.pop("", None)
        if args.get("when"):
            when = datetime.strptime(args["when"], '%m/%d/%Y-%H:%M:%S')
            if when >= datetime.now():
                print("'When' not reached")
                return

        mail = Mail(**args)
        try:
            mail.send()
            if success_message:
                print("\nEmail was sent. Please check your 'sent' folder in Outlook.\n")
            if do_on_send_success:
                do_on_send_success()
        except Exception as e:
            raise ValueError("Outlook not available. Error: {}".format(e))
            if do_on_send_failure:
                do_on_send_failure()

