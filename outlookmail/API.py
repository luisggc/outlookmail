from outlookmail import Mail
import os
from datetime import datetime
import io


class API:
    @staticmethod
    def read_from_txt(file_path, delete_after_sent=False):
        #file_data = open(file_path,'r')
        file_data = io.open(file_path, mode="r", encoding="utf-8")
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
        except Exception as e:
            raise ValueError("Outlook not available. Error: {}".format(e))
            if delete_after_sent:
                os.remove(file_path)

