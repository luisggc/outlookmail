OutlookMail
=========

Send Emails via Outlook using Python.
Usage examples in `ReadTheDocs Usage`_.

* Free software: MIT license
* Documentation: https://outlookmail.readthedocs.io.

Warning: This is a project in a alpha version.

A Quick Introduction
--------------------

Outlookmail is a python package to send emails in Windows using Outlook::

       email = om.Mail(
            email_template_path= r"path_to_template_file.msg",
            to= "johndoe@hotmail.com; another_email@email.com",
            cc=["johndoe2@hotmail.com; another_email2@email.com"],
            bcc= r"path_to_excel_with_emails.xlsx"
        )
        email.display()
        email.send()

Features
--------

- Send emails using list of contacts in Excel
- Break into multiple e-mails to avoid limit of contacts
- Use e-mail template


TODO
----

- Email scheduler
- Create variables to send custom e-mails


Installation
------------

To install outlookmail, run this command in your terminal:

.. code-block:: console

    $ pip install outlookmail

This is the preferred method to install outlookmail, as it will always install the most recent stable release.
If you don't have `pip`_ installed, this `Python installation guide`_ can guide
you through the process.

- Python installation guide: http://docs.python-guide.org/en/latest/starting/installation/
- pip: https://pip.pypa.io


Credits
-------

This package was created using:
- Cookiecutter: https://github.com/audreyr/cookiecutter
- audreyr/cookiecutter-pypackage: https://github.com/audreyr/cookiecutter-pypackage
- Pandas: https://github.com/pandas-dev/pandas
- Validate Email: https://github.com/syrusakbary/validate_email
- ReadTheDocs Usage: https://outlookmail.readthedocs.io/en/latest/usage.html