=========
Outlook Mail
=========


.. image:: https://img.shields.io/pypi/v/outlookmail.svg
        :target: https://pypi.python.org/pypi/outlookmail

.. image:: https://img.shields.io/travis/luisggc/outlookmail.svg
        :target: https://travis-ci.org/luisggc/outlookmail

.. image:: https://readthedocs.org/projects/outlookmail/badge/?version=latest
        :target: https://outlookmail.readthedocs.io/en/latest/?badge=latest
        :alt: Documentation Status




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

.. _pip: https://pip.pypa.io
.. _Python installation guide: http://docs.python-guide.org/en/latest/starting/installation/



Credits
-------

Most of vectorized calculus made with Numpy_, unit conversion with Pint_, all plots with Matplotlib_ (export to dxf with ezdxf_), detect peaks with py-findpeaks_, 
docs made with the help of Sphinx_ and Numpydoc_, analysis table with Pandas_,  
this package was created with Cookiecutter_ and the `audreyr/cookiecutter-pypackage`_ project template.

.. _Cookiecutter: https://github.com/audreyr/cookiecutter
.. _`audreyr/cookiecutter-pypackage`: https://github.com/audreyr/cookiecutter-pypackage
.. _Pandas: https://github.com/pandas-dev/pandas
.. _`ReadTheDocs Usage`: https://outlookmail.readthedocs.io/en/latest/usage.html