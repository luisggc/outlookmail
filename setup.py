#!/usr/bin/env python

"""The setup script."""

from setuptools import setup, find_packages

with open('README.rst') as readme_file:
    readme = readme_file.read()

with open('HISTORY.rst') as history_file:
    history = history_file.read()

requirements = ['Click>=7.0',
'validate_email==1.3',
'pandas==1.3.2'
]

test_requirements = [ ]

setup(
    author="Luis Gabriel GonÃ§alves Coimbra",
    author_email='luiscoimbraeng@outlook.com',
    python_requires='>=3.6',
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
    ],
    description="Make automatic many email related tasks in Outlook",
    entry_points={
        'console_scripts': [
            'outlookmail=outlookmail.cli:main',
        ],
    },
    install_requires=requirements,
    license="MIT license",
    long_description=readme + '\n\n' + history,
    include_package_data=True,
    keywords='outlookmail',
    name='outlookmail',
    packages=find_packages(include=['outlookmail', 'outlookmail.*']),
    test_suite='tests',
    tests_require=test_requirements,
    url='https://github.com/luiggc[D[D[Dsggc/outlookmail',
    version='0.0.1',
    zip_safe=False,
)
