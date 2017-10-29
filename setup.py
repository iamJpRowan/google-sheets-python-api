#!/usr/bin/env python3
import os

# See http://pythonhosted.org/distribute/ for information about the distribute
# bootstrap file imported on the next line.

from setuptools import setup, find_packages


def list_of_files_in_directory(path):
    return [
        os.path.join(path, x) for x in os.listdir(path)
    ]


scripts = list_of_files_in_directory('bin')

setup(
    name='Google Sheets Python Library',
    version='0.1dev',
    long_description=open('README.md').read(),
    author="Jp Rowan",
    author_email="jp@jprowan.com",
    url="https://github.com/LiveSlowSkateFast/google-sheets-python-api",
    license='Creative Commons Attribution-Noncommercial-Share Alike license',

    # what to include in the package
    packages=find_packages(),
    scripts=scripts,
    # dependencies (to be automatically installed or updated)
    install_requires=[
        'httplib2',
        'apiclient',
        'oauth2client',
    ]
)
