# -*- coding: utf-8 -*-
"""Installer for the ftw.flamegraph package."""

from setuptools import find_packages
from setuptools import setup


tests_require = [
    'pytest',
]

setup(
    name='docxcompose',
    version='1.1.2',
    description="Compose .docx documents",
    long_description=(open("README.rst").read() + "\n" +
                      open("HISTORY.txt").read()),
    # Get more from https://pypi.python.org/pypi?%3Aaction=list_classifiers
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
        "License :: OSI Approved :: MIT License",
    ],
    keywords='Python DOCX Word OOXML',
    author='Thomas Buchberger',
    author_email='t.buchberger@4teamwork.ch',
    url='https://github.com/4teamwork/docxcompose',
    license='MIT license',
    packages=find_packages(exclude=['ez_setup']),
    include_package_data=True,
    zip_safe=True,
    install_requires=[
        'lxml',
        'python-docx >= 0.8.8',
        'setuptools',
        'six',
    ],
    extras_require={
        'test': tests_require,
        'tests': tests_require,
    },
    entry_points={
        'console_scripts': [
            'docxcompose = docxcompose.command:main'
      ]
  },
)
