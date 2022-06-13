#!/usr/bin/env python3
import io
from glob import glob
from os.path import basename, splitext
from setuptools import find_packages, setup


setup(
    name='gvmtool',
    version='0.1.0',
    description='GVM csv merge tool',
    author='Hosub Lee',
    author_email='spyrr83@gmail.com',
    install_requires=[
        'click==8.1.3',
        'et-xmlfile==1.1.0',
        'numpy==1.22.4',
        'openpyxl==3.0.10',
        'pandas==1.4.2',
        'python-dateutil==2.8.2',
        'pytz==2022.1',
        'six==1.16.0',
    ],
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    py_modules=[splitext(basename(path))[0] for path in glob('src/*.py')],
    keywords=['GVM', 'report', 'csv'],
    python_requires='>=3.6',
    # include_package_data=True,
    # package_data={
    #     '': ['template.xlsm',],
    # },
    data_files=[('data', ['src/data/template.xlsm',]),],
    entry_points={
        'console_scripts': [
            'gvmtool = gvmtool:main', 
        ],
    },
    zip_safe=False,
    classifiers=[
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ]
)
