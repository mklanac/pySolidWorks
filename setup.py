# https://packaging.python.org/tutorials/packaging-projects/

import setuptools

with open('README.md', 'r') as fh:
    long_description = fh.read()

setuptools.setup(
    name='pysolidworks',  
    version='0.1.0',
    author='Mario Klanac',
    author_email='mario.klanac1995@gmail.com',
    description='A python interface to SolidWorks.',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/mklanac/pySolidworks',
    packages=['pysolidworks', 'pysolidworks.test'],
    include_package_data=True,
    package_data={'': ['data/*.yml']},
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)',
        'Operating System :: Windows',
    ],
    install_requires=[
      'pywin32',
	  'pythoncom',
      'numpy',
    ],
    python_requires=">=3.7"
 )
