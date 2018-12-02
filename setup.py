from distutils.core import setup

setup(
    name='o365_sharepoint_connector',
    version='20181202',
    packages=['o365_sharepoint_connector'],
    url='https://github.com/Bystroushaak/Office365SharepointConnector',
    license='MIT',
    author='Bystroushaak',
    author_email='bystrousak@kitakitsune.org',
    description='Py3 class based API for the Office365 Sharepoint.',
    install_requires=[
        "setuptools",
        "lxml",
        "requests",
    ]
)
