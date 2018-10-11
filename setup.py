from distutils.core import setup

setup(
    name='o365_sharepoint_connector',
    version='0.1.0',
    packages=['o365_sharepoint_connector'],
    url='https://github.com/Bystroushaak/Office365SharepointConnector',
    license='MIT',
    author='Bystroushaak',
    author_email='bystrousak@kitakitsune.org',
    description='Class based API for the Office365 Sharepoint',
    install_requires=[
        "lxml",
        "requests",
    ]
)
