o365_sharepoint_connector
`````````````````````````

About
+++++

o365_sharepoint_connector is a Python 3 module to perform HTTPS requests to Microsoft Office365 Sharepoint sanely, with class based ORM mapping.

Features
++++++++

o365_sharepoint_connector utilizes Microsoft Sharepoint REST architecture to work with lists, views, folders and files.

Installation::

    pip3 install --user o365_sharepoint_connector

Usage
+++++


.. code-block:: python

    import o365_sharepoint_connector
    
    connector = o365_sharepoint_connector.SharePointConnector(
        login="o365_login@your_company.cz",
        password=open("pass").read(),
        site_url="https://your_company.sharepoint.com/sites/Documents"
    )
    connector.authenticate()
    
    lib_list = connector.get_lists()["Lib"]
    view = lib_list.get_views()["Všechny dokumenty"]
    folder = view.get_folders()["Xexex"]
    
    with open("0.pdf", "wb") as f:
        remote_file = folder.get_files()["0.pdf"]
        f.write(remote_file.get_content())
    
    poc = folder.upload_file("poc2.py")
    print(poc.server_relative_path)


Credit
++++++

This fork is based on `EasySharePoint <https://github.com/krzysztofgrowinski/EasySharePoint>`_ library.
