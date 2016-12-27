import json
import os

import requests
from requests_ntlm import HttpNtlmAuth

headers = {
    "GET": {
        "Accept": "application/json;odata=verbose"
    },
    "POST": {
        "Accept": "application/json;odata=verbose",
        'X-RequestDigest': "",
        'Content-Type': "application/json;odata=verbose",
    },
    "PUT": {
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": '',
        "Content-Type": "application/json;odata=verbose",
        "X-HTTP-Method": "PATCH",
        "If-Match": "*",
    },
    "DELETE": {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": "",
        "X-HTTP-Method": "DELETE",
        "If-Match": "*"
    },
}


class SharePointConnector:
    """
    Class responsible fro performing most common SharePoint Operations.
    Use also to authenticate access to the SharepointSite and to get a digest value for POST requests.
    """

    def __init__(self, login, password, base_url, domain="eur"):
        self.session = requests.Session()
        self.base_url = base_url + "/"
        self.session.auth = HttpNtlmAuth("{}\\{}".format(domain, login), "{}".format(password))
        self.success_list = [200, 201, 202]

    def get_all_lists(self):
        """
        Gets all lists.

        :return: Returns a REST response.
        """
        get = self.session.get(
            self.base_url + "_api/web/lists?$top=5000",
            headers=headers["GET"]
        )
        print("Get all list.")
        print("GET: {}".format(get.status_code))
        if get.status_code not in self.success_list:
            print(get.content)
        else:
            return get.json()["d"]["results"]

    def create_new_list(self, data=None, list_name="new_list", description="", allow_content_types=True,
                        base_template=100, content_types_enabled=True):
        """
        Use to create new SharePoint List.
        By default creates new List of any Type with "new_list" name and blank name.

        Basic Types:
            100	Custom list
            101	Document library
            102	Survey
            103	Links
            104	Announcements
            105	Contacts
            106	Calendar
            107	Tasks
            108	Discussion board
            109	Picture library
            110	Data sources for a site
            111	Site template gallery
            112	User Information
            113	Web Part gallery


        :param data: Optional Parameter when you need to use your own data
        :param list_name: Name of new List - Optional, by default set to "new_list".
        :param description: Description of the list - Optional, by default set to blank.
        :param base_template: Optional, determines the list type
        :param allow_content_types: Optional
        :param content_types_enabled: Optional
        :return: Returns a REST response.
        """
        headers["POST"]["X-RequestDigest"] = self.digest()
        if data is None:
            data = {
                '__metadata': {'type': 'SP.List'},
                'AllowContentTypes': allow_content_types,
                'BaseTemplate': base_template,
                'ContentTypesEnabled': content_types_enabled,
                'Description': '{}'.format(description),
                'Title': '{}'.format(list_name)
            }
        post = self.session.post(
            self.base_url + "_api/web/lists",
            headers=headers["POST"],
            data=json.dumps(data)
        )
        print("Create new list - {}.".format(list_name))
        print("POST: {}".format(post.status_code))
        if post.status_code not in self.success_list:
            print(post.content)

    def create_new_list_field(self, list_name, data=None, filed_name="new_field", field_type=2):
        """
        Creates new column fields in SharepointList
        By default creates new Text field with "new_field" name.

        :param list_name: Required, provide the name of the list you want to modify as String.
        :param data: Optional Parameter when you need to use your own data
        :param filed_name: Optional, the name of new field as String, by default set to "new_field"
        :param field_type: Please choose a field type as Integer, by default set to text field.

        Field Types:
        0   Invalid             - Not used. Value = 0.
        1   Integer             - Field allows an integer value.
        2   Text                - Field allows a limited-length string of text.
        3   Note                - Field allows larger amounts of text.
        4   DateTime	        - Field allows full date and time values, as well as date-only values.
        5   Counter             - Counter is a monotonically increasing integer field, and has a unique value in
                                  relation to other values that are stored for the field in the list.
                                  Counter is used only for the list item identifier field, and not intended for use
                                  elsewhere.
        6   Choice              - Field allows selection from a set of suggested values.
                                  A choice field supports a field-level setting which specifies whether free-form
                                  values are supported.
        7   Lookup              - Field allows a reference to another list item. The field supports specification of a
                                  list identifier for a targeted list. An optional site identifier can also be
                                  specified, which specifies the site of the list which contains the target of the
                                  lookup.
        8   Boolean 	        - Field allows a true or false value.
        9   Number              - Field allows a positive or negative number.
                                  A number field supports a field level setting used to specify the number
                                  of decimal places to display.
        10  Currency            - Field allows for currency-related data. The Currency field has a
                                  CurrencyLocaleId property which takes a locale identifier of the currency to use.
        11  URL	                - Field allows a URL and optional description of the URL.
        12  Computed	        - Field renders output based on the value of other columns.
        13  Threading	        - Contains data on the threading of items in a discussion board.
        14  Guid                - Specifies that the value of the field is a GUID.
        15  MultiChoice	        - Field allows one or more values from a set of specified choices.
                                  A MultiChoice field can also support free-form values.
        16  GridChoice	        - Grid choice supports specification of multiple number scales in a list.
        17  Calculated          - Field value is calculated based on the value of other columns.
        18  File                - Specifies a reference to a file that can be used to retrieve the contents of that
                                  file.
        19  Attachments         - Field describes whether one or more files are associated with the item.
                                  See Attachments for more information on attachments.
                                  true if a list item has attachments, and false if a list item does not have
                                  attachments.
        20  User                - A lookup to a particular user in the User Info list.
        21  Recurrence	        - Specifies whether a field contains a recurrence pattern for an item.
        22  CrossProjectLink    - Field allows a link to a Meeting Workspace site.
        23  ModStat             - Specifies the current status of a moderation process on the document.
                                  Value corresponds to one of the moderation status values.
        24  Error               - Specifies errors. Value = 24.
        25  ContentTypeId       - Field contains a content type identifier for an item. ContentTypeId
                                  conforms to the structure defined in ContentTypeId.
        26  PageSeparator       - Represents a placeholder for a page separator in a survey list.
                                  PageSeparator is only intended to be used with a Survey list.
        27  ThreadIndex	        - Contains a compiled index of threads in a discussion board.
        28  WorkflowStatus      - No Information.
        29  AllDayEvent         - The AllDayEvent field is only used in conjunction with an Events list. true if the
                                  item is an all day event (that is, does not occur during a specific
                                  set of hours in a day).
        30  WorkflowEventType   - No Information.
        31  MaxItems	        - Specifies the maximum number of items. Value = 31.

        :return: Returns REST response.
        """
        # Updates headers by new Digest Value.
        headers["POST"]["X-RequestDigest"] = self.digest()
        # Sets data
        if data is None:
            data = {
                '__metadata': {'type': 'SP.Field'},
                'Title': str(filed_name),
                'FieldTypeKind': field_type
            }
        # Performs REST request
        post = self.session.post(
            self.base_url + "_api/web/lists/GetByTitle('{}')".format(list_name) + "/fields",
            headers=headers["POST"],
            data=json.dumps(data)
        )
        print("Create new list header of name {} and type {} for {}.".format(filed_name, field_type, list_name))
        print("POST: {}".format(post.status_code))
        if post.status_code not in self.success_list:
            print(post.content)

    def update_list(self, list_guid, data=None):
        """
        Updates a SharepointList Information
        By default changes only a list Title.

        :param list_guid: Required, individual id of the List you want to Modify
        :param data: Optional Parameter when you need to use your own data
        :return: Returns a REST response.
        """
        headers["PUT"]["X-RequestDigest"] = self.digest()
        put = self.session.post(
            self.base_url + "_api/web/lists(guid'{}')".format(list_guid),
            headers=headers["PUT"],
            data=json.dumps(data)
        )
        print("Update list name for list of GUID: {}".format(list_guid))
        print("PUT: {}".format(put.status_code))
        if put.status_code not in self.success_list:
            print(put.content)

    def delete_list(self, list_guid):
        """
        Deletes a Sharepoint List by its GUID.

        :param list_guid: Required, individual id of Sharepoint List.
        :return: Returns a REST response.
        """
        headers["DELETE"]["X-RequestDigest"] = self.digest()
        delete = self.session.delete(
            self.base_url + "_api/web/lists(guid'{}')".format(list_guid),
            headers=headers["DELETE"]
        )
        print("Delete list of GUID: {}".format(list_guid))
        print("DELETE: {}".format(delete.status_code))
        if delete.status_code not in self.success_list:
            print(delete.content)

    def get_list_items(self, list_name):
        """
        Gets all List Items from Sharepoint List of given Name

        :param list_name: Required, name of the list from which items will be downloaded.
        :return: Returns REST response.
        """
        get = self.session.get(
            self.base_url + "_api/web/lists/GetByTitle('{}')".format(list_name) + "/items?$top=5000",
            headers=headers["GET"]
        )
        print("Get list items from {}.".format(list_name))
        print("GET: {}".format(get.status_code))
        if get.status_code not in self.success_list:
            print(get.content)
        else:
            return get.json()["d"]["results"]

    def create_new_list_item(self, list_name, data=None):
        """
        Creates a new List item in the list of given name.

        :param list_name: Required, name of the list in which items will be created.
        :param data: Optional Parameter when you need to use your own data
        :return: Returns a REST response.
        """
        headers["POST"]['X-RequestDigest'] = self.digest()
        if data is None:
            data = {
                '__metadata': {
                    'type': 'SP.Data.{}ListItem'.format(list_name.title())
                },
                'Title': 'New_list_Item'
            }
        post = self.session.post(
            self.base_url + "_api/web/lists/GetByTitle('{}')".format(list_name) + "/items",
            data=json.dumps(data),
            headers=headers["POST"]
        )
        print("Create new list item in {}.".format(list_name))
        print("POST: {}".format(post.status_code))
        if post.status_code not in self.success_list:
            print(post.content)

    def update_list_item(self, list_name, item_id=0, data=None):
        """
        Updates already existing SharePoint list item.

        :param list_name: Required, name of the list in which item is stored.
        :param item_id: Required, an individual id of the item in the list.
        :param data: Required, provide a data by which the item will be updated
        :return: Returns a REST response.
        """
        headers["PUT"]['X-RequestDigest'] = self.digest()
        put = self.session.post(
            self.base_url + "+api/web/lists/GetByTitle('{}')".format(list_name) + "/items('{}')".format(item_id),
            data=json.dumps(data),
            headers=headers["PUT"]
        )
        print("Update list item of id {} in {}.".format(item_id, list_name))
        print("PUT: {}".format(put.status_code))
        if put.status_code not in self.success_list:
            print(put.content)

    def delete_list_item(self, list_name, item_id=0):
        """
        Deletes a list item in SharePoint list of given name.

        :param list_name: Required, name of the list in which item is stored.
        :param item_id: Required, an individual id of the item in the list.
        :return: Returns a REST response.
        """
        headers["DELETE"]["X-RequestDigest"] = self.digest()
        delete = self.session.delete(
            self.base_url + "_api/web/lists/GetByTitle('{}')".format(list_name) + "/items('{}')".format(item_id),
            headers=headers["DELETE"]
        )
        print("Delete list item of id {} in {}.".format(item_id, list_name))
        print("DELETE: {}".format(delete.status_code))
        if delete.status_code not in self.success_list:
            print(delete.content)

    # Add functions related to document libraries and lists attachments
    def get_folder_information(self, folder_name):
        """
        Gets all information about given folder directory.

        :param folder_name:  Required, name of the folder
        :return: Returns REST response
        """
        get = self.session.get(
            self.base_url + "_api/web/GetFolderByServerRelativeUrl('/{}')".format(folder_name),
            headers=headers["GET"]
        )
        print("Get information for {} folder.".format(folder_name))
        print("GET: {}".format(get.status_code))
        if get.status_code not in self.success_list:
            print(get.content)
        else:
            return get.json()["d"]

            # Add functions related to file manipulation

    def upload_file(self, file_path, destination_library):
        """
        Uploads a file to given library/folder.

        :param file_path: Required, file as path
        :param destination_library: Required, destination of upload
        :return: Returns REST response
        """
        headers["POST"]["X-RequestDigest"] = self.digest()
        file = open(file_path, "rb")
        file_as_bytes = bytearray(file.read())

        post = self.session.post(
            self.base_url + "_api/web/GetFolderByServerRalativeUrl('/{}')/Files/add(url='{}',overwrite=true)".format(
                destination_library,
                os.path.basename(file.name)
            ),
            data=file_as_bytes,
            headers=headers["POST"]
        )
        print("Add file '{}' to library '{}'.".format(
            os.path.basename(file.name),
            destination_library)
        )
        print("POST: {}".format(post.status_code))
        if post.status_code not in self.success_list:
            print(post.content)
        else:
            return post.json()["d"]

    def digest(self):
        """
        Helper function.
        Gets a digest value for POST requests.

        :return: Returns a REST response.
        """
        data = self.session.post(
            self.base_url + "_api/contextinfo",
            headers=headers["GET"]
        )
        return data.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    def authenticate(self):
        """
        Checks users authentication.
        Returns True/False dependently of user access.

        :return: Boolean
        """
        data = self.session.get(
            self.base_url,
            headers=headers["GET"]
        )
        if data.status_code == 200:
            return True
        else:
            return False


class SharePointDataParser:
    def list_item_data(self, list_name, data):
        output_data = {
            '__metadata': {
                'type': self.list_item_meta(list_name)
            },
        }
        for key, value in data.items():
            output_data[key] = value
        return output_data

    @staticmethod
    def list_data(data, allow_content_types=True, base_template=100, content_types_enabled=True):
        output_data = {
            '__metadata': {
                'type': 'SP.List'
            },
            'AllowContentTypes': allow_content_types,
            'BaseTemplate': base_template,
            'ContentTypesEnabled': content_types_enabled
        }
        for key, value in data.items():
            output_data[key] = value
        return output_data

    @staticmethod
    def folder_data(data):
        output_data = {
            '__metadata': {
                'type': 'SP.Folder'
            },
        }
        for key, value in data.items():
            output_data[key] = value
        return output_data

    @staticmethod
    def list_field_data(data):
        pass

    @staticmethod
    def list_item_meta(list_name):
        return "SP.Data." + list_name[0].upper() + list_name[1::] + "ListItem"