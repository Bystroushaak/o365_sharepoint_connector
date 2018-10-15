import os
import os.path
import json
import logging
from urllib.parse import urlparse
from urllib.parse import urlunparse

import requests
from lxml import etree


logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
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


class SharepointException(Exception):
    pass


class LoginException(SharepointException):
    pass


class CantCreateNewListException(SharepointException):
    pass


class CantCreateNewFieldException(SharepointException):
    pass


class CantChangeFieldIndexException(SharepointException):
    pass


class ListingException(SharepointException):
    pass


class UpdateException(SharepointException):
    pass


class DeleteException(SharepointException):
    pass


class UploadException(SharepointException):
    pass


class CheckInException(SharepointException):
    pass


class CheckOutException(SharepointException):
    pass


class _SharepointElementBase:
    def __init__(self):
        self._connector = None

    def _compose_url(self, *args):
        return self._connector._compose_url(*args)

    def _parse_exception(self, keywords):
        exception = None
        if "exception" in keywords:
            exception = keywords["exception"]
            del keywords["exception"]

        return exception, keywords

    def get(self, url, *url_params, **keywords):
        url = self._compose_url(url, *url_params)
        logger.debug("URL: %s", url)

        exception, keywords = self._parse_exception(keywords)

        result = self._connector.session.get(url, **keywords)
        logger.debug("GET: %s", result.status_code)

        if exception and result.status_code not in self._connector.success_list:
                raise exception(result.content)

        return result

    def post(self, url, *url_params, **keywords):
        url = self._compose_url(url, *url_params)
        logger.debug("URL: %s", url)

        exception, keywords = self._parse_exception(keywords)

        result = self._connector.session.post(url, **keywords)
        logger.debug("POST: %s", result.status_code)

        if exception and result.status_code not in self._connector.success_list:
                raise exception(result.content)

        return result

    def delete(self, url, *url_params, **keywords):
        url = self._compose_url(url, *url_params)
        logger.debug("URL: %s", url)

        exception, keywords = self._parse_exception(keywords)

        result = self._connector.session.session.delete(url, **keywords)
        logger.debug("DELETE: %s", result.status_code)

        if exception and result.status_code not in self._connector.success_list:
                raise exception(result.content)

        return result


class SharepointFile(_SharepointElementBase):
    def __init__(self):
        self.raw_data = None
        self._connector = None

        self.name = ""
        self.uuid = ""
        self.exists = True
        self.content_tag = ""
        self.linking_url = ""
        self.time_created = ""
        self.check_out_type = ""
        self.check_in_comment = ""
        self.time_last_modified = ""
        self.server_relative_url = ""
        self.folder_relative_url = ""

    @classmethod
    def from_dict(cls, connector, folder_relative_url, data):
        sfile = cls()
        sfile.raw_data = data
        sfile._connector = connector
        sfile.folder_relative_url = folder_relative_url

        sfile.name = data["Name"]
        sfile.uuid = data["UniqueId"]
        sfile.exists = data["Exists"]
        sfile.content_tag = data["ContentTag"]
        sfile.linking_url = data["LinkingUrl"]
        sfile.time_created = data["TimeCreated"]
        sfile.check_out_type = data["CheckOutType"]
        sfile.check_in_comment = data["CheckInComment"]
        sfile.time_last_modified = data["TimeLastModified"]
        sfile.server_relative_url = data["ServerRelativeUrl"]

        return sfile

    def __repr__(self):
        return "SharepointFile(%s)" % self.server_relative_url

    def check_in(self, comment="", check_in_type=0):
        logger.info(
            "CheckIn file '%s' in library '%s' with comment '%s'.",
            os.path.basename(self.name),
            self.server_relative_url,
            comment
        )

        headers["POST"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/GetFileByServerRelativeUrl('/{}/{}')/CheckIn(comment='{}',checkintype={})",
            self.folder_relative_url,
            self.name,
            comment,
            check_in_type,
            headers=headers["POST"],
            exception=CheckInException
        )

    def check_out(self):
        logger.info("CheckOut file '%s' in library '%s'.", os.path.basename(self.name), self.folder_relative_url)

        headers["POST"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/GetFileByServerRelativeUrl('/{}/{}')/CheckOut()",
            self.folder_relative_url,
            self.name,
            headers=headers["POST"],
            exception=CheckOutException
        )

    def get_content(self):
        """
        Gets file from folder/library as binary
        """
        logger.info("Get %s from %s.", self.name, self.server_relative_url)

        self.get(
            "_api/web/GetFolderByServerRelativeUrl('{}')/Files('{}')/$value",
            self.folder_relative_url,
            self.name,
            headers=headers["GET"],
            exception=ListingException
        )

        return get.content

    def delete(self, file_name, destination_library):
        """
        Deletes a file in given library/folder.

        :param file_name: Required, file name to delete
        :param destination_library: Required, folder where file exists
        :return: Returns REST response
        """
        logger.info("Delete file '%s' from library '%s'.", os.path.basename(self.name), self.folder_relative_url)

        headers["DELETE"]["X-RequestDigest"] = self._connector.digest()
        self.delete(
            "_api/web/GetFileByServerRelativeUrl('/{}/{}')",
            self.folder_relative_url,
            self.name,
            headers=headers["DELETE"],
            exception=DeleteException
        )


class SharepointFolder(_SharepointElementBase):
    def __init__(self):
        self.raw_data = None
        self._connector = None

        self.name = ""
        self.exists = True
        self.unique_id = ""
        self.item_count = 0
        self.time_created = ""
        self.time_last_modified = ""
        self.server_relative_url = ""

    @classmethod
    def from_dict(cls, connector, data):
        sdir = cls()
        sdir.raw_data = data
        sdir._connector = connector

        sdir.name = data["Name"]
        sdir.exists = data["Exists"]
        sdir.unique_id = data["UniqueId"]
        sdir.item_count = data["ItemCount"]
        sdir.time_created = data["TimeCreated"]
        sdir.time_last_modified = data["TimeLastModified"]
        sdir.server_relative_url = data["ServerRelativeUrl"]

        return sdir

    def __repr__(self):
        return "SharepointDir(%s)" % self.name

    def get_files(self):
        """
        Gets all files from given library/folder
        """
        logger.info("Get all files from %s.", self.server_relative_url)

        get = self.get(
            "_api/web/GetFolderByServerRelativeUrl('/{}')/Files",
            self.server_relative_url,
            headers=headers["GET"],
            exception=ListingException
        )

        return {
            x["Name"]: SharepointFile.from_dict(self._connector, self.server_relative_url, x)
            for x in get.json()["d"]["results"]
        }

    def upload_file(self, local_file_path):
        """
        Uploads a file.

        :return: SharepointFile
        """
        logger.info("Upload file '%s' to folder '%s'.", local_file_path, self.server_relative_url)

        headers["POST"]["X-RequestDigest"] = self._connector.digest()

        with open(local_file_path, "rb") as f:
            file_as_bytes = bytearray(f.read())

        post = self.post(
            "_api/web/GetFolderByServerRelativeUrl('/{}')/Files/add(url='{}',overwrite=true)",
            self.server_relative_url,
            os.path.basename(local_file_path),
            data=file_as_bytes,
            headers=headers["POST"],
            exception=UploadException
        )

        return SharepointFile.from_dict(
            self._connector,
            self.server_relative_url,
            post.json()["d"]
        )


class SharepointView(_SharepointElementBase):
    def __init__(self):
        self.id = ""
        self.title = ""
        self.paged = False
        self.hidden = False
        self.list_id = None
        self.read_only = False
        self.server_relative_url = ""

        self.raw_data = None
        self._connector = None

    @classmethod
    def from_dict(cls, connector, list_id, data):
        view = cls()
        view.list_id = list_id
        view.raw_data = data
        view._connector = connector

        view.id = data["Id"]
        view.paged = data["Paged"]
        view.title = data["Title"]
        view.hidden = data["Hidden"]
        view.read_only = data["ReadOnlyView"]
        view.server_relative_url = data["ServerRelativeUrl"]

        return view

    def __repr__(self):
        return "SharepointView(%s)" % self.title

    def add_field(self, field_name):
        """
        Adds a specific field to the ListView.

        :param field_name: Required, field name to be added
        """
        logging.info("Add %s field to the view.", field_name)

        headers["POST"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/lists(guid'{}')/views(guid'{}')/viewfields/addviewfield('{}')",
            self.list_id,
            self.id,
            field_name,
            headers=headers["POST"],
            exception=CantCreateNewFieldException
        )

    def change_field_index(self, field_name, field_index):
        """
        Adds a specific field to the View.

        :param field_name: Required, field name to be added
        """
        logger.info("Moved %s field to the index %s.", field_name, field_index)

        headers["POST"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/lists(guid'{}')/views(guid'{}')/viewfields/moveviewfieldto",
            self.list_id,
            self.id,
            headers=headers["POST"],
            data=json.dumps({"field": field_name, "index": field_index}),
            exception=CantChangeFieldIndexException
        )

    def remove_field(self, field_name):
        """
        Removes a specific field to the View.

        :param field_name: name of the field to be removed
        """
        logger.info("Remove %s field to the view.", field_name)

        headers["DELETE"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/lists(guid'{}')/views(guid'{}')/viewfields/removeviewfield('{}')",
            self.list_id,
            self.id,
            field_name,
            headers=headers["DELETE"],
            exception=DeleteException
        )

    def remove_all_fields(self):
        """
        Removes all fields from View.
        """
        logger.info("Remove all fields from the view.")

        headers["DELETE"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/lists(guid'{}')/views(guid'{}')/viewfields/removeallviewfields",
            self.list_id,
            self.id,
            headers=headers["DELETE"],
            exception=DeleteException
        )

    def get_folders(self):
        """
        List folders in given relative url.
        """
        relative_url = self.server_relative_url.replace("/Forms/AllItems.aspx", "/")
        logger.info("Get list of folders for %s.", relative_url)

        get = self.get(
            "_api/web/GetFolderByServerRelativeUrl('{}')/Folders",
            relative_url,
            headers=headers["GET"],
            exception=ListingException
        )

        return {
            x["Name"]: SharepointFolder.from_dict(self._connector, x)
            for x in get.json()["d"]["results"]
        }


class SharepointListItemAttachment(_SharepointElementBase):
    def __init__(self):
        self.raw_data = {}
        self._connector = None

        self.id = ""
        self.title = ""

    @classmethod
    def from_dict(cls, connector, list_title, data):
        item = cls()
        item.raw_data = data
        item._connector = connector
        item.list_title = list_title

        item.id = data["Id"]
        item.title = data["Title"]

        return item

    def __repr__(self):
        return "SharepointListItemAttachment(id=%s)" % self.id

    def update_attachment(self, file_path):
        """
        Updates list item attachment

        :param list_name: Required
        :param item_id: Required
        :param file_path: Required
        :return: Returns REST response
        """
        logger.info(
            "Update file '%s' for list item '%s' in %s.",
            os.path.basename(file_path),
            self.id,
            self.title
        )

        headers["PUT"]["X-RequestDigest"] = self._connector.digest()
        with open(file_path, "rb") as f:
            file_to_bites = bytearray(f.read())

        put = self.post(
            "_api/web/lists/GetByTitle('{}')/items({})/AttachmentFiles('{}')/$value",
            self.title,
            self.id,
            os.path.basename(file_path),
            headers=headers["POST"],
            data=file_to_bites,
            exception=UpdateException
        )

        return put.json()["d"]


class SharepointListItem(_SharepointElementBase):
    def __init__(self):
        self.raw_data = {}
        self._connector = None
        self.list_title = ""

        self.id = None
        self.guid = ""
        self.title = ""
        self.created = ""
        self.modified = ""
        self.author_id = None
        self.editor_id = None
        self.content_type_id = ""
        self.checkout_user_id = None
        self.file_system_object_type = None
        self.server_redirected_embed_uri = None
        self.server_redirected_embed_url = ""

    @classmethod
    def from_dict(cls, connector, list_title, data):
        item = cls()
        item.raw_data = data
        item._connector = connector
        item.list_title = list_title

        item.id = data["Id"]
        item.guid = data["GUID"]
        item.title = data["Title"]
        item.created = data["Created"]
        item.modified = data["Modified"]
        item.author_id = data["AuthorId"]
        item.editor_id = data["EditorId"]
        item.content_type_id = data["ContentTypeId"]
        item.checkout_user_id = data["CheckoutUserId"]
        item.file_system_object_type = data["FileSystemObjectType"]
        item.server_redirected_embed_uri = data["ServerRedirectedEmbedUri"]
        item.server_redirected_embed_url = data["ServerRedirectedEmbedUrl"]

        return item

    def __repr__(self):
        return "SharepointListItem(id=%s)" % self.id

    def delete(self):
        """
        Deletes a list item in SharePoint list of given name.

        :param list_name: Required, name of the list in which item is stored.
        :param item_id: Required, an individual id of the item in the list.
        :return: Returns a REST response.
        """
        logger.info("Delete list item of id %s in %s.", self.id, self.list_title)

        headers["DELETE"]["X-RequestDigest"] = self._connector.digest()
        self.delete(
            "_api/web/lists/GetByTitle('{}')/items('{}')",
            self.list_title,
            self.id,
            headers=headers["DELETE"],
            exception=DeleteException
        )

    def get_attachments(self):
        """
        Retrieves the list of avalible attachments for given list item

        :param list_name: Requiered
        :param item_id: Required
        :return: Returns REST response
        """
        logger.info("Get attachments for item ID: %s from %s list.", self.list_title, self.id)

        get = self.get(
            "_api/web/lists/GetByTitle('{}')/items({})/AttachmentFiles/",
            self.list_title,
            self.id,
            headers=headers["GET"],
            exception=ListingException
        )

        return get.json()["d"]["results"]

    def update_list_item(self, data):
        """
        Updates already existing SharePoint list item.

        :param list_name: Required, name of the list in which item is stored.
        :param item_id: Required, an individual id of the item in the list.
        :param data: Required, provide a data by which the item will be updated
        :return: Returns a REST response.
        """
        logger.info("Update list item of id %s in %s.", self.id, self.list_title)

        headers["PUT"]['X-RequestDigest'] = self._connector.digest()
        put = self.post(
            "+api/web/lists/GetByTitle('{}')/items('{}')",
            self.list_title,
            self.id,
            data=json.dumps(data),
            headers=headers["PUT"],
            exception=UpdateException
        )

        return SharepointListItem.from_dict(self._connector, self.title, put.json()["d"])


class SharepointList(_SharepointElementBase):
    def __init__(self):
        self.id = ""
        self.title = ""
        self.hidden = False
        self.created = None
        self.raw_data = ""
        self.description = ""
        self.entity_type_name = ""
        self.last_item_deleted_date = ""
        self.last_item_modified_date = ""
        self.last_item_user_modified_date = ""

        self._connector = None

    @classmethod
    def from_dict(cls, connector, data):
        slist = cls()
        slist.raw_data = data
        slist._connector = connector

        slist.id = data["Id"]
        slist.title = data["Title"]
        slist.hidden = data["Hidden"]
        slist.created = data["Created"]
        slist.description = data["Description"]
        slist.entity_type_name = data["EntityTypeName"]
        slist.last_item_deleted_date = data["LastItemDeletedDate"]
        slist.last_item_modified_date = data["LastItemModifiedDate"]
        slist.last_item_userModified_date = data["LastItemUserModifiedDate"]

        return slist

    def __repr__(self):
        return "SharepointList(%s)" % self.title

    def add_field(self, field_name, field_type=2):
        """
        Creates new column fields in SharepointList
        By default creates new Text field with "new_field" name.

        :param field_name: The name of new field as String.
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
        """
        logger.info("Create new list header of name %s and type %s for %s.", field_name, field_type, self.title)

        data = {
            '__metadata': {'type': 'SP.Field'},
            'Title': str(field_name),
            'FieldTypeKind': field_type
        }
        headers["POST"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/lists/GetByTitle('{}')/fields",
            self.title,
            headers=headers["POST"],
            data=json.dumps(data),
            exception=CantCreateNewFieldException
        )

    def update_list(self, data):
        """
        Updates a SharepointList Information

        :param data: Parameter when you need to use your own data
        """
        logger.info("Update list name for list of GUID: %s", self.id)

        headers["PUT"]["X-RequestDigest"] = self._connector.digest()
        self.post(
            "_api/web/lists(guid'{}')",
            self.id,
            headers=headers["PUT"],
            data=json.dumps(data),
            exception=UpdateException
        )

    def delete_list(self):
        """
        Deletes a Sharepoint List.
        """
        logger.info("Delete list of GUID: %s", self.id)

        headers["DELETE"]["X-RequestDigest"] = self._connector.digest()
        self.delete(
            "_api/web/lists(guid'{}')",
            self.id,
            headers=headers["DELETE"],
            exception=DeleteException
        )

    def get_views(self):
        """
        Gets all views for list.
        """
        logging.info("Get all list views for %s." % self.id)

        get = self.get(
            "_api/web/lists(guid'{}')/views",
            self.id,
            headers=headers["GET"],
            exception=ListingException
        )

        return {
            x["Title"]: SharepointView.from_dict(self._connector, self.id, x)
            for x in get.json()["d"]["results"]
        }

    def get_items(self):
        """
        Gets all List Items from Sharepoint List of given Name

        :return: Returns REST response.
        """
        logging.info("Get list items from %s.", self.title)

        get = self.get(
            "_api/web/lists/GetByTitle('{}')/items?$top=5000",
            self.title,
            headers=headers["GET"],
            exception=ListingException
        )

        return [
            SharepointListItem.from_dict(self._connector, self.title, x)
            for x in get.json()["d"]["results"]
        ]

    def add_item(self, title):
        """
        Creates a new List item in the list of given name.

        :param list_name: Required, name of the list in which items will be created.
        :param data: Optional Parameter when you need to use your own data
        :return: Returns a REST response.
        """
        logger.info("Create new list item %s in %s.", self.title, self.title)

        headers["POST"]['X-RequestDigest'] = self._connector.digest()
        data = {
            'Title': title,
            '__metadata': {'type': 'SP.Data.{}ListItem'.format(self.title)},
        }
        post = self.post(
            "_api/web/lists/GetByTitle('{}')/items",
            self.title,
            data=json.dumps(data),
            headers=headers["POST"],
            exception=CantCreateNewListException
        )

        return SharepointListItem.from_dict(self._connector, self.title, post.json()["d"])


class SharePointConnector:
    """
    Class responsible for performing most of common SharePoint Operations.
    Use also to authenticate access to the SharepointSite and to get a digest value for POST requests.
    """

    def __init__(self, login, password, site_url, login_url=None):
        self.session = requests.Session()
        self.success_list = [200, 201, 202]

        if not login_url:
            parsed = list(urlparse(site_url))
            parsed[2] = ""  # remove the path from the url
            login_url = urlunparse(parsed)

        self.login_url = login_url if login_url.endswith("/") else login_url + "/"
        self.base_url = site_url if site_url.endswith("/") else site_url + "/"

        self.login = login
        self.password = password

    def _compose_url(self, url, *args):
        new_url = self.base_url + url

        if not args:
            return new_url

        return new_url.format(*args)

    def digest(self):
        """
        Helper function.
        Gets a digest value for POST requests.

        :return: Returns a REST response.
        """
        data = self.session.post(
            self._compose_url("_api/contextinfo"),
            headers=headers["GET"]
        )
        return data.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    def _get_security_token(self):
        """
        Grabs a security Token to authenticate to Office 365 services.

        Inspired by shareplum; https://github.com/jasonrollins/shareplum
        """
        body = """
            <s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
                        xmlns:a="http://www.w3.org/2005/08/addressing"
                        xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
              <s:Header>
                <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>
                <a:ReplyTo>
                  <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address>
                </a:ReplyTo>
                <a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To>
                <o:Security s:mustUnderstand="1"
                   xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
                  <o:UsernameToken>
                    <o:Username>%s</o:Username>
                    <o:Password>%s</o:Password>
                  </o:UsernameToken>
                </o:Security>
              </s:Header>
              <s:Body>
                <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust">
                  <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">
                    <a:EndpointReference>
                      <a:Address>%s</a:Address>
                    </a:EndpointReference>
                  </wsp:AppliesTo>
                  <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>
                  <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>
                  <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>
                </t:RequestSecurityToken>
              </s:Body>
            </s:Envelope>""" % (self.login, self.password, self.base_url)

        response = self.session.post(
            'https://login.microsoftonline.com/extSTS.srf',
            body,
            headers={'accept': 'application/json;odata=verbose'}
        )

        xmldoc = etree.fromstring(response.content)
        token = xmldoc.find(
            './/{http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd}BinarySecurityToken'
        )

        if token is None:
            raise Exception('Check username/password and rootsite')

        return token.text

    def authenticate(self):
        """
        Checks users authentication.
        Returns True/False dependently of user access.

        :return: Boolean
        """
        token = self._get_security_token()
        url = self.login_url + '_forms/default.aspx?wa=wsignin1.0'
        response = self.session.post(url, data=token)

        if response.status_code not in self.success_list:
            raise LoginException(response.text)

        response = self.session.get(self.base_url, headers=headers["GET"])
        if response.status_code not in self.success_list:
            raise LoginException(response.text)

    def get_lists(self):
        """
        Gets all lists.

        :return: Returns a REST response.
        """
        logging.info("Called get_lists()")

        get = self.session.get(
            self._compose_url("_api/web/lists?$top=5000"),
            headers=headers["GET"]
        )

        logging.debug("GET: %s", get.status_code)
        if get.status_code not in self.success_list:
            raise ListingException(get.content)

        return {
            x["Title"]: SharepointList.from_dict(self, x)
            for x in get.json()["d"]["results"]
        }

    def create_new_list(self, list_name, data=None, description="", allow_content_types=True,
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

        :param list_name: Name of new List.
        :param data: Optional Parameter when you need to use your own data
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
            self._compose_url("_api/web/lists"),
            headers=headers["POST"],
            data=json.dumps(data)
        )
        logger.info("Create new list - {}.".format(list_name))
        logger.debug("POST: {}".format(post.status_code))
        if post.status_code not in self.success_list:
            CantCreateNewListException(post.content)

    def get_folder_by_relative_url(self, server_relative_url):
        """
        Gets all information about given folder directory.

        :param folder_name:  Required, name of the folder
        :return: Returns REST response
        """
        logger.info("Get information for %s folder.", server_relative_url)

        get = self.session.get(
            self._compose_url("_api/web/GetFolderByServerRelativeUrl('{}')", server_relative_url),
            headers=headers["GET"]
        )

        logger.debug("GET: %s", get.status_code)
        if get.status_code not in self.success_list:
            raise ListingException(get.content)

        return SharepointFolder.from_dict(self, get.json()["d"])

    def custom_query(self, query_url, request_type="GET", data=None):
        """
        Allows to provide your API end point query_url

        :param query_url: Required, url for your API end point
        :param request_type: Optional, default set to "GET" - type of your request
        :param data: Optional, default set to None. Data for POST or PUT requests
        :return: returns REST response status
        """
        if request_type == "GET":

            get = self.session.get(
                self.base_url + query_url,
                headers=headers["GET"]
            )
            logger.debug("GET: %s", get.status_code)

            if get.status_code not in self.success_list:
                raise SharepointException(get.content)

            return get.json()["d"]

        elif request_type == "POST":
            if data is None:
                raise ValueError("Data needs to be provided to perform this request.")

            headers["POST"]["X-RequestDigest"] = self.digest()
            post = self.session.post(
                self.base_url + query_url,
                headers=headers["POST"],
                data=json.dumps(data)
            )
            logger.debug("POST: %s", post.status_code)

            if post.status_code not in self.success_list:
                raise SharepointException(post.content)

            return post.json()["d"]

        elif request_type == "PUT":
            if data is None:
                raise ValueError("Data needs to be provided to perform this request.")

            headers["PUT"]["X-RequestDigest"] = self.digest()
            put = self.session.post(
                self.base_url + query_url,
                headers=headers["PUT"],
                data=json.dumps(data)
            )
            logger.debug("PUT: %s", put.status_code)

            if put.status_code not in self.success_list:
                raise SharepointException(put.content)

            return put.json()["d"]

        elif request_type == "DELETE":
            if data is None:
                raise ValueError("Data needs to be provided to perform this request.")

            headers["DELETE"]["X-RequestDigest"] = self.digest()
            delete = self.session.post(
                self.base_url + query_url,
                headers=headers["DELETE"],
            )

            logger.debug("DELETE: %s", delete.status_code)
            if delete.status_code not in self.success_list:
                raise SharepointException(delete.content)

            return delete.json()["d"]

        else:
            raise ValueError("Wrong request type.")
