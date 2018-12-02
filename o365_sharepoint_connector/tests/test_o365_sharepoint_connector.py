#! /usr/bin/env python3
import os.path
import unittest

from o365_sharepoint_connector import SharePointConnector
from o365_sharepoint_connector import CantCreateNewListException


class TestSharePointConnector(unittest.TestCase):
    TEST_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGIN_FILE = os.path.join(TEST_DIR, "login.txt")
    USERNAME, PASSWORD, SITE_URL = open(LOGIN_FILE).read().splitlines()

    @classmethod
    def setUpClass(cls):
        cls.c = SharePointConnector(
            login=cls.USERNAME,
            password=cls.PASSWORD,
            site_url=cls.SITE_URL
        )
        cls.c.authenticate()

    def test_add_list_dir_list_remove_list(self):
        try:
            new_list = self.c.add_list("unittests")
            self.assertEqual(new_list.title, "unittests")
        except CantCreateNewListException:
            pass

        lists = self.c.get_lists()
        self.assertIn("unittests", lists)

        lists = self.c.get_lists()
        lists["unittests"].delete()

        lists = self.c.get_lists()
        self.assertNotIn("unittests", lists)

    def test_get_all_folders(self):
        try:
            ut_list = self.c.add_list("unittests")
        except CantCreateNewListException:
            ut_list = self.c.get_lists()["unittests"]

        assert ut_list.get_all_folders()


if __name__ == '__main__':
    unittest.main()
