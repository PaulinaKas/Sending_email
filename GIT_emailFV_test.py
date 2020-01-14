import unittest
from unittest import mock
from GIT_emailFV import EditMenu

class EditMenuTest(unittest.TestCase):
    def setUp(self):
        self.fileToOpen = mock.mock_open(read_data=(
            """;ABC Company; 44; loading place 1 / unloading place 1; 11.01.2020 / 13.01.2020; please send 2 copies of CMR file; postal address: 00-000 Warszawa, ul. Beznazwy 1;
             ;CDE Company; 12; loading place 2 / unloading place 2; 03.01.2020 / 06.01.2020; ; postal address: 11-111 Warszawa, ul. Beznazwy 2;"""))
        self.fileToOpenEmpty = mock.mock_open(read_data='')
        self.fileMenuObject = FileMenu()

# class EditMenuTest(unittest.TestCase):
#     def test_initialization(self):
#         editMenuObject = EditMenu()
#         self.assertTrue(editMenuObject) # checks if object of class EditMenu exists
#
#     def test_(self):
#         # Given (input sitation)
#         editMenuObject = EditMenu()
#         # When (some action on this situation)
#         editMenuObject.undo()
#         # Then (results check)


if __name__ == '__main__':
    unittest.main()
