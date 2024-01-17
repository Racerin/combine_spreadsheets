import unittest
from unittest import TestCase

import main

class TestMain(TestCase):
    def test_reduce_whitespace_to_one_space(self):
        pass

    def test_dict_remove_whitespace_trim__old_to_unique_new_field(self):
        # columns = list(((str(i//3) for i in range(20))))
        columns = ['a', 'A', 'B ', 'B', 'c', 'C ', ' d', 'd', ' E', 'e']
        output = main.dict__old_to_unique_new_field(columns)
        assert len(set(output.values())) == len(columns)
        assert len(set(output.keys())) == len(columns)
        # assert output == {}, output

if __name__ == '__main__':
    unittest.main()