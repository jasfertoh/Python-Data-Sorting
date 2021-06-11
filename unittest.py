import unittest
from DA_pyproject import sortAll

class Country(unittest.TestCase):
    def test_country(self):
        visitors = 26713289.0
        for x in range(0, 21):
            if (sortAll.iat[x, 1] == visitors):
                result = sortAll.iat[x, 0]
        self.assertEqual(result, 'Indonesia')

    def test_visitors(self):
        country = 'Malaysia'
        for x in range(0, 21):
            if (sortAll.iat[x, 0] == country):
                result = sortAll.iat[x, 1]
        self.assertEqual(result, 11005817.0)

if __name__ == '__main__':
    unittest.main()