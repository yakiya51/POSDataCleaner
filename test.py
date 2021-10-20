import re


filter = re.compile('\d+.+')

str = """SOME ITEM 12OZ"""

print(re.findall(filter, str))