"""
Code used to populate dictionary by manually slicing term definition and link from Investopedia HTML

Step 1: Copy HTML from <main id="dictionary-listing_1-0" class="comp dictionary-listing mntl-block">
Found at: https://www.investopedia.com/financial-term-dictionary-4769738 ; choose see all for one of the letters

Step 2: Insert copied HTML into a 'page.txt' file (see line 19)

Step 3: Run program and import file.csv to the Excel Financial Dictionary

Step 4: Repeat process for each letter on Investopedia
"""
import re
import pandas

from itertools import islice


f = open('page.txt', 'r')
text = f.read()
f.close()

expression = '((?:http|ftp|https://)(?:[\w_-]+(?:(?:\.[\w_-]+)+))(?:[\w.,@?^=%&:/~+-]*[\w@?^=%&/~+-]))+'  # URL
expression2 = 'wrapper">(.*)</s'  # Full_Name

# Generate lists, matches = URL's  matches2 = Full_Name
matches = list(re.findall(expression, text))
matches2 = list(re.findall(expression2, text))

#Use Pandas to assemble file.csv
df = pandas.DataFrame(data={"col1": matches2, "col2": matches})
df.to_csv("./file.csv", sep=',', index=False)
