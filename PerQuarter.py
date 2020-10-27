'''
In this programm I take the number of tourists that came in greece for each Quarter of year between years 2011-2014, 
by reading Excel files from each year and month into pandas DataFrames.
Then I combine the DataFrames of each years and months into one DataFrame and group by months in same Quarter and year.
Finally for each Quarter of year and the Number of Tourists that came in greece that period,
I use the final DataFrame to extract csv file, 
insert the new table in a database and show a graph. 
'''


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sqlite3

excel_path = 'C:\\Users\\Micha\\Desktop\\Project\\'

#creation of list to put all the dataframes

dfs = []

#maps every quarter of each year to specific months. The map will be used in the for loop that reads the excel files

months = {'2011-Q1': ('ΙΑΝ', 'ΦΕΒ', 'ΜΑΡ'), '2011-Q2': ('ΑΠΡ', 'ΜΑΙ', 'ΙΟΥΝ'), '2011-Q3':('ΙΟΥΛ', 'ΑΥΓ', 'ΣΕΠ'), '2011-Q4':('ΟΚΤ', 'ΝΟΕ', 'ΔΕΚ'),
'2012-Q1': ('ΙΑΝ', 'ΦΕΒ', 'ΜΑΡ'), '2012-Q2': ('ΑΠΡ', 'ΜΑΙΟΣ', 'ΙΟΥΝ'), '2012-Q3': ('ΙΟΥΛ', 'ΑΥΓ', 'ΣΕΠΤ'), '2012-Q4': ('ΟΚΤ', 'ΝΟΕΜ', 'ΔΕΚ'),
'2013-Q1': ('ΙΑΝ', 'ΦΕΒ', 'ΜΑΡ'), '2013-Q2': ('ΑΠΡ', 'ΜΑΙΟΣ', 'ΙΟΥΝ'), '2013-Q3': ('ΙΟΥΛ', 'ΑΥΓ', 'ΣΕΠ'), '2013-Q4': ('ΟΚΤ', 'ΝΟΕ', 'ΔΕΚ'),
'2014-Q1': ('ΙΑΝ', 'ΦΕΒ', 'ΜΑΡ'), '2014-Q2': ('ΑΠΡ', 'ΜΑΙ', 'ΙΟΥΝ'), '2014-Q3': ('ΙΟΥΛ', 'ΑΥΓ', 'ΣΕΠΤ'), '2014-Q4': ('ΟΚΤ', 'ΝΟΕΜ', 'ΔΕΚ')}

#for loop that reads the values of every month from xls files into pandas dataframes by using the map 'months'

for year in range(2011, 2015):
        for i in range(1, 5):
            if year == 2013 and i < 3:
                for month in months[str(year) + '-Q' + str(i)]:
                    df1 = pd.read_excel(excel_path + str(year) + '.xls', sheet_name= month,  skiprows= 64, nrows = 1, usecols= "G", header=None, converters={6:int})
                    df1.insert(0, "Quarter", str(year) + '-Q' + str(i))
                    dfs.append(df1)
            else:
                for month in months[str(year) + '-Q' + str(i)]:
                    df1 = pd.read_excel(excel_path + str(year) + '.xls', sheet_name= month,  skiprows= 65, nrows = 1, usecols= "G", header=None, converters={6:int})
                    df1.insert(0, "Quarter", str(year) + '-Q' + str(i))
                    dfs.append(df1)

#combine the values of all dataframes in the list and groups the values of same year and Quarter

data = pd.concat(dfs).groupby(["Quarter"]).sum().reset_index()

#name each column of dataframe

data.columns = ['Quarter', 'Arrivals'] 
print(data)

#export dataframe to csv file

data.to_csv(excel_path + 'Arrivals_per_year_quarter.csv', index=False, encoding="utf-8") #export data to csv file

#create new database if it doesn't exists, if database exists uses the current database

db = sqlite3.connect(excel_path + 'db.sqlite') #creation of database

#insert new table with the dataframe in database

data.to_sql('Arrivals_per_year_quarter', db, if_exists='replace', index=False) #insert table in database 

#creation of graph

x = data['Quarter']
y = data['Arrivals']
plt.bar(x, y)
plt.title('Arrivals of tourists per Year Quarter between 2011-2014')
plt.ylabel('Arrivals')
plt.show()