'''
In this programm I take the number of tourists that came in greece for each Transport between years 2011-2014, 
by reading Excel files from each year into pandas DataFrames.
Then I combine the DataFrames of each year into one DataFrame and group by Transport.
Finally for each Transport and the Number of Tourists that used it, I use the final DataFrame to extract csv file, 
insert the new table in a database and show a graph. 
'''


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sqlite3

excel_path = 'C:\\Users\\Micha\\Desktop\\Project\\'

#creation of list to put all the dataframes

dfs = []

#reads values form xls file of the year '2011' into pandas dataframe and then inserts the dataframe into the list we created before

df1 = pd.read_excel(excel_path + '2011.xls', sheet_name= 'ΔΕΚ',  skiprows= 134, nrows= 1, usecols= "C,D,E,F", header=None, converters={2: int, 3: int, 4: int, 5: int})
dfs.append(df1)
#print(df1)

#reads values form xls files of the years '2012-2014' into pandas dataframes and then inserts the dataframes into the list we created before

for year in range(2012, 2015):
    df2 = pd.read_excel(excel_path + str(year) + '.xls', sheet_name= 'ΔΕΚ',  skiprows= 136, nrows= 1, usecols= "C,D,E,F", header=None, converters={2: int,3: int, 4: int, 5: int})
    dfs.append(df2)
    #print(df2)

#combine the values of all dataframes in the list

data = pd.concat(dfs)
#print(data)

#name each column of dataframe

data.columns = ['Airplane', 'Train', 'Ship', 'Car'] 
print(data.sum())

#add the values of each column 

data = data.sum()
#print(data)

#export dataframe to csv file

data.to_csv(excel_path + 'Arrivals_per_Transport.csv', index=True, header=['Arrivals'], index_label='Transport', encoding="utf-8")

#create new database if it doesn't exists, if database exists uses the current database

db = sqlite3.connect(excel_path + 'db.sqlite') 

#insert new table with the dataframe in database 

data.to_sql('Arrivals_per_Transport', db, if_exists='replace', index=True, index_label='Transport') 

#creation of graph

data.plot.bar()
plt.title('Arrivals of Tourists per Transport between 2011-2014')
plt.ylabel('Arrivals')
plt.show()


