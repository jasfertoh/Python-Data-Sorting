import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
dfs = pd.read_excel('outputFile.xlsx', engine='openpyxl', sheet_name='T1')  # Reads the original outputFile.xlsx
dfs = dfs.iloc[4:36] # Only selects the 5th row to the 37th row
list5 = ['Total', 'Southeast Asia', 'Greater China', 'North Asia', 'South Asia', 'West Asia', 'Other Markets In South Asia', 'Other Markets In North Asia', 'Other Markets In Greater China', 'Other Markets In Southeast Asia']
list4 = ['Date', 'Brunei Darussalam', 'Indonesia', 'Malaysia', 'Myanmar', 'Philippines', 'Thailand', 'Vietnam', 'China', 'Hong Kong SAR', 'Taiwan', 'Japan', 'South Korea', 'Bangladesh', 'India', 'Pakistan', 'Sri Lanka', 'Iran', 'Israel', 'Kuwait', 'Saudi Arabia', 'United Arab Emirates']  # Column headers
years = ['2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020']  # Filter by years
list2 = ['Brunei Darussalam', 'Indonesia', 'Malaysia', 'Myanmar', 'Philippines', 'Thailand', 'Vietnam', 'China', 'Hong Kong SAR', 'Taiwan', 'Japan', 'South Korea', 'Bangladesh', 'India', 'Pakistan', 'Sri Lanka', 'Iran', 'Israel', 'Kuwait', 'Saudi Arabia', 'United Arab Emirates']  # List of countries
list1 = []
list6 = ['Date', 'Total', 'Southeast Asia', 'Brunei Darussalam', 'Indonesia', 'Malaysia', 'Myanmar', 'Philippines', 'Thailand', 'Vietnam', 'Other Markets In Southeast Asia', 'Greater China', 'China', 'Hong Kong SAR', 'Taiwan', 'Other Markets In Greater China', 'North Asia', 'Japan', 'South Korea', 'Other Markets In North Asia', 'South Asia', 'Bangladesh', 'India', 'Pakistan', 'Sri Lanka', 'Other Markets In South Asia', 'West Asia', 'Iran', 'Israel', 'Kuwait', 'Saudi Arabia', 'United Arab Emirates']

dfs = dfs.rename(columns = {'Subject: Tourism': 'Test'}) # Renames the Subject: Tourism column to Test
dfs = dfs.T # Transposes the whole excel table.
dfs.columns = dfs.iloc[0] # Select the zeroth column
dfs = dfs.reset_index() # Resets the index
dfs = dfs.drop(dfs.index[0]) # Removing the zeroth index column
dfs = dfs.reset_index() # Resetting the index again
dfs = dfs.drop(["level_0","index"],axis = 1) 
dfs = dfs.rename(columns = {'Variables':'Date', 'Total International Visitor Arrivals By Inbound Tourism Markets':'Total'}) # Renaming first two columns to Date and Total respectively.
dfs.columns.name = None # Deletes the column name Test
dfs.columns = list6 # Sets all column names according to list6.
for x in list5:
    dfs = dfs.drop(columns=[x], axis=1) # Dropping all unnecessary columns like Total, Southeast Asia etc.
list3 = pd.DataFrame()
arr = np.array(list1)
for x in years:
    list1.append(dfs[dfs['Date'].str.contains(x)])
    arr = np.array(list1)
# For loop to filter by years
for x in range(0, 10):
    for y in range(0, 12):
        for k in range(0, 22):
            arr[x][y][k] = str(arr[x][y][k]).replace(",", "")
# Removes all commas in the numbers by converting to string then replacing the commas with empty.
df = pd.DataFrame(data = np.concatenate(arr)) # Make a new Dataframe according to the filtered years.
df.columns = list4  # Sets the column header values to list4's values
df.set_index('Date', inplace=True, drop=True)
df[list2] = df[list2].apply(pd.to_numeric, errors='coerce')  # Turn all cell types to numbers so that they can be summed up together
index = list(df.index)
df.loc['Total'] = df.sum(numeric_only=True, axis=0)  # New row named 'Total' that contains the summed up number of visitors per country.
sorted_df = df.sort_values(by ='Total', axis=1, ascending=False)  # Sorting of the values by the row 'Total' in descending order
sorted_df.to_excel('IMDA.xlsx') # Writes a new excel file with the new dataframe with the name IMDA.xlsx

print("Top 3 \n" + str(sorted_df.iloc[-1][:3]))
countries = list(sorted_df.columns)
visitors = list(sorted_df.iloc[-1][:-1][:-1])
sortTotal = pd.DataFrame({'countries': countries[:3], 'Visitors': visitors[:3]})
sortTotal = sortTotal.set_index('countries')
countries = list(sortTotal.columns)

ax = sortTotal[countries].plot(kind='bar', title="Top 3 Countries", figsize = (10, 10), legend = True, fontsize = 12)
plt.show()
my_fig = ax.get_figure()
my_fig.savefig('top3.png')

countries = list(sorted_df.columns)
sortAll = sorted_df.iloc[-1].to_frame()
sortAll = sortAll.rename(columns = {'Total':'Visitors'})
countries = list(sortAll.columns)
ax = sortAll[countries].plot(kind='bar', title="All Countries", figsize = (10, 10), legend = True, fontsize = 12)
plt.show()

my_fig = ax.get_figure()
my_fig.savefig('allcountries.png')
sortAll = sortAll.reset_index()
sortAll = sortAll.rename(columns = {'index':'Countries'})