import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Declare variables for User Input
reg = "asia"
start = 2008
end = 2017

# User Input for Region
while True:
  try:
    region = input("Enter your Region: ").lower()
    if region == reg:
      break
    print("Invalid Region entered")
  except Exception as e:
    print(e)

# User Input for 2008 - 2017 year range

while True:
  try:
    syear = int(input("Enter your starting year: "))
    if syear == start:
      break
    print("Invalid year entered")
  except Exception as e:
    print(e)

while True:
  try:
    eyear = int(input("Enter your ending year: "))
    if eyear == end:
      break
    print("Invalid year entered")
  except Exception as e:
    print(e)

# Read excel and create Class to select, sum and arrange in descending order for total of visitors
# in each country of the Asia Region from 2008 to 2017
df = pd.read_excel('Project_File.xlsx', sheet_name="International Monthly Visitor A")

class DataFunc:
  def __init__(self,df):
    self.df = df

  def select08_17(self):
    self.df08_17 = df.iloc[362: 480, 1: 19]
    self.df08_17sum = self.df08_17.iloc[:, :].sum()
    self.df08_17sum_descend = self.df08_17sum.sort_values(ascending=False)
    return self.df08_17sum_descend

if True:
  data = DataFunc(df)
  data08_17 = data.select08_17()
  print(data08_17)


  # Print top 3 countries
  print(data08_17[0:3])

  # Plot graph for all countries

  index = np.arange(len(data08_17.index))
  plt.xlabel('Countries', fontsize=5)
  plt.ylabel('Visitors (in 1e7)', fontsize=10)
  plt.xticks(index, data08_17.index, fontsize=6, rotation=90)
  plt.title('Total Visitors in Asia Region from 2008 to 2017')
  plt.bar(data08_17.index, data08_17.values)
  plt.savefig('Total visitors in Asia Region from 2008 to 2017')