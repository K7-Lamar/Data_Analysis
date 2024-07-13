import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl

# Load the Excel file
file_path = r'C:\Users\Ebenezer Sarpong\Desktop\Data_Analysis/PENG452-ASS1.xlsx'
sheet1_df = pd.read_excel(file_path, sheet_name='Sheet1', skiprows=3, header=0)

# Rename columns for better readability
sheet1_df.columns = ['Index', 'Year', 'Peak Daily Oil Rate [mbd]', 'Water Cut [%]', 'Peak Daily Gas Lift Rate [MMscfd]', 'Operating Efficiency']

# Drop the unnecessary 'Index' column
sheet1_df = sheet1_df.drop(columns=['Index'])

# Box and Whisker Plot for 'Peak Daily Oil Rate [mbd]'
plt.figure(figsize=(10, 6))
sns.boxplot(data=sheet1_df[['Peak Daily Oil Rate [mbd]']])
plt.title('Box and Whisker Plot for Peak Daily Oil Rate [mbd]')
plt.xlabel('Peak Daily Oil Rate [mbd]')
plt.show()

# Scatter Plot for 'Year' vs 'Peak Daily Oil Rate [mbd]'
plt.figure(figsize=(10, 6))
sns.scatterplot(x='Year', y='Peak Daily Oil Rate [mbd]', data=sheet1_df)
plt.title('Scatter Plot of Year vs Peak Daily Oil Rate [mbd]')
plt.xlabel('Year')
plt.ylabel('Peak Daily Oil Rate [mbd]')
plt.show()
