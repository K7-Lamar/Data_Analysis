import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = r'C:\Users\Ebenezer Sarpong\Desktop\Data_Analysis/PENG452-ASS1.xlsx'
sheet2_df = pd.read_excel(file_path, sheet_name='Sheet2', skiprows=4, header=0)

# Rename columns for better readability
sheet2_df.columns = ['Index', 'Pressure [psia]', 'Oil Volume Factor Bo [bbls/STB]', 'GOR Rs [scf/STB]', 
                     'Gas Compressibility Z', 'Gas Formation Volume Factor Bg [cf/scf]', 'Gas Gravity', 
                     'Liquid Density [g/cc]']

# Drop the unnecessary 'Index' column
sheet2_df = sheet2_df.drop(columns=['Index'])

# Convert relevant columns to numeric, forcing errors to NaN
sheet2_df['Pressure [psia]'] = pd.to_numeric(sheet2_df['Pressure [psia]'], errors='coerce')
sheet2_df['Oil Volume Factor Bo [bbls/STB]'] = pd.to_numeric(sheet2_df['Oil Volume Factor Bo [bbls/STB]'], errors='coerce')

# Drop rows with NaN values in these columns
sheet2_df = sheet2_df.dropna(subset=['Pressure [psia]', 'Oil Volume Factor Bo [bbls/STB]'])

# Display the first few rows of the cleaned data
print(sheet2_df.head())

# Box and Whisker Plot for 'Oil Volume Factor Bo [bbls/STB]'
plt.figure(figsize=(10, 6))
sns.boxplot(data=sheet2_df[['Oil Volume Factor Bo [bbls/STB]']])
plt.title('Box and Whisker Plot for Oil Volume Factor Bo [bbls/STB]')
plt.xlabel('Oil Volume Factor Bo [bbls/STB]')
plt.show()

# Scatter Plot for 'Pressure [psia]' vs 'Oil Volume Factor Bo [bbls/STB]'
plt.figure(figsize=(10, 6))
sns.scatterplot(x='Pressure [psia]', y='Oil Volume Factor Bo [bbls/STB]', data=sheet2_df)
plt.title('Scatter Plot of Pressure [psia] vs Oil Volume Factor Bo [bbls/STB]')
plt.xlabel('Pressure [psia]')
plt.ylabel('Oil Volume Factor Bo [bbls/STB]')
plt.show()
