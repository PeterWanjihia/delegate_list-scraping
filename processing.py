import pandas as pd

# Read the data from the Excel file
df = pd.read_excel('Applicants.xlsx')

# Extract the first three and six characters from 'Adm No'
df['School'] = df['Adm No'].str[:3]
df['Course'] = df['Adm No'].str[:6]

# Assign School categories based on the first three characters
df['School Category'] = pd.Categorical(df['School']).codes + 1

# Assign Course categories based on the first six characters
df['Course Category'] = pd.Categorical(df['Course']).codes + 1

# Sort the DataFrame based on the 'School' column
df_sorted = df.sort_values(by='School')

# Save the sorted DataFrame to a new Excel file
df_sorted.to_excel('Categorized_and_Sorted_Applications.xlsx', index=False)
