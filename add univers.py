import pandas as pd

# Read the Excel files into DataFrames
sku_df = pd.read_excel('sku.xlsx')
cat_df = pd.read_excel('cat.xlsx')

# Merge the DataFrames based on the 'categorie' column
merged_df = pd.merge(sku_df, cat_df, on='categorie', how='left')

# Fill missing values in 'univers' column in sku.xlsx with values from cat.xlsx
merged_df['univers_x'].fillna(merged_df['univers_y'], inplace=True)

# Drop the extra 'univers_y' column and rename 'univers_x' to 'univers'
merged_df.drop(columns=['univers_y'], inplace=True)
merged_df.rename(columns={'univers_x': 'univers'}, inplace=True)

# Write the updated DataFrame back to sku.xlsx
merged_df.to_excel('sku.xlsx', index=False)
