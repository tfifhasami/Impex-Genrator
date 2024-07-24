import pandas as pd

# Read data from Excel file
excel_data = pd.read_excel("skufiltr1.xlsx")

# Create a copy of the original data
excel_data_copy = excel_data.copy()
excel_data_copy1 = excel_data.copy()
excel_data_copy2 = excel_data.copy()
excel_data.columns = excel_data.columns.str.strip()

excel_data_copy.drop_duplicates(subset=['CODE Categorie'], keep='first', inplace=True)
excel_data_copy1.drop_duplicates(subset=['CODE famille'], keep='first', inplace=True)
excel_data_copy2.drop_duplicates(subset=['COde Sousfamille'], keep='first', inplace=True)

# Generate ImpEx script
impex_script = "## ImpEx for Importing Categories\n\n"
impex_script += "# Macros / Replacement Parameter definitions\n"
impex_script += "$productCatalog = azizaProductCatalog\n"
impex_script += "$productCatalogName = Aziza product catalog\n"
impex_script += "$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default=$productCatalog:Staged]\n"
impex_script +="# Language\n"
impex_script += "$lang = fr\n"
impex_script += "# Create Categories\n"
impex_script += "UPDATE Category; code[unique = true];  name[lang = $lang]; $catalogVersion\n"

for index, row in excel_data_copy.iterrows():
    category_code = str(row['CODE Categorie'])
    supercategory_code = row['categorie']
    impex_script += f"; {category_code}; {supercategory_code}\n"

for index, row in excel_data_copy1.iterrows():
    category_code = str(row['CODE famille'])
    supercategory_code = str(row['famille'])
    impex_script += f"; {category_code}; {supercategory_code}\n"

for index, row in excel_data_copy2.iterrows():
    category_code = str(row['COde Sousfamille'])
    supercategory_code = str(row['Sousfamille'])
    impex_script += f"; {category_code}; {supercategory_code}\n"  

# Write ImpEx script to a file
with open("categoriefr.impex", "w", encoding='utf-8') as f:
    f.write(impex_script)

print("ImpEx script generated successfully.")
