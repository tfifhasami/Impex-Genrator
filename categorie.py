import pandas as pd

# Read data from Excel file
excel_data = pd.read_excel("skufiltr1.xlsx")

# Create a copy of the original data
excel_data_copy = excel_data.copy()
excel_data_copy1 = excel_data.copy()
excel_data_copy2 = excel_data.copy()

# Generate ImpEx script
impex_script = "## ImpEx for Importing Categories\n\n"
impex_script += "# Macros / Replacement Parameter definitions\n"
impex_script += "$productCatalog = azizaProductCatalog\n"
impex_script += "$productCatalogName = Aziza product catalog\n"
impex_script += "$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default=$productCatalog:Staged]\n"
impex_script += "$supercategories = source(code, $catalogVersion)[unique=true]\n"
impex_script += "$categories = target(code, $catalogVersion)[unique=true]\n\n"
impex_script += "# Insert Categories\n"
impex_script += "INSERT_UPDATE Category; code[unique = true]; allowedPrincipals(uid)[default = '']; $catalogVersion\n"

# Drop duplicates for writing into ImpEx
excel_data_copy.drop_duplicates(subset=['CODE Categorie'], keep='first', inplace=True)
for index, row in excel_data_copy.iterrows():
    code_value = str(row['CODE Categorie'])
    impex_script += f"; {code_value}; ; \n"

excel_data_copy1.drop_duplicates(subset=['CODE famille'], keep='first', inplace=True)
for index, row in excel_data_copy1.iterrows():
    code_value = str(row['CODE famille'])
    impex_script += f"; {code_value}; ; \n"

excel_data_copy2.drop_duplicates(subset=['COde Sousfamille'], keep='first', inplace=True)
for index, row in excel_data_copy2.iterrows():
    code_value = str(row['COde Sousfamille'])
    impex_script += f"; {code_value}; ; \n"

# Restore original data
excel_data = excel_data_copy

# Write ImpEx script to a file
# with open("categorie1.impex", "w") as f:
#     f.write(impex_script)

# Continue with further processing using excel_data DataFrame
# For example, generating relationships in ImpEx script
impex_script += "## Insert Category Structure\n\n"
impex_script += "INSERT_UPDATE CategoryCategoryRelation; $categories; $supercategories\n\n"
for index, row in excel_data_copy.iterrows():
    category_code = str(row['CODE Categorie'])
    supercategory_code = str(row['Univers'])
    impex_script += f"; {category_code}; {supercategory_code}\n"

for index, row in excel_data_copy1.iterrows():
    category_code = str(row['CODE famille'])
    supercategory_code = str(row['CODE Categorie'])
    impex_script += f"; {category_code}; {supercategory_code}\n"

for index, row in excel_data_copy2.iterrows():
    category_code = str(row['COde Sousfamille'])
    supercategory_code = str(row['CODE famille'])
    impex_script += f"; {category_code}; {supercategory_code}\n"  

# Write updated ImpEx script to a file
with open("categorie_updated.impex", "w") as f:
    f.write(impex_script)

print("ImpEx script generated successfully.")
