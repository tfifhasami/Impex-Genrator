import pandas as pd

# Load Excel file
excel_file = 'skufiltr1.xlsx'
df = pd.read_excel(excel_file)

# Drop duplicate rows based on CODE Categorie
df.drop_duplicates(subset=['CODE Categorie'], keep='first', inplace=True)

# Generate ImpEx script
impex_script = "## ImpEx for Importing Category Media\n\n"
impex_script += "# Macros / Replacement Parameter definitions\n"
impex_script += "$productCatalog=azizaProductCatalog\n"
impex_script += "$productCatalogName=Aziza product catalog\n"
impex_script += "$catalogVersion=catalogversion(catalog(id[default=$productCatalog]),version[default='Staged'])[unique=true,default=$productCatalog:Staged]\n"
impex_script += "$logo=logo(code, $catalogVersion)\n"
impex_script += "$siteResource = jar:com.aziza.initialdata.setup.InitialDataSystemSetup&/azizainitialdata/import/sampledata/productCatalogs/$productCatalog\n"

impex_script += "INSERT_UPDATE Media;code[unique=true];realfilename;@media[translator=de.hybris.platform.impex.jalo.media.MediaDataTranslator];mime[default='image/jpeg'];$catalogVersion\n\n"

for index, row in df.iterrows():
    code = str(row['CODE Categorie'])
    realfilename = str(row['mediacatt'])
    media = '$siteResource/images/categories/'+""+str(row['mediacatt'])
    

impex_script += f";{code};{realfilename};{media};\n"
# Create ImpEx content
impex_script += "UPDATE Category;code[unique=true];$logo;allowedPrincipals(uid)[default='customergroup'];$catalogVersion\n\n"

# Iterate through rows and generate ImpEx lines
for index, row in df.iterrows():
    category_code = str(row['CODE Categorie'])
    image_filename = row['mediacatt']
    
    impex_script += f";{category_code};{image_filename};\n"

# Write ImpEx content to file
impex_file = 'categories-media.impex'
with open(impex_file, 'w') as file:
    file.write(impex_script)

print(f"ImpEx file generated: {impex_file}")
