import pandas as pd

# Lire le fichier Excel
df = pd.read_excel('skufiltr1.xlsx')

# Sélectionner uniquement les colonnes requises pour l'impex
impex_df = df[['CODE ARTICLE', 'EAN', 'CODE Categorie', 'COde Sousfamille']]


# Generate ImpEx script
impex_script = "## ImpEx for Importing Products Stock Levels and Warehouses\n\n"
impex_script += "# Macros / Replacement Parameter definitions\n\n"
impex_script += "$productCatalog = azizaProductCatalog\n"
impex_script += "$productCatalogName = Aziza product catalog\n"
impex_script += "$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default=$productCatalog:Staged]\n"
impex_script += "$vendor = aziza\n"
impex_script += "INSERT_UPDATE Vendor; code[unique = true]\n"
impex_script += '$vendor\n'
impex_script += "INSERT_UPDATE Warehouse; code[unique = true]; vendor(code); default[default = true]\n"
impex_script += ";azizaDarkstore;$vendor; true\n\n"
impex_script += "INSERT_UPDATE StockLevel; available; productCode[unique = true]; warehouse(code)[unique = true][default = 'tabaaAziza']; inStockStatus(code)[default = 'notSpecified']; maxPreOrder; maxStockLevelHistoryCount; overSelling; preOrder; reserved\n"
for index, row in df.iterrows():
    article_code = str(row['CODE ARTICLE'])
    stocklevel = str(row['STOCK_DISPO'])
    
    impex_script += f"; {stocklevel};{article_code};;;;;;;\n"


# Write ImpEx script to a file
with open("product-stocklevel.impex", "w") as f:
    f.write(impex_script)

print("ImpEx script generated successfully.")