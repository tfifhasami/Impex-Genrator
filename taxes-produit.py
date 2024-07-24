import pandas as pd

# Lire le fichier Excel
df = pd.read_excel('skufiltr1.xlsx')



# Generate ImpEx script
impex_script = "## ImpEx for Importing Product taxex into the Store\n\n"
impex_script += "$productCatalog = azizaProductCatalog\n"
impex_script += "$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default='$productCatalog:Staged']\n"
impex_script += "$prices = Europe1prices[translator=de.hybris.platform.europe1.jalo.impex.Europe1PricesTranslator]\n"
impex_script += "$taxGroup = Europe1PriceFactory_PTG(code)[default=tn-vat-full]\n"


impex_script += "UPDATE Product; code[unique = true]       ; $catalogVersion; $taxGroup\n"

for index, row in df.iterrows():
    article_code = str(row['CODE ARTICLE'])
    impex_script += f";{article_code};\n"

impex_script += "# Insert Product taxes for US\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"
impex_script += "# Insert Product taxes for JP\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"

impex_script += "# Insert Product taxes for GB\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"

impex_script += "# Insert Product taxes fo FR\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"

impex_script += "# Insert Product taxes for PL\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"

impex_script += "# Insert Product taxes for DE\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"

impex_script += "# Insert Product taxes for CA\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"

impex_script += "# Insert Product taxes for CN\n"
impex_script += "INSERT_UPDATE ProductTaxCode; productCode[unique = true]; taxCode; taxArea[unique = true]\n"


# Write ImpEx script to a file
with open("produit-taxes.impex", "w") as f:
    f.write(impex_script)

print("ImpEx script generated successfully.")