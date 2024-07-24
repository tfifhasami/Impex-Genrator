import pandas as pd

# Lire le fichier Excel
df = pd.read_excel('skufiltr1.xlsx')

# Sélectionner uniquement les colonnes requises pour l'impex
impex_df = df[['CODE ARTICLE', 'EAN', 'CODE Categorie', 'COde Sousfamille']]


# Generate ImpEx script
impex_script = "## ImpEx for Importing Prices\n\n"
impex_script += "# Macros / Replacement Parameter definitions\n"
impex_script += "$productCatalog = azizaProductCatalog\n"
impex_script += "$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default='$productCatalog:Staged']\n"
impex_script += "$productCatalogName = Aziza product catalog\n"
impex_script += "$prices = Europe1prices[translator=de.hybris.platform.europe1.jalo.impex.Europe1PricesTranslator]\n"
impex_script += "$product = product($catalogVersion, code)[unique=true]\n"



impex_script += "INSERT_UPDATE PriceRow; $product  ; unit(code[unique = true, default = pieces]); currency(isocode)[unique = true]; price; minqtd[default = 1]; unitFactor[default = 1]; net[default = true]; $catalogVersion\n"

for index, row in df.iterrows():
    article_code = str(row['CODE ARTICLE'])
    unit = "pieces"
    currency = "TND"
    price = str(row['PRIX_VENTE'])
    impex_script += f"; {article_code};;{currency} ;{price};;;;;\n"

# Write ImpEx script to a file
with open("produit-price.impex", "w") as f:
    f.write(impex_script)

print("ImpEx script generated successfully.")