import pandas as pd

# Function to replace line breaks with <br />
def encode_line_breaks(text):
    if isinstance(text, str):
        return text.replace('\n', '<br />').replace('\r', '<br />')
    return text

# Lire le fichier Excel
df = pd.read_excel('liste-prods.xlsx')

# Encode line breaks in specific columns
df['name'] = df['name'].apply(encode_line_breaks)
df['short_description'] = df['short_description'].apply(encode_line_breaks)
df['description'] = df['description'].apply(encode_line_breaks)

# Generate ImpEx script
impex_script = "## ImpEx for Importing Product Localisations\n\n"
impex_script += "# Macros / Replacement Parameter definitions\n"
impex_script += "$productCatalog = azizaProductCatalog\n"
impex_script += "$productCatalogName = Aziza product catalog\n"
impex_script += "$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default=$productCatalog:Staged]\n"
impex_script += "# Language\n"
impex_script += "$lang = fr\n\n"

impex_script += "# # Update allProducts with localisations\n"
impex_script += "UPDATE Product; code[unique = true]; $catalogVersion; name[lang = $lang]; description[lang = $lang]; erpShortName[lang = $lang]; erpFullName[lang = $lang]; shortDescription[lang = $lang]\n"

for index, row in df.iterrows():
    article_code = str(row['sku'])
    name = str(row['name'])
    shortdesc = str(row['short_description'])
    desc = str(row['description'])
    impex_script += f"; {article_code};;{name};{desc};{name};{name};{shortdesc}\n"

# Write ImpEx script to a file
with open("produit-fr.impex", "w", encoding="utf-8") as f:
    f.write(impex_script)

print("ImpEx script generated successfully.")
