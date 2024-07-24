import pandas as pd

# Lire le fichier Excel
df = pd.read_excel('skufiltr1.xlsx')

# Sélectionner uniquement les colonnes requises pour l'impex
impex_df = df[['CODE ARTICLE', 'EAN', 'CODE Categorie', 'COde Sousfamille']]

# Renommer les colonnes pour correspondre aux noms requis dans l'impex
impex_df.rename(columns={'CODE ARTICLE': 'code', 'EAN': 'ean', 'CODE Categorie': 'cat_code', 'COde Sousfamille': 'subfam_code'}, inplace=True)

# Convertir les colonnes cat_code et subfam_code en chaînes de caractères (str)
impex_df['cat_code'] = impex_df['cat_code'].astype(str)
impex_df['subfam_code'] = impex_df['subfam_code'].astype(str)

# Créer la colonne $supercategories en combinant cat_code et subfam_code
impex_df['$supercategories'] = impex_df['subfam_code']

# Créer la colonne externalReference qui est égale à code
impex_df['externalReference'] = impex_df['code']

# Réordonner les colonnes selon l'ordre requis dans l'impex
impex_df = impex_df[['code', 'ean', '$supercategories', 'externalReference']]

# Ajouter une ligne vide entre l'en-tête et les données
empty_row = pd.DataFrame([''] * len(impex_df.columns)).T
empty_row.columns = impex_df.columns
impex_df = pd.concat([empty_row, impex_df], ignore_index=True)

# Écrire le DataFrame au format impex dans un fichier texte
with open('product.impex', 'w') as f:
    f.write("# -----------------------------------------------------------------------\n")
    f.write("# Copyright (c) 2019 SAP SE or an SAP affiliate company. All rights reserved.\n")
    f.write("# -----------------------------------------------------------------------\n")
    f.write("# ImpEx for Importing Products\n\n")
    f.write("# Macros / Replacement Parameter definitions\n")
    f.write("$productCatalog = azizaProductCatalog\n")
    f.write("$productCatalogName = Aziza product catalog\n")
    f.write("$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default=$productCatalog:Staged]\n")
    f.write("$supercategories = supercategories(code, $catalogVersion)\n")
    f.write("$baseProduct = baseProduct(code, $catalogVersion)\n")
    f.write("$approved = approvalstatus(code)[default='approved']\n\n")
    f.write("# Insert Products\n")
    f.write("INSERT_UPDATE Product; code[unique = true]; ean           ; $supercategories; externalReference; superTag(code); salesChannel(code); $catalogVersion; $approved\n")

    # Écrire les données du DataFrame dans le fichier impex
    for index, row in impex_df.iterrows():
        f.write(f";{row['code']}; {row['ean']}; {row['$supercategories']}; {row['externalReference']}; ; ; ; ;\n")

print("Conversion terminée. Le fichier impex a été créé avec succès.")
