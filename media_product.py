import pandas as pd

# Load Excel file
excel_file = 'skufiltr1.xlsx'
df = pd.read_excel(excel_file)
df.drop_duplicates(subset=['CODE ARTICLE'], keep='first', inplace=True)

# Define the initial content of the ImpEx script
impex_content = """# -----------------------------------------------------------------------
# Copyright (c) 2019 SAP SE or an SAP affiliate company. All rights reserved.
# -----------------------------------------------------------------------
# ImPEx for Importing Product Media

# Macros / Replacement Parameter definitions
$productCatalog = azizaProductCatalog

$catalogVersion = catalogversion(catalog(id[default=$productCatalog]), version[default='Staged'])[unique=true, default=$productCatalog]
$media = @media[translator=de.hybris.platform.impex.jalo.media.MediaDataTranslator]
$thumbnail = thumbnail(code, $catalogVersion)
$picture = picture(code, $catalogVersion)
$thumbnails = thumbnails(code, $catalogVersion)
$detail = detail(code, $catalogVersion)
$normal = normal(code, $catalogVersion)
$others = others(code, $catalogVersion)
$data_sheet = data_sheet(code, $catalogVersion)
$medias = medias(code, $catalogVersion)
$galleryImages = galleryImages(qualifier, $catalogVersion)
$siteResource = jar:com.aziza.initialdata.setup.InitialDataSystemSetup&/azizainitialdata/import/sampledata/productCatalogs/$productCatalog

# Create Media
"""
impex_content += "INSERT_UPDATE Media; mediaFormat(qualifier); code[unique = true]; $media; mime[default = 'image/jpeg']; $catalogVersion; folder(qualifier)\n"


# Loop through the data frame rows to generate media entries
for index, row in df.iterrows():
    mediaFormat = '300Wx300H'
    code = str(row['images'])
    media = f'$siteResource/images/products/300Wx300H/{code}'
    folder = "images"
    impex_content += f";{mediaFormat};{code};{media};;;{folder};\n"

# Add MediaContainer entries
impex_content += """\n\nINSERT_UPDATE MediaContainer; qualifier[unique = true];$medias; $catalogVersion; conversionGroup(code[default = DefaultConversionGroup])\n"""

# Loop through the data frame rows to generate MediaContainer entries
for index, row in df.iterrows():
    qualifer = str(row['qualifer'])
    media = str(row['images'])
    impex_content += f";{qualifer};{media};;\n"

# Add Product updates
impex_content += """\n\nUPDATE Product; code[unique = true]; $picture; $thumbnail; $detail; $others; $normal; $thumbnails; $galleryImages; $catalogVersion\n"""

# Loop through the data frame rows to generate Product updates
for index, row in df.iterrows():
    codea = str(row['CODE ARTICLE'])
    media = str(row['images'])
    qualifer = str(row['qualifer'])
    impex_content += f";{codea};{media};{media};;;;;;;{qualifer};\n"

# Define the file path where the ImpEx script will be saved
file_path = "product_media.impex"

# Write the content to the file
with open(file_path, "w") as file:
    file.write(impex_content)

print(f"ImpEx script has been created at {file_path}")
