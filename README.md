# JUDS-spreadsheet-tool

This is a beta version

google add-on for JUDS  Our spreadsheet tool allows you to get the best translation from one sector code to another or to retrieve a relevant international code for a given country from directly within the spreadsheet you are working in.

Currently the tool is available as a Google Sheet Add-on. It can be used directly in a Google Sheet and results can be downloaded in Microsoft Excel format. A Microsoft Excel macro will be available in the future.

How to get the tool

    From your active Google Sheet, select “Get add-ons” from the top “Add-ons” menu.
    Search for "Joined-up Data Standard Navigator spreadsheet tool"
    Click on '+ FREE'
    Sign in to your Google Account if prompted
    Review the summary of add-on’s requirements and select 'Allow'.
    JUDS Navigator spreadsheet tool’ is now available in your 'Add-on' dropdown menu within your Google Sheet.

Function: translateSectorCode()

This function returns the best translation between codes used by data standards describing socio-economic sectors and economic activities. These include: OECD DAC CRS (crs), UN Classification of Functions of Government (cofog), UN International Standard Industrial Classification of All Economic Activities (isic), IRS National Taxonomy of Exemption Entities (ntee) or World Bank themes (world_bank_themes) and sectors (world_bank_sectors).

The abbreviations used by the function to recognise individual data standards are shown in bold red and in brackets in the list above.

To use the function, select the cells containing, or enter directly into the function:

    where to translate the code from,
    which standard to translate to
    the sector code from the data standard you wish to translate from

Example: to translate from OECD DAC CRS code 323 (“Construction”) to UN International Standard Industrial Classification of All Economic Activities, the function would be:

Cell A1 denotes the starting classification, B1 is the standard to translate to and C1 is the code I need to translate. D1 denotes the cell where the output to the above query is displayed. In this example, “F” is the code for “construction” in ISIC data standard.

This function first searches for an exact or close match. If there is not an exact or close match, the function delivers a broader search result, along with a notification.

When no match can be found the function simply returns a notification.

Function: translateGeoCode()

This function allows you to retrieve the following codes for each country: OECD DAC recipient or donor code (dac), UN FAO country code (faostat_num), FAO Global Administrative Unit Layer (gaul_num), IMF IFS code (imf_ifs), ISO 2 (iso_2), ISO 3 (iso_3), UN country code (un_num).

The abbreviations used by the function to recognise individual data standards are shown in bold red and in brackets in the list above.

If there is no code that matches user query, the function will return a notification

This function is not case sensitive and will attempt to find a given country name, even if an incomplete name is entered. For example, searching “Sao” will render results for São Tomé and Príncipe and searching for “Côte” will render results for “Côte d’Ivoire”.
