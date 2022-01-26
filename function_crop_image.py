from PIL import Image

# THIS IS THE LIST OF POSITION LEFT, TOP, RIGHT, AND BOTTOM
# FOR THE IMAGE TO CROPP
croppImagePosition = []
croppImagePODatePosition = ("PO DATE", 2630, 1150, 3050, 1250)
croppImagePosition.append(croppImagePODatePosition)
croppImagePONumberPosition = ("PO NUMBER", 2630, 1275, 3350, 1375)
croppImagePosition.append(croppImagePONumberPosition)
croppImageCompanyNamePosition = ("COMPANY NAME", 500, 975, 1500, 1125)
croppImagePosition.append(croppImageCompanyNamePosition)
croppImageCompanyAddressPosition = ("COMPANY ADDRESS", 500, 1150, 2000, 1250)
croppImagePosition.append(croppImageCompanyAddressPosition)
croppImageCompanyCityPosition = ("COMPANY CITY", 500, 1275, 2000, 1375)
croppImagePosition.append(croppImageCompanyCityPosition)
croppImageCompanyPhoneNumberPosition = ("COMPANY PHONE NUMBER", 500, 1400, 1100, 1500)
croppImagePosition.append(croppImageCompanyPhoneNumberPosition)
croppImageCompanyWebSitesPosition = ("COMPANY WEB SITES", 500, 1525, 2000, 1625)
croppImagePosition.append(croppImageCompanyWebSitesPosition)
croppImageVendorNamePosition = ("VENDOR NAME", 500, 1950, 2000, 2075)
croppImagePosition.append(croppImageVendorNamePosition)
croppImageVendorAddressPosition = ("VENDOR ADDRESS", 500, 2075, 2000, 2200)
croppImagePosition.append(croppImageVendorAddressPosition)
croppImageVendorCityPosition = ("VENDOR CITY", 500, 2200, 2000, 2325)
croppImagePosition.append(croppImageVendorCityPosition)
croppImageVendorPhoneNumberPosition = ("VENDOR PHONE NUMBER", 500, 2325, 1100, 2450)
croppImagePosition.append(croppImageVendorPhoneNumberPosition)
croppImageShipToPersonPosition = ("SHIP TO PERSON", 2100, 1950, 3600, 2075)
croppImagePosition.append(croppImageShipToPersonPosition)
croppImageShipToCompanyNamePosition = ("SHIP TO COMPANY NAME", 2100, 2075, 3600, 2200)
croppImagePosition.append(croppImageShipToCompanyNamePosition)
croppImageShipToCompanyAddressPosition = ("SHIP TO COMPANY ADDRESS", 2100, 2200, 3600, 2325)
croppImagePosition.append(croppImageShipToCompanyAddressPosition)
croppImageShipToCompanyCityPosition = ("SHIP TO COMPANY CITY", 2100, 2325, 3600, 2450)
croppImagePosition.append(croppImageShipToCompanyCityPosition)
croppImageShipToCompanyPhoneNumberPosition = ("SHIP TO COMPANY PHONE NUMBER", 2100, 2450, 2700, 2575)
croppImagePosition.append(croppImageShipToCompanyPhoneNumberPosition)
croppImageDeliveryDatePosition = ("DELIVERY DATE", 2900, 2950, 3350, 3050)
croppImagePosition.append(croppImageDeliveryDatePosition)
croppImageSubTotalPosition = ("SUBTOTAL", 3000, 4435, 3650, 4535)
croppImagePosition.append(croppImageSubTotalPosition)
croppImageTaxPosition = ("TAX", 3000, 4535, 3650, 4635)
croppImagePosition.append(croppImageTaxPosition)
croppImageShippingPosition = ("SHIPPING", 3000, 4635, 3650, 4735)
croppImagePosition.append(croppImageShippingPosition)
croppImageTotalPosition = ("TOTAL", 3000, 4735, 3650, 4850)
croppImagePosition.append(croppImageTotalPosition)

# FUNCTION TO CROPP IMAGE AND SAVE TO A FILE
def croppImage(image2cropp, file2Save, croppPosition):
    print('croppImage()')
    print('image2cropp = ', image2cropp)
    print('file2Save = ', file2Save)

    result = "SUCCESS"
    
    croppData = [data for data in croppImagePosition if croppPosition == data[0]]
    print('croppData = ', croppData)
    position = croppData[0]
    print('position = ', position)

    try:
        im = Image.open(image2cropp)
        width, height = im.size

        imcropp = im.crop((position[1], position[2], position[3], position[4]))
        imcropp.save(file2Save)
    except Exception as e:
        print('ERROR = ', e)
        result = "FAIL"

    return result
