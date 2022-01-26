import easyocr

# FUNCTION GET TEXT FROM IMAGE WITH EASYOCR
def getSimpleTextFromImage(imagePath):
    print('getSimpleTextFromImage()')

    result = ""

    try:
        reader = easyocr.Reader(['en'], gpu=False)
        resultReadImage = reader.readtext(imagePath)

        order = 0
        for data in resultReadImage:
            order+=1
            result = result + " " + str(data[1]) 
    except Exception as e:
        print('e = ', e)
    result = result.strip()

    return result