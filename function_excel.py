# FUNCTION GET COLUMN NUMBER IN EXCEL FILE
def getCellColumn(data):
    print('getCellColumn(data)')

    result = -1

    if(data == "po_number"):
        result = 1
    elif(data == "po_date"):
        result = 2
    elif(data == "company_name"):
        result = 3
    elif(data == "company_address"):
        result = 4
    elif(data == "company_city"):
        result = 5
    elif(data == "company_zip_code"):
        result = 6
    elif(data == "company_phone_number"):
        result = 7
    elif(data == "company_web_sites"):
        result = 8
    elif(data == "vendor_name"):
        result = 9
    elif(data == "vendor_address"):
        result = 10
    elif(data == "vendor_city"):
        result = 11
    elif(data == "vendor_zip_code"):
        result = 12
    elif(data == "vendor_phone_number"):
        result = 13
    elif(data == "ship_to_person"):
        result = 14
    elif(data == "ship_to_company_name"):
        result = 15
    elif(data == "ship_to_company_address"):
        result = 16
    elif(data == "ship_to_company_city"):
        result = 17
    elif(data == "ship_to_compnay_zip_code"):
        result = 18
    elif(data == "ship_to_company_phone_number"):
        result = 19
    elif(data == "delivery_date"):
        result = 20
    elif(data == "subtotal"):
        result = 21
    elif(data == "tax"):
        result = 22
    elif(data == "shipping"):
        result = 23
    elif(data == "total"):
        result = 24

    return result    
