import os

def open_doc(billtype):
    if billtype=='Payments':
        os.startfile('K:\\Client\\Service_bureau\\Staff Tools\\Python\\Documentation\\Payments.docx')
    elif billtype=='Verify zip':
        os.startfile('K:\\Client\\Service_bureau\\Staff Tools\\Python\\Documentation\\VerifyZip.docx')
    elif billtype=='Move FF':
        os.startfile('K:\\Client\\Service_bureau\\Staff Tools\\Python\\Documentation\\MoveFFInvoices.docx')
    elif billtype=='Move Secabs':
        os.startfile('K:\\Client\\Service_bureau\\Staff Tools\\Python\\Documentation\\MoveSecabs.docx')
    elif billtype=='Clli mapping':
        os.startfile('K:\\Client\\Service_bureau\\Staff Tools\\Python\\Documentation\\ClliMapping.docx')
    elif billtype=='Error files':
        os.startfile('K:\Client\Service_bureau\Staff Tools\Python\Documentation\ExtractErrorFiles.docx')