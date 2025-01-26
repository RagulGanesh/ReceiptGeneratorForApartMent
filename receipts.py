from docxtpl import DocxTemplate
import pandas as pd


def readExcel():
    excel_file = 'Apartment_Data.xlsx'
    data = pd.read_excel(excel_file, dtype={'Date_0': str, 'Date_1' : str}, sheet_name= 'Dec-2024')
    excelData = []
    data['Date_0'] = pd.to_datetime(data['Date_0']).dt.strftime('%d-%m-%Y')
    data['Date_1'] = pd.to_datetime(data['Date_1']).dt.strftime('%d-%m-%Y')
    for index, row in data.iterrows():
        receipt_no = row['ReceiptNo']
        date_0 = row['Date_0'].split(' ')[0]
        received_from = row['Received From']
        flat_no = row['FlatNo']
        towards = row['Towards']
        rupees = row['Rupees']
        rupees_in_words = row['Rupees In Words']
        modeOfPayment = row['Mode Of Payment']
        date_1 = row['Date_1'].split(' ')[0]
        excelData.append([receipt_no,date_0,received_from,flat_no,towards,rupees,rupees_in_words,modeOfPayment,date_1])
    return excelData

      
doc = DocxTemplate("receipt.docx")
excelInfo = readExcel()
i=1
arr = {}
for info in excelInfo : 
    context = {
        f'r_number_{i}' : info[0],
        f'date_{i}' : info[1],
        f'receivedFrom_{i}' : info[2],
        f'flatNo_{i}' : info[3],
        f'towards_{i}' : info[4],
        f'rupees_{i}' : info[5],
        f'rupeesInWords_{i}' : info[6],
        f'modeOfPayment_{i}' : info[7],
        f'dt_{i}' : info[8]
    }
    arr.update(context)
    # print(context,end="\n")
    i = i + 1
# context = { 'r_number' : "2423",
#            'date' : '6-7-24',
#             'receivedFrom' : 'Sri. R. Ganesh Babu',
#             'flatNo' : '304',
#             'towards' : 'maintenanace for july-24',
#             'rupees' : '1200/-',
#             'rupeesInWords' : "one thousand two hundred only )",
#             'modeOfPayment' : "By Cash",
#             'dt' : ''}
doc.render(arr)
doc.save("generated_doc.docx")
print(arr)