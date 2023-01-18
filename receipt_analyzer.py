from openpyxl import Workbook

import datetime
import PyPDF2
import os
import re

ITEM_STATUS = ['Unavailable', 'Shopped', 'Substitutions', 'Weight-adjusted',
               'Refunded', 'No need to return this item', 'Canceled']


def printRed(text):
    print(f"\033[91m {text} \033[00m")


def printGreen(text):
    print(f"\033[92m {text} \033[00m")


def returnItem(text):
    if text[1] in ITEM_STATUS:
        item = [text[0], text[1], text[2], text[3]]
        return text[4:], item
    elif text[2] in ITEM_STATUS:
        itemName = f"{text[0]} {text[1]}"
        item = [itemName, text[2], text[3], text[4]]
        return text[5:], item
    elif text[3] in ITEM_STATUS:
        itemName = f"{text[0]} {text[1]} {text[2]}"
        item = [itemName, text[3], text[4], text[5]]
        return text[6:], item
    else:
        text = printAndPop(text, 4)
        return text, False


def printAndPop(list, num):
    for _ in range(0, num):
        printRed(f"\tTrashing: |{list[0]}|, Num: {num}.")
        list.pop(0)
    return list


def cleanUp(PDFText):
    PDFText = PDFText.replace('You’re all set! ', '')

    months = ['Jan ', 'Feb ', 'Mar ', 'Apr ', 'May ', 'Jun ',
              'Jul ', 'Aug ', 'Sep ', 'Oct ', 'Nov ', 'Dec ']

    for _ in months:
        PDFText = PDFText.replace(_, f'\nskip\n')

    for _ in ['Qty', '$', ]:
        PDFText = PDFText.replace(_, '\n')

    for _ in ITEM_STATUS:
        PDFText = PDFText.replace(_, f'\n{_}')

    PDFText = PDFText.replace('Subtotal', '\nSubtotal')

    PDFText = re.sub(r'\d/\d', '\n', PDFText)
    PDFText = re.sub(r'\d?\d of \d?\d', '\n', PDFText)
    PDFText = re.sub(r'\d?, \d+:\d+', '\n', PDFText)
    PDFText = re.sub(r'(\d?\d, \d:\d\d (AM|PM) )?Order details - Walmart.com', '\n', PDFText)
    PDFText = re.sub(r'https://www ?.walmart.com/orders/\d+', '\n', PDFText)

    PDFText = PDFText.splitlines()

    PDFText = [x.strip(' ') for x in PDFText]

    PDFText = printAndPop(PDFText, 4)

    return list(filter(None, PDFText))


def listAllPDFs():
    PDFs = []
    for file in os.listdir('./pdf'):
        if file.endswith('.pdf'):
            PDFs.append(file)
    return PDFs


def main():
    wb = Workbook()
    ws = wb.active

    labels = ['Item', 'Status', 'Qty', 'Price']
    ws.append(labels)

    for PDF in listAllPDFs():
        printGreen(f"Processing PDF: {PDF}")
        with open(f'./pdf/{PDF}', 'rb') as f:

            PDFText = ''
            pdf = PyPDF2.PdfReader(f)
            for page in range(0, len(pdf.pages)):
                PDFText += pdf.pages[page].extract_text()

            PDFList = cleanUp(PDFText)

            while len(PDFList) > 3:
                PDFList, item = returnItem(PDFList)

                if item:
                    if PDFList[0] == 'skip':
                        PDFList = printAndPop(PDFList, 3)
                    elif PDFList[0] == 'Subtotal':
                        printRed(f"\tSubtotal trash: {PDFList}")
                        PDFList = []

                    printGreen(f"\tAdding {item}")
                    ws.append(item)

    destFilename = 'walmart'
    wb.save(filename=f"{destFilename} {datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")


if __name__ == '__main__':
    main()
