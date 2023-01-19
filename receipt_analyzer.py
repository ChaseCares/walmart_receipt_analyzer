from openpyxl import Workbook
from datetime import datetime

import PyPDF2
import os
import re

ITEM_STATUS = ['Unavailable', 'Shopped', 'Substitutions', 'Weight-adjusted',
               'Refunded', 'No need to return this item', 'Canceled', 'Return complete', 'Return to Walmart store', ]


def printRed(text):
    print(f"\033[91m {text} \033[00m")


def printGreen(text):
    print(f"\033[92m {text} \033[00m")


def printBlue(text):
    print(f"\033[94m {text} \033[00m")


def listAllPDFs():
    PDFs = []
    for file in os.listdir('./pdf'):
        if file.endswith('.pdf'):
            PDFs.append(file)
    return PDFs


def printAndPop(list, num, reason):
    printRed(f"\tTrashing {num} lines. Reason: {reason}.")
    for _ in range(0, num):
        printRed(f"\t\tTrashing: {list[0]}")
        list.pop(0)
    return list


def cleanUp(PDF_Text):
    PDF_Text = PDF_Text.replace('You’re all set! ', '')

    months = ['Jan ', 'Feb ', 'Mar ', 'Apr ', 'May ', 'Jun ',
              'Jul ', 'Aug ', 'Sep ', 'Oct ', 'Nov ', 'Dec ']

    for _ in months:
        PDF_Text = PDF_Text.replace(_, f'\nskip\n{_}')

    for _ in ITEM_STATUS:
        PDF_Text = PDF_Text.replace(_, f'\n{_}')

    PDF_Text = PDF_Text.replace('$', '\n')
    PDF_Text = PDF_Text.replace('Subtotal', '\nSubtotal')
    PDF_Text = PDF_Text.replace('Qty ', '\nQty\n')

    PDF_Text = re.sub(r'\d/\d', '\n', PDF_Text)
    PDF_Text = re.sub(r'\d?\d of \d?\d', '\n', PDF_Text)
    PDF_Text = re.sub(r'\d?, \d?\d:\d?\d( AM| PM)?', '\n', PDF_Text)
    PDF_Text = re.sub(r'(\d?\d, \d:\d\d (AM|PM) )?Order details - Walmart.com', '\n', PDF_Text)
    PDF_Text = re.sub(r'https://www ?.walmart.com/orders/\d+', '\n', PDF_Text)

    PDF_Text = PDF_Text.splitlines()

    PDF_Text = [x.strip(' ') for x in PDF_Text]

    return list(filter(None, PDF_Text))


def getOrderInfo(PDF_Text):
    orderNumRaw = re.search(r'Order# \d+-\d+', PDF_Text).group().replace('Order# ', '')
    orderNum = orderNumRaw.replace('-', '')

    try:
        donationRaw = re.search(r'Donation to [a-zA-Z\d ]+\$\d+\.\d\d', PDF_Text).group()
        donation = donationRaw.split('  $')
    except AttributeError:
        donation = [None, None]

    try:
        product = re.search(r'Product  ?-?\$\d+.\d\d', PDF_Text).group().replace('Product', '').replace('$', '').strip()
    except AttributeError:
        product = None

    return [
        orderNumRaw,
        re.search(r'Subtotal \$\d+\.\d\d', PDF_Text).group().replace('Subtotal $', ''),
        donation[0],
        donation[1],
        product,
        re.search(r'Taxes \$\d+\.\d\d', PDF_Text).group().replace('Taxes $', ''),
        re.search(r'Total \$\d+\.\d\d', PDF_Text).group().replace('Total $', ''),
        re.search(r'Ending in \d\d\d\d', PDF_Text).group().replace('Ending in ', ''),
        re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d\d, \d\d\d\d', PDF_Text).group(),
        f'https://www.walmart.com/orders/{orderNum}']


def returnItem(PDF_List):
    itemName = ''
    for i in range(1, 4):
        itemName += PDF_List[i-1] + ' '
        if PDF_List[i] in ITEM_STATUS:
            if PDF_List[i+1] == 'Qty':
                item = [itemName.strip(), PDF_List[i], PDF_List[i+2], PDF_List[i+3]]
                return PDF_List[i+4:], item
            else:
                item = [itemName.strip(), PDF_List[i], 'no qty', PDF_List[i+1]]
                return PDF_List[i+2:], item

        if PDF_List[i] == 'Qty':
            item = [itemName.strip(), 'no status', PDF_List[i+1], PDF_List[i+2]]
            return PDF_List[i+3:], item

    else:
        PDF_List = printAndPop(PDF_List, 4, 'No item found')
        return PDF_List, False


def main():
    wb = Workbook()

    itemsLabels = ['Item', 'Status', 'Qty', 'Price']
    wsItems = wb.active
    wsItems.append(itemsLabels)
    wsItems.title = 'Items'

    orderLabels = ['Order Number', 'Raw Subtotal', 'Subtotal', 'Donation Recipient',
                   'Donation Total', 'Product?', 'Taxes', 'Total', 'Ending in ', 'Date', 'Order Link']
    wsOrder = wb.create_sheet('Order Info')
    wsOrder.append(orderLabels)
    wsOrder.title = 'Order'

    for PDF in listAllPDFs():
        printBlue(f"Processing PDF: {PDF}")
        with open(f'./pdf/{PDF}', 'rb') as PDF_File:

            PDF_Text = ''
            pdf = PyPDF2.PdfReader(PDF_File)
            for page in range(len(pdf.pages)):
                PDF_Text += pdf.pages[page].extract_text()

            PDF_Text = PDF_Text.replace(u'\xa0', u' ')
            PDF_List = cleanUp(PDF_Text)

            rawSubtotal = 0
            while len(PDF_List) > 3:
                PDF_List, item = returnItem(PDF_List)

                if item:
                    printGreen(f"\tAdding {item}")
                    wsItems.append(item)

                    rawSubtotal += float(item[-1])

                    if PDF_List[0] == 'skip':
                        PDF_List = printAndPop(PDF_List, 3, 'Skipping unnecessary lines between pages')
                    if PDF_List[0] == 'Subtotal':
                        printRed(f"\tSubtotal trash: {PDF_List}")
                        PDF_List = []

            orderInfo = getOrderInfo(PDF_Text)
            orderInfo.insert(1, rawSubtotal)
            wsOrder.append(orderInfo)

    destFilename = 'walmart'
    currentTime = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

    wb.save(filename=f"{destFilename}_{currentTime}.xlsx")


if __name__ == '__main__':
    main()
