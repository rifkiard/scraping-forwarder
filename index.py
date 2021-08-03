from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as soup
import json
import xlsxwriter

url = 'https://idn.bizdirlib.com/taxonomy/term/901?page='
endPage = 330
company = []


def getDetailInformation(link):
    reqDetail = Request(link, headers={'User-Agent': 'Mozilla/5.0'})
    webPageDetail = urlopen(reqDetail).read()
    pageSoupDetail = soup(webPageDetail, "html.parser")
    listsDetail = pageSoupDetail.find(
        "div", {"id": "block-system-main"}).find('fieldset').find('ul').find_all('li')

    output = {}
    for liDetail in listsDetail:
        title = liDetail.find('strong').get_text()
        if liDetail.find('span') is not None:
            output[title] = liDetail.find('span').get_text()
        elif liDetail.find('a') is not None:
            output[title] = liDetail.find('a').get_text()
        else:
            detailValue = liDetail.get_text()
            listDetailValue = detailValue.split(":", 1)
            if len(listDetailValue) == 2:
                detailValue = listDetailValue[1].lstrip()
            output[title] = detailValue

    return output


def generateExcel(data):
    outWorkbook = xlsxwriter.Workbook("forwarders.xlsx")
    outSheet = outWorkbook.add_worksheet()

    outSheet.write("B2", "Company Name")
    outSheet.write("C2", "City/Province")
    outSheet.write("D2", "Country")
    outSheet.write("E2", "Address")
    outSheet.write("F2", "Zip Code")
    outSheet.write("G2", "International Area Code")
    outSheet.write("H2", "Phone")
    outSheet.write("I2", "Fax")
    outSheet.write("J2", "Category Activities")
    outSheet.write("K2", "Area")
    outSheet.write("L2", "Type")
    outSheet.write("M2", "Industry")

    count = 0
    for v in data:
        row = str(count + 3)
        companyName = str(v["Company Name"]) if "Company Name" in v else ""
        city = str(v["City/Province"]) if "City/Province" in v else ""
        country = str(v["Country"]) if "Country" in v else ""
        address = str(v["Address"]) if "Address" in v else ""
        zipCode = str(v["Zip Code"]) if "Zip Code" in v else ""
        areaCode = str(v["International Area Code"]
                       ) if "International Area Code" in v else ""
        phone = str(v["Phone"]) if "Phone" in v else ""
        fax = str(v["Fax"]) if "Fax" in v else ""
        categoryActivities = str(
            v["Category Activities"]) if "Category Activities" in v else ""
        area = str(v["Area"]) if "Area" in v else ""
        type_ = str(v["Type"]) if "Type" in v else ""
        industry = str(v["Industry"]) if "Industry" in v else ""

        count = count + 1

        outSheet.write("B" + row, companyName)
        outSheet.write("C" + row, city)
        outSheet.write("D" + row, country)
        outSheet.write("E" + row, address)
        outSheet.write("F" + row, zipCode)
        outSheet.write("G" + row, areaCode)
        outSheet.write("H" + row, phone)
        outSheet.write("I" + row, fax)
        outSheet.write("J" + row, categoryActivities)
        outSheet.write("K" + row, area)
        outSheet.write("L" + row, type_)
        outSheet.write("M" + row, industry)

    outWorkbook.close()


dataItteration = 1
for indexPage in range(endPage):
    print("Page " + str(indexPage + 1))
    currentPage = url + str(indexPage + 1)
    req = Request(currentPage, headers={'User-Agent': 'Mozilla/5.0'})
    webPage = urlopen(req).read()
    pageSoup = soup(webPage, "html.parser")
    lists = pageSoup.find_all('li', {"class": 'views-row'})

    for li in lists:
        anchor = li.find('a')
        urlDetailPage = "https://idn.bizdirlib.com" + anchor['href']
        company.append(getDetailInformation(urlDetailPage))
        print(dataItteration)
        dataItteration = dataItteration + 1

generateExcel(company)
