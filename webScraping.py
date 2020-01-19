import time

from selenium import webdriver
import xlsxwriter



# web driver open source tool for automated testing of webapps for chrome browser
driver = webdriver.Chrome(r"C:\Users\DekelRaz\PycharmProjects\webScraping\drivers\chromedriver.exe")

# set the amount of time by seconds for hte page to load before it throws an exception
driver.set_page_load_timeout(10)

# declare list of chosen companies from the website
listCompanies = ["470", "435", "3113", "15", "2356"]
company1 = "Mercaz Haformaika Averbuch INC"
company2 = "Adama Pitronot Lehaklaut INC"
company3 = "Overseas Commerce INC"
company4 = "Edgar Hashkaot Vepituach INC"
company5 = "Adri-El Israel Nechasim INC"

companiesNames = [company1, company2, company3, company4, company5]

# method that adding company's id to the url
def AddCompany(id):
    listCompanies.append(id)

# opens url in browser
driver.get("https://www.magna.isa.gov.il/Report.aspx?rid=2&eid=" + ', '.join(str(x) for x in listCompanies) + "&ifrs=2&p=42017&lang=1")

# gets basically all the numbers from the table by class name
# *************************************************GROSS MARGIN************************************
revah_golmis = []

# loop running on td number and increment to get the value of the wanted data, gets the data as string and convert it to int
for x in range(len(listCompanies)):
    index = x + 3
    url = "//*[@id=\"gvReportData\"]/tbody/tr[25]/td[" + str(index) + "]"

    text = str(driver.find_element_by_xpath(url).text)
    if text != ' ':
        revah_golmis.append(int(text.replace(',', '')))
    else:
        revah_golmis.append(0)


# ***************************************EBITDA*****************************************
ebitdas = []
# loop running on td number and increment to get the value of the wanted data, gets the data as string and convert it to int
for x in range(len(listCompanies)):
    index = x + 3
    url = "//*[@id=\"gvReportData\"]/tbody/tr[28]/td[" + str(index) + "]"

    text = str(driver.find_element_by_xpath(url).text)
    if text != ' ':
        ebitdas.append(int(text.replace(',', '')))
    else:
        ebitdas.append(0)

# **********************************************EPS************************************

epss = []
# loop running on td number and increment to get the value of the wanted data, gets the data as string and convert it to int
for x in range(len(listCompanies)):
    index = x + 3
    url = "//*[@id=\"gvReportData\"]/tbody/tr[38]/td[" + str(index) + "]"

    text = str(driver.find_element_by_xpath(url).text)
    if text != ' ':
        epss.append(float(text.replace(',', '')))
    else:
        epss.append(0)

# *******************************************PROPERTIES'S SUMMARY*****************************************
properties = []
# loop running on td number and increment to get the value of the wanted data, gets the data as string and convert it to int
for x in range(len(listCompanies)):
    index = x + 3
    url = "//*[@id=\"gvReportData\"]/tbody/tr[7]/td[" + str(index) + "]"

    text = str(driver.find_element_by_xpath(url).text)
    if text != ' ':
        properties.append(float(text.replace(',', '')))
    else:
        properties.append(-1)

# ***********************************************EARNINGS**************************************
earnings = []
# loop running on td number and increment to get the value of the wanted data, gets the data as string and convert it to int
for x in range(len(listCompanies)):
    index = x + 3
    url = "//*[@id=\"gvReportData\"]/tbody/tr[33]/td[" + str(index) + "]"

    text = str(driver.find_element_by_xpath(url).text)
    if text != ' ':
        earnings.append(float(text.replace(',', '')))
    else:
        earnings.append(-1)
# **********************************************FORTUNES******************************************
fortunes = []
# loop running on td number and increment to get the value of the wanted data, gets the data as string and convert it to int
for y in range(len(listCompanies)):
    index2 = y + 3
    url2 = "//*[@id=\"gvReportData\"]/tbody/tr[19]/td[" + str(index2) + "]"

    text2 = str(driver.find_element_by_xpath(url2).text)
    if text2 != ' ':
        fortunes.append(float(text2.replace(',', '')))
    else:
        earnings.append(-1)
# ******************************************ROE*******************************************
# calculating the vlaues of earnings and fortunes to get the value of ROE
roe = []
for z in range(len(listCompanies)):
    roe.append(float(earnings[z]/fortunes[z]))

# **************************************************Basic Earnings Power Ratio********************************
basic_earnings_power_ratio = []
# calculating the value of EBITDA and properties to get the value of BEPR
for x in range(len(listCompanies)):
    basic_earnings_power_ratio.append(float(ebitdas[x]/properties[x]))

# ******************************************Open The Project As EXCEL File***********************************

# declaring the financial ratios column
financial_ratios = []
financial_ratios.append("Gross Margin")
financial_ratios.append("ROE")
financial_ratios.append("EBITDA")
financial_ratios.append("BEPR")
financial_ratios.append("EPS")

# create file (workbook) and worksheet
outWorkbook = xlsxwriter.Workbook("magnaWebScraping.xlsx")
outsheet = outWorkbook.add_worksheet()

# write headers
outsheet.write("A1", "Financial Ratios")

# write data to file
for i in range(len(companiesNames)):
    outsheet.write(0, i + 1, companiesNames[i])
    outsheet.write(i + 1, 0, financial_ratios[i])
    outsheet.write(1, i + 1, revah_golmis[i])
    outsheet.write(2, i + 1, roe[i])
    outsheet.write(3, i + 1, ebitdas[i])
    outsheet.write(4, i + 1, basic_earnings_power_ratio[i])
    outsheet.write(5, i + 1, epss[i])


outWorkbook.close()

time.sleep(4)
driver.quit()
