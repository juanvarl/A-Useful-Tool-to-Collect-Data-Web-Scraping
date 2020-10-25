from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import pandas
import numpy
from openpyxl import Workbook
#The variable wb creates an excel workbook
wb = Workbook()
#Activate the active worksheet
ws = wb.active
#The variables options and preferences that the programmer wants enable or disable
caps = DesiredCapabilities().CHROME
caps["pageLoadStrategy"] = "normal"
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument(" - disable-infobars")
#The directory MUST be changed if the code is used in another computer
prefs={"download.default_directory" : "C:/Users/xxxxxx/Documents/Python Scripts/Driver/"}
chrome_options.add_experimental_option("prefs",prefs)
#The variables path and archive_excel have the routes to the WebDriver and the csv file with the list of ETFs
#Same as in line 15
path="C:/Users/xxxxxx/Documents/Python Scripts/Driver/chromedriver.exe"
archivo_excel='C:/Users/xxxxxx/Documents/Python Scripts/Web_Scrapping/Excel/Prueba.csv'
#df reads the file and assign into two variables, the size of the data that archivo_excel has
df=pandas.read_csv(archivo_excel, sep=',')
ren=df.shape[0]
col=df.shape[1]
#The variable nombres is a vector that has the names of all the ETFs the code is going to do web scraping
nombres=list(df.iloc[:,0])
#The variable drive has the path and all the preferences described in line 10
driver=webdriver.Chrome(path,chrome_options=chrome_options,desired_capabilities=caps)
x=ren
l=[]
y=[]
#The variable y has all the features that the code is going to extract from the web page
y=['ETF','Segmento','Score 1','Score 2','Net Asset Value (Yesterday)','Expense Ratio','Assets Under Management','Average Daily $ Volume', 'Holding_1', 'Holding_2', 'Holding_3', 'Holding_4', 'Holding_5', 'Holding_6', 'Holding_7', 'Holding_8', 'Holding_9', 'Holding_10']
#The variable y is inserted in the first row of the workbook that was created
ws.append(y)
#The code will repeat the next lines for each ETF stored in the variable nombres
for x in range(0,len(nombres)):
    y=[]
    #The variable etf will be the name of the ETF
    etf=nombres[x]
    #The WebDriver will open the requested web page with the specific ETF
    driver.get("http://www.etf.com/"+etf)
    #The WebDriver will search the text that the programmer wants to extract by searching the XPath
    segmento_html=driver.find_element_by_xpath('//*[@id="form-reports-header"]/div[1]/section[3]/div[1]/div[1]/div/a')
    #The variable segmento will store the text that the WebDriver found from the XPath
    segmento=segmento_html.text
    #The functions try and except are used in case the information the programmer wanted is not found
    try:
        score_1_html=driver.find_element_by_xpath('//*[@id="score"]/span/div[1]')
        score_1=score_1_html.text
        score_2_html =driver.find_element_by_xpath('//*[@id="score"]/span/div[2]')
        score_2=score_2_html.text
    except:
        score_1="NA"
        score_2="NA"
    try:
        NAV_html=driver.find_element_by_xpath('//*[@id="fundTradabilityData"]/div/div[16]/span')
        NAV=NAV_html.text
    except:
        NAV="NA"
    #The variable y will store all the features that were founded by the WebDriver
    y.extend([etf,segmento,score_1,score_2,NAV])
    #The features stored in summary_table had different structure inside the webpage for some ETFs. The rest of the
    #Structure remains as in line 43 of the code
    summary_table=[]
for p in range(4,7):
    #pdb.set_trace()
    try:
        if etf=="QQQ" or etf=="GDX" or etf=="VWO" or etf=="GDXJ" or etf=="VEA" or etf=="RSX" or etf=="OIH" or
        etf=="SMH" or etf=="VNQ" or etf=="VGK" or etf=="VOO":
            summary_data_html =driver.find_element_by_xpath('//*[@id="fundSummaryData"]/div/div['+str(p)+']/span')
            summary_data=summary_data_html.text
            summary_table.extend([summary_data])
            y.extend([summary_data])
        elif etf=="TQQQ" or etf=="JNUG" or etf=="NUGT" or etf=="UPRO" or etf=="SPXL" or etf=="TNA" or etf=="ERX":
            summary_data_html =driver.find_element_by_xpath(' //*[@id="fundSummaryData"]/div/div['+str(p-1)+']/span')
            summary_data=summary_data_html.text
            summary_table.extend([summary_data])
            y.extend([summary_data])
        else:
            summary_data_html =driver.find_element_by_xpath('//*[@id="fundSummaryData"]/div/div['+str(p+1)+']/span')
            summary_data=summary_data_html.text
            summary_table.extend([summary_data])
            y.extend([summary_data])
        except:
            summary_data="NA"
            summary_table.extend([summary_data])
            y.extend([summary_data])
#Same happens to variable holdings_table as in lines 64 and 65
holdings_table=[]
for p in range(1,11):
    #pdb.set_trace(), is used to track possible errors in the for loop
    try:
        holdings_html =driver.find_element_by_xpath('//*[@id="fit"]/div[1]/div[2]/div/div['+str(p)+']')
        holdings_data=holdings_html.text
        holdings_table.extend([holdings_data])
        y.extend([holdings_data])
    except:
        holdings_data="NA"
        holdings_table.extend([holdings_data])
        y.extend([holdings_data])
ws.append(y)
#Save the workbook in that route and with the name resultadoetf.csv. Same as in line 15
wb.save('C:/Users/xxxxxx/Documents/Python Scripts/Web_Scrapping/Excel/resultadoetf.csv')
#The WebDriver is closed
driver.quit()
