from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import calendar
from pathlib import Path

global DownloadDir
global LatestFile
DownloadDir='D:\\trade\\downloads\\today'
LatestFile='D:\\trade\\downloads\\latest'
# To prevent download dialog
profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2) # custom location
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', DownloadDir)
profile.set_preference('browser.helperApps.neverAsk.openFile', 'text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml,application/zip')
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml,application/zip')
profile.set_preference('browser.helperApps.alwaysAsk.force', False)
profile.set_preference('browser.download.manager.alertOnEXEOpen', False)
profile.set_preference('browser.download.manager.focusWhenStarting', False)
profile.set_preference('browser.download.manager.useWindow', False)
profile.set_preference('browser.download.manager.showAlertOnComplete', False)
profile.set_preference('browser.download.manager.closeWhenDone', False)
browser = webdriver.Firefox(profile)
def monthToNum(shortMonth):

 return{
        'Jan' : '01',
        'feb' : '02',
        'mar' : '03',
        'apr' : '04',
        'may' : '05',
        'jun' : '06',
        'jul' : '07',
        'aug' : '08',
        'sep' : '09',
        'oct' : '10',
        'nov' : '11',
        'dec' : '12'
}[shortMonth]
#download trade statistics
browser.get("https://www.nseindia.com/live_market/dynaContent/live_watch/equities_stock_watch.htm")
#extract date of last report
LastReport=browser.find_element_by_id('status1').text
timeline=LastReport.split(' ')
year=timeline[-1]
date=timeline[-2][:-1]
month=(timeline[-3])
state=(timeline[-4])
if (state=="Open"):
    browser.quit()
    exit(-1)
date_full=date + '-' + monthToNum(month.lower()) + '-' + year
FilePath=DownloadDir+'\\'+'date'
infile = open(LatestFile, 'r')
firstLine = infile.readline()
infile.close()

if (firstLine == date_full):
       browser.quit()
       exit(-7)

#create date file for reference
f=open(FilePath,"w+")
f.write(date_full+'\n')

select = Select(browser.find_element_by_id('bankNiftySelect'))
select.select_by_visible_text('FO Stocks')
browser.find_element_by_link_text('Download in csv').click()
browser.quit()
#download cm bhavcopy
cm_base_url='https://www.nseindia.com/content/historical/EQUITIES/'
cm_url=cm_base_url+year+'/'+month.upper()+'/'+'cm'+date+month.upper()+year+'bhav.csv.zip'


fo_base_url='https://www.nseindia.com/content/historical/DERIVATIVES/'
fo_url=fo_base_url+year+'/'+month.upper()+'/'+ 'fo'+date+month.upper()+year+'bhav.csv.zip'
#browser.get("https://www.nseindia.com/products/content/all_daily_reports.htm?param=equity")

mto_base_url='https://www.nseindia.com/archives/equities/mto/MTO_'
mto_url=mto_base_url+date+monthToNum(month.lower())+year+'.DAT'

f.write(cm_url+'\n')
f.write(fo_url+'\n')
f.write(mto_url+'\n')
f.close()

infile = open(LatestFile, 'w')
infile.write(date_full+'\n')
infile.close()
