from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from PIL import Image, ImageChops, ImageStat
import sys
import math, operator
from shutil import copyfile
import ConfigParser
from optparse import OptionParser
from time import gmtime, strftime
from datetime import datetime, timedelta

#Argument options
parser = OptionParser()

parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

parser.add_option("-d", "--days", dest="days",
                  help="number of previous days to download", metavar="FILE")

(options, args) = parser.parse_args()



def ConfigSectionMap(section):
    dict1 = {}
    options = Config.options(section)
    for option in options:
        try:
            dict1[option] = Config.get(section, option)
            if dict1[option] == -1:
                DebugPrint("skip: %s" % option)
        except:
            print("exception on %s!" % option)
            dict1[option] = None
    return dict1

Config = ConfigParser.ConfigParser()
Config.read("./config.ini")
Config.sections
user = ConfigSectionMap("ING")['dni']
pwd = ConfigSectionMap("ING")['pwd']

infile="./tecladoImg.png"
numbers=[None]*10
code =[None]*10
region=[None]*10
codeinv =[None]*10
box=[(0,0,23,23),(28,0,51,23),(56,0,79,23),(84,0,107,23),(112,0,135,23),(0,28,23,51),(28,28,51,51),(56,28,79,51),(84,28,107,51),(112,28,135,51)]

def getOffset(number):
	#print "Decoding number", number
	#print "decoded to:", codeinv[number]
	#print "area:",box[codeinv[number]]
	xoffset=(box[codeinv[number]][0]+box[codeinv[number]][2])/2
	yoffset=(box[codeinv[number]][1]+box[codeinv[number]][3])/2
	#print "X,Y=",xoffset,yoffset
	return xoffset,yoffset

def pressKeyPad(number,mouse,elem):
	x,y=getOffset(number)
	mouse.move_to_element_with_offset(elem,x,y)
	mouse.click()
	return

# Function to compare images. Returns 0.0 when identical.
def rmsdiff(im1, im2):
    "Calculate the root-mean-square difference between two images"
    h = ImageChops.difference(im1, im2)
    stat = ImageStat.Stat(h)
    # calculate rms
    return stat.rms[0]

# Wait until all javascript is loaded
def isPageLoaded(driver):
	return driver.execute_script("return document.readyState") == "complete";

def waitPageLoaded(driver,timeout):
	elapsed=0
	while ( isPageLoaded(driver) != True ) and elapsed < timeout:
		#print "waiting to load.."
		time.sleep(1)
		elapsed+=1
	#print "loaded"

		
###################################################################
# main program
###################################################################

print "Kutxabank excel downloader"
print " Target output file is :",options.filename
print " Number of previous days to download :", options.days
finalDate = datetime.today()
finalDate_str = time.strftime('%Y-%m-%d', time.gmtime(time.time()))
daysint=int(options.days)
initialDate = datetime.today() - timedelta(days=daysint)
print initialDate
initialDate_str = initialDate.strftime('%Y-%m-%d')
print " Dates: " + initialDate_str + " - " + finalDate_str 
print "======================================================"

# Browser configuration for downloads
fp = webdriver.FirefoxProfile()
fp.set_preference('browser.download.folderList', 2) 
fp.set_preference('browser.download.manager.showWhenStarting', False)
fp.set_preference('browser.download.dir', '/tmp/')
fp.set_preference('browser.helperApps.neverAsk.openFile', 'text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml')
fp.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv,application/x-msexcel,application/excel,application/x-excel,application/vnd.ms-excel,image/png,image/jpeg,text/html,text/plain,application/msword,application/xml')
fp.set_preference('browser.helperApps.alwaysAsk.force', False)
fp.set_preference('browser.download.manager.alertOnEXEOpen', False)
fp.set_preference('browser.download.manager.focusWhenStarting', False)
fp.set_preference('browser.download.manager.useWindow', False)
fp.set_preference('browser.download.manager.showAlertOnComplete', False)
fp.set_preference('browser.download.manager.closeWhenDone', False)


driver = webdriver.Firefox(fp)
driver.implicitly_wait(60)

# website to access
driver.get("https://portal.kutxabank.es/cs/Satellite/kb/es/particulares")
assert "Kutxabank - Particulares" in driver.title
waitPageLoaded(driver,10)

# interaction
elem = driver.find_element_by_id("usuario")
elem.send_keys(user)
waitPageLoaded(driver,10)
elem = driver.find_element_by_id("password_PAS")
elem.click()
waitPageLoaded(driver,10)
time.sleep(3)
elem = driver.find_element_by_id("tecladoImg")
elem.screenshot("./tecladoImg.png")

#Load stored numbers to be compared later
try:
	for i in range(0,10):
		infile="./numbers/"+str(i)+".png"
		numbers[i]=Image.open(infile)
except IOError:
        print "IO Error"
        pass


infile="./tecladoImg.png"
#Capture keypad image from web and crop to number images
try:
	with Image.open(infile) as im:
		#print(infile, im.format, "%dx%d" % im.size, im.mode)
		for i in range(0,10):
			#print box[i]
			global region 
			region[i] = im.crop(box[i])
			outfile="currentbox"+str(i)+".png"
			region[i].save(outfile,"PNG")
except IOError:
	print "IO Error"
	pass


#Compare images and map positions to numbers
for i in range(0,10):
	for j in range(0,10):
		if rmsdiff(region[i], numbers[j]) == 0.0:
			code[i]=j
			codeinv[j]=i
			break
#print code[0:10]
#print codeinv[0:10]

#Press code in keypad
mouse = webdriver.ActionChains(driver)
for ch in pwd:
	#print ch
	pressKeyPad(int(ch),mouse,elem)
mouse.perform()

#Press "Entrar"
elem = driver.find_element_by_id("enviar")
elem.click()
waitPageLoaded(driver,10)

#Check que s'ha obert finestra
for handle in driver.window_handles:
    driver.switch_to_window(handle)
    if driver.title == "Kutxabank":
        break

waitPageLoaded(driver,10)

#Anar a element de baixar excel
elem = driver.find_element_by_id("formMenuSuperior:PanelSuperior:2:itemMenuSuperior")
time.sleep(0.1)
elem.click()
waitPageLoaded(driver,10)

#Consultas
elem = driver.find_element_by_id("formMenuAcciones:PanelAcciones:1:Accion")
time.sleep(0.1)
elem.click()
waitPageLoaded(driver,10)


#Anar a element de baixar excel
#Account number
#elem = driver.find_element_by_xpath("//div[@label='2095 5308 30 9115766544']")
#Consultas
elem = driver.find_element_by_id("formMenuOpciones:PanelSeries:0:SelectRadioMenuContratos:_2")
time.sleep(0.1)
elem.click()
waitPageLoaded(driver,10)

#Movements
elem = driver.find_element_by_id("formMenuOpciones:PanelSeries:0:j_id77:_0")
time.sleep(0.1)
elem.click()
waitPageLoaded(driver,10)

#Entre fechas
elem = driver.find_element_by_id("formCriterios:criteriosMovimientos:_5")
time.sleep(0.1)
elem.click()
waitPageLoaded(driver,10)

iniDia=str(initialDate.day)
iniMes=str(initialDate.month)
iniAny=str(initialDate.year)
endDia=str(finalDate.day)
endMes=str(finalDate.month)
endAny=str(finalDate.year)

elem = driver.find_element_by_id("formCriterios:calendarioDesde_cmb_dias")
elem.click()
time.sleep(0.1)
waitPageLoaded(driver,10)
elem.send_keys(iniDia)
waitPageLoaded(driver,10)
elem = driver.find_element_by_id("formCriterios:calendarioDesde_cmb_mes")
waitPageLoaded(driver,10)
elem.click()
time.sleep(0.1)
waitPageLoaded(driver,10)
elem.send_keys(iniMes)
time.sleep(0.1)
waitPageLoaded(driver,10)
elem = driver.find_element_by_id("formCriterios:calendarioDesde_cmb_anyo")
waitPageLoaded(driver,10)
elem.click()
time.sleep(0.1)
waitPageLoaded(driver,10)
elem.send_keys(iniAny)
time.sleep(0.1)
waitPageLoaded(driver,10)
time.sleep(0.1)
elem = driver.find_element_by_id("formCriterios:calendarioHasta_cmb_dias")
elem.click(); waitPageLoaded(driver,10)
time.sleep(0.1)
elem.send_keys(endDia)
waitPageLoaded(driver,10)
elem = driver.find_element_by_id("formCriterios:calendarioHasta_cmb_mes")
elem.click(); waitPageLoaded(driver,10)
time.sleep(0.1)
elem.send_keys(endMes)
waitPageLoaded(driver,10)
elem = driver.find_element_by_id("formCriterios:calendarioHasta_cmb_anyo")
elem.click(); waitPageLoaded(driver,10)
time.sleep(0.1)
elem.send_keys(endAny)
waitPageLoaded(driver,10)
time.sleep(0.1)

elem = driver.find_element_by_id("formCriterios:mostrar")
elem.click()
waitPageLoaded(driver,10)

elem = driver.find_element_by_id("formListado:resourceExcel")
elem.click()
waitPageLoaded(driver,10)


#Baixar excel i processar-lo amb el python-excel
time.sleep(5)
driver.close()

copyfile("/tmp/movimientos.xls",options.filename)



