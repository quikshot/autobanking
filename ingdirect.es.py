from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from PIL import Image, ImageChops, ImageStat
import sys
import math, operator

import ConfigParser
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
driver.implicitly_wait(10)
driver.get("https://portal.kutxabank.es/cs/Satellite/kb/es/particulares")
assert "Kutxabank - Particulares" in driver.title
elem = driver.find_element_by_id("usuario")
elem.send_keys(user)
elem = driver.find_element_by_id("password_PAS")
elem.click()
#elem.send_keys(pwd)
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

#Check que s'ha obert finestra
for handle in driver.window_handles:
    driver.switch_to_window(handle)
    if driver.title == "Kutxabank":
        break


#Anar a element de baixar excel
elem = driver.find_element_by_id("formMenuSuperior:PanelSuperior:2:itemMenuSuperior")
elem.click()

#time.sleep(3)
#Consultas
elem = driver.find_element_by_id("formMenuAcciones:PanelAcciones:1:Accion")
elem.click()


#Anar a element de baixar excel
#Account number
#elem = driver.find_element_by_xpath("//div[@label='2095 5308 30 9115766544']")
#Consultas
elem = driver.find_element_by_id("formMenuOpciones:PanelSeries:0:SelectRadioMenuContratos:_2")
elem.click()
#Movements
elem = driver.find_element_by_id("formMenuOpciones:PanelSeries:0:j_id77:_0")
elem.click()

#Entre fechas
elem = driver.find_element_by_id("formCriterios:criteriosMovimientos:_5")
elem.click()

iniDia=1
iniMes=1
iniAny=2017
endDia=1
endMes=7
endAny=2017

elem = driver.find_element_by_id("formCriterios:calendarioDesde_cmb_dias")
elem.click()
elem.send_keys(iniDia)
elem = driver.find_element_by_id("formCriterios:calendarioDesde_cmb_mes")
elem.click()
elem.send_keys(iniMes)
elem = driver.find_element_by_id("formCriterios:calendarioDesde_cmb_anyo")
elem.click()
elem.send_keys(iniAny)

elem = driver.find_element_by_id("formCriterios:calendarioHasta_cmb_dias")
elem.click()
elem.send_keys(endDia)
elem = driver.find_element_by_id("formCriterios:calendarioHasta_cmb_mes")
elem.click()
elem.send_keys(endMes)
elem = driver.find_element_by_id("formCriterios:calendarioHasta_cmb_anyo")
elem.click()
elem.send_keys(endAny)

elem = driver.find_element_by_id("formCriterios:mostrar")
elem.click()

elem = driver.find_element_by_id("formListado:resourceExcel")
elem.click()



#Baixar excel i processar-lo amb el python-excel

#elem.send_keys(Keys.RETURN)
#driver.close()
