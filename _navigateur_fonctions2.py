import time
from selenium import webdriver
#import os
import sys
#from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
#from selenium.webdriver import Firefox
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import Commun.Libs._log_fonctions as journal


def OuvertureNavigateur(navigateur, tmp_url):
    global processusNavigateurExart
    
    if (tmp_url == '') :
        OuvertureNavigateur = -1
        sys.exit()
    else:
        tmp_url = tmp_url.strip()
        
    processusNavigateurExart = ""
   
    if navigateur == "IE":
        processusNavigateurExart = webdriver.Firefox()
    if navigateur == "FF":
        #processusNavigateurExart = webdriver.Firefox()
        #profile = webdriver.FirefoxProfile()
        #profile.set_preference("dom.file.createInChild",True)
        #profile.update_preferences() 
        capabilities = webdriver.DesiredCapabilities.FIREFOX.copy()
        #journal.Log(capabilities)
        capabilities['marionette'] = True
        
        """
        profile.set_preference('browser.download.folderList', 2)
        profile.set_preference('browser.download.manager.showWhenStarting', False)
        profile.set_preference('browser.download.dir', os.getcwd())
        profile.set_preference('browser.helperApps.neverAsk.saveToDisk', ('application/vnd.ms-excel'))
        #profile.setPreference("browser.helperApps.alwaysAsk.force", False)
        profile.set_preference('general.warnOnAboutConfig', False)
        profile.update_preferences() 
        """
        #processusNavigateurExart = webdriver.Firefox(capabilities)
        processusNavigateurExart = webdriver.Firefox(capabilities= capabilities, executable_path='geckodriver')
    if navigateur == "CH":
        processusNavigateurExart = webdriver.Chrome()
        #processusNavigateurExart.maximize_window
        processusNavigateurExart.set_window_size(1600,900)
	
    try:
        processusNavigateurExart.get(tmp_url)
        OuvertureNavigateur = 0
        time.sleep(3)
    except:
        processusNavigateurExart.get(tmp_url)
        OuvertureNavigateur = -1
        journal.Log("Page load Timeout Occured. Quiting !!!")
        
    return OuvertureNavigateur
	
    
def FermetureNavigateur():
    processusNavigateurExart.quit()
    return 1


def FermetureOngletActuel():
	processusNavigateurExart.close()
    
def alert():    
    try:
        WebDriverWait(processusNavigateurExart, 4).until(EC.alert_is_present(),
                                                  'Timed out waiting for PA creation ' +
                                                  'confirmation popup to appear.')
        alert = processusNavigateurExart.switch_to.alert
        journal.Log("alert : "  + alert.text)
        alert.accept()
        journal.Log("Alert accepted")
        
    except TimeoutException:
        pass

def check_exists_elementsXpath(element):

    global processusNavigateurExart       
    try:
        processusNavigateurExart.find_elements_by_xpath(element)
        return 1
    except:
        return 0  
    
def check_exists_disable(element):

    global processusNavigateurExart 
    retour = 0      
    try:
        processusNavigateurExart.find_element_by_id(element).is_enabled()
        return 1
    except:
        retour = 0  
    if retour ==0:       
        try:
            processusNavigateurExart.find_element_by_xpath(element).is_enabled()
            return 1
        except:
            return 0 
    if retour ==0:       
        try:
            processusNavigateurExart.find_element_by_name(element).is_enabled()
            return 1
        except:
            return 0 
        
def ClicSurCheckbox(element):
    
    global processusNavigateurExart
    if check_exists_disable(element) == 1:        
        try:
            checkbox = processusNavigateurExart.find_element_by_xpath(element)            
            if checkbox.is_selected() == False:
                checkbox.click()
            return 1
        except: 
            retour = 0   
    if retour == 0:
        if check_exists_disable(element) == 1:        
            try:
                checkbox = processusNavigateurExart.find_element_by_name(element)            
                if checkbox.is_selected() == False:
                    checkbox.click()
                return 1
            except: 
                retour = 0  
    if retour == 0:
        if check_exists_disable(element) == 1:        
            try:
                checkbox = processusNavigateurExart.find_element_by_id(element)            
                if checkbox.is_selected() == False:
                    checkbox.click()
                return 1
            except: 
                return 0             
            
def check_exists_element(element):
    retour = 0
    global processusNavigateurExart
    
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_id(element)
            return 1
        except: 
            retour = 0
    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_xpath(element)
            return 1
        except:
            retour = 0            
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_name(element)
            return 1
          
        except:
            retour = 0
    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_class_name(element)
            return 1
            
        except:
            retour = 0

    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_css_selector(element)
            return 1
        except:
            retour = 0
    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_link_text(element)
            return 1
        except:
            retour = 0
    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_partial_link_text(element)
            return 1
        except:
            retour = 0
    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_tag_name(element)
            return 1
        except:
            return 0          

def RemplissageChampTexte(element, valeur):
    tmpData = ''
    retour = 0
    global processusNavigateurExart
    
    tmpData = valeur.strip()
    
    if tmpData == "":
        return
        
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_id(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_id(element).send_keys(tmpData)
            retour = 1
        except:
            retour = 0
            
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_name(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_name(element).send_keys(tmpData)
            retour = 1  
        except:
            retour = 0
            
    if retour == 0:       
        try:
            processusNavigateurExart.find_element_by_class_name(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_class_name(element).send_keys(tmpData)
            retour = 1 
        except:
            retour = 0
     
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_xpath(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_xpath(element).send_keys(tmpData)
            retour = 1
        except:
            retour = 0
            
    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_css_selector(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_css_selector(element).send_keys(tmpData)
            retour = 1
        except:
            retour = 0
         
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_link_text(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_link_text(element).send_keys(tmpData)
            retour = 1
        except:
            retour = 0
            
    if retour == 0:    
        try:
            processusNavigateurExart.find_element_by_partial_link_text(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_partial_link_text(element).send_keys(tmpData)
            retour = 1
        except:
            retour = 0
            
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_tag_name(element).clear()
            time.sleep(1)
            processusNavigateurExart.find_element_by_tag_name(element).send_keys(tmpData)
            retour = 1
        except:
            retour = 0 
    return retour

            
def RemplissageDefaultvalue(element, valeur):
    tmpData = ''
    retour = 0
    imputvalue = ''
    
    global processusNavigateurExart
    
    tmpData = valeur.strip()
    
    if tmpData == "":
        return
        
    if retour == 0:
        try:
            imputvalue = processusNavigateurExart.find_element_by_id(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)
            return 1
        except:
            retour = 0
    if retour == 0:
        try:
            imputvalue = processusNavigateurExart.find_element_by_xpath(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)
            return 1
        except:
            retour = 0                        
    if retour == 0:
        try:
            imputvalue = processusNavigateurExart.find_element_by_name(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)            
            return 1  
        except:
            retour = 0
            
    if retour == 0:       
        try:
            imputvalue = processusNavigateurExart.find_element_by_class_name(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)            
            return 1 
        except:
            retour = 0
            
    if retour == 0:        
        try:
            imputvalue = processusNavigateurExart.find_element_by_css_selector(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)            
            return 1
        except:
            retour = 0
         
    if retour == 0:
        try:
            imputvalue = processusNavigateurExart.find_element_by_link_text(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)            
            return 1
        except:
            retour = 0
            
    if retour == 0:    
        try:
            imputvalue = processusNavigateurExart.find_element_by_partial_link_text(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)            
            return 1
        except:
            retour = 0
            
    if retour == 0:
        try:
            imputvalue = processusNavigateurExart.find_element_by_tag_name(element)
            processusNavigateurExart.execute_script("arguments[0].value = ''", imputvalue)
            imputvalue.send_keys(tmpData)            
            return 1
        except:
            return 0

def RemplissageSimple(element, valeur):
    tmpData = ''
    retour = 0
    global processusNavigateurExart
    
    tmpData = valeur.strip()
    
    if tmpData == "":
        return
        
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_id(element).send_keys(tmpData)
            return 1
        except:
            retour = 0
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_xpath(element).send_keys(tmpData)
            return 1
        except:
            retour = 0                        
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_name(element).send_keys(tmpData)
            return 1  
        except:
            retour = 0
            
    if retour == 0:       
        try:
            processusNavigateurExart.find_element_by_class_name(element).send_keys(tmpData)
            return 1 
        except:
            retour = 0
            
    if retour == 0:        
        try:
            processusNavigateurExart.find_element_by_css_selector(element).send_keys(tmpData)
            return 1
        except:
            retour = 0
         
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_link_text(element).send_keys(tmpData)
            return 1
        except:
            retour = 0
            
    if retour == 0:    
        try:
            processusNavigateurExart.find_element_by_partial_link_text(element).send_keys(tmpData)
            return 1
        except:
            retour = 0
            
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_tag_name(element).send_keys(tmpData)
            return 1
        except:
            return 0

def ClicSurElement(element):
    retour = 0
    global processusNavigateurExart
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_id(element).click()
            return 1
        except: 
            retour = 0
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_xpath(element).click()
            return 1
        except:
            retour = 0                        
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_name(element).click()
            return 1
          
        except:
            retour = 0
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_class_name(element).click()
            return 1
            
        except:
            retour = 0
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_css_selector(element).click()
            return 1
        except:
            retour = 0
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_link_text(element).click()
            return 1
        except:
            retour = 0
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_partial_link_text(element).click()
            return 1
        except:
            retour = 0
    if retour == 0:
        try:
            processusNavigateurExart.find_element_by_tag_name(element).click()
            return 1
        except:
            return 0
        
def SelectedOptionXpath(element):

    global processusNavigateurExart
    try:
        select = Select(processusNavigateurExart.find_element_by_xpath(element))           
        return select.first_selected_option.text
    except:
        return  ""

def SelectElement(element, valeur, defaultValue):
    tmpData = ''
    tmpData2 = ''
    retour = 0
    global processusNavigateurExart
    #print(valeur)
    tmpData = valeur.strip()
    tmpData2 = defaultValue.strip()
    
    if tmpData == "":
        return 
    """
    i = 0
    for option in select.options:                     
        if option.lower().strip() == tmpData.lower():
            select.select_by_visible_text(option)
            return 1
        else:
            if (i + 1) == len(select.options):
                select.select_by_visible_text(tmpData2)
                return -1
        i = i + 1
    """
    if retour == 0:
        try:
            select = Select(processusNavigateurExart.find_element_by_id(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1   
        except:
            retour = 0
    if retour == 0:       
        try:
            select = Select(processusNavigateurExart.find_element_by_xpath(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1
        except:
            retour = 0                       
    if retour == 0:        
        try:
            select = Select(processusNavigateurExart.find_element_by_name(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1        
        except:
            retour = 0
    if retour == 0:       
        try:
            select = Select(processusNavigateurExart.find_element_by_class_name(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1               
        except:
            retour = 0

    if retour == 0:         
        try:
            select = Select(processusNavigateurExart.find_element_by_css_selector(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1
        except:
            retour = 0
            
    if retour == 0:        
        try:
            select = Select(processusNavigateurExart.find_element_by_link_text(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1
        except:
            retour = 0
    if retour == 0:        
        try:
            select = Select(processusNavigateurExart.find_element_by_partial_link_text(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1
        except:
            retour = 0
    if retour == 0:        
        try:
            select = Select(processusNavigateurExart.find_element_by_tag_name(element))
            try:
                select.select_by_visible_text(tmpData)
                return 1
            except:
                select.select_by_visible_text(tmpData2)
                return -1
        except:
            return 0

def SelectElementByIndex(element, tmpData, tmpData2):
    retour = 0
    global processusNavigateurExart


    if retour == 0:
        try:
            select = Select(processusNavigateurExart.find_element_by_id(element))
            try:
                select.select_by_index(tmpData)
                return 1
            except:
                select.select_by_index(tmpData2)
                return -1   
        except:
            retour = 0
    if retour == 0:       
        try:
            select = Select(processusNavigateurExart.find_element_by_xpath(element))
            try:
                select.select_by_index(tmpData)
                return 1
            except:
                select.select_by_index(tmpData2)
                return -1
        except:
            retour = 0                       
    if retour == 0:        
        try:
            select = Select(processusNavigateurExart.find_element_by_name(element))
            try:
                select.select_by_index(tmpData)
                return 1
            except:
                select.select_by_index(tmpData2)
                return -1        
        except:
            retour = 0
    return retour

def ExistSelectElement(element, valeur, clef):
    tmpData = ''
    retour = 0
    global processusNavigateurExart
    
    tmpData = valeur.strip()
    
    if tmpData == "":
        return
    if retour == 0: 
        try:
            select = Select(processusNavigateurExart.find_element_by_id(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1
            except:
                return -1   
        except:
            retour = 0
    if retour == 0: 
        try:
            select = Select(processusNavigateurExart.find_element_by_xpath(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1
            except:
                return -1
        except:
            retour = 0    
    if retour == 0:        
        try:
            select = Select(processusNavigateurExart.find_element_by_name(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1
            except:
                return -1        
        except:
            retour = 0
            
    if retour == 0: 
        try:
            select = Select(processusNavigateurExart.find_element_by_class_name(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1
            except:
                return -1               
        except:
            retour = 0
    if retour == 0:
        try:
            select = Select(processusNavigateurExart.find_element_by_css_selector(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1
            except:
                return -1
        except:
            retour = 0
    if retour == 0:
                        
        try:
            select = Select(processusNavigateurExart.find_element_by_link_text(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1
            except:
                return -1
        except:
            retour = 0
    if retour == 0:
        try:
            select = Select(processusNavigateurExart.find_element_by_partial_link_text(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1                    
            except:
                return -1
        except:
            retour = 0
    if retour == 0:
        try:
            select = Select(processusNavigateurExart.find_element_by_tag_name(element))
            try:
                if clef in select.options:
                    select.select_by_visible_text(tmpData)
                    return 1
                    
            except:
                return -1
        except:
            return 0

def ExistPremierSelectElementXpath(element, valeur, clef):
    ElementNavigateur = []
    global processusNavigateurExart

    if check_exists_element(element) == 1:          
        try:
            ElementNavigateur = processusNavigateurExart.find_elements_by_xpath(element)
            #print(ElementNavigateur)
            i = 0
            while i < len(ElementNavigateur):
                try:
                    select = Select(ElementNavigateur[i])
                    try:
                        optionclef = ''
                        for option in select.options:
                            optionclef = optionclef + option.text
                            
                        if optionclef.find(clef) !=-1:
                            select.select_by_visible_text(valeur)
                            return 1
                            i = 50
                            optionclef = ''   
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor1')
                i = i + 1
        except:        
            journal.Log('erreor2') 
            return 0

def ExistPremierSelectElement(element, valeur, clef):
    retour = 0
    global processusNavigateurExart
    
    if retour == 0:
        try:
            processusNavigateurExart.find_elements_by_id(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_id(element)):
                journal.Log(len(processusNavigateurExart.find_elements_by_id(element)))
                try:
                    select = Select(processusNavigateurExart.find_elements_by_id(element)[i])
                    try:
                        if clef in select.options:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1
        except:  
            journal.Log("error1")
            
    if retour == 0:       
        try:
            processusNavigateurExart.find_elements_by_xpath(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_xpath(element)):
                
                journal.Log(len(processusNavigateurExart.find_elements_by_xpath(element)))
                try:
                    select = Select(processusNavigateurExart.find_elements_by_xpath(element)[i])
                    try:
                        optionclef = ''
                        for option in select.options:
                            optionclef = optionclef + option.text
                            
                        if optionclef.find(clef) !=-1:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50
                            optionclef = ''
                            
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1
        except:
            journal.Log('erreor2')
    if retour == 0:
        try:
            processusNavigateurExart.find_elements_by_name(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_name(element)):
                try:
                    select = Select(processusNavigateurExart.find_elements_by_name(element)[i])
                    try:
                        if clef in select.options:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50                         
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1
        except:
            journal.Log("bon3")
          
    if retour == 0:        
        try:
            processusNavigateurExart.find_elements_by_class_name(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_class_name(element)):
                try:
                    select = Select(processusNavigateurExart.find_elements_by_class_name(element)[i])
                    try:
                        if clef in select.options:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1
        except:
            journal.Log("bon3")
 
    if retour == 0:
        try:
            processusNavigateurExart.find_elements_by_css_selector(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_css_selector(element)):
                try:
                    select = Select(processusNavigateurExart.find_elements_by_css_selector(element)[i])
                    try:
                        if clef in select.options:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1                        
        except:
            journal.Log("bon")
            
    if retour == 0:      
        try:
            processusNavigateurExart.find_elements_by_link_text(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_link_text(element)):
                try:
                    select = Select(processusNavigateurExart.find_elements_by_link_text(element)[i])
                    try:
                        if clef in select.options:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1
        except:
            journal.Log()
    if retour == 0:
        try:
            processusNavigateurExart.find_elements_by_partial_link_text(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_partial_link_text(element)):
                try:
                    select = Select(processusNavigateurExart.find_elements_by_partial_link_text(element)[i])
                    try:
                        if clef in select.options:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1
        except:
            journal.Log('erreor')
            
    if retour == 0:       
        try:
            processusNavigateurExart.find_elements_by_tag_name(element)
            i = 0
            while i < len(processusNavigateurExart.find_elements_by_tag_name(element)):
                try:
                    select = Select(processusNavigateurExart.find_elements_by_tag_name(element)[i])
                    try:
                        if clef in select.options:
                            select.select_by_visible_text(valeur)
                            retour = 1
                            i = 50
                    except:
                        journal.Log('error')
                except:
                    journal.Log('erreor')
                i = i + 1
        except:
            retour = 0   
    return retour

def SelectPremierElement(element, valeur, clef):

    docseize = []
    global processusNavigateurExart
    
    if check_exists_element(element) == 1:
        docseize = processusNavigateurExart.find_elements_by_xpath(element)
        i = 0
        while i < len(docseize):
            try:
                select = Select(docseize[i])
                try:
                    optionclef = ''
                    for option in select.options:
                        optionclef = optionclef + option.text
                        
                    if optionclef.find(clef) !=-1 or clef.strip() == '':
                        try:
                            select.select_by_visible_text(valeur)
                            optionclef = ''
                            try:
                                processusNavigateurExart.find_elements_by_class_name('boutonbleuok')[i].click()
                                return 1
                            except:
                                return 0
                        except:
                            if i == (len(docseize) - 1):
                                return 2
                except:
                    return 0
            except:
                return 0
            i = i + 1
        else:
            return 2

def VerificationDernierePage(TitrePage):
    retour = 0
    all_windows = processusNavigateurExart.window_handles 

    processusNavigateurExart.switch_to.window(all_windows[(len(processusNavigateurExart.window_handles)-1)])
    time.sleep(1)
    if TitrePage in str(processusNavigateurExart.title).lower():
        journal.Log('Titre page : ' + processusNavigateurExart.title)
        retour = 1
    else:
        journal.Log('Titre page : ' + processusNavigateurExart.title)
        retour = 0
    return retour

def VerificationTitrePage(TitrePage):

    time.sleep(1)
    if TitrePage in str(processusNavigateurExart.title).lower():
        journal.Log('Titre page : ' + processusNavigateurExart.title)
        return 1
    else:
        journal.Log('Titre page : ' + processusNavigateurExart.title)
        return 0
    
def NombrePageHandle():
    all_windows = processusNavigateurExart.window_handles
    return len(all_windows) 
    
def RetourPagePrinciaple():
    retour = 0
    all_windows = processusNavigateurExart.window_handles 
    try:
        processusNavigateurExart.switch_to.window(all_windows[(len(processusNavigateurExart.window_handles)-2)])
        return 1
    except:
        journal.Log(journal.Log(all_windows))
        return retour

def PagePrinciapleRetour():
    retour = 0
    all_windows = processusNavigateurExart.window_handles 
    try:
        processusNavigateurExart.switch_to.window(all_windows[(len(processusNavigateurExart.window_handles)-1)])
        return 1
    except:
        journal.Log(journal.Log(all_windows))
        return retour

def CaptureTable(elmentID):
    retour = []
    table = ''
    succes = 0
    if check_exists_element(elmentID) == 1:
        try:
            table = processusNavigateurExart.find_element_by_id(elmentID)
    
            try:
                rows = table.find_elements_by_tag_name("tr")
                for row in rows:
                    temp = []
                    cols = row.find_elements_by_tag_name("td")
                    for col in cols:
                        #journal.Log(col.text)
                        temp.append(col.text)
                        #journal.Log(col.text)
                    retour.append(temp)
                    succes = 1
                #journal.Log(retour)
            except:
                retour = ''
                journal.Log('Erreur de caputre des données')
            return retour
        except:
            pass
    
        if succes == 0:
            try:
                table = processusNavigateurExart.find_element_by_xpath(elmentID)
        
                try:
                    rows = table.find_elements_by_tag_name("tr")
                    for row in rows:
                        temp = []
                        cols = row.find_elements_by_tag_name("td")
                        for col in cols:
                            #journal.Log(col.text)
                            temp.append(col.text)
                            #journal.Log(col.text)
                        retour.append(temp)
                        succes = 1
                    #journal.Log(retour)
                except:
                    retour = ''
                    journal.Log('table capture')
                return retour
            except:
                succes = 0
                journal.Log('sucess')
        if succes == 0:
            try:
                table = processusNavigateurExart.find_element_by_css_selector(elmentID)
        
                try:
                    rows = table.find_elements_by_tag_name("tr")
                    for row in rows:
                        temp = []
                        cols = row.find_elements_by_tag_name("td")
                        for col in cols:
                            #journal.Log(col.text)
                            temp.append(col.text)
                            #journal.Log(col.text)
                        retour.append(temp)
                        succes = 1
                    #journal.Log(retour)
                except:
                    retour = ''
                    journal.Log('table capture')
                return retour
            except:
                succes = 0
                journal.Log('sucess')

def getvalue(element):
    if check_exists_elementsXpath(element) == 1:
        try:
            return str(processusNavigateurExart.find_element_by_xpath(element).get_attribute('value')).strip()
        except:
            return ""                
    else:
        return ""
        