# -*- coding: utf-8 -*-
"""
Created on Wed Mar 18 15:20:01 2020

@author: service.wintask

data = []
tmpLignesData = []
tab_ligne = []
tmpTabCell = []
configColumn = []
configLine = []
tab_piecesJointes = [] # La liste des pieces jointes
tab_tmp_pieceJointe = [] # Contiendra les valeurs unitaires de chaque pieces jointe, sera reecrit pour chaque pieces jointe
"""

import os
import sys
import datetime
import time
from time import sleep
#import autoit as winpj
import Commun.Libs._variable_globale as var
import Commun.Libs._log_fonctions as journal
#import Commun.Libs._maths as maths
import Commun.Libs._config as config
import Commun.Libs._text_fonctions as textfonction
import Commun.Libs._fichiers_dossiers_fonctions as dossier

import Commun.Libs._navigateur_fonctions as navigateur
import Commun.Libs._navigateur_fonctions2 as navigateurExart

var.repert = ''
var.dossierSuiviPrincipal = ''
var.DossierSuivi = ''
var.fichierSuivi
var.fichierLog = ''
var.fichierConfig = ''


var.environnement = "PROD" # 'PROD'/'RCT'
var.navigatur = "CH"


date = ''
annee = ''
mois = ''
date = str(datetime.datetime.now())[0:10]
annee = str(datetime.datetime.now())[0:4]
mois = str(datetime.datetime.now())[5:7]

start_time = time.time()	

var.repert = os.path.abspath(os.path.split(__file__)[0])
var.launchTimeStamp = date + '-' + str(datetime.datetime.now())[11:19].replace(":", "") 
#print(str(datetime.datetime.now()))
#print(var.launchTimeStamp)
var.dossierSuiviPrincipal = dossier.GenereDossierSiNecessaire('Suivis') + '\\'
#DossierSuivi = dossier.GenereDossierSiNecessaire('Suivis' + '\\' + annee)
#DossierSuivi = dossier.GenereDossierSiNecessaire('Suivis' + '\\' + annee + '\\' + mois)
filePath = dossier.GenereDossierSiNecessaire('OM_EXART') + '\\'
filePath = var.repert + "\\" + dossier.GenereDossierSiNecessaire(filePath + date + "\\")

var.dossierSuivi = dossier.GenereDossierSiNecessaire(var.dossierSuiviPrincipal + date + "\\")
var.fichierSuivi = var.dossierSuivi + "Suivi - " + date + ".csv"
#print(var.fichierSuivi)

var.dossierSuivi = dossier.GenereDossierSiNecessaire(var.dossierSuivi + var.launchTimeStamp + "-" + var.environnement + "\\")
var.fichierLog = var.dossierSuivi + "Journal.log"
var.fichierConfig = "_config.xlsx"

import Commun.Libs.EnvoiMail as function_mail
from zipfile import ZipFile
import os
from os.path import basename
# Zip the files from given directory that matches the filter
def zipFilesInDir(dirName, zipFileName):
   # create a ZipFile object
   with ZipFile(zipFileName, 'w') as zipObj:
       # Iterate over all the files in directory
       for folderName, subfolders, filenames in os.walk(dirName):
           for filename in filenames:
               if filename.upper().find(".PDF") != -1:
                   # create complete filepath of file in directory
                   fileP = os.path.join(folderName, filename)
                   # Add file to zip
                   zipObj.write(fileP, basename(fileP))
       zipObj.close()

journal.Log("---------------------------------------------Début d'xécution du programme------------------------------------------------------------------------")

var.url = config.GetConfigValue("LETTRAGE-RECUP", "URL-DARVA", "")
var.url = var.url.replace("{0}", config.GetConfigValue("LETTRAGE-RECUP", "ID-ABONNE", ""))
var.url = var.url.replace("{1}", config.GetConfigValue("LETTRAGE-RECUP", "PWD-ABONNE", ""))	

def TrouverReferenceInterne(tab, sinsitre) :
    
    for ref in tab:

        temp_ref_interne = ""
        temp_ref_interne = " ".join(str(ref[0]).strip().split())
        
        tab_ref_interne = []
        tab_ref_interne = temp_ref_interne.split(" ")
        internal = ""
        external = ""
        internal = str(tab_ref_interne[2]).strip()
        external = str(tab_ref_interne[3]).strip()
        
        if textfonction.CompleteTexteAvecCharactere(external, "0", 25) == textfonction.CompleteTexteAvecCharactere(sinsitre, "0", 25):
            return internal
        else:
            continue
    return ""
boucle = 0
while boucle <= 5:                                
    ret = 0
    ret = navigateur.OuvertureNavigateur(var.navigatur, var.url.strip())
    
    if (ret != 0):
        journal.LogErreur(0, "Impossible d'ouvrir le navigateur a l'adresse '" + var.url + "'")
        sys.exit()
    else:
        navigateur.ClicSurElement("onglets__IrdWeb")
        
        #frame = navigateur.processusNavigateur.find_element_by_id("docsInWaitFilterActionForm")
        navigateur.processusNavigateur.switch_to.frame(1) 
        
        time.sleep(1)
        if navigateur.SelectElement("//select[@name='docType']", "Ordre de mission", "Ordre de mission") == 1:
            time.sleep(2)
            navigateur.ClicSurElement("//*[@href='javascript:if(checkDocsInWaitFilterForm()==true){envoiFormulaire(); }']")
            sleep(3)
            tab_om = navigateur.CaptureTable("/html/body/table[2]/tbody/tr/td/table[2]/tbody")
            
            
            ligne_om = 1
            #try:
            tab_pdfs = navigateur.processusNavigateur.find_elements_by_xpath("//*[@src='./images/IconePDF.gif']")
                #try:
            internal_imput = navigateur.processusNavigateur.find_elements_by_name("internal")
                    #try:
            chkRAD_coche = navigateur.processusNavigateur.find_elements_by_name("chkRAD")

            len(tab_om)
            macif = ""
            covea = ""
            aviva = ""
            autres = ""
            objet_mail_macif = ""
            objet_mail_covea = ""
            objet_mail_aviva = ""
            objet_mail_autres = ""
            
            referenceSinistre = []
            for tab_pdf in tab_pdfs:
                 
                filePath1 = ""
                filePath2 = ""
                if str(tab_om[ligne_om][6]).strip().upper().find("MACIF") != -1:
                    objet_mail_macif = "MACIF_" + var.launchTimeStamp
                    filePath1 = dossier.GenereDossierSiNecessaire(filePath + objet_mail_macif)
                    macif = filePath1
                elif str(tab_om[ligne_om][6]).strip().upper().find("GMF") != -1:
                    objet_mail_covea = "COVEA_" + var.launchTimeStamp
                    filePath1 = dossier.GenereDossierSiNecessaire(filePath + objet_mail_covea)
                    covea = filePath1
                elif str(tab_om[ligne_om][6]).strip().upper().find("MAAF") != -1:
                    objet_mail_covea = "COVEA_" + var.launchTimeStamp
                    filePath1 = dossier.GenereDossierSiNecessaire(filePath + objet_mail_covea)
                    covea = filePath1
                elif str(tab_om[ligne_om][6]).strip().upper().find("MMA") != -1:
                    objet_mail_covea = "COVEA_" + var.launchTimeStamp
                    filePath1 = dossier.GenereDossierSiNecessaire(filePath + objet_mail_covea)
                    covea = filePath1
                elif str(tab_om[ligne_om][6]).strip().upper().find("BPCE") != -1:
                    objet_mail_covea = "COVEA_" + var.launchTimeStamp
                    filePath1 = dossier.GenereDossierSiNecessaire(filePath + objet_mail_covea)
                    covea = filePath1               
                elif str(tab_om[ligne_om][6]).strip().upper().find("ABEILLE") != -1:
                    objet_mail_aviva = "AVIVA_" + var.launchTimeStamp
                    filePath1 = dossier.GenereDossierSiNecessaire(filePath + objet_mail_aviva)
                    aviva = filePath1
                else:
                    objet_mail_autres = "AUTRES" + var.launchTimeStamp
                    filePath1 = dossier.GenereDossierSiNecessaire(filePath + objet_mail_autres)
                    autres = filePath1
                referenceSinistre.append(str(tab_om[ligne_om][1]).strip())    
                filePath2 = filePath1 + "\\" + str(tab_om[ligne_om][1]).strip() + ".pdf"
                
                

                if not os.path.isfile(filePath2) :
                    """
                    print("B")
                    os.path.abspath(filePath2)
                    print("C")
                    
                    tab_pdf.send_keys(filePath2)
                    print(ligne_om)
                else:
                    print("KO")

                    #tab_pdf.send_keys(filePath2)
                    import urllib.request
                    print(tab_pdf.url)
                    urllib.request.urlretrieve(tab_pdf.url , filePath2)
                    sys.exit()
                    """
                    tab_pdf.click()
                    sleep(2)
                    if winpj.win_wait_active(str(tab_om[ligne_om][1])) == 1:
                        sleep(1)            
                        winpj.send("^s")
                        sleep(1)
                        if winpj.win_wait_active("Save As") == 1:
                            sleep(1)
                            winpj.send(filePath2)
                            sleep(1)
                            winpj.send("{ENTER}")
                            sleep(1)
                        winpj.win_close(str(tab_om[ligne_om][1]))

                ligne_om = ligne_om + 1
                sleep(2)
            var.launchTimeStamp = ""    
            traitementexart = 0       
            if objet_mail_macif != "":
                
                var.launchTimeStamp = date + '-' + str(datetime.datetime.now())[11:19].replace(":", "")
                objet_mail_macif = "MACIF_" + var.launchTimeStamp
                zipFilesInDir(macif, macif + ".zip")
                function_mail.envoie_mail_simple(pj = macif + ".zip" , sujet = objet_mail_macif, destinataire=config.GetConfigValue("LETTRAGE-RECUP", "DEST-EMAIL-MACIF", ""))
                traitementexart = 1
            if objet_mail_covea != "":
                var.launchTimeStamp = date + '-' + str(datetime.datetime.now())[11:19].replace(":", "")
                objet_mail_covea = "COVEA_" + var.launchTimeStamp
                zipFilesInDir(covea, covea + ".zip")
                function_mail.envoie_mail_simple(pj = covea + ".zip" , sujet = objet_mail_covea, destinataire=config.GetConfigValue("LETTRAGE-RECUP", "DEST-EMAIL-COVEA", ""))
                traitementexart = 1
                
            if objet_mail_aviva != "":
                var.launchTimeStamp = date + '-' + str(datetime.datetime.now())[11:19].replace(":", "")
                objet_mail_aviva = "AVIVA_" + var.launchTimeStamp
                zipFilesInDir(aviva, aviva + ".zip")
                function_mail.envoie_mail_simple(pj = aviva + ".zip", sujet = objet_mail_aviva, destinataire=config.GetConfigValue("LETTRAGE-RECUP", "DEST-EMAIL-AVIVA", ""))
                traitementexart = 1
                
            if objet_mail_autres != "":
                zipFilesInDir(autres, autres + ".zip")   
                
            if traitementexart == 1: 
                
                nav = 0        
                nav = navigateurExart.OuvertureNavigateur("CH", config.GetConfigValue("LETTRAGE-RECUP", "URL_EXART", "http://10.240.40.4."))
                
                if nav == 0:
                    sleep(2)
                    navigateurExart.RemplissageChampTexte("UserName", "renfort")
                    sleep(1)
                    navigateurExart.RemplissageChampTexte("Password", "renfort")
                    sleep(3)
                    if navigateurExart.ClicSurElement("//*[@id='login']/form/table/tbody/tr[3]/td/input") == 1:
                        
                        if navigateurExart.VerificationTitrePage("dossiers") == 1:
                            
                            sleep(1)
                            
                            if navigateurExart.ClicSurElement("//*[@href='/AcquisitionArea/Acquisition']")   == 1 :
                                sleep(50)
                                if objet_mail_macif != "":
                                    if navigateurExart.ClicSurElement("//*[starts-with(@href, '/AcquisitionArea/Acquisition') and contains(text(),'MACIF SOP')]") == 1: 
                                        sleep(3)
                                        navigateurExart.ClicSurElement("//*[contains(text(),'" + objet_mail_macif + "')]")  
                                        sleep(2)
                                        navigateurExart.alert()
                                        sleep(1)
                                if objet_mail_aviva != "":
                                    if navigateurExart.ClicSurElement("//*[starts-with(@href, '/AcquisitionArea/Acquisition') and contains(text(),'AVIVA')]") == 1: 
                                        sleep(3)
                                        navigateurExart.ClicSurElement("//*[contains(text(),'" + objet_mail_aviva + "')]")  
                                        sleep(2)
                                        navigateurExart.alert()
                                        sleep(1)
                                        
                                   
                                if navigateurExart.ClicSurElement("//*[@href='/ReportArea/Report/GetReports?status=1']") == 1:
                                    
                                    sleep(2)
                                    countnombrepage = 1
                                    while navigateurExart.check_exists_element('//*[@id="center"]/div[1]/div[2]/div[2]/span[' + str(countnombrepage) + ']') != 0:
                                        countnombrepage = countnombrepage + 1
                                        
                                    countnombrepage = countnombrepage - 1
                                
                                    if countnombrepage > 10:
                                        countnombrepage = 10
                                        
                                    temptab = []
                                    while countnombrepage != 0:
                                        sleep(1)
                                        if navigateurExart.ClicSurElement('//*[@id="center"]/div[1]/div[2]/div[2]/span[' + str(countnombrepage) + ']') == 1:
                                            sleep(2)
                                            table = navigateurExart.CaptureTable("//*[@id='center']/div[1]/div[2]/table")
                                            temptab = temptab + table
                                        countnombrepage = countnombrepage - 1

                                    time.sleep(1)

                    if navigateurExart.FermetureNavigateur() == 1:
                        Avalider = False
                        ligneOMDarva = 0
                        for sinsitre in referenceSinistre:
                            ref_interne_exart = ""
                            ref_interne_exart = TrouverReferenceInterne(temptab, sinsitre)
                            
                            if ref_interne_exart != "":
                                internal_imput[ligneOMDarva].clear()
                                sleep(1)
                                internal_imput[ligneOMDarva].send_keys(ref_interne_exart)
                                sleep(1)
                                chkRAD_coche[ligneOMDarva].click()
                                Avalider = True
                                
                            ligneOMDarva = ligneOMDarva + 1
                            
                        sleep(2)
                        if Avalider:
                            navigateur.ClicSurElement("//*[@title = 'Validation de la sélection']")
                        sleep(2)
            """                           
                    except:
                        journal.Log("Pas de resultat disponible 1")
                except:
                    journal.Log("Pas de resultat disponible 2")
            except:
                journal.Log("Pas de resultat disponible 3")
            """
                
        navigateur.FermetureNavigateur()
        boucle = boucle + 1
    
sys.exit()

    