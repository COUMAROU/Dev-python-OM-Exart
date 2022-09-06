# -*- coding: utf-8 -*-
"""
Created on Thu Apr  2 11:09:33 2020

@author: service.wintask
"""
#import os.path
import Commun.Libs._variable_globale as var
import Commun.Libs._log_fonctions as journal
import Commun.Libs._powershell as powershell
import Commun.Libs._config as config
import Commun.Libs._text_fonctions as textfonction

def EnvoiEchangeInterne(numDossier, codeAssistante, message):
    tmpResulat = ''
    cheminInput = ''
    tmpScript = ''
    cheminScript = ''

    tmpResulat = var.templateEchangeInterne.replace("{NUM_DOSSIER_REPLACE}", numDossier)
    tmpResulat = tmpResulat.replace("{CODE_ASSISTANTE_REPLACE}", textfonction.CompleteTexteAvecCharactere(codeAssistante, "0", 5))
    tmpResulat = tmpResulat.replace("{LIBELLE_REPLACE}", message)
    
    cheminInput = var.repert + "/WebService/EchangeInterne/echange_interne-input.xml"
        
    powershell.Write(cheminInput, tmpResulat)
    
    #REM Completion et generation du Script
    if (var.environnement == "PROD"):
        tmpScript = var.templateScriptEchangeInterne.replace("{ADRESSE_SERVEUR_REPLACE}", config.GetConfigValue("WEBSERVICE", "PROD", "http://esb-prod.groupe-prunay.fr:81/gexsi_metier/Gexsi_Metier_Retour.svc?singleWsdl"))
    else:
        tmpScript = var.templateScriptEchangeInterne.replace("{ADRESSE_SERVEUR_REPLACE}", config.GetConfigValue("WEBSERVICE", "RCT", "http://gexsi-rec.groupe-prunay.fr:81/Gexsi_Metier/Gexsi_Metier_Retour.svc?singleWsdl"))
        
    cheminScript = var.repert + "/WebService/EchangeInterne/echange_interne.ps1"
        
    powershell.Write(cheminScript, tmpScript)

    
    try:
        powershell.Shell(var.repert + "/WebService/EchangeInterne/run_echange_interne.bat")
        journal.Log("Processus d'echange interne execute avec succes pour le dossier '" + numDossier)
        journal.Log("Script lance avec succes.") 
    except:
        journal.Log("Processus d'echange interne non execute pour le dossier '" + numDossier)
        journal.Log("Le WebService d'envoi de message de suivi a echoue.")
