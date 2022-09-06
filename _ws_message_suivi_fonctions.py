# -*- coding: utf-8 -*-
"""
Created on Thu Apr  2 09:48:30 2020

@author: service.wintask
"""
import Commun.Libs._variable_globale as var
import Commun.Libs._log_fonctions as journal
import Commun.Libs._powershell as powershell
import Commun.Libs._config as config

def EnvoiMessageSuivi(tmpNumDossier, detailMessage):
    tmpResulat = ''
    cheminInput = ''
    tmpScript = ''
    cheminScript = ''
    
    tmpResulat = var.templateMessageSuivi.replace("{NUM_DOSSIER_REPLACE}", tmpNumDossier)
    tmpResulat = tmpResulat.replace("{DETAIL_REPLACE}", detailMessage)
    
    cheminInput = var.repert + "/WebService/MessageSuivi/message_suivi-input.xml"
        
    powershell.Write(cheminInput, tmpResulat)
    
    #REM Completion et generation du Script
    if (var.environnement == "PROD"):
        tmpScript = var.templateScriptMessageSuivi.replace("{ADRESSE_SERVEUR_REPLACE}", config.GetConfigValue("WEBSERVICE", "PROD", "http://esb-prod.groupe-prunay.fr:81/gexsi_metier/Gexsi_Metier_Retour.svc?singleWsdl"))
    else:
        tmpScript = var.templateScriptMessageSuivi.replace("{ADRESSE_SERVEUR_REPLACE}", config.GetConfigValue("WEBSERVICE", "RCT", "http://gexsi-rec.groupe-prunay.fr:81/Gexsi_Metier/Gexsi_Metier_Retour.svc?singleWsdl"))
        
    cheminScript = var.repert + "/WebService/MessageSuivi/message_suivi.ps1"
        
    powershell.Write(cheminScript, tmpScript)
    
    try:
        powershell.Shell(var.repert + "/WebService/MessageSuivi/run_message_suivi.bat")
        journal.Log("Processus d'envoi de message de suivi execute avec succes pour le dossier '" + tmpNumDossier)
        journal.Log("Script lance avec succes.") 
    except:
        journal.Log("Processus d'envoi de message de suivi non execute pour le dossier '" + tmpNumDossier)
        journal.Log("Le WebService d'envoi de message de suivi a echoue.")

	
