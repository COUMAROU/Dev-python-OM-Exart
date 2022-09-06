# -*- coding: utf-8 -*-
"""
Created on Thu Apr  2 10:52:39 2020

@author: service.wintask
"""
import Commun.Libs._variable_globale as var
import Commun.Libs._log_fonctions as journal
import Commun.Libs._powershell as powershell
import Commun.Libs._config as config

def SoldeTravailAdministratif(numDossierComplet, compteurRobot, tmpPayloadSoldeTravailAdmin, tmpScriptSoldeTravailAdmin):
    tmpResulat = ''
    cheminInput = ''
    tmpScript = ''
    cheminScript = ''
    
    tmpResulat = tmpPayloadSoldeTravailAdmin.replace("{NUM_DOSSIER_REPLACE}", numDossierComplet)
    tmpResulat = tmpResulat.replace("{COMPTEUR_SUIVI_REPLACE}", compteurRobot)
    
    cheminInput = var.repert + "/WebService/SolderTravailAdministratif/solder_travail_administratif-input.xml"
        
    powershell.Write(cheminInput, tmpResulat)
    
    #REM Completion et generation du Script
    if (var.environnement == "PROD"):
        tmpScript = tmpScriptSoldeTravailAdmin.replace("{ADRESSE_SERVEUR_REPLACE}", config.GetConfigValue("WEBSERVICE", "PROD", "http://esb-prod.groupe-prunay.fr:81/gexsi_metier/Gexsi_Metier_Retour.svc?singleWsdl"))
    else:
        tmpScript = tmpScriptSoldeTravailAdmin.replace("{ADRESSE_SERVEUR_REPLACE}", config.GetConfigValue("WEBSERVICE", "RCT", "http://gexsi-rec.groupe-prunay.fr:81/Gexsi_Metier/Gexsi_Metier_Retour.svc?singleWsdl"))
        
    cheminScript = var.repert + "/WebService/SolderTravailAdministratif/solder_travail_administratif.ps1"
        
    powershell.Write(cheminScript, tmpScript)
    
    try:
        powershell.Shell(var.repert + "/WebService/SolderTravailAdministratif/run_solder_travail_administratif.bat")
        journal.Log("Processus de solde de travail administratif execute avec succes pour le dossier '" + numDossierComplet)
        journal.Log("Script lance avec succes.") 
        return "succes"
    except:
        journal.Log("Processus de solde de travail administratif non execute pour le dossier '" + numDossierComplet)
        journal.Log("Le WebService d'envoi de message de suivi a echoue.")

