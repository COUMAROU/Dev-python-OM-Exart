# -*- coding: utf-8 -*-

import unicodedata
import os, os.path
import subprocess
import Commun.Libs._variable_globale as var
import Commun.Libs._log_fonctions as journal

def ShellSys(FilePath):
    try:
        os.systeme(FilePath)
    except:        
        journal.Log('erreur ' + subprocess.CalledProcessError.returncode)
        
def Shell(FilePath):
    try:
        subprocess.check_call(FilePath)
    except:        
        journal.Log('erreur ' + subprocess.CalledProcessError.returncode)
    
def Write(FilePath, ContenuAEcrire):
    if os.path.isfile(FilePath):
        os.remove(FilePath)
    
    f = open(FilePath, "w+", encoding = "utf-8")
    CodeRetour = 0
    try:
        f.write(ContenuAEcrire)

        CodeRetour = 1
    except:
        CodeRetour = 0
    finally:
        f.close()
    return CodeRetour

def Read(FilePath):
    
    ContenuFichier = ''
    if os.path.isfile(FilePath):
        f = open(FilePath, "r", encoding = "utf-8")
        try:
            ContenuFichier = str(f.read())
        except:
            ContenuFichier = ''
        finally:
            f.close()
    return ContenuFichier

def RemoveAccents(text):
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore')     
    return text

def WriteLineMonthlyReport(text):
    cheminTmpScript = ''
    tmpScriptContent = ''
    
    cheminTmpScript = var.repert + "Commun/PWS/rapport_mensuel-tmp.ps1"
    
    tmpScriptContent = var.templateRapportMensuel.replace("{DATA_TO_WRITE}", text)
    tmpScriptContent = tmpScriptContent.replace("{FILE_PATH}", var.dossierSuiviPrincipal)
    journal.Log(tmpScriptContent)

    Write(cheminTmpScript, tmpScriptContent)
    
    try:
        Shell(var.repert + "Commun\PWS\rapport_mensuel-run.bat")
        journal.Log("Script lance avec succes.") 
    except:
        journal.Log("Echec lors du lancement du script.")
