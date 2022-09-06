# -*- coding: utf-8 -*-
import datetime
import unicodedata
import os.path
import openpyxl
import Commun.Libs._variable_globale as var
import Commun.Libs._log_fonctions as journal
#import _variable_globale as a

def RemoveAccentsSuivi(text):
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore')     
    return text
    
def WriteSuivi(FilePath, ContenuAEcrire):
    
    f = open(FilePath, "a+")
    CodeRetour = 0
    try:
        #f.EcrireLigne(ContenuAEcrire)
        f.writelines(ContenuAEcrire + '\n')
        CodeRetour = 1
    except:
        CodeRetour = 0
    finally:
        f.close()
    return CodeRetour

def GenereHeaderSuivi():
	ligneSuivi = ''
    #global fichierSuivi
	if not os.path.isfile(var.fichierSuivi):	
		ligneSuivi = "DATE ET HEURE;DATE RAPPORT;ENVIRONNEMENT;TYPE DE RAPPORT;TYPE D'EXPERTISE;NUMERO DE DOSSIER;NUMERO DE SINISTRE;AGENCE;COMPAGNIE;ETAT DOSSIER;ECHANGE INTERNE;COMMENTAIRE;URL"
		WriteSuivi(var.fichierSuivi, ligneSuivi)

def AjoutSuivi(contenu, typeTraitement):
    #global fichierSuivi
    ligneSuivi = ''
    flatContent = ''
    
    GenereHeaderSuivi()
    
    flatContent = RemoveAccentsSuivi(contenu)
    #typeTraitement = RemoveAccentsSuivi(typeTraitement)
    #ligneSuivi = datetime.datetime.now()[0:19] + ";" + "Date" + ";" + "RCT" + ";" + typeTraitement + ";" + "ESP" + ";" + "00273372" + ";" + "sinistre" + ";" + "txt_agence" + ";" + "txt_compagnie" + ";" + "etat" + ";" + "echange_interne" + ";" + flatContent + ";" + "url"
    ligneSuivi = str(datetime.datetime.now())[0:19] + ";" + str(var.date_rapport) + ";" + str(var.environnement) + ";" + str(typeTraitement) + ";" + str(var.type_expertise) + ";" + str(var.txt_num_dossier) + ";" + str(var.sinistre) + ";" + str(var.txt_agence) + ";" + str(var.txt_compagnie) + ";" + str(var.etat) + ";" + str(var.echange_interne) + ";" + str(flatContent) + ";" + str(var.url)
    WriteSuivi(var.fichierSuivi, ligneSuivi)
    
def AjoutSuiviSimple(contenu, typeTraitement):
    #global fichierSuivi
    ligneSuivi = ''
    flatContent = ''
    
    GenereHeaderSuivi()
    
    flatContent = RemoveAccentsSuivi(contenu)
    ligneSuivi = str(datetime.datetime.now())[0:19] + ";" + str(var.date_rapport) + ";" + str(var.environnement) + ";" + str(typeTraitement) + ";" + str(flatContent) + ";" 
    
    WriteSuivi(var.fichierSuivi, ligneSuivi)

def GenereExcel(filepath):
    retour = ""
    try:
        wb = openpyxl.Workbook() 
        wb.save(filepath)
        retour = "true"
    except:
        retour = ""
    finally:
        wb.close()
    return retour
    journal.Log('Fichier Excel generé avec succès')
