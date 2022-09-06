# -*- coding: utf-8 -*-
import datetime
import Commun.Libs._variable_globale as var
"""
Created on Tue Feb 25 21:00:08 2020

@author: service.wintask
"""
def EcrireLigne(FilePath, ContenuAEcrire):
    
    f = open(FilePath, "a+", encoding = "utf-8")
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


def LogErreur(typeErreur, tmpTxtErreur):
	buffer = ''
	
	buffer = str(datetime.datetime.now())[0:10] + "|" + str(datetime.datetime.now())[11:19]
	
	if typeErreur == 1:
		buffer = buffer + "|ERREUR CRITIQUE|"
	else:
		buffer = buffer + "|ERREUR MINEURE|"
	
	buffer = buffer + tmpTxtErreur
	EcrireLigne(var.fichierLog, buffer)

def LogWarning(tmpTxtErreur):
	buffer = ''
	
	buffer = str(datetime.datetime.now())[0:10] + "|" + str(datetime.datetime.now())[11:19] + "|WARNING|" + tmpTxtErreur
	EcrireLigne(var.fichierLog, buffer)
    
def LogWintaskError(tmpTxtErreur):
	buffer = ''
	
	buffer = str(datetime.datetime.now())[0:10] + "|" + str(datetime.datetime.now())[11:19] + "|ERREUR WINTASK|" + tmpTxtErreur
	EcrireLigne(var.fichierLog, buffer)


def Log(tmpTxtErreur):
	buffer = ''
	
	buffer = str(datetime.datetime.now())[0:10] + "|" + str(datetime.datetime.now())[11:19] + "|INFO|" + tmpTxtErreur
	EcrireLigne(var.fichierLog, buffer)
