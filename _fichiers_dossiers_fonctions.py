import os
import Commun.Libs._log_fonctions as journal

def GenereDossierSiNecessaire(chemin) :
    
    if not os.path.exists(chemin):
        os.makedirs(chemin, exist_ok=True)
    return chemin

def RecupereDossierActuel():
    dossier = ''
    dossier = os.path.abspath(os.path.split(__file__)[0])

    if (dossier[-1]).find("/" ) == -1 :
        dossier = dossier + "/"
    return dossier

def GenereExcel(nomComplet):
    if os.path.isfile(nomComplet):
        os.remove(nomComplet)    
    wb = openpyxl.Workbook() 
    try:
        wb.save(nomComplet)
    except:
        journal.Log('Erreur de generation du fichier excel')
    finally:
        wb.close()

