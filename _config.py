import pandas as pd
import numpy as np
import Commun.Libs._log_fonctions as journal

def GetConfigValue (Sheet, ValeurRechercher, ValeurSiNonTrouver) : 
    
    #data=pd.read_excel("_config.xlsx", Sheet, keep_default_na=True)
    data=pd.read_excel("_config.xlsx", Sheet, keep_default_na=True, name = ['A','B'])
    ReturnValue = ''
    colonnes = []
    colonnes = list(data)
    #df.loc[:,['A', 'C']]
    if len(np.where(data[colonnes[0]]== ValeurRechercher)[0]) > 0:
        nb=np.where(data[colonnes[0]]== ValeurRechercher)[0]
        ReturnValue = data[colonnes[1]][nb[0]]
        if ReturnValue is np.nan:
            ReturnValue = ValeurSiNonTrouver
        journal.Log("Valeur trouvee pour la cle '" + ValeurRechercher + "' dans la feuille '" + Sheet + "' du fichier de configuration. Resultat: '" + str(ReturnValue) + "'")
    else:    
        ReturnValue = ValeurSiNonTrouver
        journal.LogErreur(0, "Impossible de trouver la valeur correspondant a la cle '" + ValeurRechercher + "' dans la feuille '" + Sheet + "' du fichier de configuration.")       
    return ReturnValue       

#print (GetConfigValue('REPARATION', 'AR', 'NON'))

