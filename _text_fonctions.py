# -*- coding: utf-8 -*-
def CompleteTexteAvecCharactere(texte, charactere, longueurDesiree):
    longueurActuelle = '' 
    tmpResultat = ''
    longueurActuelle = len(texte)
    tmpResultat = texte
    
    if (longueurActuelle >= longueurDesiree):
        tmpResultat = texte
    else:
        while (longueurActuelle < longueurDesiree):
            tmpResultat = charactere + tmpResultat
            longueurActuelle = len(tmpResultat)
            
    return tmpResultat

def majuscsa(ch):
    #Convertit la chaine ch en majuscules non accentuées 
    
    alpha1 = u"aàÀâÂäÄåÅæbcçÇdeéÉèÈêÊëËfghiîÎïÏjklmnoôÔöÖœpqrstuùÙûÛüÜvwxyÿŸz"
    alpha2 = u"AAAAAAAAAÆBCCCDEEEEEEEEEFGHIIIIIJKLMNOOOOOŒPQRSTUUUUUUUVWXYYYZ"
    x = ""
    for c in ch:
        k = alpha1.find(c) # k = indice de "c" dans alpha1
        if k >= 0:
            # "c" est dans alpha1: on remplace par le car. correspondant de alpha2
            x += alpha2[k]
        else:
            # "c" n'est pas dans alpha1: on le laisse passer
            x += c
    return x

def dezip(filezip, pathdst = ''):
    if pathdst == '': pathdst = os.getcwd()  ## on dezippe dans le repertoire locale
    zfile = zipfile.ZipFile(filezip, 'r')
    for i in zfile.namelist():  ## On parcourt l'ensemble des fichiers de l'archive
        #print (i)
        if majuscsa(str(i)).find('.PDF') != -1:
            if majuscsa(str(i)).find('RAPP') != -1:                    
                logger.info('La pièce jointe est rapport ou pas un pdf' + str(i))
                #os.remove(i)
            else:
                if os.path.isdir(i):   ## S'il s'agit d'un repertoire, on se contente de creer le dossier
                    try: os.makedirs(pathdst + os.sep + i) 
                    except: pass
                else:
                    try: os.makedirs(pathdst + os.sep + os.path.dirname(i))
                        #EcrireInfoRecupere(pathdst + os.sep + os.path.dirname(i))
                    except: pass
                    data = zfile.read(i)                   ## lecture du fichier compresse
                    fp = open(pathdst + os.sep + i, "wb")  ## creation en local du nouveau fichier
                    fp.write(data)                         ## ajout des donnees du fichier compresse dans le fichier local
                    fp.close()     
    zfile.close()
    os.remove(filezip)
#print(CompleteTexteAvecCharactere('20008', '0', 20))