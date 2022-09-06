import pyodbc 
import Commun.Libs._variable_globale as var
import Commun.Libs._log_fonctions as journal


#wb = Workbook()
#test = 'D:/REFONTE ROBOT DARVA/Suivis/2020-03-19/2020-03-19-184805-RCT-NH/Data.xlsx'
        
def OuvertureConnexionSQL():
    global cnxn, cursor, row
    db_serveur = ''
    db_nomBase = ''
    db_user = ''
    db_password = ''
    connexion = ''
    tmpResult = ''
    cnxn= ''
    cursor = ''
    row = ''    
    
    db_nomBase = 'GEXSI'
    #environnement= 'RCT'
    
    if (var.environnement == 'PROD'):
        db_serveur = "gexsidb.groupe-prunay.fr"
        db_user = "metro_gexsi"
        db_password = "PiU8uf98"
    else:
        db_serveur = "rct-gexsidb123.groupe-prunay.fr"
        db_user = "robot_darva"
        db_password = "Robot_478"
        
    connexion = 'Driver={SQL Server Native Client 11.0}'+';SERVER='+db_serveur+';DATABASE='+db_nomBase+';UID='+db_user+';PWD='+ db_password             
    try:
        cnxn = pyodbc.connect(connexion)  
        cursor = cnxn.cursor() 
        tmpResult = 1
    except Exception as e:
        tmpResult = 0
        journal.Log ("Erreur de connexion à la base sql: " + str(e))
        
    return tmpResult

def FermetureConnexionSQL():  
    cursor.close() #Close the cursor and connection objects
    cnxn.close()
   
def GetString (fieldName):
    tmpResult  = ''

    try:
        tmpResult =  getattr(row, fieldName)
    except:
        tmpResult =  '' 
    return tmpResult


"""
def DbSelect(requetesql):
    try:
        cursor.execute(requetesql)
    except: 
        journal.Log('erreur sql')    
 

if OuvertureConnexionSQL() == 1:
   cursor.execute("SELECT TOP 10 [ANN_AGE_NUM_DE_SITE] AS num, [ANN_AGE_NOM_AGENCE], [ANN_AGE_CODIFICATION_DU_SITE],[ANN_AGE_ADRESSE_LIGNE_1],[ANN_AGE_ADRESSE_LIGNE_2], [ANN_AGE_ADRESSE_LIGNE_3] ,[ANN_AGE_CP] FROM [GEXSI].[dbo].[ANN_AGENCE]")
   row = cursor.fetchone()
   resultatretoursql = 1
   #row.count
   journal.Log(str(row))
   while row:
       #journal.Log (str(row[0]) + " " + str(row[1]))
       resultatretoursql = resultatretoursql + 1
       journal.Log(GetString("ANN_AGE_NOM_AGENCE"))
       row = cursor.fetchone()
   journal.Log (resultatretoursql)
   FermetureConnexionSQL() 
"""  