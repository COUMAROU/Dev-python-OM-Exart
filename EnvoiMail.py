# -*- coding: utf-8 -*-
"""
Created on Fri Jul 17 14:51:14 2020

@author: service.wintask
"""

# -*-coding:utf-8 -*


import email, smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders
import time
import datetime
import locale
import os
import Commun.Libs._powershell as powershell
import Commun.Libs._config as config
import Commun.Libs._variable_globale as var


def envoie_mail_simple(serveur=config.GetConfigValue("LETTRAGE-RECUP", "serveur", ""), port=25, adresse_exp=str(config.GetConfigValue("LETTRAGE-RECUP", "MAIL_EXPEDITEUR", "")), mdp="",
                       destinataire="", destinataire_copie=config.GetConfigValue("LETTRAGE-RECUP", "DEST-EMAIL", ""), pj="", destinaaireCache= "", sujet = ""):
    pieces_jointes = pj.split(",")
    msg = MIMEMultipart()
    text = ''
    #tab = "oui"
        
    codage = 'ISO-8859-15'
    typetexte = 'html'
    msg['From'] = adresse_exp
    if "@" in destinataire:
        msg['To'] = destinataire
    if "@" in destinataire_copie:
        msg['CC'] = destinataire_copie
    if "@" in destinaaireCache:
        msg['CCI'] = destinaaireCache

    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = sujet
    msg['Charset'] = codage
    msg['Content-Type'] = 'text/' + typetexte + '; charset=' + codage

    body = text

    if pj != "":
        for f in pieces_jointes:
            print(f)
            filename = f.split('/')
            filename = f.split('\\')
            filename = filename[(len(filename) - 1)]
            #attachment = open(pj, "rb")
            attachment = open(f, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
            msg.attach(part)
            
    msg.attach(MIMEText(body, 'html'))
    text = msg.as_string()
    server = smtplib.SMTP(serveur, port)
    #server.login(adresse_exp, mdp)

    

    if ";" in destinataire:
        destinataire2 = destinataire.split(";")
        for mail in destinataire2:
            if mail == "":
                pass
            else:
                server.sendmail(adresse_exp, mail, text)
    elif destinataire != '':
        destinataire2 = destinataire
        server.sendmail(adresse_exp, destinataire2, text)
    else:
        pass

    if ";" in destinataire_copie:
        destinataire_copie2 = destinataire_copie.split(";")
        for mail in destinataire_copie2:
            if mail == "":
                pass
            else:
                server.sendmail(adresse_exp, mail, text)
    elif destinataire_copie != '':
        destinataire_copie2 = destinataire_copie
        server.sendmail(adresse_exp, destinataire_copie2, text)
    else :
        pass
    if ";" in destinaaireCache:
        destinataire_copie2 = destinaaireCache.split(";")
        for mail in destinataire_copie2:
            if mail == "":
                pass
            else:
                server.sendmail(adresse_exp, mail, text)
    elif destinaaireCache != '':
        destinataire_copie2 = destinaaireCache
        server.sendmail(adresse_exp, destinataire_copie2, text)
    else :
        pass
    server.quit()
       
       