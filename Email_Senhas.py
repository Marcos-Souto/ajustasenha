# -*- coding: utf-8 -*-
"""
Created on Tue Sep 11 14:40:22 2018

@author: marcos.souto
"""

import mimetypes
import os
import smtplib
from log import logger
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def adiciona_anexo(msgX, filename, strTitulo): # Envia e-mail com anexo

    de = 'riscocase@gmail.com'
    para = ['marcos.souto@grupocase.com.br',
            "adriano.velloso@grupocase.com.br",
            "viviane.costa@grupocase.com.br"]

# --- Cria corpo do e-mail ---------------------------#
    msg = MIMEMultipart()
    msg['From'] = de
    msg['To'] = ', '.join(para)
    msg['Subject'] = strTitulo #'Gestão de Senhas Valid'

    # Corpo da mensagem
    msg.attach(MIMEText(msgX, 'html', 'utf-8'))
# ----------------------------------------------------#
# --- Verifica se esxite anexo -----------------------#    
    if not os.path.isfile(filename):
        return

    ctype, encoding = mimetypes.guess_type(filename)

    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'

    maintype, subtype = ctype.split('/', 1)

    if maintype == 'text':
        with open(filename) as f:
            mime = MIMEText(f.read(), _subtype=subtype)
    elif maintype == 'image':
        with open(filename, 'rb') as f:
            mime = MIMEImage(f.read(), _subtype=subtype)
    elif maintype == 'audio':
        with open(filename, 'rb') as f:
            mime = MIMEAudio(f.read(), _subtype=subtype)
    else:
        with open(filename, 'rb') as f:
            mime = MIMEBase(maintype, subtype)
            mime.set_payload(f.read())
        encoders.encode_base64(mime)
# --------------------------------------------------------------#
# --- Monta o e-mail ------------------------------------------#
    mime.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(mime)
    raw = msg.as_string()
    # -- envia o e-mail e fecha o servido
    smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    smtp.login('riscocase@gmail.com', 'dinamica')
    smtp.sendmail(de, para, raw)
    smtp.quit()
    logger.error("E-mail enviado com senhas de gestão para casos ")
# ----------------------------------------------------------------------#


def email_sem_anexo(msgX, strTitulo):  # envia e-mail sem anexo
    de = 'riscocase@gmail.com'
    para = ['marcos.souto@grupocase.com.br',
            "adriano.velloso@grupocase.com.br",
            "viviane.costa@grupocase.com.br"]
# --- Cria corpo do e-mail ---------------------------#
    msg = MIMEMultipart()
    msg['From'] = de
    msg['To'] = ', '.join(para)
    msg['Subject'] = strTitulo

    # Corpo da mensagem
    msg.attach(MIMEText(msgX, 'html', 'utf-8'))
    raw = msg.as_string()
# -- envia o e-mail e fecha o servido
    smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    smtp.login('riscocase@gmail.com', 'dinamica')
    smtp.sendmail(de, para, raw)
    smtp.quit()

