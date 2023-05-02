import win32serviceutil
import jenkins
import logging
from datetime import datetime
import requests
from requests.auth import HTTPBasicAuth
import json
import os
from os.path import exists
import win32com.client


class nodePipeline:

    def __init__(self):
        pass

    def serviceValidator(self):

        # Log
        cwd = os.getcwd()
        nowHour = datetime.now().strftime("%m%d%Y-%H%M%S")
        log_file = "node_status_"+nowHour+".txt"
        logging.basicConfig(filename=cwd+"\\"+log_file, level=logging.INFO, format='%(asctime)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

        # Conexión al servidor Jenkins
        server = jenkins.Jenkins('http://192.168.18.33:8080/', username='admin', password='admin')

        # Nombre del servicio a reiniciar
        # service_name = "jenkins8080agent"
        node = 'MasterWinPro'

        # Obtención de información del nodo
        node_info = server.get_node_info(node)
        node_name = node_info['displayName']

        # Verificar el estado del nodo
        if node_info['offline']:

            # Logging
            node_status = f"El nodo {node_name} està desconectado"
            # service_status = f"Reiniciando servicio {service_name}..."
            logging.info(node_status)

            ####################################
            # Manejando Outlook
            cwd = os.getcwd()
            ol = win32com.client.Dispatch('Outlook.Application')
            olmailitem = 0x0
            newmail = ol.CreateItem(olmailitem)

            # Receptor de email y contenido de el
            newmail.Subject = 'Jenkins node '+node_name+' is down. Sending report.'
            newmail.To = 'claudio.torres.burgos@gmail.com'
            newmail.Body = 'Jenkins node '+node_name+' isn\'t functioning normally. Adjuncting information.\n\n'+json.dumps(node_info, indent = 4)

            # Archivo adjunto
            fileToAttach = cwd+"\\"+log_file
            print("File to attach: " + fileToAttach)
            attach = fileToAttach
            newmail.Attachments.Add(attach)
            newmail.Send()
            #####################################

            # win32serviceutil.RestartService(service_name)
            # logging.info(service_status)
            print(node_status)    

        else:

            node_status = f"El nodo {node_name} està conectado"
            logging.info(node_status)
            print(node_status)

        # Retorno de info, usada en posteo a Confluence
        return log_file, json.dumps(node_info, indent=4)

    def confluencePoster(self, log_file, info):

        # Creando dirección
        cwd = os.getcwd()
        fileDir = cwd+"\\"+log_file
        pageIdFileName = cwd+"\\"+datetime.now().strftime("%d-%m-%Y")+".txt"

        # Creación de body de llamada POST
        with open(fileDir, "r") as fileIn:
            data = fileIn.read()
            fileIn.close()
        data = data + "\n" + info + "\n\n"
        spaceId=163841

        # Información de conexión
        url = "https://claudio1992.atlassian.net/wiki/api/v2/pages"
        auth = HTTPBasicAuth("claudio.torres.burgos@gmail.com", "Api Token (hidden)")

        # Creación de request POST
        headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
        }
        payload = json.dumps( {
        "spaceId": spaceId,
        "status": "current",
        "title": log_file,
        "parentId": "",
        "body": {
            "representation": "storage",
            "value": data
        } })

        # Creando llamadas de acuerdo a existencia de archivo diario con ID de página de Confluence
        if not exists(pageIdFileName):

            # Llamada POST
            response = requests.request(
            "POST",
            url,
            data=payload,
            headers=headers,
            auth=auth
            )

            # Extracción de información y creación de archivo con ID de página creada en Confluence
            res = response.json()
            file = open(pageIdFileName, "w")
            id = str(res["id"])
            file.write(id)
            file.close()

        else:

            # Extracción de ID y creación de request
            file = open(pageIdFileName, "r")
            id = file.readline()
            file.close()
            bodyFormat="storage"
            params={
                "id": id,
                "body-format": bodyFormat
            }

            # Llamada GET
            response = requests.request(
            "GET",
            url,
            params=params,
            headers=headers,
            auth=auth
            )

            # Construcción de llamada PUT
            res = response.json()
            url=url+"/"+id
            title=res["results"][0]["title"]
            newBody=res["results"][0]["body"]["storage"]["value"]+data
            version=res["results"][0]["version"]["number"]+1
            payload = json.dumps({
                "id": id,
                "status": "current",
                "title": title,
                "spaceId": spaceId,
                "body": {
                    "representation": "storage",
                    "value": newBody
                },
                "version": {
                    "number": version,
                    "message": "Update"
                }
            })
            
            # Actualizando página
            response = requests.request(
            "PUT",
            url,
            data=payload,
            headers=headers,
            auth=auth
            )

        # Imprimiendo status
        print(json.dumps(json.loads(response.text), sort_keys=True, indent=4, separators=(",", ": ")))  
      

pipeline = nodePipeline()
log_file, info = pipeline.serviceValidator()
offline = json.loads(info)["offline"]
if offline:
    pipeline.confluencePoster(log_file, info)
