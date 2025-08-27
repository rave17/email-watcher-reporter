#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#python V 3.10.5
"""
Sistema de Monitoreo de Correos de ejecuciones para Azure Boards.
Monitorea múltiples fuentes: Azure DevOps, os-certificacionoperaciones, otros.
Crea Work Items automáticamente con detalles consumidos de los mails
"""

import imaplib
import email
import requests
import base64
import time
import json
import os
import re
from email.header import decode_header
from urllib.parse import quote
from dotenv import load_dotenv

# Cargar configuraciones y variables
load_dotenv()

# Configuración desde variables de entorno
config = {
    "imap_server": os.getenv("IMAP_SERVER", "imap.gmail.com"),
    "imap_user": os.getenv("IMAP_USER"),
    "imap_pass": os.getenv("IMAP_PASS"),
    "azure_org": os.getenv("AZURE_ORG"),
    "azure_project": os.getenv("AZURE_PROJECT"),
    "azure_pat": os.getenv("AZURE_PAT"),
    "log_file": os.getenv("LOG_FILE", "monitor_correos.log"),
    "monitored_senders": os.getenv("MONITORED_SENDERS", "azuredevops@microsoft.com").split(","),
    "check_interval": int(os.getenv("CHECK_INTERVAL", "60"))
}

# Validar configuración esencial
config_requerida = ["imap_user", "imap_pass", "azure_org", "azure_project", "azure_pat"]
for clave in config_requerida:
    if not config[clave]:
        raise ValueError(f"Configuración faltante: {clave.upper()}")

# Mapeos personalizados para el tablero
MAPEO_TABLERO = {
    "columnas_estados": {
        "Bugs creados": "To Do",
        "En revision": "Doing", 
        "Ejecucion existosa": "Done"
    }
}

# Mapeo de remitentes a columnas y patrones
MAPEO_REMITENTES = {
    "azuredevops@microsoft.com": {
        "failed": "Bugs creados",
        "succeeded": "Ejecucion existosa",
        "warning": "En revision"
    },
    "os-certificacionoperaciones@osde.com.ar": {
        "failed": "Bugs creados",
        "success": "Ejecucion existosa", 
        "unstable": "En revision"
    }
}

# Plantillas para detalles específicos por herramienta
PLANTILLAS_DETALLES = {
    "azuredevops@microsoft.com": {
        "failed": "🚨 Pipeline de Azure DevOps falló",
        "succeeded": "✅ Pipeline de Azure DevOps exitoso",
        "warning": "⚠️ Advertencia en Pipeline de Azure DevOps"
    },
    "os-certificacionoperaciones@osde.com.ar": {
        "failed": "🚨 Prueba fallida",
        "success": "✅ Prueba exitosa"
    }
}


class Logger:
    """Manejador de logs"""
    def __init__(self, archivo_log):
        self.archivo_log = archivo_log
        
    def registrar(self, mensaje, emoji="📝"):
        
        marca_tiempo = time.strftime("%d-%m-%Y %H:%M:%S")
        linea = f"[{marca_tiempo}] {emoji} {mensaje}"
        
        with open(self.archivo_log, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
        print(linea)


class ClienteAzureDevOps:
    """Cliente para interactuar con Azure DevOps"""
    def __init__(self, organizacion, proyecto, pat):
        self.org = organizacion.rstrip('/')
        self.proyecto = proyecto
        self.pat = pat
        self.encabezados = {
            "Authorization": "Basic " + base64.b64encode((":" + pat).encode()).decode(),
            "Content-Type": "application/json-patch+json"
        }
    
    def obtener_tipos_elementos(self):
        """Obtiene los tipos de elementos de trabajo disponibles"""
        try:
            proyecto_codificado = quote(self.proyecto)
            url = f"{self.org}/{proyecto_codificado}/_apis/wit/workitemtypes?api-version=6.0"
            
            respuesta = requests.get(url, headers=self.encabezados, timeout=30)
            
            if respuesta.status_code == 200:
                tipos = [tipo['name'] for tipo in respuesta.json()['value']]
                return tipos
            return ["Issue", "Task"] 
        except Exception as error:
            print(f"Error obteniendo tipos: {error}")
            return ["Issue", "Task"]  
    
    def obtener_estados_elemento(self, tipo_elemento):
        """Obtiene los estados disponibles para un tipo de elemento"""
        try:
            proyecto_codificado = quote(self.proyecto)
            url = f"{self.org}/{proyecto_codificado}/_apis/wit/workitemtypes/{tipo_elemento}/states?api-version=6.0"
            
            respuesta = requests.get(url, headers=self.encabezados, timeout=30)
            
            if respuesta.status_code == 200:
                estados = [estado['name'] for estado in respuesta.json()['value']]
                # print(f"🎯 Estados REALES para '{tipo_elemento}': {estados}") 
                return estados
            return ["To Do", "Doing", "Done"]
        except Exception as error:
            print(f"Error obteniendo estados: {error}")
            return ["To Do", "Doing", "Done"]
    
    def extraer_detalles_correo(self, mensaje):
        """Extrae detalles específicos del cuerpo del correo"""
        detalles = {}
        
        try:
            # Obtener el cuerpo mail
            cuerpo = ""
            if mensaje.is_multipart():
                for part in mensaje.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        cuerpo = part.get_payload(decode=True).decode(errors='ignore')
                        break
            else:
                cuerpo = mensaje.get_payload(decode=True).decode(errors='ignore')
            
            # Extraer información para detalle
            detalles['cuerpo_completo'] = cuerpo[:1000] + "..." if len(cuerpo) > 1000 else cuerpo
            
            # Patrones para información (revisar para extraer datos y completar tarjetas)
            patrones = {
                'tiempo_ejecucion': r'(time|duration|tiempo|duracion)[:\s]*([0-9:\.]+)\s*(seconds|secs|minutos|minutes|ms|s)',
                'error': r'(error|exception|failed|failure)[:\s]*(.+)',
                'resultado': r'(result|status|estado)[:\s]*(success|failed|passed|completed|completado)',
                'url_reporte': r'(https?://[^\s<>"]+|www\.[^\s<>"]+)'
            }
            
            for clave, patron in patrones.items():
                coincidencias = re.findall(patron, cuerpo, re.IGNORECASE)
                if coincidencias:
                    detalles[clave] = coincidencias[0] if isinstance(coincidencias[0], str) else ' '.join(coincidencias[0])
            
            return detalles
            
        except Exception as error:
            print(f"Error extrayendo detalles: {error}")
            return {'cuerpo_completo': 'Error al extraer contenido del correo'}
    
    def crear_elemento_trabajo(self, titulo, tipo_elemento, columna_destino, detalles=None, remitente=""):
        """Crea un nuevo elemento de trabajo en Azure DevOps con detalles"""
        try:
            proyecto_codificado = quote(self.proyecto)
            url = f"{self.org}/{proyecto_codificado}/_apis/wit/workitems/${tipo_elemento}?api-version=6.0"
            
            # Determinar estado según la columna destino
            estado = MAPEO_TABLERO["columnas_estados"].get(columna_destino, "To Do")
            
            # Verificar que el estado existe para este tipo de elemento
            estados_disponibles = self.obtener_estados_elemento(tipo_elemento)
            if estado not in estados_disponibles:
                print(f"⚠️ Estado '{estado}' no disponible para {tipo_elemento}. Estados disponibles: {estados_disponibles}")
                # Usar el primer estado disponible como fallback
                estado = estados_disponibles[0] if estados_disponibles else "To Do"
                print(f"⚠️ Usando estado por defecto: {estado}")
            
            # Construir descripción con detalles
            descripcion = self._construir_descripcion(columna_destino, detalles, remitente)
            
            # Datos para crear el elemento
            datos = [
                {"op": "add", "path": "/fields/System.Title", "value": titulo},
                {"op": "add", "path": "/fields/System.Description", "value": descripcion},
                {"op": "add", "path": "/fields/System.State", "value": estado},
                {"op": "add", "path": "/fields/System.Tags", "value": "Auto-Generado"}
            ]
            
            # Agregar campo personalizado para remitente si existe
            if remitente:
                datos.append({
                    "op": "add", 
                    "path": "/fields/System.History", 
                    "value": f"Generado desde: {remitente}"
                })
            
            respuesta = requests.post(url, headers=self.encabezados, json=datos, timeout=30)
            
            if respuesta.status_code in [200, 201]:
                id_elemento = respuesta.json().get('id', 'N/A')
                url_elemento = f"{self.org}/{self.proyecto}/_workitems/edit/{id_elemento}"
                return True, id_elemento, url_elemento, estado
            else:
                print(f"Error API: {respuesta.status_code} - {respuesta.text}")
                return False, None, None, None
                
        except Exception as error:
            print(f"Error creando elemento: {error}")
            return False, None, None, None
    
    def _construir_descripcion(self, columna_destino, detalles, remitente):
        """Construye la descripción según el tipo de elemento"""
        descripcion = f"<h3>📧 Elemento generado automáticamente</h3>"
        descripcion += f"<p><strong>Remitente:</strong> {remitente}</p>"
        
        if columna_destino == "Bugs creados":
            descripcion += self._descripcion_error(detalles)
        elif columna_destino == "Ejecucion existosa":
            descripcion += self._descripcion_exitosa(detalles)
        elif columna_destino == "En revision":
            descripcion += self._descripcion_advertencia(detalles)
        else:
            descripcion += "<p>Notificación de sistema CI/CD</p>"
        
        if detalles and 'cuerpo_completo' in detalles:
            descripcion += "<h4>📋 Contenido completo:</h4>"
            descripcion += f"<pre>{detalles['cuerpo_completo']}</pre>"
        
        descripcion += "<p><em>🔄 Creado automáticamente desde monitoreo de correo</em></p>"
        return descripcion
    
    def _descripcion_error(self, detalles):
        """Construye descripción para errores"""
        descripcion = "<h3>🚨 Error en Ejecución</h3>"
        descripcion += "<p>Se ha detectado un error durante la ejecución.</p>"
        
        if detalles:
            descripcion += "<h4>🔍 Detalles del error:</h4>"
            descripcion += "<ul>"
            
            if 'error' in detalles:
                descripcion += f"<li><strong>Error:</strong> {detalles['error']}</li>"
            
            if 'tiempo_ejecucion' in detalles:
                descripcion += f"<li><strong>Tiempo de ejecución:</strong> {detalles['tiempo_ejecucion']}</li>"
            
            descripcion += "</ul>"
        
        return descripcion
    
    def _descripcion_exitosa(self, detalles):
        """Construye descripción para ejecuciones exitosas"""
        descripcion = "<h3>✅ Ejecución Exitosa</h3>"
        descripcion += "<p>La ejecución se ha completado sin errores.</p>"
        
        if detalles:
            descripcion += "<h4>📊 Métricas de ejecución:</h4>"
            descripcion += "<ul>"
            
            if 'tiempo_ejecucion' in detalles:
                descripcion += f"<li><strong>Tiempo de ejecución:</strong> {detalles['tiempo_ejecucion']}</li>"
            
            if 'resultado' in detalles:
                descripcion += f"<li><strong>Resultado:</strong> {detalles['resultado']}</li>"
            
            if 'url_reporte' in detalles:
                descripcion += f"<li><strong>Reporte:</strong> <a href='{detalles['url_reporte']}'>Ver reporte</a></li>"
            
            descripcion += "</ul>"
        
        return descripcion
    
    def _descripcion_advertencia(self, detalles):
        """Construye descripción para advertencias"""
        descripcion = "<h3>⚠️ Ejecución con Advertencias</h3>"
        descripcion += "<p>La ejecución se completó pero con advertencias que requieren revisión.</p>"
        
        if detalles:
            descripcion += "<h4>📝 Detalles:</h4>"
            descripcion += "<ul>"
            
            if 'error' in detalles:
                descripcion += f"<li><strong>Advertencia:</strong> {detalles['error']}</li>"
            
            if 'tiempo_ejecucion' in detalles:
                descripcion += f"<li><strong>Tiempo de ejecución:</strong> {detalles['tiempo_ejecucion']}</li>"
            
            descripcion += "</ul>"
        
        return descripcion


class ProcesadorCorreos:
    """Procesa correos electrónicos y extrae información relevante"""
    def __init__(self, servidor, usuario, contraseña):
        self.servidor = servidor
        self.usuario = usuario
        self.contraseña = contraseña
    
    def conectar(self):
        """Establece conexión con el servidor IMAP"""
        try:
            cliente = imaplib.IMAP4_SSL(self.servidor)
            cliente.login(self.usuario, self.contraseña)
            cliente.select("inbox")
            return cliente
        except Exception as error:
            print(f"Error conectando al servidor: {error}")
            return None
    
    def buscar_correos_monitoreados(self, cliente, remitentes):
        """Busca correos no leídos de múltiples remitentes monitoreados"""
        try:
            todos_correos = []
            
            for remitente in remitentes:
                remitente_limpio = remitente.strip()
                if not remitente_limpio:
                    continue
                    
                criterio = f'(UNSEEN FROM "{remitente_limpio}")'
                estado, mensajes = cliente.search(None, criterio)
                
                if estado == "OK" and mensajes[0]:
                    correos_remitente = mensajes[0].split()
                    todos_correos.extend([(msg_id, remitente_limpio) for msg_id in correos_remitente])
            
            return todos_correos
        except Exception as error:
            print(f"Error buscando correos: {error}")
            return []
    
    def decodificar_asunto(self, asunto_codificado):
        """Decodifica el asunto del mail"""
        try:
            partes_decodificadas = decode_header(asunto_codificado)
            asunto = ""
            for parte, codificacion in partes_decodificadas:
                if isinstance(parte, bytes):
                    asunto += parte.decode(codificacion if codificacion else 'utf-8', errors='ignore')
                else:
                    asunto += parte
            return asunto
        except:
            return str(asunto_codificado)
    
    def determinar_accion_por_remitente(self, asunto, remitente):
        """Determina la accion segun el remitente y el asunto"""
        asunto_lower = asunto.lower()
        remitente_limpio = remitente.strip().lower()
        
        # Buscar mapeo por remitente
        mapeo_remitente = None
        for remitente_mapeo in MAPEO_REMITENTES:
            if remitente_mapeo.lower() in remitente_limpio:
                mapeo_remitente = MAPEO_REMITENTES[remitente_mapeo]
                break
        
        # Si no hay mapeo específico, usar mapeo por defecto
        if not mapeo_remitente:
            mapeo_remitente = MAPEO_REMITENTES.get("azuredevops@microsoft.com", {})
        
        # Buscar patrones en el asunto
        for patron, columna in mapeo_remitente.items():
            if patron in asunto_lower:
                return columna, patron
                
        # Patrones genéricos si no hay coincidencia específica
        if any(p in asunto_lower for p in ["failed", "failure", "error", "falló", "fallo", "fallida"]):
            return "Bugs creados", "failed"
        elif any(p in asunto_lower for p in ["succeeded", "success", "exitoso", "completado", "exitosa"]):
            return "Ejecucion existosa", "success"
        elif any(p in asunto_lower for p in ["warning", "unstable", "advertencia", "inestable"]):
            return "En revision", "warning"
            
        return None, None

    def procesar_correo(self, cliente, id_mensaje, remitente, cliente_azure, logger):
        """Procesa un correo individual considerando el remitente"""
        try:
            estado, datos = cliente.fetch(id_mensaje, "(RFC822)")
            if estado != "OK":
                return
            
            mensaje = email.message_from_bytes(datos[0][1])
            asunto = self.decodificar_asunto(mensaje["subject"])
            
            logger.registrar(f"Procesando correo de {remitente}: {asunto}", "📧")
            
            # Extraer detalles del correo
            detalles = cliente_azure.extraer_detalles_correo(mensaje)
            detalles['remitente'] = remitente
            
            # Marcar como leído
            cliente.store(id_mensaje, '+FLAGS', '\\Seen')
            
            # Determinar acción basada en remitente y asunto
            columna, tipo_evento = self.determinar_accion_por_remitente(asunto, remitente)
            
            if not columna:
                logger.registrar(f"Correo de {remitente} no requiere acción: {asunto}", "📨")
                return
            
            # Determinar tipo de elemento
            tipos_disponibles = cliente_azure.obtener_tipos_elementos()
            
            tipo_elemento = "Issue"
            
            # Verificar que el tipo seleccionado existe
            if tipo_elemento not in tipos_disponibles:
                logger.registrar(f"⚠️ Tipo {tipo_elemento} no disponible. Usando primer tipo disponible", "⚠️")
                tipo_elemento = tipos_disponibles[0] if tipos_disponibles else "Issue"
            
            # Crear título apropiado
            titulo_prefijo = PLANTILLAS_DETALLES.get(remitente, {}).get(tipo_evento, "")
            if not titulo_prefijo:
                if tipo_evento == "failed":
                    titulo_prefijo = "🚨 Error en ejecución"
                elif tipo_evento == "success":
                    titulo_prefijo = "✅ Ejecución exitosa"
                else:
                    titulo_prefijo = "⚠️ Notificación"
            
            titulo = f"{titulo_prefijo}: {asunto[:100]}{'...' if len(asunto) > 100 else ''}"
            
            # Crear workItem con detalles
            exito, id_elemento, url, estado = cliente_azure.crear_elemento_trabajo(
                titulo, tipo_elemento, columna, detalles, remitente
            )
            
            if exito:
                logger.registrar(f"Elemento #{id_elemento} creado en '{columna}'", "✅")
                logger.registrar(f"Remitente: {remitente}", "👤")
                logger.registrar(f"Tipo evento: {tipo_evento}", "🎯")
                logger.registrar(f"URL: {url}", "🔗")
            else:
                logger.registrar("No se pudo crear el elemento", "❌")
                
        except Exception as error:
            logger.registrar(f"Error procesando correo de {remitente}: {error}", "❌")


def main():
    """Inicio"""
    logger = Logger(config["log_file"])
    logger.registrar("🚀 Iniciando Monitor de ejecucucion de pruebas automaticas", "🚀")
    logger.registrar(f"📧 Remitentes monitoreados: {', '.join(config['monitored_senders'])}", "👀")
    
    # Muestra informacion del mapeo al tablero
    logger.registrar("🎯 Configuración del tablero:", "⚙️")
    for columna, estado in MAPEO_TABLERO["columnas_estados"].items():
        logger.registrar(f"  {columna} → {estado}")
    
    # Inicializar clientes
    cliente_azure = ClienteAzureDevOps(config["azure_org"], config["azure_project"], config["azure_pat"])
    procesador_correos = ProcesadorCorreos(config["imap_server"], config["imap_user"], config["imap_pass"])
    
    while True:
        try:
            # Conectar y procesar correos
            cliente_imap = procesador_correos.conectar()
            if cliente_imap:
                # Buscar correos de todos los remitentes monitoreados
                correos = procesador_correos.buscar_correos_monitoreados(
                    cliente_imap, config["monitored_senders"]
                )
                
                if correos:
                    logger.registrar(f"📬 Encontrados {len(correos)} correos nuevos de {len(config['monitored_senders'])} remitentes", "📬")
                    
                    for id_correo, remitente in correos:
                        procesador_correos.procesar_correo(
                            cliente_imap, id_correo, remitente, cliente_azure, logger
                        )
                else:
                    logger.registrar(f"📭 No hay correos nuevos de {len(config['monitored_senders'])} remitentes monitoreados", "📭")
                
                cliente_imap.close()
                cliente_imap.logout()
            else:
                logger.registrar("❌ No se pudo conectar al servidor IMAP", "❌")
            
            # Esperar antes de volver a revisar
            logger.registrar(f"⏰ Esperando {config['check_interval']} segundos para siguiente verificación", "⏰")
            time.sleep(config["check_interval"])
            
        except Exception as error:
            logger.registrar(f"❌ Error en el bucle principal: {error}", "❌")
            time.sleep(config["check_interval"])


if __name__ == "__main__":
    main()