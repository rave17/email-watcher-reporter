# 📧 Monitor de Correos CI/CD para Azure Boards

Sistema automatizado que monitorea correos electrónicos de ejecuciones CI/CD y crea work items automáticamente en Azure Boards.

## 🚀 Características

- ✅ Monitoreo de múltiples remitentes (Azure DevOps, Jenkins, GitLab, etc.)
- ✅ Creación automática de work items en Azure Boards
- ✅ Clasificación inteligente por tipo de evento (éxito, fallo, advertencia)
- ✅ Extracción automática de detalles desde el cuerpo del correo
- ✅ Logging detallado con timestamps y emojis

## 📋 Requisitos Previos

- **Python 3.10+**
- Cuenta de **Azure DevOps** con permisos para crear work items
- Acceso **IMAP** a cuenta de correo electrónico
- **Personal Access Token (PAT)** de Azure DevOps

## ⚙️ Configuración Rápida

### 1. Clonar y preparar entorno
```bash
git clone <tu-repositorio>
cd email-watcher-reporter
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# o
.venv\Scripts\activate      # Windows
2. Instalar dependencias
pip install -r requirements.txt
```
3. Configurar variables de entorno
Crear archivo .env en la raíz del proyecto:

## Configuración IMAP
```
IMAP_SERVER=imap.gmail.com
IMAP_USER=tu_email@gmail.com
IMAP_PASS=tu_password_app
```

## Configuración Azure DevOps
```
AZURE_ORG=https://dev.azure.com/tu_organizacion
AZURE_PROJECT=Nombre_Del_Proyecto
AZURE_PAT=tu_personal_access_token
```

## Remitentes a monitorear (separados por coma)
```
MONITORED_SENDERS=azuredevops@microsoft.com
```
## Opcional: Configuración avanzada
```
CHECK_INTERVAL=60
LOG_FILE=monitor_correos.log
```
4. Configurar mapeos (opcional)
Editar MAPEO_REMITENTES en el código para agregar nuevos remitentes.

## 🏃‍♂️ Ejecución
```bash
python main.py
```
El sistema iniciará y mostrará logs en consola:

```
[2024-01-15 10:30:45] 🚀 Iniciando Monitor de ejecución de pruebas automáticas
[2024-01-15 10:30:45] 👀 Remitentes monitoreados: azuredevops@microsoft.com
[2024-01-15 10:30:46] ✅ Conexión IMAP exitosa
📊 Configuración de Azure Boards
Estados soportados:
To Do → Bugs creados

Doing → En revision

Done → Ejecucion existosa
```
## Tipos de work items:
Issue - Para errores y advertencias

Task - Para ejecuciones exitosas


## 🎨 Personalización
Agregar nuevo remitente:
* Agregar email a  MONITORED_SENDERS en .env

* Agregar mapeo en MAPEO_REMITENTES

* Agregar plantilla en PLANTILLAS_DETALLES (opcional)

Modificar columnas/estados:
Editar MAPEO_TABLERO para reflejar tu configuración de Azure Boards.

## 📋 Logs
Los logs se guardan en monitor_correos.log e incluyen:

Timestamps exactos

Emojis para fácil visualización

Detalles de cada operación

Errores y advertencias

## ⚠️ Notas Importantes
El sistema corre en bucle infinito con intervalos configurables

Los correos procesados se marcan como leídos

Los work items creados se taggean como "Auto-Generado"

Verificar periodicamente que el PAT de Azure no haya expirado
