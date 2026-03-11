# ScannerPerito

**Generador de Evidencia Digital para Peritaje Informatico**

Script de PowerShell para la recoleccion forense de evidencia digital de discos y carpetas. Genera un listado completo de archivos con hashes de integridad, metadatos del sistema y un informe pericial formal en formato TXT, util como soporte documental en procesos de peritaje informatico.

---

## Caracteristicas

- **Numero de caso auto-incremental** - Cada ejecucion genera un identificador unico (`SP-000001`, `SP-000002`...) persistido en `case_seed.json` con historial completo
- **Escaneo profundo** - Recorre recursivamente unidades o carpetas, incluyendo archivos ocultos y de sistema
- **Hash SHA256** - Calculo de hash para cada archivo escaneado, garantizando integridad de la evidencia
- **Informacion forense completa** del equipo: hostname, SO, CPU, RAM, BIOS, interfaces de red (MAC/IP), zona horaria
- **Informacion del disco** - Modelo, serial, interfaz, estado de salud, particiones, sistema de archivos, espacio libre, datos SMART
- **Exportacion a Excel** (`.xlsx`) con formato profesional o **CSV delimitado por pipes** como respaldo automatico
- **Informe pericial TXT** con cabecera de entidad, datos del autor, descripcion del caso y hashes de integridad de los archivos generados
- **Streaming a disco** - Procesamiento archivo por archivo sin acumular en memoria, soporta millones de archivos
- **Barra de progreso** en tiempo real durante el escaneo
- **Log de ejecucion** con timestamps para auditoria
- **Configurable** mediante archivo JSON para datos de entidad y autor

---

## Requisitos

| Componente | Version minima |
|---|---|
| **Sistema Operativo** | Windows 10 (1809+), Windows 11, Windows Server 2019+ |
| **PowerShell** | 5.1 o superior |
| **ImportExcel** *(opcional)* | Cualquier version - si no esta instalado, el script ofrece instalarlo o usa CSV |

> El script valida la compatibilidad del sistema al inicio y se detiene si no cumple los requisitos.

---

## Instalacion

1. Clone el repositorio:

```powershell
git clone https://github.com/orlandobcr/ScannerPerito.git
cd ScannerPerito
```

2. Edite `config.json` con los datos de su entidad y autor (ver seccion [Configuracion](#configuracion)).

3. Ejecute el script:

```powershell
.\ScannerPerito.ps1
```

> **Nota:** Se recomienda ejecutar como Administrador para acceder a todos los archivos del sistema. Si no se ejecuta como administrador, el script continuara pero algunos archivos podrian ser inaccesibles.

---

## Uso

### Flujo de ejecucion

El script guia al usuario paso a paso:

```
1. Validacion de compatibilidad del SO
2. Carga de configuracion (config.json)
3. Asignacion automatica del numero de caso (SP-XXXXXX)
4. Solicitud de descripcion del caso al usuario
5. Seleccion de unidad a procesar (menu interactivo)
6. Seleccion de alcance: unidad completa o carpeta especifica
7. Recopilacion de informacion del equipo y disco
8. Confirmacion de parametros antes de iniciar
9. Escaneo de archivos con barra de progreso
10. Exportacion a Excel o CSV
11. Generacion del informe pericial TXT
12. Calculo de hashes de integridad de los archivos generados
```

### Ejemplo de ejecucion

```
  ================================================================
    ScannerPerito - Generador de Evidencia Digital
    Version 1.0.0
  ================================================================

  SO detectado: Microsoft Windows 11 Pro (Build 26200)

  ============================================
    DATOS DEL CASO
  ============================================

  Numero de caso asignado: SP-000001

  Ingrese la descripcion del caso.
  Ejemplo: Extraccion de evidencia digital del disco duro del equipo
           de escritorio asignado al area de contabilidad, solicitada
           por la Fiscalia General dentro del proceso 2026-00123.

  Descripcion: Recoleccion de evidencia del disco USB entregado por el area juridica

  ============================================
    SELECCION DE UNIDAD A PROCESAR
  ============================================

    [1] C: - OS (NTFS, 237.00 GB, Libre: 98.50 GB)
    [2] D: - DATOS (NTFS, 500.00 GB, Libre: 312.00 GB)
    [3] E: - USB_EVIDENCIA (exFAT, 32.00 GB, Libre: 28.00 GB)

  Seleccione el numero de la unidad: 3

  ============================================
    ALCANCE DEL ESCANEO
  ============================================

    [1] Unidad completa (E:\)
    [2] Carpeta especifica

  Seleccione una opcion (1 o 2): 1
```

---

## Configuracion

### config.json

Archivo de configuracion con los datos de la entidad y el autor del informe. Debe editarse antes de la primera ejecucion:

```json
{
  "entity_name": "Nombre de la Entidad / Empresa",
  "nit": "NIT o Identificacion Fiscal",
  "address": "Direccion de la entidad",
  "city": "Ciudad",
  "phone": "+57 XXX XXX XXXX",
  "email": "contacto@entidad.com",
  "author": "Nombre del Perito / Autor del Informe",
  "author_id": "Cedula o identificacion",
  "notes": ""
}
```

| Campo | Descripcion |
|---|---|
| `entity_name` | Nombre de la empresa o entidad que realiza el peritaje |
| `nit` | Numero de identificacion tributaria o fiscal |
| `address` | Direccion fisica de la entidad |
| `city` | Ciudad |
| `phone` | Telefono de contacto |
| `email` | Correo electronico de contacto |
| `author` | Nombre completo del perito o autor del informe |
| `author_id` | Cedula o documento de identidad del autor |
| `notes` | Notas adicionales que se incluyen en el informe (opcional) |

> Si `config.json` no existe al ejecutar el script, se crea automaticamente una plantilla.

### case_seed.json

Archivo generado automaticamente en el directorio del script. Almacena el ultimo numero de caso y un historial de todas las ejecuciones:

```json
{
  "last_case_number": 3,
  "last_updated": "2026-03-11 14:30:00 -05:00",
  "history": [
    {
      "case_id": "SP-000001",
      "case_number": 1,
      "description": "Recoleccion de evidencia del disco USB...",
      "timestamp": "2026-03-11 10:00:00 -05:00"
    }
  ]
}
```

> El numero de caso solo se incrementa y persiste cuando el usuario confirma el escaneo. Si cancela, no se consume el numero.

---

## Archivos de salida

Cada ejecucion genera una carpeta con el formato:

```
ScannerPerito_SP-000001_20260311_143022/
├── ScannerPerito_FileList_20260311_143022.xlsx   (o .csv)
├── ScannerPerito_Report_20260311_143022.txt
└── ScannerPerito_Execution.log
```

### Listado de archivos (Excel/CSV)

Contiene una fila por cada archivo y carpeta encontrada, con las siguientes columnas:

| Columna | Descripcion |
|---|---|
| `RutaCompleta` | Ruta absoluta del archivo o carpeta |
| `Nombre` | Nombre del archivo o carpeta |
| `Extension` | Extension del archivo |
| `EsDirectorio` | SI / NO |
| `EsOculto` | SI / NO - indica si tiene atributo oculto |
| `EsSistema` | SI / NO - indica si tiene atributo de sistema |
| `EsSoloLectura` | SI / NO - indica si tiene atributo de solo lectura |
| `TamanoBytes` | Tamano en bytes (vacio para directorios) |
| `HashSHA256` | Hash SHA256 del archivo (vacio para directorios) |
| `FechaCreacion` | Fecha y hora de creacion |
| `FechaModificacion` | Fecha y hora de ultima modificacion |
| `FechaAcceso` | Fecha y hora de ultimo acceso |
| `Propietario` | Usuario propietario del archivo |
| `Atributos` | Cadena completa de atributos del sistema de archivos |

**Formato Excel:** Si el modulo `ImportExcel` esta disponible y el numero de filas no excede 1,048,575, se genera un archivo `.xlsx` con formato profesional (encabezado congelado, negrita, tabla con estilo).

**Formato CSV:** Si no hay modulo Excel o se supera el limite de filas, se genera un `.csv` delimitado por pipes (`|`) en codificacion UTF-8 con BOM.

### Informe pericial (TXT)

Documento formal que incluye:

- Cabecera con datos de la entidad y el autor
- Numero y descripcion del caso
- Informacion completa del equipo examinado (hardware, SO, red)
- Informacion detallada del disco y volumen
- Parametros del escaneo (ruta, fechas, algoritmo hash)
- Estadisticas (total archivos, directorios, errores, tamano)
- Referencia al archivo de datos con su hash SHA256 de integridad
- Hash SHA256 del propio informe

### Log de ejecucion

Registro cronologico de todas las operaciones realizadas, errores de acceso y eventos relevantes durante la ejecucion.

---

## Manejo de errores

El script esta disenado para ser resiliente:

- **Archivos bloqueados** (pagefile.sys, registry hives): se registra `ERROR` en la columna de hash y se continua
- **Acceso denegado**: se registra `ACCESS_DENIED` en la columna de hash y el propietario se marca como `DESCONOCIDO`
- **Reparse points / junctions**: se omiten automaticamente para evitar bucles infinitos
- **Rutas largas (>260 caracteres)**: limitacion de PowerShell 5.1; se registra el error y se continua

Todos los errores se contabilizan en el informe final y se detallan en el log de ejecucion.

---

## Consideraciones de rendimiento

- El escaneo usa **streaming**: cada archivo se procesa y escribe a disco inmediatamente, sin acumular en memoria
- El cuello de botella principal es el **calculo de hash SHA256** para archivos grandes
- La barra de progreso se actualiza cada 100 elementos para minimizar el impacto en rendimiento
- Para drives con millones de archivos, el proceso puede tomar varias horas; el progreso se muestra en tiempo real

---

## Estructura del proyecto

```
ScannerPerito/
├── ScannerPerito.ps1    # Script principal
├── config.json          # Configuracion de entidad y autor
├── case_seed.json       # Numero de caso (generado automaticamente)
├── .gitignore           # Excluye carpetas de salida
└── README.md            # Este archivo
```

---

## Licencia

Uso interno. Todos los derechos reservados.
