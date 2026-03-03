# Ficha de Datos Personales - DGSAM

Aplicación de escritorio para la recopilación de datos personales del personal militar.

# Ficha de Datos Personales - DGSAM

Aplicación de escritorio para la recopilación de datos personales del personal militar.

## Instalación y Ejecución

### Opción 1: Aplicación Portable (Recomendado)

1. Ejecuta el script `build_portable.ps1` para generar la aplicación portable.
2. Ve a la carpeta `dist/FormularioDGSAM/`.
3. Copia todo el directorio `FormularioDGSAM` a la máquina destino.
4. Haz doble clic en `FormularioDGSAM.exe`.

**Ventaja**: Java está incluido en la aplicación. No requiere instalación de Java en la máquina destino.

### Opción 2: Ejecutable .exe (Requiere Java)

1. Ejecuta `mvn package` para generar el JAR.
2. Ve a la carpeta `target/`.
3. Haz doble clic en `formularioDeCarga.exe`.

**Requisito**: Java 17+ instalado en el sistema.

### Opción 3: Línea de Comandos

```bash
java -jar target/formulario-1.0.0.jar
```

**Requisito**: Java 17+ instalado.

## Desarrollo

### Compilación
```bash
mvn clean compile
```

### Empaquetado
```bash
mvn clean package
```

### Generar Aplicación Portable
```bash
.\build_portable.ps1
```

Esto crea un directorio `dist/FormularioDGSAM/` con la aplicación portable.

## Características

- ✅ **Cascading Dropdowns**: Escalafón → Especialidad Básica → Especialidad Avanzada
- ✅ **Carga de datos desde Excel**: Carga automática de especialidades desde archivo de referencia
- ✅ **Exportación a Excel**: Genera archivo `ficha_datos.xlsx` con tabla formateada
- ✅ **Interfaz moderna**: Utiliza FlatLaf para una apariencia moderna
- ✅ **Geolocalización**: Obtiene ubicación automáticamente desde IP
- ✅ **Validación de datos**: DNI requerido antes de guardar

## Estructura del Formulario

### Pestaña 1: Datos Personales
- Nombre Completo
- DNI
- Teléfono
- Email
- Escalafón (combo cascada)
- Especialidad Básica (combo cascada, depende de Escalafón)
- Especialidad Avanzada (combo cascada, depende de Especialidad Básica)

### Pestaña 2: Domicilio
- Dirección
- Código Postal
- Ciudad
- Provincia
- País
- Geolocalización (automática)

### Pestaña 3: Preferencias
- Disponibilidad
- Notas

## Requisitos del Sistema

**Para Ejecutable .exe**:
- Windows 7 o superior
- No requiere Java instalado

**Para Script Batch o JAR**:
- Windows 7 o superior
- Java 8 o superior instalado

## Construcción desde Código Fuente

### Flujo Completo: Código → JAR → EXE

Si haces cambios en el código, sigue estos pasos para regenerar los distribuciones:

#### Paso 1: Editar el código
Modifica los archivos Java que necesites:
- `src/main/java/dgsam/FichaDatosPersonalesFrame.java` (interfaz principal)
- `src/main/java/dgsam/App.java` (punto de entrada)
- Otros archivos según sea necesario

#### Paso 2: Compilar y empaquetar con Maven
```bash
cd formulario
mvn clean package -DskipTests
```

Esto:
- ✅ Compila el código Java
- ✅ Genera `target/formularioDeCarga.jar` (19.8 MB con todas las dependencias)
- ✅ Valida que no hay errores

**Tiempo esperado**: 30-60 segundos

#### Paso 3: Generar el EXE desde el JAR
```powershell
cd formulario

$jpackage = "C:\Program Files\Java\jdk-17.0.5\bin\jpackage.exe"
& $jpackage --input target --name "formularioDeCarga" --main-jar formulario-1.0.0.jar --main-class dgsam.App --type app-image
```

Esto genera:
- Nueva carpeta `formularioDeCarga/` con `formularioDeCarga.exe` adentro
- Se copia automáticamente el JAR actualizado

**Tiempo esperado**: 10-15 segundos

#### Paso 4: Ejecutar para verificar
```powershell
# Ejecutable directo
Start-Process ".\formularioDeCarga\formularioDeCarga.exe"

# O JAR (si tienes Java)
java -jar target/formularioDeCarga.jar

# O script batch
.\formularioDeCarga.bat
```

### Automatizar el proceso completo

Crea un archivo `build.ps1` en la raíz del proyecto:

```powershell
# Script: build.ps1
Write-Host "=== Compilando JAR ===" -ForegroundColor Green
mvn clean package -DskipTests

if ($LASTEXITCODE -ne 0) {
    Write-Host "Error en compilación" -ForegroundColor Red
    exit 1
}

Write-Host "=== Generando EXE ===" -ForegroundColor Green
$jpackage = "C:\Program Files\Java\jdk-17.0.5\bin\jpackage.exe"
& $jpackage --input target --name "formularioDeCarga" --main-jar formularioDeCarga.jar --main-class dgsam.App --type app-image

if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Completado. EXE en: formularioDeCarga/formularioDeCarga.exe" -ForegroundColor Green
} else {
    Write-Host "Error en generación del EXE" -ForegroundColor Red
}
```

Luego simplemente ejecuta:
```powershell
.\build.ps1
```

### Checklist de construcción

| Paso | Comando | Salida esperada |
|------|---------|-----------------|
| 1️⃣ Editar código | Edita en VS Code | Sin errores de sintaxis |
| 2️⃣ Compilar | `mvn clean package -DskipTests` | `BUILD SUCCESS` |
| 3️⃣ Generar EXE | `jpackage ...` | `formularioDeCarga.exe` creado |
| 4️⃣ Probar | `Start-Process .\formularioDeCarga\...exe` | App se abre correctamente |

## Archivos de Datos

- **Entrada**: `src/main/resources/ESPECIALIDADES DEL PERSONAL MILITAR SUBALTERNO (EN PROCESO).xlsx`
- **Salida**: `ficha_datos.xlsx` (se crea en el directorio de ejecución)
- **Referencia**: `src/main/resources/grados.csv`

## Contacto

Para reportar errores o sugerencias, contacta al equipo de desarrollo.

---

**Versión**: 1.0  
**Última actualización**: Marzo 2026
