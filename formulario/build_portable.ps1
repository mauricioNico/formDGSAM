# Script para generar la aplicación portable con jpackage
# Ejecuta Maven para compilar y empaquetar el JAR
Write-Host "Compilando y empaquetando el proyecto con Maven..."
& mvn clean package -DskipTests
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error en Maven. Abortando."
    exit 1
}

# Limpiar directorio anterior
Write-Host "Limpiando directorio anterior..."
if (Test-Path "dist") {
    Remove-Item -Recurse -Force "dist"
}

# Ejecuta jpackage para crear la aplicación portable
Write-Host "Generando aplicación portable con jpackage..."
& jpackage --input target --main-jar formulario-1.0.0.jar --main-class dgsam.App --name FormularioDGSAM --type app-image --dest dist

if ($LASTEXITCODE -eq 0) {
    Write-Host "¡Aplicación portable generada exitosamente en 'dist/FormularioDGSAM'!"
    
    # Ocultar archivos que no son el EXE principal en la carpeta bin
    Write-Host "Ocultando archivos de soporte..."
    $binDir = "dist\FormularioDGSAM\bin"
    if (Test-Path $binDir) {
        # Eliminar archivo .ico
        # Eliminar cualquier .ico que jpackage pueda haber dejado en bin o en subcarpetas
        Get-ChildItem -Path $binDir -Recurse -Filter *.ico -File -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
            Write-Host "Eliminado icono: $($_.FullName)"
        }

        # Además eliminar .ico en otras carpetas de la app (si existen)
        $appRoot = "dist\FormularioDGSAM"
        if (Test-Path $appRoot) {
            Get-ChildItem -Path $appRoot -Recurse -Filter *.ico -File -ErrorAction SilentlyContinue | ForEach-Object {
                Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
                Write-Host "Eliminado icono: $($_.FullName)"
            }
        }
        
        # Ocultar archivos que no sean el EXE
        Get-ChildItem -Path $binDir -File | Where-Object { $_.Extension -ne ".exe" } | ForEach-Object {
            & attrib +h "$($_.FullName)"
            Write-Host "Ocultado: $($_.Name)"
        }
    }
    
    # Ocultar carpetas de soporte (app, runtime, etc)
    $supportDirs = @("app", "runtime", "lib")
    foreach ($dir in $supportDirs) {
        $dirPath = "dist\FormularioDGSAM\$dir"
        if (Test-Path $dirPath) {
            & attrib +h "$dirPath"
            Write-Host "Carpeta ocultada: $dir"
        }
    }
    
    Write-Host "Copia todo el directorio 'dist/FormularioDGSAM' a la máquina destino."
} else {
    Write-Host "Error en jpackage."
}