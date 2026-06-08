# PPT2PDF

Conversor por lotes de presentaciones de PowerPoint (`.ppt` y `.pptx`) a PDF mediante PowerShell y automatización COM de Microsoft PowerPoint.

Este proyecto está pensado para Windows y para ejecutarse localmente con Microsoft Office instalado.

## Características

- Convierte todos los archivos `.ppt` y `.pptx` de una carpeta de entrada.
- Procesa subcarpetas de forma recursiva.
- Genera los PDFs en una carpeta de salida dedicada.
- Crea automáticamente la carpeta de salida si no existe.
- Muestra el progreso por archivo y continúa ante errores individuales.

## Requisitos

- Windows
- PowerShell 5.1 o PowerShell 7+
- Microsoft PowerPoint instalado
- Permisos para crear objetos COM en la sesión actual

## Estructura del proyecto

```text
PPT2PDF/
├── Convert.ps1
├── input/
└── output/
```

- `input/`: coloca aquí los `.ppt` y `.pptx` a convertir.
- `output/`: se crearán aquí los archivos PDF generados.

## Uso rápido

1. Coloca tus archivos PowerPoint dentro de `input/`.
2. Abre una terminal en la raíz del proyecto.
3. Ejecuta:

```powershell
.\Convert.ps1
```

Si todo va bien, verás mensajes como:

```text
Convertido: Presentacion1.pptx -> C:\ruta\PPT2PDF\output\Presentacion1.pdf
```

## Cómo funciona

El script:

1. Resuelve rutas absolutas basadas en la ubicación de `Convert.ps1`.
2. Busca archivos `*.ppt*` de forma recursiva en `input/`.
3. Inicia `PowerPoint.Application` por COM.
4. Abre cada presentación en modo no interactivo.
5. Guarda cada archivo como PDF (`ppSaveAsPDF = 32`) en `output/`.
6. Cierra PowerPoint y libera el objeto COM al finalizar.

## Solución de problemas

### Error al iniciar PowerPoint vía COM

Mensaje típico:

```text
No se pudo iniciar PowerPoint vía COM...
```

Revisa:

- PowerPoint está instalado y activado.
- La sesión de PowerShell tiene permisos suficientes.
- No hay políticas corporativas bloqueando automatización COM.

### No se generan PDFs

Revisa:

- Existen archivos `.ppt` o `.pptx` dentro de `input/`.
- No hay archivos protegidos con contraseña.
- No hay cuadros de diálogo de Office pendientes de interacción.

## Personalización

Puedes editar en `Convert.ps1`:

- `$inputFolder` para cambiar la carpeta de entrada.
- `$outputFolder` para cambiar la carpeta de salida.

## Limitaciones

- Depende de Microsoft PowerPoint (no funciona sin Office).
- No está pensado para Linux/macOS.
- Archivos con macros, enlaces externos o fuentes no disponibles pueden diferir visualmente en el PDF.

## Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.