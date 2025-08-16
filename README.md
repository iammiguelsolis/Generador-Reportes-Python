# 📊 Generador de Reportes Automático

Una aplicación de escritorio con interfaz gráfica que permite generar reportes personalizados en Word (.docx) a partir de datos en CSV de forma masiva y automatizada.

![Python](https://img.shields.io/badge/python-v3.7+-blue.svg)
![Platform](https://img.shields.io/badge/platform-windows%20%7C%20macOS%20%7C%20linux-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

## ✨ Características

- 🖥️ **Interfaz gráfica moderna** con tkinter
- 📄 **Procesamiento masivo** de datos desde CSV
- 📝 **Templates personalizables** en formato Word
- 👁️ **Vista previa** de datos antes del procesamiento
- 📊 **Barra de progreso** en tiempo real
- 🔄 **Procesamiento en segundo plano** (no bloquea la interfaz)
- 🛡️ **Manejo robusto de errores**
- 📂 **Gestión automática** de carpetas de salida
- 🧹 **Sanitización automática** de nombres de archivos

## 📸 Capturas de Pantalla

```
┌─────────────────────────────────────┐
│ 📊 Generador de Reportes Automático │
├─────────────────────────────────────┤
│ 📄 Archivo CSV: [clientes.csv    ] │
│ 📝 Template DOCX: [template.docx ] │
│ 📂 Carpeta salida: [reportes     ] │
│ ┌─────────────────────────────────┐ │
│ │ 👁️ Vista previa de datos        │ │
│ │ ┌─────┬─────┬─────┬─────┬─────┐ │ │
│ │ │Name │Email│Phone│City │...  │ │ │
│ │ └─────┴─────┴─────┴─────┴─────┘ │ │
│ └─────────────────────────────────┘ │
│ [🚀 Generar] [📂 Abrir] [🗑️ Limpiar] │
│ ████████████████░░░░ 80%          │
└─────────────────────────────────────┘
```

## 🚀 Instalación

### Prerrequisitos

- Python 3.7 o superior
- pip (gestor de paquetes de Python)

### Pasos de instalación

1. **Clona el repositorio**
```bash
git clone https://github.com/tu-usuario/generador-reportes.git
cd generador-reportes
```

2. **Instala las dependencias**
```bash
pip install -r requirements.txt

3. **Ejecuta la aplicación**
```bash
python generador_reportes.py
```

## 📋 Uso

### 1. Preparación de archivos

#### CSV de datos
Crea un archivo CSV con los datos de tus clientes/registros:
```csv
nombre,email,telefono,ciudad,empresa
Juan Pérez,juan@email.com,555-1234,Madrid,TechCorp
María García,maria@email.com,555-5678,Barcelona,DataSoft
Carlos López,carlos@email.com,555-9012,Valencia,CloudSys
```

#### Template de Word
Crea un documento Word (.docx) con placeholders que serán reemplazados:
```
Estimado/a $nombre$,

Nos complace informarle que su empresa $empresa$ 
ha sido seleccionada para nuestro programa especial.

Datos de contacto:
- Email: $email$
- Teléfono: $telefono$
- Ciudad: $ciudad$

Atentamente,
El equipo de ventas
```

### 2. Usar la aplicación

1. **Ejecuta la aplicación**
   ```bash
   python generador_reportes.py
   ```

2. **Selecciona archivos**
   - Haz clic en "📁 Examinar" junto a "Archivo CSV"
   - Selecciona tu archivo CSV con los datos
   - Haz clic en "📁 Examinar" junto a "Template DOCX"
   - Selecciona tu template de Word

3. **Revisa la vista previa**
   - Los datos del CSV aparecerán en la tabla
   - Verifica que las columnas sean correctas

4. **Configura la salida**
   - La carpeta por defecto es "reportes"
   - Puedes cambiarla con "📁 Cambiar"

5. **Genera los reportes**
   - Haz clic en "🚀 Generar Reportes"
   - Observa el progreso en la barra
   - Al finalizar, haz clic en "📂 Abrir Carpeta"

### 3. Resultado

La aplicación generará un archivo Word por cada fila del CSV:
```
reportes/
├── reporte_Juan_Pérez.docx
├── reporte_María_García.docx
└── reporte_Carlos_López.docx
```

## 🔧 Configuración Avanzada

### Formato de placeholders

Los placeholders en el template deben tener el formato `$nombre_columna$`:
- ✅ `$nombre$` → Reemplaza con el valor de la columna "nombre"
- ✅ `$email$` → Reemplaza con el valor de la columna "email"
- ❌ `{nombre}` → No será reconocido
- ❌ `%nombre%` → No será reconocido

### Caracteres especiales

Los nombres de archivos se sanitizan automáticamente:
- Se eliminan caracteres especiales
- Se reemplazan espacios por guiones bajos
- Se mantienen solo caracteres alfanuméricos, guiones y guiones bajos

## 🐛 Solución de Problemas

### Error: "El archivo CSV no existe"
- Verifica que la ruta del archivo sea correcta
- Asegúrate de que el archivo no esté abierto en otro programa

### Error: "El template DOCX no existe"
- Verifica que la ruta del template sea correcta
- Asegúrate de que sea un archivo .docx válido

### Los placeholders no se reemplazan
- Verifica que uses el formato `$columna$`
- Asegúrate de que el nombre de la columna coincida exactamente
- Ten en cuenta que es sensible a mayúsculas/minúsculas

### La aplicación se congela
- La generación se ejecuta en segundo plano
- Espera a que termine o cierra la aplicación si hay un error

## 🛠️ Estructura del Proyecto

```
generador-reportes/
├── generador_reportes.py    # Aplicación principal
├── requirements.txt         # Dependencias
├── README.md               # Este archivo
├── examples/               # Archivos de ejemplo
│   ├── clientes.csv        # CSV de ejemplo
│   └── template.docx       # Template de ejemplo
└── reportes/              # Carpeta de salida (se crea automáticamente)
```

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## ⭐ Reconocimientos

- [python-docx](https://python-docx.readthedocs.io/) - Manipulación de documentos Word
- [pandas](https://pandas.pydata.org/) - Análisis de datos
- [tkinter](https://docs.python.org/3/library/tkinter.html) - Interfaz gráfica

---

⭐ **¡Si este proyecto te fue útil, dale una estrella!** ⭐