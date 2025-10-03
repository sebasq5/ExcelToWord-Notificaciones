# 📑 Generador de Notificaciones desde Excel a Word

Aplicación en **Python** con interfaz gráfica (**Tkinter**) que automatiza la creación de documentos **Word (.docx)** a partir de datos en un archivo **Excel (.xlsx)**.  
Ideal para generar **notificaciones personalizadas** en entornos académicos o administrativos, reduciendo tiempo y errores en la creación manual de documentos.

---

## 🚀 Características
- Interfaz gráfica sencilla y amigable con **Tkinter**.
- Lectura de datos desde Excel con **Pandas**.
- Uso de **plantillas Word (.docx)** con **marcadores** para personalizar la información.
- Generación de:
  - ✅ Documentos individuales (uno por estudiante).
  - ✅ Documento único consolidado con todos los registros.
- Manejo de rutas flexible y compatible con ejecución en `.exe` (PyInstaller).
- Barra de progreso en la interfaz para seguimiento del proceso.

---

## 🛠️ Tecnologías usadas
- **Python 3.x**
- [Tkinter](https://docs.python.org/3/library/tkinter.html) → Interfaz gráfica.
- [Pandas](https://pandas.pydata.org/) → Lectura y procesamiento de Excel.
- [python-docx](https://python-docx.readthedocs.io/) → Manipulación de documentos Word.
- [Pathlib](https://docs.python.org/3/library/pathlib.html) → Manejo moderno de rutas.

---

## 📂 Estructura del proyecto



---

## 📊 Ejemplo de Excel esperado
El archivo Excel debe contener columnas con encabezados como:

| CÉDULA DEL ESTUDIANTE | ID ESTUDIANTE | APELLIDOS | NOMBRES | CARRERA | TEMA | TRL1 | TRL2 | TRL3 |
|------------------------|---------------|-----------|---------|---------|------|------|------|------|
| 1234567890            | 001           | Pérez     | Juan    | Sistemas| IA   | Dr. A| Dr. B| Dr. C|

---

## 📝 Ejemplo de plantilla Word
El archivo `plantilla.docx` debe contener **marcadores** entre llaves dobles, por ejemplo:





Durante la ejecución, estos marcadores se reemplazan automáticamente con los valores del Excel.

---





