# Sistema de Filtrado y Descarga de Reportes Automatizado

Este proyecto es un script automatizado desarrollado en Python que permite:
- Descargar reportes desde un sitio web utilizando Selenium.
- Filtrar y organizar los datos de los reportes en formato Excel.
- Enviar los reportes filtrados por correo electrónico con un archivo adjunto.

### Características principales

1. **Automatización de descargas:** Utiliza Selenium para iniciar sesión y descargar reportes desde una URL específica.
2. **Filtrado de datos:** Aplica filtros basados en la fecha actual y ajusta columnas para una mejor visualización.
3. **Notificaciones por correo:** Envía el archivo filtrado automáticamente como adjunto por correo electrónico.

### Tecnologías utilizadas

- **Python**
  - Selenium
  - Pandas
  - Openpyxl
  - Python-dotenv
- **Herramientas adicionales:** Selenium WebDriver (ChromeDriver)

---

### Configuración

#### Clonar el repositorio
```bash
git clone https://github.com/renners22/RPA-filtro-orden-y-envio-de-reporte.git
```

#### Instalación de dependencias
```bash
pip install -r requirements.txt
```

#### Convertir el script en un ejecutable
Si deseas convertir el script en un archivo ejecutable (.exe), puedes usar `pyinstaller`:

1. Instala PyInstaller si no lo tienes:
   ```bash
   pip install pyinstaller
   ```

2. Ejecuta el siguiente comando en la raíz del proyecto:
   ```bash
   pyinstaller --onefile --windowed main.py
   ```

   Esto generará un archivo ejecutable en la carpeta `dist/`.

#### Ejecución
Navega a la raíz del proyecto en la consola y ejecuta:
```bash
python main.py
```

---

### Flujo de trabajo

1. Verifica si los archivos `(1)` y `(2)` están en la carpeta base:
   - Si existen, se eliminan automáticamente para evitar errores.
2. Inicia sesión en el sitio web utilizando Selenium.
3. Navega a la URL donde se encuentra el archivo `(1)` y lo descarga.
4. El archivo descargado `(1)` se mueve desde la carpeta de descargas a la carpeta base.
5. Se ejecuta una función que ordena y filtra los datos del archivo `(1)`, generando un archivo `(2)`.
6. El archivo `(2)` se adjunta a un correo electrónico y es enviado al destinatario especificado.

---

### Resultado esperado

El archivo filtrado `(2)` tendrá columnas ordenadas y ajustadas, como se muestra en el ejemplo:

![Ejemplo de archivo filtrado](docs/ejemplo_filtrado.png)

---

### Licencia

Este proyecto está licenciado bajo la [MIT License](LICENSE).

---

### Contribuciones

Las contribuciones son bienvenidas. Si tienes sugerencias o mejoras, no dudes en abrir un issue o enviar un pull request.
