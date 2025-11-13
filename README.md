# Automatización de Exportación de Tickets JIRA

Script creado para **Decal** que automatiza la descarga y envío de reportes desde **JIRA**.  
Accede automáticamente al portal, exporta el filtro **“Tickets Pendent”** en formato CSV, lo convierte a Excel y lo envía por correo a un destinatario configurado.  
Tras el envío, elimina los archivos descargados para mantener limpia la carpeta de descargas.

---

## ⚙️ Funcionalidad

1. Inicia sesión en JIRA mediante Selenium.  
2. Exporta los tickets filtrados en formato CSV.  
3. Convierte el CSV a Excel (.xlsx).  
4. Envía el archivo por correo usando Outlook (SMTP).  
5. Elimina los archivos generados.

---

## 🧩 Requisitos

- **Python 3.13+**  
- **Google Chrome** y **ChromeDriver** (en el PATH)  
- Librerías:
  ```bash
  pip install selenium pandas openpyxl
