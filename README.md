# Conciliación Bancaria Automatizada

Sistema de punteo automático entre el extracto bancario y el libro mayor (cuenta 572.1).  
Interfaz web local construida con **Streamlit**.

---

## ¿Qué hace este programa?

Compara automáticamente los movimientos del banco con los apuntes del libro mayor y genera un informe Excel indicando:

- Qué movimientos están **conciliados** (coinciden en ambas fuentes).
- Cuáles quedan **pendientes de revisión manual**.

> **Política:** máxima precisión, cero falsos positivos. Ante cualquier duda, el movimiento se deja como pendiente.

---

## Requisitos previos

| Requisito | Detalle |
|-----------|---------|
| **Python** | Versión 3.10 o superior |
| **Sistema operativo** | macOS, Linux o Windows |
| **Archivos Excel** | Extracto bancario (`.xls` / `.xlsx`) y libro mayor (`.xls` / `.xlsx`) |

---

## Instalación paso a paso

### 1. Instalar Python (si no lo tienes)

- **macOS:** Abre Terminal y ejecuta:
  ```bash
  brew install python
  ```
  Si no tienes Homebrew, instálalo primero desde [brew.sh](https://brew.sh).

- **Windows:** Descarga el instalador desde [python.org](https://www.python.org/downloads/) y durante la instalación **marca la casilla "Add Python to PATH"**.

- **Linux (Debian/Ubuntu):**
  ```bash
  sudo apt update && sudo apt install python3 python3-venv python3-pip
  ```

### 2. Descargar el proyecto

```bash
git clone https://github.com/TU_USUARIO/punteo_automatizado.git
cd punteo_automatizado
```

> Si no tienes Git, puedes descargar el ZIP desde GitHub haciendo clic en **Code → Download ZIP**, descomprimir y abrir la carpeta.

### 3. Crear el entorno virtual

```bash
python3 -m venv .venv
```

### 4. Activar el entorno virtual

- **macOS / Linux:**
  ```bash
  source .venv/bin/activate
  ```

- **Windows (PowerShell):**
  ```powershell
  .venv\Scripts\Activate.ps1
  ```

- **Windows (CMD):**
  ```cmd
  .venv\Scripts\activate.bat
  ```

> Sabrás que está activo porque verás `(.venv)` al inicio de la línea de tu terminal.

### 5. Instalar las dependencias

```bash
pip install -r requirements.txt
```

### 6. Ejecutar la aplicación

```bash
streamlit run app.py
```

Se abrirá automáticamente tu navegador en `http://localhost:8502` con la interfaz de conciliación.

---

## Uso

1. **Sube el extracto bancario** (archivo `.xls` o `.xlsx` del banco).
2. **Sube el libro mayor** (archivo `.xls` o `.xlsx` exportado de contabilidad).
3. Ajusta los parámetros en la barra lateral si lo necesitas (normalmente los valores por defecto funcionan bien).
4. Pulsa **"Ejecutar conciliación"**.
5. Revisa los resultados y **descarga el informe Excel** generado.

---

## Uso por terminal (sin interfaz web)

Si prefieres ejecutar la conciliación directamente desde la terminal:

1. Coloca los dos archivos Excel en la misma carpeta que `conciliacion.py`.
   - El archivo del banco debe contener la palabra **"BANCO"** en el nombre.
   - El archivo del mayor debe contener la palabra **"MAYOR"** en el nombre.
2. Ejecuta:
   ```bash
   python3 conciliacion.py
   ```
3. El informe se generará en la misma carpeta.

---

## Estructura del proyecto

```
punteo_automatizado/
├── app.py                 # Interfaz web (Streamlit)
├── conciliacion.py        # Motor de conciliación (6 pases)
├── requirements.txt       # Dependencias de Python
├── .streamlit/
│   └── config.toml        # Configuración del servidor Streamlit
├── LEEME_conciliacion.txt # Documentación técnica detallada
└── README.md              # Este archivo
```

---

## Solución de problemas

| Problema | Solución |
|----------|----------|
| `command not found: python3` | Instala Python siguiendo el paso 1 |
| `No module named 'streamlit'` | Activa el entorno virtual (paso 4) e instala dependencias (paso 5) |
| El navegador no se abre | Ve manualmente a `http://localhost:8502` |
| Error al leer el Excel | Comprueba que el archivo no esté abierto en otro programa |
