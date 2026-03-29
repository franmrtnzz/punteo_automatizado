# 🏦 Conciliación Bancaria Automatizada

Una aplicación que:

- ✅ Compara el **extracto bancario** con el **libro mayor**
- ✅ Detecta automáticamente **coincidencias** y **errores**
- ✅ Funciona en tu propio ordenador (**sin internet**)
- ✅ Genera un **Excel con los resultados**

> **Política clave:** cero falsos positivos → ante la duda, no concilia.  
> Esto es exactamente lo que necesitas en contabilidad.

---

## 🧩 PASO 1 — Instalar Python

### ¿Qué es Python?

Python es el "motor" que hace funcionar el programa. Necesitas instalarlo una sola vez.

### En Windows

1. Ve a 👉 [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Descarga Python (pulsa el botón grande amarillo)
3. Ejecuta el instalador

   > ⚠️ **MUY IMPORTANTE:** marca esta casilla antes de instalar:
   >
   > ☑️ **Add Python to PATH**

4. Pulsa **"Install Now"**

### En macOS

Abre la app **Terminal** (búscala en Spotlight con `Cmd + Espacio`) y ejecuta:

```bash
brew install python
```

> Si no tienes Homebrew, instálalo primero desde [brew.sh](https://brew.sh).

### Comprobar que funciona

Abre una terminal (CMD en Windows, Terminal en macOS) y escribe:

```bash
python --version
```

Si ves algo como `Python 3.x.x` → ✅ correcto.

---

## 📦 PASO 2 — Descargar el programa

1. Ve al repositorio en GitHub
2. Pulsa el botón verde **Code → Download ZIP**
3. Descomprime el archivo ZIP
4. Renombra la carpeta resultante a:

   ```
   punteo_automatizado
   ```

> 💡 Si tienes Git instalado, también puedes clonar directamente:
> ```bash
> git clone https://github.com/franmrtnzz/punteo_automatizado.git
> ```

---

## 💻 PASO 3 — Abrir la terminal en la carpeta del proyecto

### En Windows

1. Abre la carpeta `punteo_automatizado` en el Explorador de archivos
2. Haz clic en la **barra de ruta** (donde pone la dirección de la carpeta)
3. Escribe `cmd` y pulsa **Enter**

👉 Se abrirá una ventana negra (terminal). Ya estás dentro del proyecto.

### En macOS

1. Abre la app **Terminal**
2. Escribe `cd ` (con un espacio después) y arrastra la carpeta `punteo_automatizado` encima de la ventana
3. Pulsa **Enter**

---

## ⚙️ PASO 4 — Crear el entorno virtual

### ¿Qué es?

Un "entorno limpio" para que el programa funcione sin interferir con nada más de tu ordenador.

Ejecuta este comando en la terminal:

**Windows:**
```bash
python -m venv .venv
```

**macOS:**
```bash
python3 -m venv .venv
```

👉 Esto crea una carpeta llamada `.venv` dentro del proyecto (puede tardar unos segundos).

---

## 🔌 PASO 5 — Activar el entorno virtual

**Windows (CMD):**
```cmd
.venv\Scripts\activate.bat
```

**Windows (PowerShell):**
```powershell
.venv\Scripts\Activate.ps1
```

**macOS / Linux:**
```bash
source .venv/bin/activate
```

✅ Si funciona, verás **`(.venv)`** al inicio de la línea de la terminal. Eso significa que el entorno está activo.

---

## 📚 PASO 6 — Instalar las dependencias

### ¿Qué es esto?

Instalar todas las "piezas" adicionales que el programa necesita para funcionar.

Ejecuta:

```bash
pip install -r requirements.txt
```

⏳ Puede tardar entre 1 y 3 minutos. Espera a que termine.

---

## 🚀 PASO 7 — Ejecutar la aplicación

Ejecuta:

```bash
streamlit run app.py
```

---

## 🌐 PASO 8 — Abrir la aplicación en el navegador

Normalmente se abrirá sola. Si no, abre tu navegador y ve a:

👉 [http://localhost:8502](http://localhost:8502)

---

## 🧠 PASO 9 — Usar el programa

1. **Sube el archivo del banco** (extracto bancario en `.xls` o `.xlsx`)
2. **Sube el archivo del mayor** (libro mayor en `.xls` o `.xlsx`)
3. Pulsa **"Ejecutar conciliación"**
4. **Descarga el Excel** con los resultados

El informe generado indicará para cada movimiento:

| Estado | Significado |
|--------|-------------|
| ✅ **OK** | Coincide en ambas fuentes |
| ⏳ **Pendiente** | Necesita revisión manual |
| ❌ **No conciliado** | No se encontró correspondencia |

---

## ⚠️ Errores típicos (y cómo solucionarlos)

### ❌ `"python" no se reconoce como un comando`

**Causa:** Python no está bien instalado o no se añadió al PATH.  
**Solución:** Reinstala Python y asegúrate de marcar **☑️ Add Python to PATH** durante la instalación.

---

### ❌ La carpeta `.venv` no existe

**Causa:** No se ha creado el entorno virtual.  
**Solución:** Ejecuta:

```bash
python -m venv .venv
```

---

### ❌ `No module named 'streamlit'`

**Causa:** No se instalaron las dependencias.  
**Solución:** Asegúrate de tener el entorno activo `(.venv)` y ejecuta:

```bash
pip install -r requirements.txt
```

---

### ❌ No se abre la web

**Solución:** Abre tu navegador manualmente y entra en:

👉 [http://localhost:8502](http://localhost:8502)

---

### ❌ Error al leer el archivo Excel

**Solución:** Cierra el archivo Excel en tu ordenador antes de subirlo al programa.

---

## 🏎️ Resumen rápido (6 comandos)

Para usuarios con algo más de soltura técnica:

```bash
cd punteo_automatizado
python -m venv .venv
.venv\Scripts\activate.bat        # Windows
source .venv/bin/activate          # macOS / Linux
pip install -r requirements.txt
streamlit run app.py
```

---

## 📁 Estructura del proyecto

```
punteo_automatizado/
├── app.py                 # 🖥 Interfaz web (Streamlit)
├── conciliacion.py        # 🧠 Motor de conciliación
├── requirements.txt       # 📚 Dependencias de Python
├── LEEME_conciliacion.txt # 📝 Documentación técnica detallada
└── README.md              # 📖 Este archivo
```
