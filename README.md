# Belleza Lua: ingresos, gastos y stock

App web (HTML/JS) en GitHub Pages + Firebase (Authentication + Realtime Database, **plan gratuito Spark**).
Herramienta interna: **no emite facturas ni tickets**.

## Qué hace

| Pestaña | Función |
|---|---|
| **Ingresos** | Hojas del talonario por foto (con IA gratuita) o a mano. Total, base, IVA, cremas y colonias del mes. Avisa de números de hoja sin registrar. Exporta el Excel del mes con el formato de Declarando |
| **Gastos** | Facturas en PDF (se leen sin IA y sin coste) o en foto (con IA). Gastos sin factura: SS, modelos 115/130/303, letra de la máquina. Totales por categoría |
| **Stock** | Catálogo, escáner de código de barras, pedido, historial. Las facturas de Germaine dan entrada en stock automáticamente |
| **Resumen** | Resultado del mes y del año, gastos por categoría y exportación del Excel anual |
| **Ajustes** | Clave de IA, NIF del negocio, próximo nº de Declarando, factor PVP y proveedores aprendidos |

### Excel del mes (Declarando)
Se genera sobre `plantilla_mes.xlsx` (el Excel mensual de siempre, vacío de datos) con las mismas reglas:
- Hojas "INGRESOS A+ B" y "FACTURAS GASTOS.", mismos anchos, bordes, colores y formatos (fechas como fecha `DD/MM/YYYY`).
- Una fila por hoja del talonario; una factura de Declarando por día (NIF `2026-000xx` encadenado); si el día llega a 400 €, las hojas se reparten en varias facturas que quedan por debajo de 400 € (factura simplificada).
- Cremas en amarillo, colonias en verde, comisiones SumUp en rojo.
- Gastos: una fila por gasto con su concepto (automático o el que escribas en "Concepto en el Excel").
- La columna M y las filas de TOTAL B quedan vacías / con fórmula: la app no las rellena.
- Nombre del archivo: `MES_<MES>_<AÑO>_MIRIAM_LORENZO.xlsx`.

## Archivos (súbelos todos a la raíz del repositorio)

`index.html`, `lector.js`, `plantilla_mes.xlsx`, `catalogo.json`, `database.rules.json`, `manifest.json`, `icon-180.png`, `icon-512.png`, `README.md` y la carpeta `vendor/` (contiene `exceljs.min.js`, licencia MIT)

## Puesta en marcha (unos 30 minutos)

### 1. Firebase (gratis)
1. console.firebase.google.com → **Añadir proyecto** → `belleza-lua-stock` (sin Analytics).
2. **Authentication** → Comenzar → **Correo electrónico/contraseña** → Habilitar.
3. Authentication → **Users** → *Add user*: Miriam y tú.
4. Authentication → **Settings** → *User actions*: desmarca **Enable create (sign-up)**.
5. Authentication → Settings → **Authorized domains** → añade `polmm4810.github.io`.
6. **Realtime Database** → Crear base de datos → **Bélgica (europe-west1)** → modo bloqueado.
7. Realtime Database → **Reglas**: pega `database.rules.json` → Publicar.
8. Realtime Database → **Datos**: crea `usuarios` con un hijo por usuario: clave = **UID** (Authentication → Users → *User UID*), valor = `true`.
9. Configuración del proyecto → **Tus apps** → icono web `</>` → registra "Belleza Lua" → copia `firebaseConfig`.

### 2. `index.html`
Busca `PEGA AQUÍ LA CONFIGURACIÓN` y sustituye el bloque `FIREBASE_CONFIG`. Comprueba que `databaseURL` coincide con la de tu base de datos.

### 3. GitHub Pages
1. Repositorio nuevo `belleza-lua-stock` (POLMM4810) y sube todos los archivos, incluida la carpeta `vendor`.
2. Settings → Pages → *Deploy from a branch* → `main` / `root`.
3. Disponible en `https://polmm4810.github.io/belleza-lua-stock/`.

### 4. IA gratuita para las fotos (opcional pero recomendado)
1. aistudio.google.com → **Get API key** → *Create API key* (no pide tarjeta).
2. En la app: **Ajustes → Clave API** → pega la clave → Guardar → **Probar**.
3. Recomendado: en console.cloud.google.com → *APIs y servicios → Credenciales* → la clave → *Restricciones de aplicaciones: sitios web* → `polmm4810.github.io/*`.

La clave se guarda en tu base de datos de Firebase (solo la leen los dos usuarios autorizados), no en el código público.
Si el modelo deja de existir, en Ajustes → Modelo prueba `gemini-2.5-flash`.

### 5. Primer uso
1. Entra, ve a **Ajustes**: NIF del negocio y próximo nº de Declarando (p. ej. `2026-00183`).
2. **Stock → Cargar catálogo inicial** (solo una vez).
3. iPhone: Safari → Compartir → **Añadir a pantalla de inicio**.

## Consejos para las fotos del talonario
- Una hoja por foto, de frente y con buena luz (la foto oscura o movida falla).
- Que se vean el número rojo, la fecha y el total.
- Siempre aparece una pantalla de revisión antes de guardar: comprueba el total.
- Se pueden seleccionar varias fotos a la vez desde la galería.

## Copia de seguridad
- **Ajustes → Copia de seguridad → Descargar copia**: un archivo `COPIA_BELLEZA_LUA_AAAA-MM-DD.json` con ingresos, gastos, stock, movimientos, proveedores y ajustes (sin la clave de IA). En el iPhone: Guardar en Archivos / iCloud.
- Hazla una vez al mes, después de sacar el Excel. Si pasan más de 30 días sin copia, la pestaña Ingresos avisa.
- El plan gratuito de Firebase no hace copias automáticas: este archivo es la única copia.

### Restaurar una copia (no hay botón en la app)
El archivo de copia guarda los datos dentro de `"datos"`, separados por nodos: `ingresos`, `gastos`, `productos`, `movimientos`, `proveedores` y `config`. Se restaura **nodo a nodo** desde la consola de Firebase:
1. Saca del archivo el contenido de un nodo, por ejemplo todo lo que hay dentro de `"datos" → "ingresos"`, y guárdalo como `ingresos.json` (o pídeselo a Claude con el archivo de copia).
2. Firebase → Realtime Database → **Datos** → entra en el nodo `ingresos` (haz clic en él para que la ruta de arriba termine en `/ingresos`).
3. Menú ⋮ → **Importar JSON** → elige `ingresos.json`. Sustituye todo el contenido de ese nodo.
4. Repite con los demás nodos que quieras recuperar.

⚠ **Nunca importes en la raíz** (`/`): borraría el nodo `usuarios` y nadie podría entrar. Si pasara, vuelve a crear `usuarios` con los UID y valor `true`.
⚠ Al restaurar `config`, la clave de IA no viene en la copia: vuelve a pegarla en Ajustes.

## Numeración de Declarando
- La app guarda qué números ha usado cada mes exportado. Si al exportar un mes los números se pisan con otro mes ya exportado (p. ej. al volver a sacar un mes con una hoja olvidada), avisa antes de generar y propone el siguiente número libre.

## Notas
- Las facturas en PDF con texto se leen en el propio móvil, sin IA y sin enviar nada fuera.
- Las fotos se envían a Google solo si hay clave de IA. Con usuarios del EEE, Google aplica a los datos el trato de los servicios de pago (no se usan para entrenar); compruébalo en las condiciones de la API de Gemini.
- La app no guarda las fotos: conserva el talonario y las facturas originales.
- Proveedores reconocidos de serie: Germaine, Yodeyma, O2, la luz, SumUp, Depilshop y Alcampo. El resto se aprende al guardar la primera factura.
