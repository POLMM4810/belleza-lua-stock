# Belleza Lua: ingresos, gastos y stock

App web (HTML/JS) en GitHub Pages + Firebase (Authentication + Realtime Database, **plan gratuito Spark**).
Herramienta interna: **no emite facturas ni tickets**.

## Qué hace

| Pestaña | Función |
|---|---|
| **Ingresos** | Hojas del talonario por foto (con IA gratuita) o a mano, **con los productos vendidos** (descuentan stock en el mismo paso). Total, base, IVA, cremas y colonias del mes. Avisa de números de hoja sin registrar. Exporta el Excel del mes con el formato de Declarando |
| **Gastos** | Facturas en PDF (se leen sin IA y sin coste) o en foto (con IA). Recibos de autónomos y pagos de modelos 115/130/303 del banco en PDF. Gastos sin factura a mano. Totales por categoría |
| **Stock** | Catálogo, escáner de código de barras, pedido, historial. Las facturas de Germaine y de Yodeyma dan entrada en stock automáticamente y las hojas del talonario descuentan lo vendido |
| **Resumen** | Resultado del mes y del año, gastos por categoría y botón **Actualizar Excel anual** (explicado y con gráficas) |
| **Ajustes** | Clave de IA, NIF del negocio, próximo nº de Declarando, factor PVP, copia de seguridad (descargar y restaurar), histórico de años anteriores y proveedores aprendidos |

### Excel del mes (Declarando)
Se genera sobre `plantilla_mes.xlsx` (el Excel mensual de siempre, vacío de datos) con las mismas reglas:
- Hojas "INGRESOS A+ B" y "FACTURAS GASTOS.", mismos anchos, bordes, colores y formatos (fechas como fecha `DD/MM/YYYY`).
- Una fila por hoja del talonario; una factura de Declarando por día (NIF `2026-000xx` encadenado); si el día llega a 400 €, las hojas se reparten en varias facturas que quedan por debajo de 400 € (factura simplificada).
- Cremas en amarillo, colonias en verde, comisiones SumUp en rojo.
- Gastos: una fila por gasto con su concepto (automático o el que escribas en "Concepto en el Excel").
- La columna M y las filas de TOTAL B quedan vacías / con fórmula: la app no las rellena.
- Nombre del archivo: `MES_<MES>_<AÑO>_MIRIAM_LORENZO.xlsx`.

### Excel anual (para entender cómo va el negocio)
**Resumen → Actualizar Excel anual** genera `Belleza_Lua_Anual_<AÑO>.xlsx` del año que se ve arriba. Se puede sacar cuantas veces quieras: siempre sale con lo último apuntado. Lo hace `excel_anual.js` en el propio móvil, sin librerías ni internet, y se ve bien en la vista previa del iPhone (números y gráficas).

| Hoja | Qué tiene |
|---|---|
| EMPIEZA AQUÍ | Instrucciones, qué significa cada palabra (beneficio, margen, base, IVA…), colores y de dónde sale cada dato |
| RESUMEN | Las cifras clave del año, cuánto hay que facturar al mes y al día para cubrir gastos (los días de trabajo se pueden cambiar en la casilla amarilla), gráficas de ingresos/gastos y de en qué se va el dinero |
| MES A MES | Ingresos, gastos, beneficio, margen y beneficio acumulado de cada mes |
| GASTOS POR CATEGORÍA | 22 categorías en 5 grupos (el local, producto para vender, trabajo y equipo, impuestos y autónomo, servicios y otros), con gráfica |
| HISTÓRICO | Los años comparados desde 2023 |
| GASTOS DETALLE | Cada gasto apuntado en la app |

**Histórico de años anteriores:** los meses de antes de usar la app salen del Excel antiguo. Se cargan una sola vez en **Ajustes → Histórico de años anteriores → Cargar histórico…** con el archivo `historico.json` (se guarda en Firebase, en `config/historico`; no está en este repositorio porque son cifras del negocio). El archivo dice desde qué mes cuenta la app (`"desde": "2026-11"`): antes de ese mes, histórico; desde ese mes, lo apuntado en la app.

## Archivos (súbelos todos a la raíz del repositorio)

`index.html`, `lector.js`, `excel_anual.js`, `plantilla_mes.xlsx`, `catalogo.json`, `database.rules.json`, `manifest.json`, `icon-180.png`, `icon-512.png`, `README.md` y la carpeta `vendor/` (contiene `exceljs.min.js`, licencia MIT)

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

## Productos vendidos en la hoja del talonario
Al dar de alta, revisar (foto) o editar una hoja hay una sección **Productos vendidos**:
- **+ Producto**: busca por nombre, código de barras o referencia, o pulsa **📷 Escanear** (el mismo escáner de Stock).
- Cada línea: producto, unidades (1 por defecto), precio por unidad **con IVA** (el PVP por defecto, se puede cambiar) y ✕ para quitarla. Si añades dos veces el mismo producto, se suman las unidades.
- **Cremas** y **Colonias** se rellenan solas con la suma de las líneas: lo de **Yodeyma** va a Colonias y el resto a Cremas (mientras haya líneas, esos dos campos no se escriben a mano). Sin líneas, se escriben a mano como siempre.
- Cremas + colonias no pueden superar el total de la hoja.
- Para que una colonia cuente como colonia, tiene que tener el proveedor **Yodeyma** (las que se dan de alta solas desde la factura ya lo tienen).
- Si vendes más unidades de las que hay en stock (el stock quedaría en negativo), avisa al guardar: «X tiene 0 en stock; ¿guardar igualmente?». Lo mismo con **− Venta** y **− Cabina** en la ficha del producto.

Al guardar:
- La hoja guarda sus productos y, por cada línea, se registra una **venta** en el historial de stock con la referencia `Hoja Nº X`.
- **Editar** la hoja deshace sus ventas anteriores y aplica las nuevas (solo se mueve la diferencia). **Eliminar** la hoja devuelve las unidades al stock.
- Esas ventas **no se deshacen desde el Historial**: al tocarlas se abre la hoja para corregirla allí.
- **− Venta** en la ficha del producto sigue disponible para ventas sueltas.

Con foto y IA: si la IA lee las líneas de la hoja, propone el producto del catálogo más parecido como **sugerencia** («¿DUO COMFORT SKIN?») con botón **Añadir**. Nada se descuenta sin que lo confirmes.

El Excel de Declarando no cambia: sigue usando el total de cada hoja y las cifras de cremas y colonias.

## Facturas de Yodeyma: colonias en stock
Al subir el PDF de una factura de Yodeyma, la pantalla de revisión ofrece **Dar entrada en stock** (igual que con Germaine):
- Cada colonia de la factura que no exista se da de alta sola: proveedor Yodeyma, uso Venta, coste = precio de la factura sin IVA y PVP = coste × factor PVP (Ajustes). Si ya existe, suma unidades y actualiza el coste.
- Las facturas de Yodeyma no traen código de barras ni referencia: la colonia se reconoce por su descripción («EAU TOILET. POUR FEMME 100 ML»). Por eso el stock va **por tipo y tamaño**, no por fragancia.
- Las líneas a 0 € (probadores, expositores, kits promocionales) no entran en stock. Un producto regalado (precio con descuento total) entra como bonificado.
- **Separar por fragancia:** en Stock, abre la colonia → **Separar por fragancia** → escribe la fragancia y cuántas unidades son. Se crea una ficha para esa fragancia (mismo coste y PVP) y las unidades pasan a ella; si la fragancia ya existe, se suman. Queda en el historial como «Separación» y se puede deshacer el mismo día (se deshacen las dos partes a la vez).
- Para usar el escáner con una colonia: en Stock, abre la colonia → Editar datos del producto → escribe su código de barras.

## Facturas en PDF de proveedores
- Se leen en el móvil sin IA: proveedor, NIF, nº, fecha, base, IVA, retención y total. Siempre hay pantalla de revisión antes de guardar.
- Reconoce de serie a los proveedores habituales (empresas): Germaine de Capuccini, Yodeyma, The Beauty Corner (Postquam), Factoría Cosmética, Estética Directo, Yosemite (láser), O2, la luz, AXA, SumUp, Declarando, Ubeira, Alcampo, Mercadona, Leroy Merlin, IKEA y otros. Los particulares (por ejemplo, el casero) se aprenden al guardar su primera factura con «Recordar este proveedor».
- Abonos (facturas rectificativas) en negativo; facturas con suplidos sin IVA; importes con 4 decimales (O2).
- Cuando la factura trae «TOTAL FACTURA» y «TOTAL A PAGAR» distintos (Germaine con bonificaciones o pagos a cuenta), se toma el total de la factura (base + IVA) y se avisa del importe a pagar.
- Justificantes del banco (ING): los **recibos de autónomos** y los **pagos de modelos 115/130/303** se registran como gasto sin factura con su importe, fecha y periodo. Los **adeudos y transferencias** que pagan una factura, los extractos y las declaraciones presentadas muestran un aviso ⚠ porque no son facturas (guardarlos contaría el gasto dos veces).
- Probado con las 511 facturas en PDF del chat de facturas (2022-2026): en las 413 facturas, el total sale en el 94 % y base + IVA cuadran en el 92 %; los 19 recibos de autónomos y los 4 pagos de modelos salen completos.

## Consejos para las fotos del talonario
- Una hoja por foto, de frente y con buena luz (la foto oscura o movida falla).
- Que se vean el número rojo, la fecha y el total.
- Siempre aparece una pantalla de revisión antes de guardar: comprueba el total.
- Se pueden seleccionar varias fotos a la vez desde la galería.

## Copia de seguridad
- **Ajustes → Copia de seguridad → Descargar copia**: un archivo `COPIA_BELLEZA_LUA_AAAA-MM-DD.json` con ingresos, gastos, stock, movimientos, proveedores y ajustes (sin la clave de IA). En el iPhone: Guardar en Archivos / iCloud.
- Hazla una vez al mes, después de sacar el Excel. Si pasan más de 30 días sin copia, la pestaña Ingresos avisa.
- El plan gratuito de Firebase no hace copias automáticas: este archivo es la única copia.

### Restaurar una copia
**Ajustes → Copia de seguridad → Restaurar una copia…** → elige el archivo `COPIA_BELLEZA_LUA_….json`.
1. La app enseña de qué fecha es la copia y cuántas hojas, gastos y productos tiene frente a los de ahora.
2. Al confirmar, **primero descarga una copia de lo que hay ahora** (`ANTES_DE_RESTAURAR_BELLEZA_LUA_….json`), por si te equivocas de archivo.
3. Después sustituye todos los datos (hojas, gastos, stock, movimientos, proveedores y ajustes) por los de la copia. Lo apuntado después de la fecha de la copia se pierde.

La clave de IA y el histórico se conservan si la copia no los trae. Los usuarios no se tocan nunca.

⚠ Si restauras a mano desde la consola de Firebase (Importar JSON), hazlo nodo a nodo y **nunca en la raíz** (`/`): borraría el nodo `usuarios` y nadie podría entrar.

## Numeración de Declarando
- La app guarda qué números ha usado cada mes exportado. Si al exportar un mes los números se pisan con otro mes ya exportado (p. ej. al volver a sacar un mes con una hoja olvidada), avisa antes de generar y propone el siguiente número libre.

## Notas
- Las devoluciones (abonos de Germaine) restan stock y quedan en el historial como «Devolución».
- Las facturas en PDF con texto se leen en el propio móvil, sin IA y sin enviar nada fuera.
- Las fotos se envían a Google solo si hay clave de IA. Con usuarios del EEE, Google aplica a los datos el trato de los servicios de pago (no se usan para entrenar); compruébalo en las condiciones de la API de Gemini.
- La app no guarda las fotos: conserva el talonario y las facturas originales.
- Los proveedores no reconocidos de serie se aprenden al guardar la primera factura (Ajustes → Proveedores aprendidos).
