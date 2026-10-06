# Viandas Fiesta Familiar — cómo instalar la nueva versión

## 1. Backend (Google Apps Script)

1. Abre el proyecto de Apps Script que ya tienes vinculado a tu Google Sheet.
2. Borra el contenido del archivo `.gs` actual y pega todo el contenido de **Codigo.gs**.
3. Guarda (ícono de disquete o Ctrl+S).
4. Ve a **Implementar → Administrar implementaciones → ✏️ Editar → Nueva versión → Implementar**.
   - Importante: usa "Nueva versión", no crear una implementación nueva, así la URL
     `.../exec` que ya tienes sigue siendo la misma y no hay que cambiarla en el frontend.
5. El backend crea solo las hojas que le falten (`Usuarios`, `Pedidos`) la primera vez que se usan.
   Tu hoja `Productos` ya existente se sigue usando tal cual (columnas: Producto, Precio, Categoría, Estado).

## 2. Frontend (las páginas)

Archivos incluidos:
- `bienvenida.html` — pantalla de entrada (nombre + forma de pago)
- `menu.html` — pantalla principal de pedidos
- `usuario.html` — cambiar de usuario / forma de pago
- `productos.html` — agregar y editar productos y precios
- `resumen.html` — ventas netas por producto, precio, usuario y forma de pago
- `cambios.html` — registrar cambios, devoluciones y diferencias
- `historial.html` — consultar las últimas 15 operaciones
- `reportes.html` — pantalla de reportes
- `styles.css` y `app-common.js` — estilos y lógica que comparten todas las páginas

**Sube todos los archivos HTML, CSS y JS juntos, en la misma carpeta**, al lugar donde hoy tienes tu `index.html`
(por ejemplo GitHub Pages, Firebase Hosting, o cualquier hosting de archivos estáticos).
El punto de entrada para tus usuarios ahora es `bienvenida.html` (o puedes renombrarlo a
`index.html` si tu hosting espera ese nombre exacto).

### Instalar en el celular

La app incluye `manifest.webmanifest`, `sw.js` y los iconos PNG. Sube también estos archivos a la misma carpeta del hosting. La instalación requiere que el sitio esté publicado con HTTPS.

- Android: abre el sitio en Chrome, toca el menú ⋮ y elige **Instalar app** o **Agregar a pantalla principal**.
- iPhone: abre el sitio en Safari, toca **Compartir** y elige **Agregar a pantalla de inicio**.

Al abrirla desde el icono instalado, se muestra como app y no como pestaña del navegador. La interfaz puede abrirse desde caché, pero los pedidos y cambios requieren conexión para comunicarse con Apps Script.

La URL de tu Apps Script ya está puesta dentro de `app-common.js`, en la constante `SCRIPT_URL`.
Si alguna vez vuelves a implementar el backend como una implementación *nueva* (no como
nueva versión), esa URL cambia y hay que actualizarla ahí.

## 3. Cómo funciona la identificación por dispositivo

Cada navegador guarda el nombre y la forma de pago elegidos en `localStorage` (no se borra
al cerrar la pestaña, solo si el usuario limpia datos del navegador o toca "Cerrar sesión").
Por eso, la próxima vez que la misma persona abra la app desde el mismo celular, su nombre
ya aparece preseleccionado en la pantalla de bienvenida — pero igual debe confirmar
"Continuar al menú" antes de entrar; no se salta la pantalla.

## 4. Cambios, devoluciones e historial

Las ventas normales se guardan como movimientos `venta`. Los cambios y devoluciones agregan
filas nuevas en `Pedidos`: los productos devueltos usan cantidades y subtotales negativos,
los productos nuevos usan valores positivos y una ofrenda se registra como movimiento separado.
Las filas antiguas sin tipo se interpretan como ventas. El stock de los productos devueltos
solo se repone cuando el cajero lo marca.
