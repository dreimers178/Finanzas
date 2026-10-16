# Zulú · Catálogo de relojes

Catálogo web de relojes para el proyecto **Zulú** (relojes desde la aviación).
Sitio estático, sin servidor ni base de datos: la web lee los relojes desde `data.js`.

## Estructura

```
relojes/
├── index.html   → el catálogo (diseño + lógica). No se edita para el día a día.
├── data.js      → LOS DATOS: config + lista de relojes. Este es el que se edita.
└── README.md    → este archivo.
```

## Cómo mantenerlo

Todo se hace en **`data.js`**.

### Cambiar datos de contacto
En el bloque `CONFIG`, al inicio:
- `whatsapp` → tu número en formato internacional, sin `+` ni espacios (MX: `521` + 10 dígitos).
- `instagram` → tu usuario (sin la @).
- `agente` → tu nombre.

### Agregar un reloj
Copia un bloque `{ ... }` dentro de `RELOJES` y cámbialo. Campos:

| Campo            | Qué es                                                        |
|------------------|---------------------------------------------------------------|
| `id`             | identificador único, sin espacios (ej. `"rolex-sub-2023"`)    |
| `marca`          | marca                                                         |
| `modelo`         | modelo                                                        |
| `referencia`     | número de referencia                                          |
| `anio`           | año (número)                                                  |
| `caja_mm`        | diámetro de caja en mm (número)                               |
| `material`       | material de la caja                                           |
| `precio`         | precio en número, sin símbolos (ej. `142000`)                 |
| `estado`         | condición: `"Nuevo"`, `"Como nuevo"`, `"Excelente"`, etc.     |
| `incluye`        | qué incluye (ej. caja y papeles)                              |
| `disponibilidad` | `"disponible"` · `"apartado"` · `"vendido"`                   |
| `destacado`      | `true` / `false`                                              |
| `nota_aviacion`  | dato breve que conecta la pieza con la aviación               |
| `descripcion`    | descripción completa                                          |
| `fotos`          | lista de rutas/URLs de fotos; vacía `[]` muestra un marcador  |

### Marcar como vendido
Cambia `disponibilidad: "disponible"` → `disponibilidad: "vendido"`.
El reloj se queda visible (sirve como prueba social), atenuado y con botón
"Preguntar por similares".

### Fotos
Por ahora las piezas muestran un marcador "foto pendiente". Cuando haya fotos,
se suben a la carpeta (ej. `relojes/fotos/`) y se referencian en `fotos`, p. ej.
`fotos: ["fotos/navitimer-1.jpg", "fotos/navitimer-2.jpg"]`.

## Ver el sitio
Es HTML estático: se abre `index.html` en el navegador. Para publicarlo en
internet con dominio propio, cualquier hosting estático gratis sirve
(GitHub Pages, Netlify, Vercel) — pendiente de definir.

## Nota
Este proyecto vive junto a **FinPro** (`/index.html` en la raíz), que es una app
distinta y no se toca.
