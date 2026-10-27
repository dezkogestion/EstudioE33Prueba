# Estudio E33 · Medellín

Landing de reservas del estudio de Entr3t3nimiento: configurador de sesión, agenda, galería, rider y solicitud de reserva. Bilingüe (ES/EN), tema claro y oscuro.

## Estructura

```
index.html              Página completa (HTML + CSS + JS, sin dependencias)
assets/img/             Fotos de la galería
assets/img/thumbs/      Miniaturas de los formatos del inicio
apps-script/Codigo.gs   Script de la hoja "E33 Estudio — Control landing"
netlify.toml            Configuración de Netlify
```

## Previsualizar con el equipo (Netlify)

1. En Netlify: **Add new site → Import an existing project → GitHub** y elige este repositorio.
2. Build command: vacío. Publish directory: `.`
3. En **Site configuration → Build & deploy → Branches and deploy contexts**, activa **Branch deploys** para `develop`.
4. Trabaja en `develop`. Cada push genera un link de vista previa del tipo `develop--<sitio>.netlify.app` para compartir con el equipo.
5. Cuando esté aprobado, haz merge de `develop` a `main` para publicar la versión final.

> El link de vista previa es público para quien lo tenga. No lo compartas fuera del equipo hasta publicar.

## Conectar la hoja de Google (datos en vivo)

1. Abre la hoja **E33 Estudio — Control landing** → Extensiones → Apps Script.
2. Pega el contenido de `apps-script/Codigo.gs` y ejecuta `instalar()` (autoriza los permisos).
3. Implementar → Nueva implementación → **Aplicación web** · Ejecutar como: tú · Acceso: **Cualquier usuario**. Copia la URL.
4. En `index.html`, busca `const CONFIG` y completa:
   - `API_URL`: la URL del paso 3
   - `WHATSAPP`: solo números, ej. `573001234567`
   - `EMAIL`: correo de reservas
5. Haz commit en `develop` y revisa la vista previa.

Sin `API_URL`, la página funciona con los datos de ejemplo incluidos en el HTML y la agenda se muestra como "Agenda de ejemplo".

## Qué se edita desde la hoja

| Pestaña | Qué controla |
| --- | --- |
| Config | TRM (automática), recargo nocturno, WhatsApp, correos de aviso, calendario, carpeta de fotos, capacidad |
| Tarifas | Formatos y precios con y sin ingeniero (ES/EN) |
| Produccion | Mezcla, máster y precio combo |
| Equipos | Canales de la consola (solo estado Nuevo u Operativo) |
| Rider | Rider técnico público |
| Media | Fotos y videos de la galería (se llena con "Sincronizar fotos y videos") |
| Solicitudes | Cada solicitud recibida; el estado actualiza el calendario |

Los cambios aparecen en la página en unos 5 minutos (caché del script). Para verlos al instante: menú **E33 Estudio → Limpiar caché de la landing**.

## App de gestión (Santi, Isis, Pepe)

Vista previa con datos de ejemplo: `gestion/index.html` (en Pages: `/gestion/`).

La app real corre en el mismo Apps Script de la hoja y usa las pestañas **Inventario, Compras, Ingresos, Usuarios y Solicitudes**.

1. En el proyecto de Apps Script agrega dos archivos: **Admin.gs** (Archivo → Nuevo → Script) y **Admin** (Archivo → Nuevo → HTML; nombre exacto `Admin`). Pega `apps-script/Admin.gs` y `apps-script/Admin.html`. Reemplaza también `Codigo.gs` por la versión de esta carpeta.
2. En la pestaña **Usuarios** completa el correo de Google de cada persona y su rol:
   - `admin`: todo, incluido aprobar compras (Pepe)
   - `operaciones`: reservas, inventario, compras e ingresos (Isis)
   - `local`: ver todo, marcar sesiones realizadas, estado de equipos, solicitar compras (Santi)
3. Comparte la hoja y el calendario "Reserva de Estudio" con esos correos (como editores).
4. Implementar → **Nueva implementación** → Aplicación web · Ejecutar como: **Usuario que accede** · Acceso: **Cualquier usuario con cuenta de Google**.
5. Abre la URL de esa implementación agregando `?app=admin` al final. Guárdala como acceso directo en el celular.

La implementación de la landing (Ejecutar como: yo · Acceso: cualquier usuario) sigue funcionando igual; son dos implementaciones del mismo proyecto.
