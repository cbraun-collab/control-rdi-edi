# Control RDI / EDI — Senercom

PWA-COM-005 REV 0 (01-09-2026). Unifica FOR-COM-008 (RDI) y FOR-COM-009 (EDI) en una sola aplicación.

## Paso 1 completado
- Estructura PWA base (manifest, service worker, ícono en 3 tamaños)
- Identidad visual Senercom (teal #003B49, naranjo #CF4520, Space Grotesk + JetBrains Mono)
- Pantalla de acceso con clave de app (placeholder `senercom2026`, se reemplaza en el paso del backend)
- Navegación por pestañas: Bandeja / Pendientes de firma / Historial
- Nota de Ley 21.719 ya integrada

## Cómo publicar (GitHub Pages)
1. Crea un repositorio nuevo en la cuenta `cbraun-collab`, por ejemplo `control-rdi-edi`.
2. Sube todo el contenido de esta carpeta a la raíz del repo.
3. En Settings → Pages, activa GitHub Pages desde la rama `main` / carpeta raíz.
4. La app quedará en `https://cbraun-collab.github.io/control-rdi-edi/`.

## Pendiente (próximos pasos)
- Backend Apps Script (Sheet maestro, correlativo por proyecto/tipo, búsqueda de carpeta de obra)
- Formulario de creación (Sección A) con carga de adjuntos
- Envío con link de firma remota + página pública de firma (Sección B)
- Generación automática del PDF final y guardado en carpeta del proyecto
- Respaldo por monitoreo de correo
