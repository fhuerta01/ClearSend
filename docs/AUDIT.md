# Revisión de ClearSend — 18 de septiembre de 2026

## Alcance y estado

Repositorio de origen: `fhuerta01/ClearSend`, revisión inicial `6d36600`. Copia creada en `/Users/fhuerta/Proyectos/ClearSend`, rama `codex/privacy-quality-analytics`.

Revisión de los dos puntos de entrada (panel y cinta), transformaciones de destinatarios, lectura/escritura Office.js, preferencias, exportaciones, HTML/CSS, manifiestos, compilación, dependencias, CI y documentación pública. Los cambios están preparados localmente; no se ha modificado el despliegue público ni una base de datos de producción.

No constituye una certificación de ausencia de vulnerabilidades. Las pruebas de Outlook usan simulaciones; queda pendiente verificar los clientes reales y la configuración efectiva del alojamiento.

## Tráfico del repositorio

La API autenticada de GitHub devolvió, para el periodo disponible del 4 al 17 de septiembre de 2026:

- 3 visitas, 1 visitante único.
- 3 clones, 3 clonadores únicos.

Son métricas del repositorio y no permiten saber cuántas personas usan el complemento ni cuántas veces pulsan sus acciones. La clonación realizada durante esta revisión puede afectar a futuras estadísticas.

## Hallazgos y correcciones

| Prioridad | Hallazgo | Corrección |
| --- | --- | --- |
| Alta | Promesas de «los datos nunca salen del dispositivo» incompatibles con Microsoft RoamingSettings y la carga automática de analítica. | Política, README, instrucciones, comentarios e interfaz distinguen procesamiento local, sincronización de Microsoft, exportación y contadores opcionales. |
| Alta | Validación y comportamiento diferentes entre panel, atajos y Quick clean; algunas rutas eliminaban direcciones inválidas. | Motor compartido. La validación bloquea antes de aplicar cambios; revisar formatos no verifica entregabilidad. |
| Alta | Escrituras múltiples podían dejar campos parcialmente actualizados; límite incorrecto de destinatarios. | Comprobación previa de 100 por campo modificado, detección de cambios desde la lectura, escritura secuencial y recuperación con errores explícitos. |
| Alta | Interpolación de direcciones y dominios en HTML en rutas de la interfaz. | Valores de buzón renderizados como texto/valor de formulario; retirada de HTML dinámico y manejadores inline. |
| Alta | CSV sin escape ni protección contra fórmulas. | Escape de comillas/delimitadores/saltos y neutralización de prefijos de fórmulas. |
| Alta | 82 avisos iniciales de dependencias: 8 críticos, 41 altos, 24 moderados y 9 bajos. | Actualización de dependencias y herramientas, retirada de paquetes sin uso y override de `adm-zip`. Auditoría final: 0 avisos conocidos. Los avisos iniciales incluyen herramientas de desarrollo; no equivalen a 82 vulnerabilidades explotables del complemento. |
| Media | Scripts incluidos dos veces, botones con manejadores duplicados y bloqueo de procesamiento ligado al identificador equivocado. | Inclusión única del compilador, un manejador por acción y exclusión de operaciones concurrentes. |
| Media | Orden configurable no aplicado consistentemente y Undo desactivado tras ordenar. | Secuencia compartida con validación previa; Undo disponible tras cualquier cambio real del panel, incluida la reordenación. |
| Media | Eventos de destinatarios registrados sobre objetos equivocados. | Registro en el elemento de Outlook y alternativa de consulta periódica si no está disponible. |
| Media | Análisis de duplicados por texto completo, sensible al nombre mostrado. | Comparación por dirección normalizada; el resumen cuenta apariciones adicionales. |
| Media | Configuración sin normalización, dominios de ejemplo tratados como datos y listas guardadas sin límite. | Normalización de preferencias, dominios vacíos seguros, límites de almacenamiento y eliminación explícita de entradas. |
| Media | Manifiestos ofrecían modificación en modo lectura, enlaces de soporte inexistentes y atajos globales no implementados. | Activación solo al redactar, enlace de soporte correcto y descripción real de atajos dentro del panel. |
| Media | La documentación de analítica describía archivos inexistentes, contadores efímeros y garantías de privacidad incorrectas. | Una guía vigente con implementación Supabase real, eventos cerrados, activación explícita y límites de exactitud. |
| Media | La configuración de CI ignoraba fallos de lint y usaba versiones antiguas de Node. | Comprobaciones obligatorias, Node 22/24, pruebas de base de datos y auditoría de dependencias. |
| Media | La comprobación visual de esta rama detectó una referencia CSS rota tras la compilación. | Copia estable de CSS y comprobación de existencia de todos los recursos HTML generados. |

## Diseño de analítica

Se ha sustituido Vercel Web Analytics por un endpoint de Vercel que solo acepta nombres de acción predeterminados. Supabase recibe ese nombre y almacena un contador agregado por día UTC y acción. No almacena destinatarios, dominios, identificadores, IP, user-agent, hashes de personas, sesiones ni filas de eventos individuales.

La preferencia está activada por defecto cuando no existe una preferencia guardada y el despliegue de producción está configurado. Las preferencias desactivadas ya guardadas se conservan, también al restaurar ajustes. Se informa en el panel y se puede desactivar en Configuración; no se presenta como consentimiento. DNT/GPC, desarrollo local y previews impiden el envío. No hay reintentos que puedan duplicar una entrega incierta. Los incrementos son atómicos, pero los contadores no equivalen al total exacto de uso ni a usuarios únicos: opt-outs, fallos de red y tráfico automatizado pueden sesgar las cifras.

La aplicación no registra metadatos de conexión ni los copia a Supabase; Vercel y otros proveedores sí reciben metadatos técnicos al atender HTTP. Por ello no se afirma anonimato absoluto de toda la infraestructura. Tampoco se afirma cumplimiento legal automático.

## Evidencia de validación

- `npm ci` completado en Node 22.23.1, con instalación normal de dependencias.
- `npm run check`: lint sin advertencias, 27 pruebas aprobadas, compilación y comprobación de recursos generados, ambos manifiestos válidos.
- `npm audit`: 0 vulnerabilidades conocidas, incluyendo dependencias de desarrollo, a fecha de revisión.
- PostgreSQL 16 en contenedor aislado: migración ejecutada; permisos públicos restringidos, incremento, rechazo de eventos desconocidos y retención verificados.
- 24 incrementos concurrentes adicionales: los 24 persistieron sin pérdidas.
- Interfaz compilada con Office simulado y destinatarios ficticios: procesamiento, Undo y configuración comprobados a 360 y 320 píxeles; sin errores/advertencias de consola en la vista revisada.
- `git diff --check` sin errores.

## Pendiente antes de publicar

1. Probar Outlook web y los clientes Windows/macOS objetivo; comprobar CSP, APIs Office.js, resolución de destinatarios y comportamientos de error. El validador XML no comprueba esos flujos.
2. Aplicar la migración en un proyecto Supabase dedicado y configurar los secretos/orígenes en Vercel. No se han creado servicios ni provisionado credenciales de producción.
3. Revisar registros, región, copias de seguridad, retención durante inactividad y límites contra abuso del endpoint. Un Origin permitido no autentica a un script malicioso.
4. Publicar la rama revisada y confirmar con eventos sintéticos que el despliegue y la política coinciden. Las instrucciones están en `ANALYTICS_README.md` y `docs/RELEASE_CHECKLIST.md`.

La garantía principal queda así: **ClearSend no envía datos de Outlook a sus servicios de alojamiento o analítica; el procesamiento permanece en Outlook y Microsoft puede sincronizar las preferencias y las direcciones inválidas guardadas dentro del entorno del usuario.**
