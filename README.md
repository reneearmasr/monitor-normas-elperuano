# Monitor de normas – El Peruano

Revisa las normas legales publicadas en <https://diariooficial.elperuano.pe/Normas>,
selecciona las que tratan de **la ATU / transporte urbano de Lima y Callao**,
**normatividad presupuestal** e **inversiones** (incluye contrataciones públicas),
genera un resumen en Word y lo envía por Telegram.

## Uso

```
ejecutar_monitor.bat                 (normas de hoy)
ejecutar_monitor.bat --dias 7        (últimos 7 días)
ejecutar_monitor.bat --desde 01/10/2026 --hasta 22/10/2026
ejecutar_monitor.bat --no-enviar     (solo genera el Word)
```

Los Word quedan en la carpeta `resumenes`.

## Cómo decide qué normas incluir

- Lee la lista del portal (entidad, título y sumilla) y descarga el texto completo de cada norma.
- Busca palabras clave por tema (`docs/config/palabras_clave.json`, editable desde el panel).
- Si el texto completo menciona a la ATU (aunque la sumilla no lo diga, p. ej. en anexos de una
  transferencia del MEF), la norma se marca en rojo con el fragmento donde aparece.
- Se descartan designaciones, renuncias, encargos y viajes, salvo que sean de la ATU.

## Envío por Telegram

Envía el listado de normas (con enlaces) y el Word adjunto a su chat o grupo de Telegram.
Pasos para crear el bot y obtener el `chat_id` en `config.ini`. Si prefiere no dejar el token en el
archivo, guárdelo como variable de entorno de su usuario (una sola vez, en PowerShell):
`[Environment]::SetEnvironmentVariable("MONITOR_TELEGRAM_TOKEN", "123456789:AAH...", "User")`

## Panel web

<https://reneearmasr.github.io/monitor-normas-elperuano/> (carpeta `docs/`, GitHub Pages):

- **Normas**: historial de normas detectadas con buscador, filtros por tema, entidad y fechas,
  detalle (lo que dispone, montos, menciones a la ATU), enlaces y CSV.
- **Indicadores**: normas por semana y tema, entidades más frecuentes.
- **Palabras clave**: agregar o quitar términos (se guardan en `docs/config/palabras_clave.json`
  con un token de GitHub de permisos mínimos; instrucciones en la misma pestaña).

Cada ejecución con `--publicar` agrega las normas a `docs/data/normas.json` y copia el Word a
`docs/resumenes/`.

## Envío por correo institucional (Power Automate)

Cada ejecución publica una copia del último Word con nombre fijo:
<https://reneearmasr.github.io/monitor-normas-elperuano/resumenes/Normas_ElPeruano_ultimo.docx>

Un flujo programado de Power Automate (cuenta @atu.gob.pe, 08:30 hora de Lima) lo descarga a OneDrive
con "Cargar un archivo desde una dirección URL", obtiene su contenido y lo envía adjunto con
"Enviar un correo electrónico (V2)" de Office 365 Outlook. Ninguna clave institucional sale de Microsoft 365.

## Ejecución diaria en GitHub Actions (recomendado)

El workflow `.github/workflows/monitor.yml` corre todos los días a las 07:30 (hora de Lima) con
`--diario` (normas de hoy + extraordinarias de ayer) y envía por Telegram. No depende de la PC ni
del firewall de la oficina. Usa los secrets del repositorio `TELEGRAM_BOT_TOKEN` y `TELEGRAM_CHAT_ID`.

- Ejecutar a mano / otro rango: pestaña **Actions → Monitor normas El Peruano → Run workflow**
  (campo de opciones, p. ej. `--dias 7`).
- El Word de cada ejecución queda también como artefacto de la ejecución (30 días).

## Ejecución diaria automática en la PC

Programador de tareas de Windows → Crear tarea básica → Diariamente (p. ej. 9:00) →
Iniciar un programa → `ejecutar_monitor.bat` (iniciar en esta carpeta).
