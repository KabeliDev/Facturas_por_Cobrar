# Facturas_Por_Cobrar

Robot Rocketbot que reúne las facturas electrónicas pendientes de pago de **Aula Editores**. Las saca del ERP Manager, descarga sus PDF y XML desde el servidor de documentos tributarios electrónicos (DTE) y envía a cobranza un Excel con el detalle de cada factura junto con los PDF.

> Documentación basada en `bot_facturas_por_Cobrar_Aula.db` (análisis del 2026-11-13).
> Análisis técnico y de riesgos: [`analisis.md`](analisis.md). Flujo paso a paso: [`flujos/Facturas_Por_Cobrar_completo.md`](flujos/Facturas_Por_Cobrar_completo.md).

---

## Qué hace

1. Consulta en la base de Manager las facturas electrónicas (`e-FAV`) en estado **Pendiente** emitidas en los **últimos 12 días**.
2. Les suma las facturas que quedaron pendientes en el Excel de control del mes, porque en ejecuciones anteriores no se encontró su PDF.
3. Descarta las que el Excel de control ya registra como enviadas.
4. Descarga el PDF de cada factura desde el servidor DTE y la anota en el Excel de control con la fecha de envío.
5. Si alguna factura no tiene PDF, avisa por correo ("Facturas sin emitir").
6. Descarga el XML de cada factura encontrada y extrae los datos del cliente: razón social, RUT, folio, emisión, vencimiento, monto, orden de compra, contacto, teléfono, dirección y comuna.
7. Arma el Excel **"Reporte-facturas-por-cobrar-…-DD-MM-AAAA.xlsx"** y lo envía por correo a cobranza, con todos los PDF adjuntos.
8. Vacía las carpetas locales de PDF y XML.

Si no hay facturas que procesar, envía el correo **"Sin facturas por cobrar en Aula"** y termina.

## Robots que lo componen

```text
Facturas_Por_Cobrar                       (principal)
├── conexion_base_datos                   consulta SQL Server y filtra los últimos 12 días
├── Descargar_pdf_dte_AulaK               arma la lista, descarga PDF y actualiza el Excel de control
│   ├── enviar_correo                     aviso "Facturas sin emitir"
│   └── enviar_correo_sin_facturas        aviso "Sin facturas por cobrar"
├── Descargar_xml_dte_AulaK               descarga los XML de las facturas encontradas
├── Extraer_datos_XML                     lee los XML
│   └── ingresar_datos_excel              genera el Excel del reporte
└── enviar_reporte_facturas_por_cobrar    envía el reporte y limpia carpetas
```

La base también contiene `descarga_excel_email`, `enviar_reporte`, `extraer_datos_xml` y `verificar_id_email`, pero este flujo no los usa.

## Requisitos del equipo donde corre

| Recurso | Detalle |
|---|---|
| Rocketbot | Con los módulos `SQLServer_`, `sftp_`, `AdvancedExcel`, `readXML`, `emailAdvanced` y `System++`. |
| Python (el de Rocketbot) | `openpyxl` y `pandas`. El `locale` `en_US` debe estar disponible. |
| Red | Acceso a la base SQL Server de Manager (red interna), al servidor SFTP del ERP DTE y al servidor de correo SMTP (puerto 465, SSL). |
| Usuario de Windows | `Administrator`: las rutas locales están fijas bajo `C:/Users/Administrator/Desktop/AULA/`. |

## Archivos y carpetas

Todas las rutas cuelgan de `C:/Users/Administrator/Desktop/AULA/Facturas por Cobrar/`.

| Ruta | Uso | ¿Persiste? |
|---|---|---|
| `plantilla_reporte_f.xlsx` | Plantilla del Excel de control. Se copia el primer día de cada mes. | Sí, **no borrar** |
| `Reporte-Facturas-enviadas-AAAA-MM.xlsx` | **Excel de control mensual.** Desde la fila 6: A = n° factura, B = fecha de emisión, C = fecha de envío (vacía = pendiente). Es la única "memoria" del robot entre ejecuciones. | Sí, **no borrar ni editar a mano** |
| `reporte facturas por cobrar/planilla vacia.xlsx` | Plantilla del reporte que se envía. | Sí, **no borrar** |
| `pdf/` | PDF descargados y Excel del reporte del día. | Se vacía al final de cada ejecución |
| `xml/` | XML descargados. | Se vacía al final de cada ejecución |

## Configuración

El robot **no usa un archivo .ini**. Los parámetros se guardan en el **panel de variables** de cada robot, dentro de Rocketbot:

| Variable | Robot | Para qué sirve |
|---|---|---|
| `servidor`, `puerto`, `usuario`, `password` | Facturas_Por_Cobrar / Descargar_pdf_dte_AulaK | Conexión SFTP para los PDF |
| `directorio`, `carpeta_pdf_sftp` | Descargar_pdf_dte_AulaK | Ruta de los PDF en el servidor DTE |
| `ruta_carpeta`, `ruta_desc_pdf` | Facturas_Por_Cobrar / Descargar_pdf_dte_AulaK | Carpeta del Excel de control y carpeta de descarga de PDF |
| `usuario`, `password` | enviar_reporte_facturas_por_cobrar | Cuenta de correo que envía el reporte |
| `name` | ingresar_datos_excel / enviar_reporte_facturas_por_cobrar | Prefijo del nombre del Excel del reporte |
| `tipo` | Extraer_datos_XML | Texto fijo de la columna "tipo" ("Factura electrónica") |

Otros parámetros están **escritos directamente en los comandos**:

- conexión a SQL Server (`conexion_base_datos`);
- SFTP para los XML (`Descargar_xml_dte_AulaK`);
- cuenta de correo de los avisos (`enviar_correo`, `enviar_correo_sin_facturas`);
- destinatarios y textos de todos los correos.

Para cambiarlos hay que editar el comando en Rocketbot Studio.

> 🔒 Las credenciales no se documentan aquí. Están en el panel de variables y en los comandos del robot.

## Correos que envía

| Correo | Cuándo | Destinatario (rol) | Adjuntos |
|---|---|---|---|
| **AULA Facturas por cobrar DD-MM-AAAA** | Cuando hay facturas con PDF y XML procesados | Cobranza, con copia a finanzas de Aula y a soporte de Kabeli | Excel del reporte y todos los PDF de la carpeta `pdf/` |
| **AULA Facturas sin emitir DD/MM/AAAA** | Cuando no se encontró el PDF de alguna factura en el servidor DTE | Responsable del ERP, con copia a Aula (info, finanzas) y soporte de Kabeli | Ninguno; el cuerpo lista los números de factura |
| **Sin facturas por cobrar en Aula** | Cuando no hay facturas que procesar | Cobranza, con copia a Aula (info, finanzas) y soporte de Kabeli | Ninguno |

## Excel del reporte

Se genera desde `planilla vacia.xlsx`. Cada fila es una factura y se escribe desde la columna A, en este orden:

| A | B | C | D | E | F | G | H | I | J–L | M | N | O | P | Q |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Razón social | RUT | *(vacío)* | N° factura | Tipo | Emisión | Vencimiento | Monto | Orden de compra | *(vacías)* | Dirección | Comuna | Teléfono | Contacto | Email |

> La columna **email** sale siempre con el texto `falta agregar`. El robot no obtiene ese dato.

## Cómo se ejecuta

- Ejecutar el robot **`Facturas_Por_Cobrar`**. Los demás robots se llaman solos.
- La frecuencia y el horario de ejecución no están definidos en la base. Si hay una tarea programada (Rocketbot Orchestrator o el Programador de tareas de Windows), se configura fuera del robot.
- Solo se consideran las facturas emitidas en los **12 días anteriores** a cada ejecución, así que el robot debe correr al menos una vez dentro de ese plazo.

## Cómo saber si funcionó

El robot **no escribe un archivo de log**. Para revisar una ejecución:

1. Confirmar que llegó uno de los tres correos de la tabla anterior. **Si no llegó ninguno, la ejecución falló o el envío no salió.**
2. Abrir el Excel de control del mes: las facturas enviadas tienen fecha en la columna C, y las que siguen pendientes la tienen vacía.
3. Revisar la consola o el historial de ejecución de Rocketbot.
