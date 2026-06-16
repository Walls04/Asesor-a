# Proyecto: AsesorIA — Copiloto de análisis de carteras para asesores financieros

## Qué es
Herramienta web B2B para asesores financieros, agentes y EAFs en España.
El asesor sube un Excel con las posiciones de la cartera de un cliente y la app
genera un informe profesional en PDF con análisis de costes, riesgos y eficiencia.

## Qué NO es (crítico — restricción legal)
- NO da recomendaciones de inversión ni asesoramiento al cliente final.
- Es una herramienta de productividad PARA el asesor (él firma y decide).
- Todo output debe incluir disclaimer: "Documento de uso interno para el asesor.
  No constituye recomendación de inversión."

## Estado del proyecto
✅ **MVP COMPLETADO el 11/06/2026** (las 4 fases probadas por Gonzalo en el navegador):
1. ✅ COMPLETADA — Resumen de cartera (subida de Excel, valor total, posiciones, gráfico de tarta).
2. ✅ COMPLETADA — Análisis de costes (TER ponderado, coste anual, proyección 10 años, coste de oportunidad).
3. ✅ COMPLETADA — Banderas rojas (TER > 1,5%, concentración > 40%, solapamiento) con contador.
4. ✅ COMPLETADA — Informe PDF descargable (reportlab; portada, secciones, disclaimer en cada página).

Todo el MVP vive en un único archivo: `app.py`.
Siguientes pasos candidatos: `requirements.txt` + deploy en Streamlit Cloud; endurecer el parser con Excels imperfectos.

## Entorno de desarrollo
- Windows 11 con **PowerShell** (no usar sintaxis bash/Linux: ni heredocs, ni `/dev/null`, ni `export`).
- Para probar código Python: crear un archivo `.py` temporal y ejecutarlo con `python archivo.py`, no one-liners complejos.
- Python 3.13. Librerías instaladas: streamlit, pandas, openpyxl, matplotlib, reportlab.
- La app la arranca Gonzalo manualmente (`streamlit run app.py`); Claude no ejecuta streamlit.

## MVP (versión 1 — lo ÚNICO que construimos ahora)
Flujo: subir Excel → analizar → descargar informe PDF.

### Input esperado (Excel del asesor)
Columnas: ISIN, Nombre del fondo, Categoría, Divisa, Nº Participaciones,
Valor Liquidativo, Valor de Mercado, TER (%), Clase.
Ver `cartera_ejemplo_cliente.xlsx` como referencia. El parser debe ser tolerante:
nombres de columna aproximados (fuzzy match), valores de mercado calculables si faltan.

### Análisis que genera (v1)
1. **Resumen de cartera**: valor total, nº posiciones, distribución por categoría (gráfico de tarta).
2. **Análisis de costes**: TER medio ponderado de la cartera, coste anual en €,
   coste proyectado a 10 años, y comparación con un TER de referencia del 0,40%
   (cartera indexada). Mostrar el "coste de oportunidad" en € a 10 años.
3. **Banderas rojas**:
   - Fondos con TER > 1,5% → marcar como "clase cara, revisar si existe clase limpia".
   - Concentración: cualquier categoría > 40% del total.
   - Solapamiento: 2+ fondos en la misma categoría (ej. dos fondos de tecnología).
4. **Informe PDF**: portada con nombre del cliente y fecha, las secciones anteriores,
   diseño sobrio estilo banca privada (azul oscuro #1F3864, tipografía limpia),
   disclaimer al pie de cada página.

### Stack (mantener simple, el desarrollador no es programador)
- Python + Streamlit (una sola app, sin frontend separado, deploy gratis en Streamlit Cloud).
- pandas + openpyxl para leer el Excel.
- matplotlib para gráficos.
- reportlab o weasyprint para el PDF.
- SIN base de datos en v1. SIN login en v1. Nada se guarda (ventaja RGPD: stateless).

### Fuera de alcance en v1 (NO construir todavía)
- Login/usuarios, pagos, comparador de fondos alternativos, lectura de PDFs bancarios,
  conexión a APIs de datos de fondos (Morningstar etc.), multi-divisa.

## Reglas de trabajo para Claude Code
- Explica cada paso en español y de forma simple antes de ejecutarlo.
- Una funcionalidad cada vez. Probar antes de pasar a la siguiente.
- Código comentado en español.
- Tras cada cambio, indicar el comando exacto para probar la app (`streamlit run app.py`).

## Roadmap (después del MVP, por orden)
v2: lectura de extractos PDF de bancos españoles (Santander, CaixaBank, BBVA).
v3: sugerencia de clases limpias / alternativas indexadas por categoría.
v4: login + planes de pago (Stripe), informe con logo del asesor (white-label).
