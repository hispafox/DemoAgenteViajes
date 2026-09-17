---
name: planificar-viaje-equipaje
description: "Planifica viajes completos y verificables con especial atención a viajeros, equipaje, rutas, requisitos oficiales, presupuesto, enlaces y documentación final. Úsalo cuando el usuario pida planificar, presupuestar o documentar un viaje, itinerario, ruta o escapada."
argument-hint: "Indica origen, destino, fechas, viajeros, presupuesto, preferencias y formatos de salida."
user-invocable: true
disable-model-invocation: false
---

# Planificar viaje y equipaje

## Objetivo

Preparar una propuesta de viaje coherente, verificable y visualmente cuidada que cubra viajeros, equipaje, rutas, actividades, costes, condiciones del destino, enlaces útiles y documentación final.

## Capacidades y fuentes

- Usa navegación web para consultar fuentes públicas actuales.
- Usa fuentes oficiales y los datos facilitados por el usuario como fuentes principales.
- Usa las habilidades especializadas `docx`, `xlsx` y `pptx` cuando se soliciten esos formatos.
- Usa generación de imágenes únicamente para recursos visuales útiles y etiquetados; no añadas decoración sin valor informativo.

## Procedimiento

1. **Recoger requisitos.** Solicita origen, destino o destinos, fechas, duración, flexibilidad, viajeros, nacionalidad o residencia cuando afecte a trámites, actividades, presupuesto, moneda, zona preferida, alojamiento, alimentación, movilidad, ritmo, transporte y equipaje. No cierres la propuesta si falta un dato que cambie rutas, costes o trámites.
2. **Confirmar condiciones.** Consulta meteorología actual cuando las fechas estén dentro del horizonte disponible. Si no lo están, usa clima histórico y márcalo como orientativo. Incluye diferencia horaria, divisas, tipo de cambio con fecha, roaming y conectividad.
3. **Verificar requisitos oficiales.** Comprueba visados, documentación, vacunas y recomendaciones sanitarias según nacionalidad, residencia y destino. Distingue requisitos obligatorios de recomendaciones e indica fuente y fecha.
4. **Investigar enlaces.** Añade enlaces HTTPS directos, vigentes y no duplicados para turismo oficial, transporte, aeropuertos, compañías, mapas, alojamientos, actividades, restaurantes, emergencias, consulados, entrada, sanidad, meteorología, divisas y roaming cuando sean relevantes. Usa títulos descriptivos, explica para qué sirve cada enlace y repite los recursos en una sección final.
5. **Diseñar rutas.** Para cada trayecto indica salida, llegada, medio, duración, distancia cuando proceda, transbordos, horarios o frecuencia disponible, margen recomendado, accesibilidad, equipaje, coste estimado y alternativa. En desplazamientos urbanos, incluye estaciones, líneas, sentido, número de paradas y último tramo a pie. Nunca inventes tarifas ni horarios: enlaza la fuente consultada.
6. **Representar rutas.** Incluye un esquema legible de cada ruta principal y, cuando sea posible, mapa o diagrama con origen, destino, conexiones y puntos de interés. Añade texto alternativo, pie, fuente y fecha. Si el recurso no es verificable, etiquétalo como diagrama orientativo.
7. **Planificar equipaje.** Compara las normas de las compañías consideradas: dimensiones, peso, cantidad, tarifa y posibles costes. Si no se conoce compañía o tarifa, presenta alternativas sin afirmar una franquicia concreta.
8. **Crear itinerario.** Organiza actividades por día y zona considerando horarios, desplazamientos, descansos y meteorología. Sugiere restauración con rango de precio, ubicación, cocina y adecuación a restricciones. Verifica disponibilidad antes de afirmar que un establecimiento está abierto.
9. **Calcular presupuesto.** Normaliza importes a la moneda solicitada y conserva moneda original y tipo de cambio. Separa confirmado, estimado y pendiente. Incluye transporte, alojamiento, actividades, comidas, equipaje, comunicaciones, seguros y contingencia. Comprueba subtotales y total.
10. **Seleccionar recursos visuales.** Añade imágenes relevantes del destino, alojamiento, actividades, gastronomía y transporte. Cada imagen debe tener resolución y proporción adecuadas, pie, atribución cuando corresponda y texto alternativo. No reutilices imágenes salvo que aporten contexto distinto.
11. **Generar Word.** Cuando se solicite, incluye portada, resumen ejecutivo, datos del viaje, itinerario diario, rutas, equipaje, trámites, presupuesto, recomendaciones, enlaces y fuentes. Usa tablas legibles, saltos de página, hipervínculos descriptivos y recursos visuales útiles.
12. **Generar Excel.** Cuando se solicite, crea hojas para resumen, presupuesto, itinerario, rutas, equipaje, enlaces y fuentes. Usa fórmulas para subtotales y total, formatos de moneda y fecha, filtros, anchos legibles, leyenda de estados e hipervínculos funcionales.
13. **Generar PowerPoint.** Cuando se solicite, crea una narrativa con portada, resumen, mapa general, ruta por etapas, experiencias, alojamiento, gastronomía, equipaje, presupuesto, trámites, enlaces y próximos pasos. Mantén una idea principal por diapositiva, poco texto, imágenes o diagramas útiles y enlaces clicables.
14. **Aplicar control editorial.** Revisa ortografía, estilos, contraste, paginación, tablas, recortes, resolución, enlaces, pies, textos alternativos y solapamientos.
15. **Validar coherencia.** Contrasta fechas, duraciones, rutas, actividades, importes, equipaje y enlaces entre Word, Excel y PowerPoint. Marca todos los datos no verificados o pendientes.
16. **Entregar.** Proporciona los archivos solicitados con nombres claros y un resumen de contenido, fuentes, fecha de consulta, supuestos, estimaciones, enlaces y puntos que deben verificarse antes de reservar.

## Estados obligatorios

Usa siempre estas etiquetas cuando correspondan:

- **Confirmado:** respaldado por una fuente vigente o por datos explícitos del usuario.
- **Estimado:** calculado u orientativo, sujeto a cambios.
- **Pendiente:** requiere dato del usuario o verificación posterior.

## Salida mínima

- Propuesta de viaje en texto.
- Itinerario diario con rutas, tiempos, conexiones y alternativas.
- Equipaje y requisitos oficiales aplicables.
- Presupuesto desglosado con estado, moneda original y moneda solicitada.
- Enlaces agrupados, contextualizados y verificables.
- Fuentes consultadas, fecha de consulta y puntos por verificar.
