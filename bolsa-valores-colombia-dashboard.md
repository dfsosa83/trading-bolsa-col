# Dashboard de Renta Fija — Bolsa de Valores de Colombia
### Resumen Ejecutivo para Stakeholders

---

## ¿Qué es?

Una plataforma de análisis en tiempo real para el escritorio de trading de renta fija, enfocada en los **Bonos de Deuda Pública Externa USD (BGLT)** negociados en la BVC. La herramienta transforma el Boletín Diario de la BVC —un archivo Excel técnico— en dashboards interactivos que el equipo puede consultar desde cualquier navegador, sin instalar software.

**Acceso:** https://trading-bolsa-col-zqt8cbw7elgkwdfc3blntx.streamlit.app

---

## Problema que Resuelve

Antes de este proyecto, el análisis diario del mercado de bonos BGLT requería:

- Descargar manualmente el boletín de la BVC cada día
- Abrir y filtrar el Excel buscando los bonos externos USD
- Cruzar información de múltiples hojas para entender quién compró y quién vendió
- Reconstruir a mano la tabla de cruces vendedor → comprador → bono
- Consultar por separado la TRM, las tasas de cierre y las posiciones de los fondos de pensiones

**Este proceso tomaba tiempo valioso del escritorio y estaba expuesto a errores humanos.**

---

## Solución Implementada

La plataforma automatiza el ciclo completo:

```
BVC publica el boletín  →  GitHub descarga automáticamente  →  Dashboard actualizado
        (diario ~8am)              (sin intervención humana)         (en tiempo real)
```

El equipo simplemente abre el link y toda la información está disponible, organizada en cinco módulos.

---

## Cinco Módulos del Dashboard

### 1. 📅 Vista Diaria
Resumen operativo de la sesión seleccionada: listado de bonos BGLT negociados con monto en COP y su equivalente en USD (calculado con la TRM oficial del día), tasas de cierre, y distribución por tipo de inversionista.

### 2. 📈 Posiciones Históricas
Evolución de las posiciones netas por sector (AFPs, Fiduciarias, Extranjeros, Pinas, etc.) a lo largo de todas las sesiones disponibles. Incluye gráfico de líneas, mapa de calor y posición acumulada en el período.

### 3. 📉 Curva de Tasas
Snapshot de la curva de rendimientos BGLT para cualquier fecha seleccionada (tasa vs. plazo al vencimiento), evolución histórica de tasas por bono, y tabla de últimas tasas registradas.

### 4. 🏦 Fondos de Pensiones
Tenencias de bonos BGLT de los cuatro fondos obligatorios (Porvenir, Colfondos, Protección, Skandia), con evolución mensual, composición por bono y distribución por tipo de fondo. Datos de la Superfinanciera.

### 5. 🔀 Triangulación de Cruces *(módulo diferencial)*
El módulo más sofisticado. Reconstruye automáticamente la tabla que el escritorio construía a mano: **quién le vendió a quién, y qué bono**. Ver sección siguiente.

---

## Triangulación: El Motor Analítico Central

### El desafío
El boletín BVC reporta por separado:
- *Por bono:* monto total negociado (pero no las contrapartes)
- *Por sector:* cuánto compró y vendió cada tipo de inversionista (pero no qué bonos)

Reconstruir el cruce completo requería razonamiento manual.

### La solución algorítmica

El sistema implementa tres niveles de análisis complementarios:

| Nivel | Método | Confianza | Cobertura |
|---|---|---|---|
| 🎯 Cruces exactos | Coincidencia de montos sector ↔ bono (±1.5%) | Alta | Parcial |
| 🧩 Atribución completa | Backtracking con restricción de suma cero | Alta / Determinística | 100% de bonos |
| 📊 Modelo proporcional | Matriz de flujos por entropía máxima | Estimada | 100% del volumen |

**Principio matemático clave:** el monto total de bonos negociados = total comprado por sectores = total vendido por sectores (suma cero). Esto garantiza que siempre existe una solución válida y el algoritmo la encuentra.

**Restricción BVC aplicada automáticamente:** Extranjeros → Extranjeros = 0 (las operaciones entre dos entidades extranjeras no se registran en BVC).

### Resultado: la misma tabla que construye el escritorio, en segundos

| Vendedor | Comprador | Bono | Monto (M COP) |
|---|---|---|---|
| Fdo. Pensiones y Cesantías | Extranjeros | BGLT12141135 | 71,844 |
| Fdo. Pensiones y Cesantías | Extranjeros | BGLT30071154 | 72,868 |
| Pna Colombiana | Fdo. Pensiones y Cesantías | BGLT30150645 | 53,114 |
| Fiduciarias | Extranjeros | BGLT32180141 *(1 de 2 op.)* | 2,852 |
| Extranjeros | Fiduciarias | BGLT32180141 *(2 de 2 op.)* | 2,851 |

El sistema incluso detecta automáticamente cuándo un bono fue transado en múltiples operaciones con distintas contrapartes, etiquetándolas como "1 de 2 op." / "2 de 2 op."

---

## Ventajas para el Escritorio de Trading

| Antes | Ahora |
|---|---|
| Proceso manual de 30–60 min por sesión | Disponible al abrir el link |
| Riesgo de error en el cruce de datos | Algoritmo auditado y reproducible |
| Sin visibilidad histórica inmediata | 14+ sesiones analizadas en segundos |
| Curva de tasas construida manualmente | Snapshot y evolución automáticos |
| Posiciones de pensiones consultadas por separado | Integradas en el mismo dashboard |
| Información atrapada en archivos locales | Accesible desde cualquier dispositivo |

**Sin fricción para el equipo:** no requiere instalación, no requiere credenciales, no hay archivos que compartir. El link siempre muestra el estado más reciente.

---

## Arquitectura y Seguridad

- **Infraestructura:** Streamlit Community Cloud (plataforma gestionada, sin servidores propios que mantener)
- **Código fuente:** repositorio privado en GitHub (`dfsosa83/trading-bolsa-col`)
- **Actualización automática:** GitHub Actions descarga el boletín diario a las 8:15am (hora Colombia), sin intervención humana
- **Datos:** toda la información proviene de fuentes públicas oficiales (BVC, Superfinanciera, datos.gov.co)
- **Sin escritura:** el dashboard es 100% de solo lectura; no modifica ni almacena datos del cliente

---

## Estado Actual y Próximos Pasos

**Operativo:** los cinco módulos están en producción y el escritorio los usa diariamente.

**En agenda:**
- Módulo de bonos comparables Colombia (pendiente lista de Roy Garuz)
- Análisis de correlación: flujos de trading ↔ tenencias de pensiones ↔ movimientos de tasas
- Alertas automáticas cuando se detectan cruces de alta confianza en sesiones nuevas

---

*Desarrollado para Pina Colombiana · Agosto 2026*
