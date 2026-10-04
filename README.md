# Automatización de Consolidación y Monitoreo de Mermas

Herramienta desarrollada en Python para consolidar automáticamente los reportes de volumen exportados en formato CSV por el software volumétrico, integrando las compras mensuales y generando tableros de control para el seguimiento de mermas de combustible.

## Objetivo del Proyecto
Eliminar la captura manual y los errores humanos en el registro de compras de combustible, centralizando la información dispersa en un único archivo de control automatizado que permita auditar las diferencias entre inventarios teóricos y reales.

## Diseño Analítico y Reglas de Negocio (Mi Contribución Principal)
Como Analista de Finanzas e Ingeniero Matemático, estructuré la lógica financiera y las validaciones de control:
*   **Estandarización de Fuentes:** Diseño del flujo para procesar y unificar archivos CSV heterogéneos provenientes de los sistemas volumétricos de las estaciones.
*   **Control de Cumplimiento Fiscal (SAT):** Implementación de la regla de validación de umbrales de merma (límite permitido por el SAT $\le$ 0.5% de la venta).
*   **Semaforización y Alertas:** Configuración de umbrales automáticos para detectar anomalías operativas:
    *   **Verde:** $\vert{}\text{merma}\vert{} \le 0.40\%$
    *   **Amarillo:** $0.40\% \text{ a } 0.50\%$
    *   **Rojo:** $> 0.50\%$ (Fuera de norma SAT)
*   **Modelado de Balance Volumétrico:** Cálculo de inventarios: 
    $$\text{Saldo Final Teórico} = \text{Saldo Inicial} + \text{Compras} - \text{Ventas} + \text{Ajustes}$$
    $$\text{Merma (L)} = \text{Saldo Real} - \text{Saldo Final Teórico}$$

##  Implementación Técnica
*Nota de transparencia: La definición de la lógica financiera, el diseño de las fórmulas de balance volumétrico y las reglas de control fueron de desarrolo propio. La optimización del código en Python (`Mermas_Colab.ipynb`) para la automatización del archivo Excel final fue desarrollado con asistencia de Inteligencia Artificial.*

*   **Lenguaje:** Python (Jupyter Notebook / Google Colab).
*   **Librerías principales:** `pandas` para limpieza y consolidación de estructuras de datos.
*   **Entregable:** Generación automatizada del reporte ejecutivo consolidado (`Reportes_Mermas_2026-08.xlsx`).
