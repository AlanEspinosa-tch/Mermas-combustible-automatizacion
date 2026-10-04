# Automatización de consolidación y monitoreo de mermas

Herramienta desarrollada en Python para consolidar reportes volumétricos, integrar información de compras y generar indicadores para el seguimiento de diferencias de inventario.

## Objetivo

Reducir el trabajo manual asociado a la integración de archivos provenientes de diferentes fuentes y facilitar la revisión de diferencias entre inventarios teóricos y reales.

## Diseño analítico

La lógica del proyecto contempla:

- estandarización de archivos de entrada;
- integración de compras y movimientos volumétricos;
- cálculo de balances de inventario;
- cálculo de diferencias en litros y porcentaje;
- semaforización para facilitar la revisión de casos que requieren atención.

Ejemplo de balance:

```
Saldo final teórico =
Saldo inicial + Compras - Ventas + Ajustes
```

## Mi contribución

Definí el problema, las reglas de cálculo, las validaciones y los criterios utilizados para interpretar las diferencias.

La implementación y optimización de parte del código Python se realizó con asistencia de herramientas de Inteligencia Artificial.

## Tecnologías

Python · Pandas · Jupyter / Google Colab · Excel

> Los archivos operativos y resultados reales de la empresa no forman parte de la versión pública del repositorio. Para una demostración pública se deben utilizar datos sintéticos o anonimizados.
