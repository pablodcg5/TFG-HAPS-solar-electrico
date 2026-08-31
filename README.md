# TFG - Diseño conceptual de un HAPS solar-eléctrico

Este repositorio contiene el código y los archivos auxiliares empleados para el dimensionado preliminar de una plataforma UAV tipo HAPS solar-eléctrica.

El script principal permite reproducir los cálculos de:
- balance energético solar;
- dimensionado preliminar de baterías;
- distribución de placas solares;
- geometría preliminar de fuselaje, boom y cola;
- volumen interno del ala para integración de baterías;
- sistema propulsivo de referencia;
- dimensionado preliminar de hélices mediante teoría de cantidad de movimiento;
- generación de tablas de resultados para la memoria.

## Estructura del repositorio

- `codigo/`: script principal de cálculo.
- `datos/`: archivos de entrada, perfil SD7032 y polares XFLR5.
- `resultados/`: tablas CSV generadas por el programa.
- `openvsp/`: modelo geométrico preliminar del HAPS en OpenVSP.

## Ejecución

Instalar dependencias:

```bash
pip install -r requirements.txt
