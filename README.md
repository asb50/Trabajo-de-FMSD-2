# Sistema de Gestión y Consulta de Datos Meteorológicos (AEMET & SQLite)

Proyecto para la asignatura **Fundamentos Matemáticos de los Sistemas de Datos**. El script integra llamadas a la API de **AEMET OpenData** y descargas remotas desde UCA Drive para construir, alimentar y consultar una base de datos relacional en SQLite (`datos/meteo.db`) con información geográfica y meteorológica de España.

## Características Principales

- **Gestión de bases de datos SQLite (`datos/meteo.db`):** Creación y estructuración de tablas relacionales con claves primarias, foráneas y borrado en cascada.
- **Integración con AEMET OpenData API:**
  - Descarga de estaciones meteorológicas y municipios.
  - Obtención de datos observados de las últimas 12 horas.
  - Extracción de datos climatológicos diarios por rango de fechas (con validación de intervalos de máx. 15 días).
- **Consumo de Fuentes Auxiliares (UCA Drive):** Descarga y extracción en memoria (vía `zipfile` e `io`) de ficheros Excel y paquetes ZIP con históricos de datos (`datosHora`, `datosDias`).
- **Cálculo de Distancias Geodésicas:** Conversión de coordenadas sexagesimales (GMS) a radianes decimales para calcular la distancia en kilómetros entre municipios y estaciones meteorológicas.
- **Interfaz Interactiva:** Menús en consola para la navegación entre funciones y consultas personalizadas por municipio y rango de distancia.

## Requisitos Previos e Instalación

### Requisitos
* **Python 3.8** o superior.
* Una **API Key de AEMET OpenData** (solicítala gratuitamente en el [Centro de Descargas de AEMET](https://opendata.aemet.es/centrodedescargas/inicio)).
