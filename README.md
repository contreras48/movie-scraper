
# Trabajo Final

Este es el trabajo final del Curso de Business Intelligence y Big Data. El proyecto consiste en realizar web scraping a sitios web de cines regionales para obtener información sobre las funciones de cada película del día actual. El objetivo es recopilar, procesar y presentar estos datos de manera eficiente utilizando herramientas de Python.

## Tabla de Contenidos

- [Descripción](#descripción)
- [Requisitos Previos](#requisitos-previos)
- [Instalación](#instalación)
- [Uso](#uso)
- [Características](#características)
- [Autores](#autores)

## Descripción

El proyecto  tiene como finalidad aplicar técnicas de web scraping para extraer información sobre las funciones de cine del día actual desde sitios web de cines regionales. Para ello, se han utilizado herramientas como Selenium para la automatización de la navegación, Pandas para la manipulación de datos, y Openpyxl para el manejo de archivos Excel.

## Requisitos Previos

Para este proyecto se requiere tener las siguientes dependencias:

- Python 3.x
- Selenium
- Pandas
- Openpyxl
- Un navegador web (Edge y Firefox) y el respectivo WebDriver

## Instalación

1. Clona el repositorio:

    ```bash
    git clone https://github.com/contreras48/movie-scraper.git
    cd movie-scraper
    ```

2. Instala las dependencias necesarias:

    ```bash
    pip install -r requirements.txt
    ```

3. Configura el WebDriver:

    - Descarga el WebDriver para el navegador ([Edge](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/?form=MA13LH) y [FireFox](https://github.com/mozilla/geckodriver/releases)).
    - Asegúrate de que el WebDriver esté en el `PATH` o especifica su ubicación en el script.

## Uso

Para ejecutar el script de web scraping, utiliza el siguiente comando:

### Linux
```bash
./run.sh
```
### Windows
Para ejecutarlo en Windows se debe ejecutar el archivo run.bat haciendo doble click sobre el


Esto iniciará el proceso de scraping y almacenará la información recopilada en un archivo Excel.

## Características

- **Web Scraping Automatizado**: Extracción de datos de funciones de cine del día actual.
- **Manejo de Datos**: Procesamiento de la información obtenida usando Pandas.
- **Exportación a Excel**: Los datos se almacenan en un archivo Excel para su análisis y presentación.

## Autores

- **Daniel Ernesto Nerio Hernández**
- **Javier Antonio Rivas Avalos**
- **Mateo Abrahan Figueroa Contreras**
- **José Miguel Chávez Hernández**
- **William Alfredo Cubas Monroy**
