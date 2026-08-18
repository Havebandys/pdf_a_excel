# Extractor Bancario IA - Beta educativa

Aplicación web Streamlit para extraer movimientos de PDF de BTF, Banco Patagonia, BBVA y Banco Comafi.
Procesa un documento por vez y genera archivos normalizados para Excel, Power BI y Google Sheets.

**Autoría:** Profesor Carlos Freyberguer - Universidad del Litoral - Cátedra Emprendedurismo (IA).

Beta de uso exclusivamente educativo. Prohibido su uso comercial.

## Uso local

1. Ejecutar `pip install -r requirements.txt`.
2. Copiar `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml` y cambiar la clave si corresponde.
3. Ejecutar `streamlit run app.py`.
4. Abrir la dirección indicada por Streamlit, cargar un PDF y descargar el resultado.

## Uso con Docker

```bash
docker build -t extractos-excel .
docker run --rm -p 8501:8501 extractos-excel
```

Abrir `http://localhost:8501`.

## Publicación en Streamlit Community Cloud

1. Crear un repositorio privado en GitHub.
2. Subir el contenido de esta carpeta a la raíz del repositorio.
3. Conectar GitHub con Streamlit Community Cloud.
4. Crear una aplicación seleccionando el repositorio y `app.py` como archivo principal.
5. En **Advanced settings > Secrets**, agregar `ACCESS_CODE = "1234"`.
6. Elegir la versión de Python y desplegar.

`requirements.txt` instala todas las dependencias. No requiere Poppler, Homebrew ni ejecutables externos.
No subir `.streamlit/secrets.toml`; la clave debe permanecer en los secretos privados de Streamlit.

## Clave privada

La clave inicial es `1234`. Para cambiarla sin modificar el código, configurar `ACCESS_CODE` en los
secretos privados del servidor o la variable de entorno `EXTRACTOR_ACCESS_CODE`. No se debe publicar
un archivo `secrets.toml` real en GitHub.

## Privacidad

El archivo subido se mantiene durante el procesamiento de la sesión y no se guarda en una base de datos.
La aplicación genera Excel y CSV en memoria para su descarga local. El servidor utilizado para desplegar
Streamlit también debe configurarse sin logs de contenido ni copias persistentes de los archivos.

## Salida

- `Movimientos`: datos numéricos y filtrables.
- `Control`: cantidades y totales.
- `Revisar`: líneas que requieren control manual.

La primera versión está calibrada con los cuatro extractos modelo recibidos. Antes de usarla como
único respaldo contable se deben comparar los totales mensuales contra el resumen de cada banco.
