# Sistema de Comparacion SIAF

Aplicacion local en Python y Streamlit para comparar archivos Excel de Devengados, Girados y Pagados, identificar pendientes y generar reportes profesionales.

## Ejecucion local

```bash
streamlit run app.py
```

## Archivos requeridos

La aplicacion solicita tres archivos Excel:

- Devengados.xlsx
- Girados.xlsx
- Pagados.xlsx

## Configurar secrets.toml local

Crear el archivo `.streamlit/secrets.toml` con este formato:

```toml
APP_USER = "usuario"
APP_PASSWORD = "contrasena"
```

No publiques este archivo ni lo subas a repositorios.

## Streamlit Cloud

En Streamlit Cloud, configurar las variables desde la seccion Secrets del proyecto:

```toml
APP_USER = "usuario"
APP_PASSWORD = "contrasena"
```

No incluyas credenciales reales en documentacion publica.
