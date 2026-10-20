# Stock por pallets

App web (Flask) que recibe el Excel "STOCK POR VENTAS" y devuelve el resumen de pallets
mono / multiproducto por rangos de cajas y kg, con descarga en Excel.

## Local
    pip install -r requirements.txt
    python app.py

## Render
1. Subir este repositorio a GitHub.
2. En Render: New > Blueprint (usa `render.yaml`) o New > Web Service con:
   - Build: `pip install -r requirements.txt`
   - Start: `gunicorn app:app --timeout 180 --workers 1`
3. Variable de entorno `APP_PASSWORD` para proteger el acceso (usuario: cualquiera, contraseña: la definida).
