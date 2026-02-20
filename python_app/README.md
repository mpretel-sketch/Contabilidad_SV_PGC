# Proyecto 100% Python (Streamlit + SQLite)

## Ejecutar

```bash
cd /Users/miguelpretelpozo/Contabilidad_SV_PGC/python_app
python3 -m pip install -r requirements.txt
streamlit run app.py
```

## Funcionalidades

- Carga mensual de archivo (`xlsx/xls/csv`) y analisis completo.
- Edicion de partidas al maximo detalle (sumas y saldos).
- Mapeo manual de lineas sin equivalencia PGC.
- Filtro por estado (`sin mapear`, `mapeadas`, `sumatorias`).
- Balance, P&G colapsable y total del periodo visible.
- Guardado en SQLite por mes/anio (sobrescribe periodo existente).
- Bloqueo de guardado si hay sin mapear o balanza final no cuadra.
- Exportacion a XLSX.

## Base de datos

Se crea automaticamente en:

`/Users/miguelpretelpozo/Contabilidad_SV_PGC/python_app/data/contabilidad.db`
