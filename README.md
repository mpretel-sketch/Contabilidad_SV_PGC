# Conversor Contable SNC Portugal -> PGC Espana

Aplicacion full-stack para convertir balanzas CONTPAQi/SNC Portugal a estructura PGC 2007 (Espana), con:

- Persistencia en BBDD SQLite por periodo (mes/anio).
- Sobrescritura automatica al guardar si el periodo ya existe.
- Flujo mensual: subir archivo -> revisar/mapear/cuadrar -> guardar en BBDD.
- Mapeo automatico de cuentas por prefijos.
- Edicion manual de partidas al maximo detalle (SID/SIA/Cargos/Abonos/SFD/SFA).
- Override manual de mapeo PGC por linea (cuenta, nombre, grupo y subgrupo).
- Conversion EUR/EUR por tipo de cambio editable (por defecto 1.0).
- Balance y P&G agregados por grupo/subgrupo PGC.
- Validaciones (cuadre de balanza, descuadre de balance, cobertura de mapeo).
- Exportacion a Excel.

## Estructura

- `server`: API Express para conversion, carga de archivos y exportacion.
- `web`: Frontend React + Vite.

## Ejecutar en local

1. Instalar dependencias:

```bash
npm install
npm --prefix server install
npm --prefix web install
```

2. Levantar backend + frontend:

```bash
npm run dev
```

- Frontend: `http://localhost:5173`
- Backend: `http://localhost:4000`

## Formato esperado de archivo de entrada

Columnas equivalentes (admite variaciones de nombre):

- `code` o `codigo` o `cuenta`
- `name` o `nombre` o `descripcion`
- `sid` (saldo inicial deudor)
- `sia` (saldo inicial acreedor)
- `cargos` (debe)
- `abonos` (haber)
- `sfd` (saldo final deudor)
- `sfa` (saldo final acreedor)

## Endpoints principales

- `GET /api/health`
- `GET /api/periods`
- `GET /api/periods/:year/:month`
- `POST /api/periods/upload` (multipart, analiza sin persistir)
- `POST /api/periods/save` (persiste el periodo solo si todo esta mapeado y cuadra)
- `GET /api/sample`
- `GET /api/mapping/meta`
- `POST /api/convert` (JSON)
- `POST /api/export` (devuelve `.xlsx`)

## Nota tecnica

El parser/exportador usa la libreria `xlsx` (SheetJS community), que actualmente publica avisos de seguridad sin fix oficial. Para entornos de produccion con politicas estrictas, se recomienda migrar a una alternativa endurecida y acotar formatos de entrada.
