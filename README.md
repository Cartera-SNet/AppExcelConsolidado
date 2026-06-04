# Resumen de cartera

Aplicación local en Flask para cargar varios archivos Excel `.xlsx`, generar un resumen en el mismo formato de la tabla y validar si existe `Fecha de radicación` en 2026.

## Requisitos

- Python 3.10 o superior
- pip

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
python app.py
```

Luego abre en el navegador:

```text
http://127.0.0.1:5000
```

## Qué hace

- Detecta la fila de encabezados por la columna `Fecha de radicación`
- Toma la `Razón social del prestador`
- Cuenta facturas por `Número de reclamo`
- Suma `Valor neto de la reclamación`
- Suma `Valor saldo de la reclamación`
- Genera un Excel con el mismo formato del resumen
- Indica si hay fechas de radicación en 2026

## Supuesto principal

Los Excel deben venir con una estructura parecida a la de tus archivos actuales, incluyendo estas columnas:

- `Razón social del prestador`
- `Número de reclamo`
- `Fecha de radicación`
- `Valor neto de la reclamación`
- `Valor saldo de la reclamación`
