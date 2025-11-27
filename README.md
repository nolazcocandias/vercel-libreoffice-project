# vercel-libreoffice-project

API en Python para Vercel que actualiza simulacion.xlsx y devuelve datos.

## Endpoint
POST https://<tu-dominio-vercel>/api/calcular

### Body
{
  "cantidad_pallets": 120,
  "meses_operacion": 12
}

### Respuesta
{
  "tarjetas": {"pallet_parking": 123, "tradicional": 456, "ahorro": 789},
  "tabla": [{"mes": 1, "in": 10, "out": 7, "stock": 3}],
  "costos": {"pallet_parking": [...], "tradicional": [...]} 
}
