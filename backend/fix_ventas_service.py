content = open('app/modules/ventas/service.py', 'r', encoding='utf-8').read()
content = content.replace(
    'costo_total = linea.cantidad * p["costo_unitario"]',
    'costo_total = float(linea.cantidad) * float(p["costo_unitario"])'
).replace(
    'margen = subtotal_neto - costo_total',
    'margen = round(subtotal_neto - costo_total, 2)'
).replace(
    'if p["maneja_stock"] and p["stock_actual"] < linea.cantidad:',
    'if p["maneja_stock"] and float(p["stock_actual"]) < float(linea.cantidad):'
).replace(
    '"stock_resultante": linea["stock_antes"] - linea["cantidad"],',
    '"stock_resultante": float(linea["stock_antes"]) - float(linea["cantidad"]),'
).replace(
    '"stock_antes": p["stock_actual"],',
    '"stock_antes": float(p["stock_actual"]),'
)

with open('app/modules/ventas/service.py', 'w', encoding='utf-8') as f:
    f.write(content)
print('OK')