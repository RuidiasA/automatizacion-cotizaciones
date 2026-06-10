def calcular_interpolacion(cantidad, margenes_dict):
    # Convertir las llaves a enteros y ordenarlas
    cantidades_ordenadas = sorted([int(k) for k in margenes_dict.keys()])
    
    # Caso 1: La cantidad es menor o igual al escalón más bajo
    if cantidad <= cantidades_ordenadas[0]:
        return margenes_dict[str(cantidades_ordenadas[0])]
        
    # Caso 2: La cantidad es mayor o igual al escalón más alto
    if cantidad >= cantidades_ordenadas[-1]:
        return margenes_dict[str(cantidades_ordenadas[-1])]
        
    # Caso 3: La cantidad está entre dos escalones (Interpolación Lineal)
    for i in range(len(cantidades_ordenadas) - 1):
        x1 = cantidades_ordenadas[i]
        x2 = cantidades_ordenadas[i+1]
        
        if x1 <= cantidad <= x2:
            y1 = margenes_dict[str(x1)]
            y2 = margenes_dict[str(x2)]
            
            # Fórmula matemática: Y = Y1 + ((Y2 - Y1) / (X2 - X1)) * (X - X1)
            pendiente = (y2 - y1) / (x2 - x1)
            margen_interpolado = y1 + pendiente * (cantidad - x1)
            
            return round(margen_interpolado, 2)

# ==========================================
# PRUEBA DE CONTROL (Tus datos de Chaleco)
# ==========================================
margenes_chaleco = {
     "10": 72.0,
      "25": 52.0,
      "50": 46.0,
      "100": 37.0, #26% para 160 unidades aprox
      "500": 32.5,
      "1000": 25.5
}

cantidad_test = 160
resultado = calcular_interpolacion(cantidad_test, margenes_chaleco)

print(f"--- Verificación de Algoritmo ---")
print(f"Cantidad a cotizar: {cantidad_test}")
print(f"Margen obtenido por interpolación: {resultado}%")