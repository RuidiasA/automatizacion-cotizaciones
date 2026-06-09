class QuoteCalculator:
    def __init__(self, config_data):
        self.config_data = config_data

    def _interpolate_margin(self, qty, margin_scale):
        # Convertimos las llaves a enteros ordenados para evaluar los tramos
        sorted_quantities = sorted([int(q) for q in margin_scale.keys()])

        # CASO CRÍTICO: Si piden MÁS del máximo registrado en tu JSON
        if qty > sorted_quantities[-1]:
            # Retornamos el margen del tope máximo y una alerta de exceso de volumen
            return float(margin_scale[str(sorted_quantities[-1])]), "A REVISAR"

        # Caso Límite Inferior: Si piden menos o igual a la cantidad mínima
        if qty <= sorted_quantities[0]:
            return float(margin_scale[str(sorted_quantities[0])]), "OK"

        # Caso Límite Superior Exacto
        if qty == sorted_quantities[-1]:
            return float(margin_scale[str(sorted_quantities[-1])]), "OK"

        # Caso Intermedio estándar: Buscamos entre qué dos nodos se encuentra la cantidad
        for i in range(len(sorted_quantities) - 1):
            q_lower = sorted_quantities[i]
            q_upper = sorted_quantities[i+1]

            if q_lower <= qty <= q_upper:
                m_lower = float(margin_scale[str(q_lower)])
                m_upper = float(margin_scale[str(q_upper)])

                # Fórmula matemática de interpolación lineal
                fraction = (qty - q_lower) / (q_upper - q_lower)
                interpolated_margin = m_lower + fraction * (m_upper - m_lower)
                return interpolated_margin, "OK"

        return 0.0, "OK"

    def calculate_client_cost(self, product_key, quantity, provider_cost):
        """
        Calcula el precio final para el cliente y devuelve el estado de auditoría.
        """
        if product_key not in self.config_data:
            raise ValueError(
                f"El producto '{product_key}' no existe en las matrices de configuración.")

        margin_scale = self.config_data[product_key].get("margenes", {})

        # 1. Obtenemos el margen exacto y el estado de la alerta (OK o A REVISAR)
        margin_percentage, status = self._interpolate_margin(
            quantity, margin_scale)

        # 2. Convertimos el porcentaje a factor multiplicador
        margin_factor = 1.0 + (margin_percentage / 100.0)

        # 3. Aplicamos la fórmula directa sobre el costo base del proveedor
        client_cost_unit = margin_factor * float(provider_cost)

        # Retornamos el costo unitario redondeado, el porcentaje y el estado de control
        return round(client_cost_unit, 2), margin_percentage, status