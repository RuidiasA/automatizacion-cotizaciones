import re


class ProductMatcher:
    def __init__(self, config_data):
        self.config_data = config_data

    def _normalize_text(self, text):
        if not text or not isinstance(text, str):
            return ""
        # Convertimos a minúsculas y limpiamos saltos de línea/espacios extra
        text_normalized = text.lower().strip()
        # Quitamos tildes básicas para evitar fallos humanos comunes (á, é, í, ó, ú)
        replacements = (("á", "a"), ("é", "e"),
                        ("í", "i"), ("ó", "o"), ("ú", "u"))
        for target, replacement in replacements:
            text_normalized = text_normalized.replace(target, replacement)
        return text_normalized

    def find_matched_product(self, detail_text):
        normalized_detail = self._normalize_text(detail_text)

        # Recorremos cada producto definido en nuestro JSON
        for product_key, properties in self.config_data.items():
            prenda_tokens = properties.get("filtros_prenda", [])
            material_tokens = properties.get("filtros_material", [])

            # Verificamos si al menos uno de los sinónimos de prenda existe en el texto
            match_prenda = any(
                token in normalized_detail for token in prenda_tokens)

            # Verificamos si al menos uno de los sinónimos del material existe en el texto
            match_material = any(
                token in normalized_detail for token in material_tokens)

            # Si ambos criterios se cumplen en la celda, tenemos un MATCH exitoso
            if match_prenda and match_material:
                return product_key

        return None
