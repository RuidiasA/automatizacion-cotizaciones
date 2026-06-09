import sys
from src.config_loader import ConfigLoader
from src.excel_processor import ExcelProcessor


def main():
    print("==================================================")
    # 1. Cargamos las matrices de configuración desde el JSON
    try:
        config_loader = ConfigLoader()
        json_config = config_loader.load_config()
    except Exception as e:
        print(f"❌ Error crítico al cargar la configuración JSON: {str(e)}")
        sys.exit(1)

    # 2. Inicializamos el orquestador de archivos Excel
    # Por defecto buscará en 'data/input' y guardará en 'data/output'
    processor = ExcelProcessor(json_config=json_config)

    # 3. Lanzamos el proceso por lotes (Batch Processing)
    processor.process_all_excels()
    print("==================================================")


if __name__ == "__main__":
    main()
