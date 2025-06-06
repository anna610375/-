# shared_utils/config.py
# Настройки доступа читаются из переменных окружения
import os

PRINTOFFICE_USER = os.environ.get("PRINTOFFICE_USER", "")
PRINTOFFICE_PASS = os.environ.get("PRINTOFFICE_PASS", "")
FIN_TABLO_PATH = os.environ.get("FIN_TABLO_PATH", "C:\\Users\\anna6\\путь\\к\\FinTablo.xlsx")
