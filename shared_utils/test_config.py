# shared_utils/test_config.py
# Проверка, что config.py импортируется без ошибок

from shared_utils.config import PRINTOFFICE_USER, PRINTOFFICE_PASS, FIN_TABLO_PATH

print("PRINTOFFICE_USER =", PRINTOFFICE_USER)
print("PRINTOFFICE_PASS =", PRINTOFFICE_PASS)
print("FIN_TABLO_PATH =", FIN_TABLO_PATH)

