from shared_utils.utils import parse_amount, extract_deal_id


def test_parse_amount():
    assert parse_amount('1 234,56 руб') == 1234.56
    assert parse_amount('') == 0.0


def test_extract_deal_id():
    assert extract_deal_id('Сделка №8426') == '8426'
    assert extract_deal_id('Без номера') == ''
