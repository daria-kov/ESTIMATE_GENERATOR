from decimal import Decimal


def apply_price_data(obj, price_data):
    """Применение цен к Resource / AbstractResource"""
    if not price_data:
        return

    if 'base_price' in price_data:
        obj.base_price = Decimal(str(price_data['base_price']))

    if 'index' in price_data:
        obj.index = Decimal(str(price_data['index']))

    if 'current_price' in price_data:
        obj.current_price = Decimal(str(price_data['current_price']))
    elif obj.base_price and obj.index:
        obj.current_price = obj.base_price * obj.index
