
from datetime import datetime

def padronizar_data(data):
    """Padroniza a data para o formato datetime.date."""
    if isinstance(data, datetime):
        return data.date()
    elif isinstance(data, str):
        data = data.strip()
        try:
            if len(data.split('/')) == 2 or len(data.split('-')) == 2:
                data_com_ano = f"{data}/{datetime.now().year}"
                return datetime.strptime(data_com_ano, "%d/%m/%Y").date()
            else:
                return datetime.strptime(data, "%d/%m/%Y").date()
        except ValueError:
            try:
                return datetime.strptime(data, "%d-%m-%Y").date()
            except ValueError:
                return None
    elif isinstance(data, (int, float)):
        try:
            return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(data) - 2).date()
        except Exception:
            return None
    return None



