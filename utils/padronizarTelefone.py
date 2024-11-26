import re

def formatar_telefone(telefone, codigo_pais='+55'):
    """Formata o número de telefone, removendo caracteres especiais e adicionando o código do país."""
    if telefone:
        telefone_formatado = re.sub(r'\D', '', telefone)
        if len(telefone_formatado) == 11 or len(telefone_formatado) == 12:
            return f'{codigo_pais}{telefone_formatado}'
        else:
            raise ValueError(f"Número de telefone inválido: {telefone}")
    return None