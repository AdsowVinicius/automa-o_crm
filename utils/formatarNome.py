def formatar_nome(nome):
    """Formata o nome para ter a primeira letra maiúscula de cada palavra."""
    return " ".join(p.capitalize() for p in nome.strip().split()) if nome else "Cliente"



