# cnpj_formatter.py

def format_cnpj(cnpj_str):
    """
    Formata uma string de CNPJ para o padrão X.XXX.XXX/YYYY-ZZ
    
    Args:
        cnpj_str (str): String contendo apenas números do CNPJ
        
    Returns:
        str: CNPJ formatado no padrão X.XXX.XXX/YYYY-ZZ ou string original se inválida
    """
    if not cnpj_str:
        return ""
    
    # Remove todos os caracteres não numéricos
    cnpj_numbers = ''.join(filter(str.isdigit, cnpj_str))
    
    # Verifica se tem 14 dígitos
    if len(cnpj_numbers) != 14:
        return cnpj_str  # Retorna original se não tiver 14 dígitos
    
    # Aplica a formatação X.XXX.XXX/YYYY-ZZ
    formatted = f"{cnpj_numbers[0]}.{cnpj_numbers[1:4]}.{cnpj_numbers[4:7]}/{cnpj_numbers[7:11]}-{cnpj_numbers[11:13]}"
    
    return formatted

def validate_cnpj_format(cnpj_str):
    """
    Valida se o CNPJ está no formato correto X.XXX.XXX/YYYY-ZZ
    
    Args:
        cnpj_str (str): String do CNPJ a ser validada
        
    Returns:
        bool: True se está no formato correto, False caso contrário
    """
    import re
    
    # Padrão regex para X.XXX.XXX/YYYY-ZZ
    pattern = r'^\d\.\d{3}\.\d{3}\/\d{4}-\d{2}$'
    
    return bool(re.match(pattern, cnpj_str))

def clean_cnpj(cnpj_str):
    """
    Remove formatação do CNPJ, deixando apenas números
    
    Args:
        cnpj_str (str): CNPJ formatado ou não
        
    Returns:
        str: Apenas os números do CNPJ
    """
    if not cnpj_str:
        return ""
    
    return ''.join(filter(str.isdigit, cnpj_str))

