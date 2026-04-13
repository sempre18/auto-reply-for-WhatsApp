import re
import pandas as pd


def normalize_column_name(col: str) -> str:
    return str(col).strip().lower()


def normalize_phone(phone: str) -> str:
    if pd.isna(phone):
        return ""

    phone = str(phone).strip()
    digits = re.sub(r"\D", "", phone)

    if not digits:
        return ""

    digits = digits.lstrip("0")

    # Se vier com 10 ou 11 dígitos nacionais, adiciona 55
    if len(digits) in (10, 11):
        digits = "55" + digits

    # Se tiver muito número e contiver 55, tenta aproveitar a parte correta
    if len(digits) > 13 and "55" in digits:
        pos = digits.find("55")
        digits = digits[pos:]
        if len(digits) > 13:
            digits = digits[:13]

    if len(digits) in (12, 13) and digits.startswith("55"):
        return digits

    return ""


def is_valid_phone(phone: str) -> bool:
    digits = re.sub(r"\D", "", str(phone))
    return len(digits) in (12, 13) and digits.startswith("55")


def parse_money(value):
    if pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    text = text.replace("R$", "").replace(" ", "")

    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except Exception:
        return 0.0


def format_money_br(value) -> str:
    try:
        number = float(value)
        return f"R$ {number:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def clean_dataframe(df: pd.DataFrame, phone_column: str) -> pd.DataFrame:
    """
    Mantém todas as colunas originais.
    Cria:
    - telefone_original
    - telefone
    - telefone_valido
    """

    df = df.copy()

    if phone_column not in df.columns:
        raise ValueError(f"A coluna de telefone '{phone_column}' não existe na planilha.")

    df["telefone_original"] = df[phone_column].fillna("").astype(str).str.strip()
    df["telefone"] = df[phone_column].apply(normalize_phone)
    df["telefone_valido"] = df["telefone"].apply(is_valid_phone)

    return df


def render_template(template: str, row: dict) -> str:
    """
    Substitui {NomeDaColuna} pelo valor da coluna correspondente.
    Também aceita:
    - {telefone} => telefone normalizado
    - {telefone_original} => valor original da planilha
    """

    text = template

    # Substitui todas as colunas originais
    for key, value in row.items():
        placeholder = "{" + str(key) + "}"
        text = text.replace(placeholder, "" if pd.isna(value) else str(value))

    # Extras padronizados
    text = text.replace("{telefone}", str(row.get("telefone", "")))
    text = text.replace("{telefone_original}", str(row.get("telefone_original", "")))

    return text


def generate_messages(df: pd.DataFrame, template: str):
    messages = []

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        msg = render_template(template, row_dict)

        messages.append({
            "nome": row_dict.get("Historico", "") or row_dict.get("nome", ""),
            "documento": row_dict.get("Documento", "") or row_dict.get("documento", ""),
            "telefone": row_dict.get("telefone", ""),
            "telefone_original": row_dict.get("telefone_original", ""),
            "telefone_valido": row_dict.get("telefone_valido", False),
            "mensagem": msg,
            "row_data": row_dict,
        })

    return messages