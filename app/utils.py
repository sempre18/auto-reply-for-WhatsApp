import re
import pandas as pd


REAL_COLUMNS = {
    "nome": "Historico",
    "documento": "Documento",
    "vencimento": "Vencimento",
    "valor": "Vl.Documento",
    "telefone": "Telefone Cliente",
}


def normalize_column_name(col: str) -> str:
    return str(col).strip().lower()


def validate_columns(df: pd.DataFrame):
    df_cols = {normalize_column_name(c): c for c in df.columns}
    missing = []

    for _, real_col in REAL_COLUMNS.items():
        if normalize_column_name(real_col) not in df_cols:
            missing.append(real_col)

    return missing


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renomeia as colunas reais da planilha para nomes internos do sistema.
    """
    rename_map = {}
    df_cols = {normalize_column_name(c): c for c in df.columns}

    for internal_name, real_name in REAL_COLUMNS.items():
        col_found = df_cols.get(normalize_column_name(real_name))
        if col_found:
            rename_map[col_found] = internal_name

    df = df.rename(columns=rename_map)
    return df


def normalize_phone(phone: str) -> str:
    if pd.isna(phone):
        return ""

    phone = str(phone).strip()

    # pega todos os grupos numéricos
    digits = re.sub(r"\D", "", phone)

    if not digits:
        return ""

    # remove zeros iniciais
    digits = digits.lstrip("0")

    # 10 ou 11 dígitos nacionais -> adiciona 55
    if len(digits) in (10, 11):
        digits = "55" + digits

    # se tiver mais de 13, tenta usar os últimos 13 ou 12 com 55
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


def parse_date(value):
    if pd.isna(value):
        return None
    try:
        return pd.to_datetime(value, dayfirst=True, errors="coerce")
    except Exception:
        return None


def format_date_br(value) -> str:
    dt = parse_date(value)
    if dt is None or pd.isna(dt):
        return ""
    return dt.strftime("%d/%m/%Y")


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
        return "R$ 0,00"


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = map_columns(df).copy()

    # garante colunas internas
    for col in ["nome", "documento", "vencimento", "valor", "telefone"]:
        if col not in df.columns:
            df[col] = ""

    df["nome"] = df["nome"].fillna("").astype(str).str.strip()
    df["documento"] = df["documento"].fillna("").astype(str).str.strip()
    df["telefone"] = df["telefone"].apply(normalize_phone)

    df["vencimento_dt"] = df["vencimento"].apply(parse_date)
    df["vencimento_fmt"] = df["vencimento"].apply(format_date_br)

    df["valor_num"] = df["valor"].apply(parse_money)
    df["valor_fmt"] = df["valor_num"].apply(format_money_br)

    df["telefone_valido"] = df["telefone"].apply(is_valid_phone)

    return df


def filter_dataframe(df: pd.DataFrame, mode="todos", dias=3, only_valid_phone=False) -> pd.DataFrame:
    result = df.copy()
    today = pd.Timestamp.now().normalize()

    if mode == "vencidos":
        result = result[result["vencimento_dt"].notna()]
        result = result[result["vencimento_dt"].dt.normalize() < today]

    elif mode == "a_vencer":
        limit = today + pd.Timedelta(days=int(dias))
        result = result[result["vencimento_dt"].notna()]
        result = result[
            (result["vencimento_dt"].dt.normalize() >= today) &
            (result["vencimento_dt"].dt.normalize() <= limit)
        ]

    if only_valid_phone:
        result = result[result["telefone_valido"] == True]

    return result.reset_index(drop=True)


def render_template(template: str, row: dict) -> str:
    values = {
        "nome": row.get("nome", ""),
        "documento": row.get("documento", ""),
        "vencimento": row.get("vencimento_fmt", ""),
        "valor": row.get("valor_fmt", ""),
        "telefone": row.get("telefone", ""),
    }

    text = template
    for key, value in values.items():
        text = text.replace("{" + key + "}", str(value))
    return text


def generate_messages(df: pd.DataFrame, template: str):
    messages = []

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        msg = render_template(template, row_dict)
        messages.append({
            "nome": row_dict.get("nome", ""),
            "documento": row_dict.get("documento", ""),
            "vencimento": row_dict.get("vencimento_fmt", ""),
            "valor": row_dict.get("valor_fmt", ""),
            "telefone": row_dict.get("telefone", ""),
            "telefone_valido": row_dict.get("telefone_valido", False),
            "mensagem": msg
        })

    return messages