import json
import random
import re
from typing import Optional

import pandas as pd


DEFAULT_ALIASES = {
    "nome": ["nome", "cliente", "historico", "razão social", "razao social"],
    "documento": ["documento", "doc", "nf", "nota", "fatura", "pedido"],
    "vencimento": ["vencimento", "dt vencimento", "data vencimento"],
    "valor": ["valor", "vl.documento", "vl documento", "vl_documento", "total"],
    "telefone": ["telefone", "telefone cliente", "celular", "whatsapp", "fone"],
}


def normalize_column_name(col: str) -> str:
    return str(col).strip().lower()


def normalize_key_name(name: str) -> str:
    return re.sub(r"\s+", "_", str(name).strip().lower())


def normalize_phone(phone) -> str:
    if pd.isna(phone):
        return ""
    digits = re.sub(r"\D", "", str(phone)).lstrip("0")
    if not digits:
        return ""
    if len(digits) in (10, 11):
        digits = "55" + digits
    if len(digits) > 13 and "55" in digits:
        pos = digits.find("55")
        digits = digits[pos:pos + 13]
    return digits if len(digits) in (12, 13) and digits.startswith("55") else ""


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
    return dt.strftime("%d/%m/%Y") if dt is not None and not pd.isna(dt) else ""


def parse_money(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("R$", "").replace(" ", "")
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
        n = float(value)
        return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def extract_template_variables(template_manager, template_id: Optional[str] = None) -> list[str]:
    if template_manager is None:
        return []
    return template_manager.get_template_variables(template_id)


def resolve_aliases(template_manager=None, template_id: Optional[str] = None) -> dict[str, list[str]]:
    aliases = {k: list(v) for k, v in DEFAULT_ALIASES.items()}
    if template_manager is not None:
        tpl_aliases = template_manager.get_template_aliases(template_id)
        for key, vals in tpl_aliases.items():
            aliases.setdefault(key, [])
            for val in vals:
                if val not in aliases[key]:
                    aliases[key].append(val)
    return aliases


def resolve_column_mapping(
    df: pd.DataFrame,
    required_vars: list[str],
    template_manager=None,
    template_id: Optional[str] = None,
) -> dict[str, str]:
    df_cols = {normalize_column_name(c): c for c in df.columns}
    aliases = resolve_aliases(template_manager, template_id)
    mapping: dict[str, str] = {}

    for var in required_vars:
        var_norm = normalize_column_name(var)

        if var_norm in df_cols:
            mapping[var] = df_cols[var_norm]
            continue

        for alias in aliases.get(var_norm, []):
            alias_norm = normalize_column_name(alias)
            if alias_norm in df_cols:
                mapping[var] = df_cols[alias_norm]
                break

    return mapping


def validate_template_columns(
    df: pd.DataFrame,
    required_vars: list[str],
    template_manager=None,
    template_id: Optional[str] = None,
) -> tuple[list[str], dict[str, str]]:
    mapping = resolve_column_mapping(df, required_vars, template_manager, template_id)
    missing = [var for var in required_vars if var not in mapping]
    return missing, mapping


def clean_dataframe_dynamic(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    result = df.copy()

    for var, original_col in mapping.items():
        result[var] = result[original_col]

    for col in mapping.keys():
        if col not in result.columns:
            result[col] = ""

    for col in result.columns:
        if result[col].dtype == object:
            result[col] = result[col].fillna("").astype(str).str.strip()

    if "telefone" in result.columns:
        result["telefone"] = result["telefone"].apply(normalize_phone)
        result["telefone_valido"] = result["telefone"].apply(is_valid_phone)
    else:
        result["telefone"] = ""
        result["telefone_valido"] = False

    if "vencimento" in result.columns:
        result["vencimento_dt"] = result["vencimento"].apply(parse_date)
        result["vencimento_fmt"] = result["vencimento"].apply(format_date_br)
    else:
        result["vencimento_dt"] = pd.NaT
        result["vencimento_fmt"] = ""

    if "valor" in result.columns:
        result["valor_num"] = result["valor"].apply(parse_money)
        result["valor_fmt"] = result["valor_num"].apply(format_money_br)
    else:
        result["valor_num"] = 0.0
        result["valor_fmt"] = ""

    return result


def filter_dataframe(
    df: pd.DataFrame,
    mode: str = "todos",
    dias: int = 3,
    only_valid_phone: bool = False,
) -> pd.DataFrame:
    result = df.copy()
    today = pd.Timestamp.now().normalize()

    if "vencimento_dt" in result.columns:
        if mode == "vencidos":
            result = result[result["vencimento_dt"].notna()]
            result = result[result["vencimento_dt"].dt.normalize() < today]

        elif mode == "a_vencer":
            limit = today + pd.Timedelta(days=int(dias))
            result = result[result["vencimento_dt"].notna()]
            result = result[
                (result["vencimento_dt"].dt.normalize() >= today)
                & (result["vencimento_dt"].dt.normalize() <= limit)
            ]

    if only_valid_phone and "telefone_valido" in result.columns:
        result = result[result["telefone_valido"]]

    return result.reset_index(drop=True)


def build_context(row: dict) -> dict:
    ctx = {}
    for key, value in row.items():
        if key.endswith("_dt") or key.endswith("_num"):
            continue
        if pd.isna(value):
            ctx[key] = ""
        else:
            ctx[key] = str(value)

    if "vencimento_fmt" in row:
        ctx["vencimento"] = row.get("vencimento_fmt", "")
    if "valor_fmt" in row:
        ctx["valor"] = row.get("valor_fmt", "")
    if "telefone" in row:
        ctx["telefone"] = row.get("telefone", "")

    return ctx


def has_unresolved_placeholders(text: str) -> list[str]:
    return re.findall(r"{\s*([a-zA-Z0-9_]+)\s*}", text or "")


def evaluate_row_preparation(row: dict, mensagem: str, required_vars: list[str]) -> tuple[str, list[str], list[str]]:
    missing_fields = []
    for var in required_vars:
        value = row.get(var, "")
        if value is None or str(value).strip() == "":
            missing_fields.append(var)

    placeholders_left = has_unresolved_placeholders(mensagem)

    if not row.get("telefone_valido", False):
        return "telefone_invalido", missing_fields, placeholders_left
    if missing_fields:
        return "campo_vazio", missing_fields, placeholders_left
    if placeholders_left:
        return "placeholder_pendente", missing_fields, placeholders_left
    return "pronto", missing_fields, placeholders_left


def generate_messages(
    df: pd.DataFrame,
    template_manager=None,
    template_id: Optional[str] = None,
    fallback_template: str = "",
) -> list[dict]:
    messages = []
    rows = df.to_dict("records")
    required_vars = extract_template_variables(template_manager, template_id)

    if len(rows) > 3:
        _soft_shuffle(rows, strength=0.25)

    for row in rows:
        ctx = build_context(row)

        if template_manager is not None:
            if template_id:
                msg, tid = template_manager.render(template_id, ctx)
            else:
                msg, tid = template_manager.render_random(ctx)
        else:
            msg = fallback_template
            for key, value in ctx.items():
                msg = msg.replace("{" + key + "}", str(value))
            tid = "manual"

        prep_status, missing_fields, placeholders_left = evaluate_row_preparation(
            row=row,
            mensagem=msg,
            required_vars=[v for v in required_vars if v != "telefone"],
        )

        messages.append(
            {
                "nome": row.get("nome", row.get("cliente", "")),
                "documento": row.get("documento", row.get("pedido", "")),
                "vencimento": row.get("vencimento_fmt", row.get("vencimento", "")),
                "valor": row.get("valor_fmt", row.get("valor", "")),
                "telefone": row.get("telefone", ""),
                "telefone_valido": row.get("telefone_valido", False),
                "mensagem": msg,
                "template_id": tid,
                "contexto": ctx,
                "preparation_status": prep_status,
                "missing_fields": missing_fields,
                "placeholders_left": placeholders_left,
                "row_json": json.dumps(ctx, ensure_ascii=False),
            }
        )

    return messages


def _soft_shuffle(lst: list, strength: float = 0.2) -> None:
    n = len(lst)
    if n < 2:
        return
    swaps = max(1, int(n * strength))
    for _ in range(swaps):
        i = random.randint(0, n - 1)
        j = min(n - 1, i + random.randint(1, max(1, int(n * 0.1))))
        lst[i], lst[j] = lst[j], lst[i]