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
    template_manager=None,
    template_id: Optional[str] = None,
) -> dict[str, str]:
    """
    Mapeia automaticamente colunas da planilha para nomes mais comuns,
    mas NÃO exige nenhuma.
    """
    df_cols = {normalize_column_name(c): c for c in df.columns}
    aliases = resolve_aliases(template_manager, template_id)
    mapping: dict[str, str] = {}

    # tenta mapear aliases conhecidos
    for var_name, alias_list in aliases.items():
        if normalize_column_name(var_name) in df_cols:
            mapping[var_name] = df_cols[normalize_column_name(var_name)]
            continue

        for alias in alias_list:
            alias_norm = normalize_column_name(alias)
            if alias_norm in df_cols:
                mapping[var_name] = df_cols[alias_norm]
                break

    # também copia todas as colunas reais como contexto direto
    for original_col in df.columns:
        col_norm = normalize_column_name(original_col)
        if col_norm not in mapping:
            mapping[col_norm] = original_col

    return mapping


def clean_dataframe_dynamic(
    df: pd.DataFrame,
    template_manager=None,
    template_id: Optional[str] = None,
    phone_column: Optional[str] = None,
) -> tuple[pd.DataFrame, dict[str, str]]:
    """
    Nunca falha por coluna ausente.
    Cria contexto com tudo que existir.
    Se phone_column for informado, usa essa coluna como telefone.
    """
    result = df.copy()
    mapping = resolve_column_mapping(result, template_manager, template_id)

    # cria colunas internas com base no mapeamento encontrado
    for internal_name, original_col in mapping.items():
        try:
            result[internal_name] = result[original_col]
        except Exception:
            result[internal_name] = ""

    # limpa strings
    for col in result.columns:
        if result[col].dtype == object:
            result[col] = result[col].fillna("").astype(str).str.strip()

    # coluna manual de telefone tem prioridade
    if phone_column and phone_column in result.columns:
        result["telefone"] = result[phone_column].apply(normalize_phone)
        mapping["telefone"] = phone_column
    elif "telefone" in result.columns:
        result["telefone"] = result["telefone"].apply(normalize_phone)
    else:
        result["telefone"] = ""

    result["telefone_valido"] = result["telefone"].apply(is_valid_phone)

    # vencimento opcional
    if "vencimento" in result.columns:
        result["vencimento_dt"] = result["vencimento"].apply(parse_date)
        result["vencimento_fmt"] = result["vencimento"].apply(format_date_br)
    else:
        result["vencimento_dt"] = pd.NaT
        result["vencimento_fmt"] = ""

    # valor opcional
    if "valor" in result.columns:
        result["valor_num"] = result["valor"].apply(parse_money)
        result["valor_fmt"] = result["valor_num"].apply(format_money_br)
    else:
        result["valor_num"] = 0.0
        result["valor_fmt"] = ""

    return result, mapping

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
    """
    Joga tudo da linha para o contexto.
    O template usa qualquer chave que existir.
    """
    ctx = {}
    for key, value in row.items():
        if key.endswith("_dt") or key.endswith("_num"):
            continue
        if pd.isna(value):
            ctx[key] = ""
        else:
            ctx[key] = str(value)

    # prioriza formatos amigáveis se existirem
    if "vencimento_fmt" in row:
        ctx["vencimento"] = row.get("vencimento_fmt", "")
    if "valor_fmt" in row:
        ctx["valor"] = row.get("valor_fmt", "")
    if "telefone" in row:
        ctx["telefone"] = row.get("telefone", "")

    return ctx


def has_unresolved_placeholders(text: str) -> list[str]:
    return re.findall(r"{\s*([a-zA-Z0-9_]+)\s*}", text or "")


def evaluate_row_preparation(row: dict, mensagem: str) -> tuple[str, list[str]]:
    """
    Não existe mais campo obrigatório.
    Só marca pendente se sobrou placeholder no texto.
    """
    placeholders_left = has_unresolved_placeholders(mensagem)

    if not row.get("telefone_valido", False):
        return "telefone_invalido", placeholders_left

    if placeholders_left:
        return "placeholder_pendente", placeholders_left

    return "pronto", placeholders_left


def safe_replace_template(text: str, context: dict) -> str:
    """
    Substitui qualquer {campo} por valor se existir,
    ou string vazia se não existir.
    """
    def repl(match):
        key = match.group(1).strip()
        return str(context.get(key, ""))

    return re.sub(r"{\s*([a-zA-Z0-9_]+)\s*}", repl, text or "")


def generate_messages(
    df: pd.DataFrame,
    template_manager=None,
    template_id: Optional[str] = None,
    fallback_template: str = "",
) -> list[dict]:
    messages = []
    rows = df.to_dict("records")

    if len(rows) > 3:
        _soft_shuffle(rows, strength=0.25)

    for row in rows:
        ctx = build_context(row)

        if template_manager is not None:
            if template_id:
                tpl = template_manager.get_template_by_id(template_id)
                if tpl:
                    raw = template_manager.pick_variant(tpl)
                    msg = safe_replace_template(raw, ctx)
                    tid = tpl["id"]
                else:
                    msg = safe_replace_template(fallback_template, ctx)
                    tid = "manual"
            else:
                actives = template_manager.get_active_templates()
                if actives:
                    tpl = random.choice(actives)
                    raw = template_manager.pick_variant(tpl)
                    msg = safe_replace_template(raw, ctx)
                    tid = tpl["id"]
                else:
                    msg = safe_replace_template(fallback_template, ctx)
                    tid = "manual"
        else:
            msg = safe_replace_template(fallback_template, ctx)
            tid = "manual"

        prep_status, placeholders_left = evaluate_row_preparation(row, msg)

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