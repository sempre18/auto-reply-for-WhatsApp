import json
import os
import random
import re
from typing import Optional

DEFAULT_PATH = "templates.json"

_BUILTIN_FALLBACK = {
    "version": "3.0",
    "meta": {
        "description": "Template padrão embutido",
        "variables": ["{nome}", "{documento}", "{vencimento}", "{valor}", "{telefone}"],
    },
    "templates": [
        {
            "id": "aviso 1",
            "name": "aviso 1",
            "category": "cobranca",
            "tone": "formal",
            "active": True,
            "weight": 1,
            "aliases": {
                "nome": ["nome", "cliente", "historico"],
                "documento": ["documento", "doc", "pedido"],
                "vencimento": ["vencimento", "data_vencimento"],
                "valor": ["valor", "vl.documento", "vl_documento"],
                "telefone": ["telefone", "telefone cliente", "celular", "whatsapp"],
            },
            "variants": [
                ( 
                    "📢 1º AVISO – TÍTULO VENCIDO 📢\n\n"
                    "Olá, {nome}!! Tudo bem?\n"
                    "Verificamos em nosso sistema que há um título vencido em seu nome. Seguem os dados para sua conferência:\n"
                    "📄 Nota Fiscal nº: {documento}\n"
                    "💰 Valor: R$ {valor}\n"
                    "🗓️ Vencimento: {vencimento}\n\n"
                    "Solicitamos, por gentileza, que nos retorne dentro do prazo de 2 dias úteis a contar do recebimento deste comunicado, informando uma das opções abaixo:\n"
                    "✔️ Confirmação do pagamento com o envio do comprovante;\n"
                    "✔️ Solicitação da 2ª via da nota ou boleto;\n"
                    "✔️ Justificativa ou informação sobre o motivo da pendência.\n\n"
                    "Entendemos que situações assim podem ocorrer por equívoco ou por ter passado despercebido, e estamos à disposição para auxiliar da melhor forma possível.\n"
                    "Aguardamos seu retorno.\n"
                    "Atenciosamente,\n"
                    "Grupo Gs Trator 🚜\n"
                    "Equipe - Contas a Receber\n"
                )
            ],
        },
            {
            "id": "aviso 2",
            "name": "aviso 2",
            "category": "cobranca",
            "tone": "formal",
            "active": True,
            "weight": 1,
            "aliases": {
                "nome": ["nome", "cliente", "historico"],
                "documento": ["documento", "doc", "pedido"],
                "vencimento": ["vencimento", "data_vencimento"],
                "valor": ["valor", "vl.documento", "vl_documento"],
                "telefone": ["telefone", "telefone cliente", "celular", "whatsapp"],
            },
            "variants": [
                ( 
                    "📢 2º AVISO – TÍTULO VENCIDO 📢\n\n"
                    "Olá, {nome}!! Tudo bem?\n"
                    "Verificamos em nosso sistema que há um título vencido em seu nome. Seguem os dados para sua conferência:\n"
                    "📄 Nota Fiscal nº: {documento}\n"
                    "💰 Valor: R$ {valor}\n"
                    "🗓️ Vencimento: {vencimento}\n\n"
                    "Solicitamos, por gentileza, que nos retorne dentro do prazo de 2 dias úteis a contar do recebimento deste comunicado, informando uma das opções abaixo:\n"
                    "✔️ Confirmação do pagamento com o envio do comprovante;\n"
                    "✔️ Solicitação da 2ª via da nota ou boleto;\n"
                    "✔️ Justificativa ou informação sobre o motivo da pendência.\n\n"
                    "Entendemos que situações assim podem ocorrer por equívoco ou por ter passado despercebido, e estamos à disposição para auxiliar da melhor forma possível.\n"
                    "Aguardamos seu retorno.\n"
                    "Atenciosamente,\n"
                    "Grupo Gs Trator 🚜\n"
                    "Equipe - Contas a Receber\n"
                )
            ],
        },
            {
            "id": "aviso 3",
            "name": "aviso 3",
            "category": "cobranca",
            "tone": "formal",
            "active": True,
            "weight": 1,
            "aliases": {
                "nome": ["nome", "cliente", "historico"],
                "documento": ["documento", "doc", "pedido"],
                "vencimento": ["vencimento", "data_vencimento"],
                "valor": ["valor", "vl.documento", "vl_documento"],
                "telefone": ["telefone", "telefone cliente", "celular", "whatsapp"],
            },
            "variants": [
                ( 
                    "📢 3º AVISO – TÍTULO VENCIDO 📢\n\n"
                    "Olá, {nome}!! Tudo bem?\n"
                    "Verificamos em nosso sistema que há um título vencido em seu nome. Seguem os dados para sua conferência:\n"
                    "📄 Nota Fiscal nº: {documento}\n"
                    "💰 Valor: R$ {valor}\n"
                    "🗓️ Vencimento: {vencimento}\n"
                    "Solicitamos, por gentileza, que nos retorne dentro do prazo de 2 dias úteis a contar do recebimento deste comunicado, informando uma das opções abaixo:\n\n"
                    "✔️ Confirmação do pagamento com o envio do comprovante;\n"
                    "✔️ Solicitação da 2ª via da nota ou boleto;\n"
                    "✔️ Justificativa ou informação sobre o motivo da pendência.\n\n"
                    "Entendemos que situações assim podem ocorrer por equívoco ou por ter passado despercebido, e estamos à disposição para auxiliar da melhor forma possível.\n\n"
                    "Aguardamos seu retorno.\n"
                    "Atenciosamente,\n"
                    "Grupo Gs Trator 🚜\n"
                    "Equipe - Contas a Receber\n"
                )
            ],
        },
            {
            "id": "ultimo aviso",
            "name": "ultimo aviso",
            "category": "cobranca",
            "tone": "formal",
            "active": False,
            "weight": 1,
            "aliases": {
                "nome": ["nome", "cliente", "historico"],
                "documento": ["documento", "doc", "pedido"],
                "vencimento": ["vencimento", "data_vencimento"],
                "valor": ["valor", "vl.documento", "vl_documento"],
                "telefone": ["telefone", "telefone cliente", "celular", "whatsapp"],
            },
            "variants": [
                ( 
                    "📢 ÚLTIMO AVISO – URGENTE: PENDÊNCIA DE PAGAMENTO E ENCAMINHAMENTO PARA O SERASA 📢\n\n"
                    "Olá, {nome}!! Tudo bem?\n"
                    "Estamos entrando em contato pela última vez para alertá-los sobre a pendência no pagamento da seguinte nota fiscal:\n"
                    "📄 Nota Fiscal nº: {documento}\n"
                    "💰 Valor: R$ {valor}\n"
                    "🗓️ Vencimento: {vencimento}\n"
                    "Até o momento, não identificamos o pagamento nem recebemos retorno referente aos avisos anteriores.\n"
                    "Este comunicado representa a última tentativa de contato.  Caso não haja manifestação ou regularização da pendência, informamos que a nota fiscal será encaminhada automaticamente para negativação no dia 07/03, para as devidas providências legais, o que poderá gerar custos adicionais e impactos no crédito da empresa.\n"
                    "Para evitar esse encaminhamento, solicitamos que nos retornem antes da data informada, com uma das opções abaixo:\n"
                    "✔️ Envio do comprovante de pagamento (caso já tenha sido efetuado)\n"
                    "✔️ Solicitação da segunda via do boleto ou da nota fiscal\n"
                    "✔️ Justificativa ou esclarecimentos sobre a pendência\n"
                    "Contamos com sua compreensão e reforçamos a necessidade de regularização imediata para evitar maiores transtornos.\n\n"
                    "Atenciosamente,\n"
                    "Grupo Gs Trator 🚜\n"
                    "Equipe - Contas a Receber\n"
                )
            ],
        }
    ],
}


class TemplateManager:
    """Gerencia templates, aliases e placeholders."""

    def __init__(self, path: str = DEFAULT_PATH):
        self.path = path
        self._data: dict = {}
        self._last_used: dict[str, int] = {}
        self.load()

    def load(self) -> None:
        if os.path.exists(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    self._data = json.load(f)
            except (json.JSONDecodeError, OSError):
                self._data = json.loads(json.dumps(_BUILTIN_FALLBACK))
        else:
            self._data = json.loads(json.dumps(_BUILTIN_FALLBACK))
            self.save()

    def save(self) -> None:
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get_all_templates(self) -> list[dict]:
        return self._data.get("templates", [])

    def get_active_templates(self, category: Optional[str] = None) -> list[dict]:
        tpls = [t for t in self.get_all_templates() if t.get("active", True)]
        if category:
            tpls = [t for t in tpls if t.get("category") == category]
        return tpls

    def get_template_by_id(self, tid: str) -> Optional[dict]:
        for t in self.get_all_templates():
            if t.get("id") == tid:
                return t
        return None

    def get_template_by_name(self, name: str) -> Optional[dict]:
        for t in self.get_all_templates():
            if t.get("name") == name:
                return t
        return None

    def get_template_names(self) -> list[str]:
        return [t.get("name", "") for t in self.get_all_templates()]

    def get_active_names(self) -> list[str]:
        return [t.get("name", "") for t in self.get_active_templates()]

    def get_template_aliases(self, tid: Optional[str] = None) -> dict[str, list[str]]:
        templates = []
        if tid:
            tpl = self.get_template_by_id(tid)
            if tpl:
                templates = [tpl]
        else:
            templates = self.get_active_templates()

        aliases: dict[str, list[str]] = {}
        for tpl in templates:
            for key, values in tpl.get("aliases", {}).items():
                aliases.setdefault(key, [])
                for v in values:
                    if v not in aliases[key]:
                        aliases[key].append(v)
        return aliases

    @staticmethod
    def extract_placeholders_from_text(text: str) -> list[str]:
        return sorted(set(re.findall(r"{\s*([a-zA-Z0-9_]+)\s*}", text or "")))

    def get_template_variables(self, tid: Optional[str] = None) -> list[str]:
        templates = []
        if tid:
            tpl = self.get_template_by_id(tid)
            if tpl:
                templates = [tpl]
        else:
            templates = self.get_active_templates()

        vars_found = set()
        for tpl in templates:
            for variant in tpl.get("variants", []):
                vars_found.update(self.extract_placeholders_from_text(variant))
        return sorted(vars_found)

    def pick_variant(self, template: dict, avoid_repeat: bool = True) -> str:
        variants = template.get("variants", [])
        if not variants:
            return ""

        if len(variants) == 1:
            return variants[0]

        tid = template["id"]
        last_idx = self._last_used.get(tid, -1)

        if avoid_repeat:
            candidates = [i for i in range(len(variants)) if i != last_idx]
            if not candidates:
                candidates = list(range(len(variants)))
        else:
            candidates = list(range(len(variants)))

        chosen_idx = random.choice(candidates)
        self._last_used[tid] = chosen_idx
        return variants[chosen_idx]

    def render(self, template_id: str, context: dict) -> tuple[str, str]:
        tpl = self.get_template_by_id(template_id)
        if not tpl:
            actives = self.get_active_templates()
            if not actives:
                return self._fallback_render(context), "fallback"
            tpl = actives[0]

        raw = self.pick_variant(tpl)
        rendered = self._substitute(raw, context)
        return rendered, tpl["id"]

    def render_random(self, context: dict, category: Optional[str] = None) -> tuple[str, str]:
        actives = self.get_active_templates(category)
        if not actives:
            return self._fallback_render(context), "fallback"

        weights = [t.get("weight", 1) for t in actives]
        tpl = random.choices(actives, weights=weights, k=1)[0]
        raw = self.pick_variant(tpl)
        rendered = self._substitute(raw, context)
        return rendered, tpl["id"]

    @staticmethod
    def _substitute(text: str, ctx: dict) -> str:
        result = text
        for key, val in ctx.items():
            result = result.replace("{" + key + "}", str(val))
        return result

    @staticmethod
    def _fallback_render(context: dict) -> str:
        return (
            f"Olá {context.get('nome', '')}, tudo bem?\n\n"
            f"Doc. {context.get('documento', '')} — "
            f"venc. {context.get('vencimento', '')} — "
            f"{context.get('valor', '')}.\n\n"
            "Qualquer dúvida é só falar!"
        )

    def validate_template(self, template: dict) -> tuple[bool, str]:
        if not template.get("id", "").strip():
            return False, "ID obrigatório."
        if not template.get("name", "").strip():
            return False, "Nome obrigatório."
        variants = template.get("variants", [])
        if not isinstance(variants, list) or not variants:
            return False, "É necessário pelo menos uma variante."
        return True, ""

    def add_template(self, template: dict) -> None:
        ok, err = self.validate_template(template)
        if not ok:
            raise ValueError(err)

        existing_ids = [t["id"] for t in self.get_all_templates()]
        if template.get("id") in existing_ids:
            raise ValueError(f"Template com id '{template['id']}' já existe.")
        self._data.setdefault("templates", []).append(template)
        self.save()

    def update_template(self, tid: str, updates: dict) -> bool:
        for i, t in enumerate(self._data.get("templates", [])):
            if t["id"] == tid:
                merged = dict(t)
                merged.update(updates)
                ok, err = self.validate_template(merged)
                if not ok:
                    raise ValueError(err)
                self._data["templates"][i] = merged
                self.save()
                return True
        return False

    def delete_template(self, tid: str) -> bool:
        before = len(self._data.get("templates", []))
        self._data["templates"] = [
            t for t in self._data.get("templates", []) if t["id"] != tid
        ]
        if len(self._data["templates"]) < before:
            self.save()
            return True
        return False

    def toggle_active(self, tid: str) -> Optional[bool]:
        for t in self._data.get("templates", []):
            if t["id"] == tid:
                t["active"] = not t.get("active", True)
                self.save()
                return t["active"]
        return None