"""
QUALITY_VALIDATOR
Valida qualidade de documentos DOCX/TXT gerados (atas e ofícios)
antes de enviá-los à Amanda para importação no SEI.
"""

import json

try:
    import language_tool_python
    LT_AVAILABLE = True
except ImportError:
    LT_AVAILABLE = False

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


class QualityValidator:
    """Valida ortografia, gramática, entidades jurídicas e formatação CJF."""

    def __init__(self, documento_path: str):
        self.documento_path = documento_path
        self.erros: list = []
        self.avisos: list = []
        self._tool = None

    def _get_tool(self):
        if self._tool is None and LT_AVAILABLE:
            try:
                self._tool = language_tool_python.LanguageTool("pt-BR")
            except Exception as e:
                print(f"[AVISO] LanguageTool indisponível: {e}")
        return self._tool

    def _ler_documento(self) -> str:
        path = self.documento_path
        if path.endswith(".docx") and DOCX_AVAILABLE:
            doc = Document(path)
            return "\n".join(p.text for p in doc.paragraphs)
        elif path.endswith(".txt") or not path.endswith(".docx"):
            for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
                try:
                    with open(path, "r", encoding=enc) as f:
                        return f.read()
                except UnicodeDecodeError:
                    continue
        return ""

    # ------------------------------------------------------------------
    # VALIDAÇÕES
    # ------------------------------------------------------------------

    def validar_ortografia_gramatica(self) -> list:
        tool = self._get_tool()
        if not tool:
            self.avisos.append({
                "tipo": "Ortografia/Gramática",
                "mensagem": "LanguageTool não disponível. Instale language-tool-python.",
                "severidade": "Baixa",
            })
            return self.erros

        texto = self._ler_documento()
        if not texto:
            return self.erros

        try:
            matches = tool.check(texto)
            for match in matches:
                self.erros.append({
                    "tipo": "Ortografia/Gramática",
                    "mensagem": match.message,
                    "contexto": match.context,
                    "sugestoes": match.replacements[:3],
                    "severidade": "Média",
                })
        except Exception as e:
            self.avisos.append({
                "tipo": "Ortografia/Gramática",
                "mensagem": f"Erro ao verificar ortografia: {e}",
                "severidade": "Baixa",
            })
        return self.erros

    def validar_entidades_juridicas(self, dados_json: dict = None) -> list:
        if not dados_json:
            return self.avisos
        ata = dados_json.get("ata", {})

        normas = ata.get("normas_citadas", [])
        for norma in normas:
            if not norma.get("tipo") or not norma.get("numero"):
                self.avisos.append({
                    "tipo": "Entidade Jurídica",
                    "mensagem": f"Norma com dados incompletos: {norma}",
                    "severidade": "Baixa",
                })

        julgados = ata.get("julgados_citados", [])
        for julgado in julgados:
            if not julgado.get("numero"):
                self.avisos.append({
                    "tipo": "Julgado",
                    "mensagem": f"Julgado sem número identificado: {julgado}",
                    "severidade": "Baixa",
                })

        return self.avisos

    def validar_interlocutores(self, dados_json: dict = None) -> list:
        if not dados_json:
            return self.erros
        participantes = dados_json.get("ata", {}).get("participantes", [])
        if not participantes:
            self.avisos.append({
                "tipo": "Participantes",
                "mensagem": "Nenhum participante identificado na ata.",
                "severidade": "Média",
            })
            return self.erros

        for p in participantes:
            if not p.get("nome"):
                self.erros.append({
                    "tipo": "Dados de Participante",
                    "mensagem": f"Participante sem nome: {p}",
                    "severidade": "Alta",
                })
            elif not p.get("cargo"):
                self.avisos.append({
                    "tipo": "Dados de Participante",
                    "mensagem": f"Participante sem cargo: {p.get('nome')}",
                    "severidade": "Baixa",
                })

        return self.erros

    def validar_deliberacoes(self, dados_json: dict = None) -> list:
        if not dados_json:
            return self.avisos
        temas = dados_json.get("ata", {}).get("temas", [])
        if not temas:
            self.avisos.append({
                "tipo": "Conteúdo da Ata",
                "mensagem": "Nenhum tema/deliberação encontrado na ata.",
                "severidade": "Média",
            })
            return self.avisos

        for tema in temas:
            if not tema.get("deliberacao"):
                self.avisos.append({
                    "tipo": "Deliberação",
                    "mensagem": f"Tema sem deliberação registrada: {tema.get('titulo', '?')}",
                    "severidade": "Média",
                })

        return self.avisos

    def validar_encaminhamentos(self, dados_json: dict = None) -> list:
        if not dados_json:
            return self.erros
        encaminhamentos = dados_json.get("ata", {}).get("encaminhamentos", [])
        for enc in encaminhamentos:
            if not enc.get("responsavel"):
                self.erros.append({
                    "tipo": "Encaminhamento",
                    "mensagem": f"Encaminhamento sem responsável: '{enc.get('descricao', '?')}'",
                    "severidade": "Alta",
                })
            if not enc.get("prazo") or enc.get("prazo") == "A definir":
                self.avisos.append({
                    "tipo": "Encaminhamento",
                    "mensagem": f"Encaminhamento sem prazo definido: '{enc.get('descricao', '?')}'",
                    "severidade": "Baixa",
                })

        return self.erros

    def validar_formatacao(self) -> list:
        texto = self._ler_documento()
        if not texto:
            return self.avisos

        if "CONSELHO DA JUSTIÇA FEDERAL" not in texto.upper():
            self.avisos.append({
                "tipo": "Formatação",
                "mensagem": "Cabeçalho padrão 'CONSELHO DA JUSTIÇA FEDERAL' não detectado.",
                "severidade": "Baixa",
            })
        if "CENTRO NACIONAL DE INTELIGÊNCIA" not in texto.upper():
            self.avisos.append({
                "tipo": "Formatação",
                "mensagem": "Cabeçalho 'CENTRO NACIONAL DE INTELIGÊNCIA' não detectado.",
                "severidade": "Baixa",
            })

        return self.avisos

    # ------------------------------------------------------------------
    # RELATÓRIO COMPLETO
    # ------------------------------------------------------------------

    def gerar_relatorio_validacao(self, dados_json: dict = None) -> dict:
        self.erros = []
        self.avisos = []

        self.validar_ortografia_gramatica()
        self.validar_entidades_juridicas(dados_json)
        self.validar_interlocutores(dados_json)
        self.validar_deliberacoes(dados_json)
        self.validar_encaminhamentos(dados_json)
        self.validar_formatacao()

        erros_criticos = [e for e in self.erros if e.get("severidade") == "Alta"]

        if erros_criticos:
            status = "ERROS CRÍTICOS"
        elif self.erros:
            status = "ERROS"
        elif self.avisos:
            status = "AVISOS"
        else:
            status = "OK"

        return {
            "status": status,
            "total_erros": len(self.erros),
            "total_avisos": len(self.avisos),
            "erros": self.erros,
            "avisos": self.avisos,
            "recomendacoes": (
                "Revisar todos os itens marcados como ERRO antes de enviar à Amanda."
                if self.erros else
                "Documento aprovado para envio. Revisar avisos se necessário."
            ),
        }

    def imprimir_relatorio(self, dados_json: dict = None):
        relatorio = self.gerar_relatorio_validacao(dados_json)
        print(f"\n{'='*60}")
        print(f"  RELATÓRIO DE VALIDAÇÃO — Status: {relatorio['status']}")
        print(f"{'='*60}")
        print(f"  Erros: {relatorio['total_erros']}  |  Avisos: {relatorio['total_avisos']}")
        if relatorio["erros"]:
            print("\n  ERROS:")
            for e in relatorio["erros"]:
                print(f"  [{e['severidade']}] {e['tipo']}: {e['mensagem']}")
        if relatorio["avisos"]:
            print("\n  AVISOS:")
            for a in relatorio["avisos"]:
                print(f"  [{a['severidade']}] {a['tipo']}: {a['mensagem']}")
        print(f"\n  {relatorio['recomendacoes']}")
        print(f"{'='*60}\n")
        return relatorio


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python quality_validator.py <documento.docx>")
        sys.exit(1)
    validator = QualityValidator(sys.argv[1])
    validator.imprimir_relatorio()
