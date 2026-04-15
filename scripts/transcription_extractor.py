"""
TRANSCRIPTION_EXTRACTOR
Lê arquivo TXT de transcrição de reunião e extrai dados estruturados em JSON,
conforme o schema oficial do CJF (Resolução CJF nº 499/2018 e alterações).
"""

import json
import re
from datetime import datetime

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


MESES_PT = {
    "janeiro": "01", "fevereiro": "02", "março": "03", "marco": "03",
    "abril": "04", "maio": "05", "junho": "06", "julho": "07",
    "agosto": "08", "setembro": "09", "outubro": "10",
    "novembro": "11", "dezembro": "12"
}


class TranscriptionExtractor:
    """Extrai dados estruturados de transcrições TXT de reuniões do CJF."""

    def __init__(self, arquivo_txt: str, xlsx_participantes: str = None):
        self.arquivo = arquivo_txt
        self.xlsx_participantes = xlsx_participantes
        self.conteudo = self._ler_arquivo()

    def _ler_arquivo(self) -> str:
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                with open(self.arquivo, "r", encoding=enc) as f:
                    return f.read()
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Não foi possível decodificar o arquivo: {self.arquivo}")

    # ------------------------------------------------------------------
    # METADADOS
    # ------------------------------------------------------------------

    def extrair_metadados(self) -> dict:
        meta = {
            "data": "",
            "horario_inicio": "",
            "horario_fim": "",
            "local": "",
            "orgao": "Centro Nacional de Inteligência da Justiça Federal",
            "tipo_reuniao": "Ordinária",
            "numero_sei": "[A PREENCHER]",
        }

        # Data — padrão: "04 de março de 2026" ou "04/03/2026"
        m = re.search(
            r"Data[:\s]+(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})",
            self.conteudo, re.IGNORECASE
        )
        if m:
            dia, mes_txt, ano = m.group(1), m.group(2).lower(), m.group(3)
            mes = MESES_PT.get(mes_txt, "??")
            meta["data"] = f"{dia.zfill(2)}/{mes}/{ano}"
        else:
            m2 = re.search(r"Data[:\s]+(\d{2}/\d{2}/\d{4})", self.conteudo, re.IGNORECASE)
            if m2:
                meta["data"] = m2.group(1)

        # Horário — padrão: "10:00 - 17:30" ou "10:00 às 17:30"
        m = re.search(
            r"Hor[aá]rio[:\s]+(\d{1,2}:\d{2})\s*[-–àas]+\s*(\d{1,2}:\d{2})",
            self.conteudo, re.IGNORECASE
        )
        if m:
            meta["horario_inicio"] = m.group(1)
            meta["horario_fim"] = m.group(2)

        # Local
        m = re.search(r"Local[:\s]+(.+?)(?:\n|$)", self.conteudo, re.IGNORECASE)
        if m:
            meta["local"] = m.group(1).strip()

        # Tipo de reunião
        if re.search(r"extraordin[áa]ria", self.conteudo, re.IGNORECASE):
            meta["tipo_reuniao"] = "Extraordinária"

        return meta

    # ------------------------------------------------------------------
    # COORDENADOR
    # ------------------------------------------------------------------

    def extrair_coordenador(self) -> dict:
        coordenador = {"nome": "", "cargo": "", "tribunal": ""}

        # Busca padrão: "Coordenador(a): Nome (Cargo) - TRIBUNAL"
        m = re.search(
            r"[Cc]oordenador[ae]?\w*[:\s]+([^\n(]+?)\s*\(([^)]+)\)\s*[-–]\s*(TRF\d+|TJ\w+|STJ|STF|CJF)",
            self.conteudo
        )
        if m:
            coordenador["nome"] = m.group(1).strip()
            coordenador["cargo"] = m.group(2).strip()
            coordenador["tribunal"] = m.group(3).strip()
        else:
            # Busca pelo nome padrão da coordenadora atual
            m2 = re.search(r"(V[âa]nila\s+Cardoso[^\n,]*)", self.conteudo)
            if m2:
                coordenador["nome"] = m2.group(1).strip()
                coordenador["cargo"] = "Juíza Federal Auxiliar da Corregedoria-Geral"
                coordenador["tribunal"] = "TRF6"

        return coordenador

    # ------------------------------------------------------------------
    # PARTICIPANTES
    # ------------------------------------------------------------------

    def extrair_interlocutores(self) -> list:
        participantes = []

        # Tenta carregar XLSX se fornecido
        if self.xlsx_participantes and OPENPYXL_AVAILABLE:
            participantes = self._carregar_xlsx_participantes()

        if not participantes:
            participantes = self._extrair_participantes_txt()

        return participantes

    def _carregar_xlsx_participantes(self) -> list:
        participantes = []
        try:
            wb = openpyxl.load_workbook(self.xlsx_participantes, read_only=True)
            ws = wb.active
            headers = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    headers = [str(c).lower().strip() if c else "" for c in row]
                    continue
                if not any(row):
                    continue
                p = {}
                for j, val in enumerate(row):
                    if j < len(headers):
                        p[headers[j]] = str(val).strip() if val else ""
                # Normaliza chaves comuns
                participantes.append({
                    "nome": p.get("nome completo", p.get("nome", "")),
                    "cargo": p.get("cargo", ""),
                    "tribunal": p.get("tribunal", ""),
                    "regiao": p.get("região", p.get("regiao", "")),
                    "status": "Presente",
                })
        except Exception as e:
            print(f"[AVISO] Não foi possível ler o XLSX: {e}")
        return participantes

    def _extrair_participantes_txt(self) -> list:
        participantes = []
        # Bloco PARTICIPANTES:
        m = re.search(
            r"PARTICIPANTES[:\s]*(.*?)(?=\n[A-ZÁÉÍÓÚÂÊÔ]{3,}|\nPAUTA|\nDISCUSS|$)",
            self.conteudo, re.DOTALL | re.IGNORECASE
        )
        if not m:
            return participantes

        bloco = m.group(1)
        linhas = [l.strip() for l in bloco.splitlines() if l.strip()]

        for linha in linhas:
            if not linha.startswith("-"):
                continue
            linha = linha.lstrip("-").strip()

            # Padrão: "Nome (Cargo) - TRFx"  ou  "Nome - Cargo - TRFx"
            m_trib = re.search(
                r"(TRF\d+|TJ\w+|STJ|STF|CNJ|CJF)",
                linha, re.IGNORECASE
            )
            tribunal = m_trib.group(1).upper() if m_trib else ""

            m_cargo = re.search(r"\(([^)]+)\)", linha)
            cargo = m_cargo.group(1).strip() if m_cargo else ""

            nome = re.sub(r"\s*\([^)]*\)", "", linha)
            nome = re.sub(r"\s*-\s*(TRF\d+|TJ\w+|STJ|STF|CNJ|CJF)\s*", "", nome, flags=re.IGNORECASE)
            nome = nome.strip()

            if nome:
                participantes.append({
                    "nome": nome,
                    "cargo": cargo,
                    "tribunal": tribunal,
                    "regiao": "",
                    "status": "Presente",
                })

        return participantes

    # ------------------------------------------------------------------
    # TEMAS / NOTAS TÉCNICAS
    # ------------------------------------------------------------------

    def extrair_temas(self) -> list:
        temas = []

        # Bloco de discussões — divide por seções numeradas
        bloco_discussoes = re.search(
            r"DISCUSS[OÕ]ES?[:\s]*(.*?)(?=ENCAMINHAMENTOS|$)",
            self.conteudo, re.DOTALL | re.IGNORECASE
        )
        if not bloco_discussoes:
            return temas

        texto = bloco_discussoes.group(1)

        # Separa cada tema pelo número ou título
        secoes = re.split(r"\n(?=\d+\.\s+[A-ZÁÉÍÓÚ])", texto)

        for secao in secoes:
            if not secao.strip():
                continue

            # Título do tema
            m_titulo = re.match(r"(\d+)\.\s+(.+?)(?:\n|$)", secao)
            titulo = m_titulo.group(2).strip() if m_titulo else secao.split("\n")[0].strip()

            # Número de Nota Técnica
            m_nt = re.search(
                r"(Nota T[ée]cnica(?:\s+n\.?\s*[\d/]+)?(?:\s*\([^)]*\))?)",
                secao, re.IGNORECASE
            )
            numero_nt = m_nt.group(1).strip() if m_nt else titulo[:60]

            # Relator
            m_rel = re.search(
                r"(?:Relator[ae]?|Apresenta[çc][aã]o)[:\s]+([^\n]+)",
                secao, re.IGNORECASE
            )
            relator = m_rel.group(1).strip() if m_rel else ""

            # Resumo da apresentação
            m_apres = re.search(
                r"Apresenta[çc][aã]o[:\s]*(.*?)(?=Discuss[aã]o|Delibera[çc][aã]o|$)",
                secao, re.DOTALL | re.IGNORECASE
            )
            resumo_apres = ""
            if m_apres:
                resumo_apres = " ".join(m_apres.group(1).split())[:300]

            # Discussões por participante
            discussoes = self._extrair_discussoes_secao(secao)

            # Deliberação
            m_delib = re.search(
                r"Delibera[çc][aã]o[:\s]*(.*?)(?=\n\n|\nVota[çc][aã]o|$)",
                secao, re.DOTALL | re.IGNORECASE
            )
            deliberacao = ""
            if m_delib:
                deliberacao = " ".join(m_delib.group(1).split())

            # Votação
            votacao = "Unânime"
            if re.search(r"maioria", secao, re.IGNORECASE):
                votacao = "Por maioria"
            elif re.search(r"unanim", secao, re.IGNORECASE):
                votacao = "Unânime"

            temas.append({
                "numero": numero_nt,
                "titulo": titulo,
                "relator": relator,
                "resumo_apresentacao": resumo_apres,
                "discussoes": discussoes,
                "deliberacao": deliberacao,
                "votacao": votacao,
            })

        return temas

    def _extrair_discussoes_secao(self, secao: str) -> list:
        discussoes = []
        bloco = re.search(
            r"Discuss[aã]o[:\s]*(.*?)(?=Delibera[çc][aã]o|$)",
            secao, re.DOTALL | re.IGNORECASE
        )
        if not bloco:
            return discussoes

        texto = bloco.group(1)
        linhas = [l.strip() for l in texto.splitlines() if l.strip().startswith("-")]

        for linha in linhas:
            linha = linha.lstrip("-").strip()
            # Padrão: "Nome (Cargo) (TRFx): fala"
            m = re.match(
                r"([^(:\n]+?)\s*\(([^)]+)\)\s*(?:\(([^)]+)\))?[:\s]+(.*)",
                linha
            )
            if m:
                nome = m.group(1).strip()
                cargo = m.group(2).strip()
                tribunal = m.group(3).strip() if m.group(3) else ""
                fala = m.group(4).strip()
                # Trunca fala em 3 linhas (~240 chars)
                resumo = fala[:240] + ("..." if len(fala) > 240 else "")
                discussoes.append({
                    "participante": nome,
                    "cargo": cargo,
                    "tribunal": tribunal,
                    "resumo": resumo,
                })
            else:
                # Fallback: linha sem padrão estruturado
                resumo = linha[:240] + ("..." if len(linha) > 240 else "")
                discussoes.append({
                    "participante": "",
                    "cargo": "",
                    "tribunal": "",
                    "resumo": resumo,
                })

        return discussoes

    # ------------------------------------------------------------------
    # ENCAMINHAMENTOS
    # ------------------------------------------------------------------

    def extrair_encaminhamentos(self) -> list:
        encaminhamentos = []

        m = re.search(
            r"ENCAMINHAMENTOS[:\s]*(.*?)(?=\n[A-ZÁÉÍÓÚ]{4,}|\Z)",
            self.conteudo, re.DOTALL | re.IGNORECASE
        )
        if not m:
            return encaminhamentos

        bloco = m.group(1)
        # Divide por linhas numeradas ou com hífen
        itens = re.split(r"\n(?=\d+\.|[-•])", bloco)

        for i, item in enumerate(itens, start=1):
            item = item.strip()
            if not item:
                continue

            # Remove numeração inicial
            item = re.sub(r"^\d+\.\s*", "", item).strip()
            item = item.lstrip("-•").strip()

            # Responsável
            m_resp = re.search(
                r"[Rr]espons[aá]vel[:\s]+([^\.\n-]+)",
                item
            )
            responsavel = m_resp.group(1).strip() if m_resp else ""

            # Prazo
            m_prazo = re.search(
                r"[Pp]razo[:\s]+([^\.\n]+)",
                item
            )
            prazo = m_prazo.group(1).strip() if m_prazo else "A definir"

            # Descrição (remove campos extraídos)
            descricao = item
            descricao = re.sub(r"[Rr]espons[aá]vel[:\s]+[^\.\n-]+", "", descricao)
            descricao = re.sub(r"[Pp]razo[:\s]+[^\.\n]+", "", descricao)
            descricao = re.sub(r"\s{2,}", " ", descricao).strip().rstrip(".-")

            # Prioridade inferida
            prioridade = "Média"
            if re.search(r"urgente|imediato|priorit", descricao, re.IGNORECASE):
                prioridade = "Alta"

            encaminhamentos.append({
                "numero": i,
                "descricao": descricao,
                "responsavel": responsavel,
                "cargo": "",
                "prazo": prazo,
                "prioridade": prioridade,
                "status": "Pendente",
            })

        return encaminhamentos

    # ------------------------------------------------------------------
    # NORMAS E JULGADOS
    # ------------------------------------------------------------------

    def extrair_normas_julgados(self) -> tuple:
        normas = []
        julgados = []

        # Resoluções, Leis, Decretos
        for m in re.finditer(
            r"(Resolu[çc][aã]o|Lei|Decreto|Portaria|Instrução Normativa)\s+"
            r"([A-Z]*\s*n[º°\.]?\s*[\d.]+(?:/\d+)?)",
            self.conteudo, re.IGNORECASE
        ):
            tipo = m.group(1).strip()
            numero = m.group(2).strip()
            entrada = {"tipo": tipo, "numero": numero, "descricao": ""}
            if entrada not in normas:
                normas.append(entrada)

        # Temas Repetitivos e julgados
        for m in re.finditer(
            r"(Tema(?:\s+Repetitivo)?|REsp|AREsp|RE|AI)\s*[n°º\.]?\s*([\d.,]+(?:/\d+)?)"
            r"(?:\s+(?:do|da)\s+(STJ|STF|TRF\d+|CNJ))?",
            self.conteudo, re.IGNORECASE
        ):
            tribunal = m.group(3).upper() if m.group(3) else "STJ"
            entrada = {
                "tribunal": tribunal,
                "numero": f"{m.group(1)} {m.group(2)}".strip(),
                "tema": "",
            }
            if entrada not in julgados:
                julgados.append(entrada)

        return normas, julgados

    # ------------------------------------------------------------------
    # JSON FINAL
    # ------------------------------------------------------------------

    def gerar_json_estruturado(self) -> str:
        metadata = self.extrair_metadados()
        coordenador = self.extrair_coordenador()
        participantes = self.extrair_interlocutores()
        temas = self.extrair_temas()
        encaminhamentos = self.extrair_encaminhamentos()
        normas, julgados = self.extrair_normas_julgados()

        dados = {
            "ata": {
                "metadata": {**metadata, "coordenador": coordenador},
                "participantes": participantes,
                "temas": temas,
                "encaminhamentos": encaminhamentos,
                "normas_citadas": normas,
                "julgados_citados": julgados,
            }
        }
        return json.dumps(dados, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python transcription_extractor.py <arquivo.txt> [participantes.xlsx]")
        sys.exit(1)
    xlsx = sys.argv[2] if len(sys.argv) > 2 else None
    extractor = TranscriptionExtractor(sys.argv[1], xlsx)
    print(extractor.gerar_json_estruturado())
