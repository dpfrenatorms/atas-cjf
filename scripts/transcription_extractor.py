"""
TRANSCRIPTION_EXTRACTOR
Lê arquivo TXT de transcrição de reunião e extrai dados estruturados em JSON,
conforme o schema oficial do CJF (Resolução CJF nº 499/2018 e alterações).

Suporta dois formatos:
  1. Exportação SEI "Memória de Reunião" (formato real CJF)
  2. Transcrição livre estruturada (formato legado)
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

NUMEROS_EXTENSO = {
    "um": 1, "uma": 1, "dois": 2, "duas": 2, "três": 3, "quatro": 4,
    "cinco": 5, "seis": 6, "sete": 7, "oito": 8, "nove": 9, "dez": 10,
    "onze": 11, "doze": 12, "treze": 13, "quatorze": 14, "quinze": 15,
    "dezesseis": 16, "dezessete": 17, "dezoito": 18, "dezenove": 19,
    "vinte": 20, "vinte e um": 21, "vinte e dois": 22, "vinte e três": 23,
    "vinte e quatro": 24, "vinte e cinco": 25, "vinte e seis": 26,
    "vinte e sete": 27, "vinte e oito": 28, "vinte e nove": 29,
    "trinta": 30, "trinta e um": 31,
}

# Prefixos de cargo em ordem do mais específico para o menos
CARGOS_PREFIXOS = [
    "Juíza Federal Auxiliar da Corregedoria-Geral e Coordenadora do Grupo Operacional do CIn",
    "Juíza Federal Auxiliar da Corregedoria-Geral e Coordenadora",
    "Juíza Auxiliar da Corregedoria-Geral e Coordenadora",
    "Juíza Federal Auxiliar da Corregedoria-Geral",
    "Juíza Auxiliar da Corregedoria-Geral",
    "Desembargadora Federal",
    "Desembargador Federal",
    "Juíza de Direito",
    "Juiz de Direito",
    "Juíza Federal",
    "Juiz Federal",
    "Assessora-Chefe",
    "Assessor-Chefe",
    "Assessora da Secretaria do Centro de Estudos Judiciário",
    "servidor da Secretaria do Centro de Estudos Judiciário",
    "Secretária da Turma Nacional de Uniformização de Jurisprudência",
    "Secretário da Turma Nacional de Uniformização de Jurisprudência",
    "servidora",
    "servidor",
    "Assessora",
    "Assessor",
]


class TranscriptionExtractor:
    """Extrai dados estruturados de transcrições TXT de reuniões do CJF."""

    def __init__(self, arquivo_txt: str, xlsx_participantes: str = None):
        self.arquivo = arquivo_txt
        self.xlsx_participantes = xlsx_participantes
        self.conteudo = self._ler_arquivo()
        self._formato_sei = self._detectar_formato_sei()

    def _ler_arquivo(self) -> str:
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                with open(self.arquivo, "r", encoding=enc) as f:
                    return f.read()
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Não foi possível decodificar o arquivo: {self.arquivo}")

    def _detectar_formato_sei(self) -> bool:
        """Detecta se o arquivo é uma exportação SEI 'Memória de Reunião'."""
        return bool(re.search(
            r'MEMÓRIA DE REUNIÃO|Estavam presentes os \(as\) seguintes participantes',
            self.conteudo, re.IGNORECASE
        ))

    def _limpar_paginacao(self, texto: str) -> str:
        """Remove artefatos de paginação da exportação SEI (URLs, cabeçalhos de página)."""
        # Remove linhas com URLs SEI
        texto = re.sub(r'https?://sei\.[^\n]+\n?', '\n', texto)
        # Remove cabeçalhos de página: "DD/MM/AAAA, HH:MM SEI/CJF - ..."
        texto = re.sub(r'\d{2}/\d{2}/\d{4},\s+\d{2}:\d{2}\s+SEI/[^\n]*\n?', '\n', texto)
        # Remove múltiplas linhas em branco consecutivas
        texto = re.sub(r'\n{3,}', '\n\n', texto)
        return texto

    # ------------------------------------------------------------------
    # METADADOS
    # ------------------------------------------------------------------

    def extrair_metadados(self) -> dict:
        meta = {
            "data": "",
            "horario_inicio": "",
            "horario_fim": "",
            "local": "Videoconferência",
            "orgao": "Centro Nacional de Inteligência da Justiça Federal",
            "tipo_reuniao": "Ordinária",
            "numero_sei": "[A PREENCHER]",
            "numero_processo": "[A PREENCHER]",
        }

        conteudo_limpo = self._limpar_paginacao(self.conteudo)

        if self._formato_sei:
            # Número SEI do header: "MEMÓRIA DE REUNIÃO n. 0840493"
            m = re.search(r'MEMÓRIA DE REUNIÃO n\.\s*(\d+)', conteudo_limpo, re.IGNORECASE)
            if m:
                meta["numero_sei"] = m.group(1)

            # Número SEI também na última linha: "SEI nº0840493"
            m2 = re.search(r'SEI\s+nº\s*(\d+)', conteudo_limpo)
            if m2:
                meta["numero_sei"] = m2.group(1)

            # Número do processo: "Processo nº0000325-13.2025.4.90.8000"
            m3 = re.search(r'Processo\s+nº\s*([\d\-\.]+)', conteudo_limpo)
            if m3:
                meta["numero_processo"] = m3.group(1)

            # Data por extenso: "Aos quatro dias do mês de março de 2026"
            m_data = re.search(
                r'Aos\s+([\w\s]+?)\s+dias?\s+do\s+mês\s+de\s+(\w+)\s+de\s+(\d{4})',
                conteudo_limpo, re.IGNORECASE
            )
            if m_data:
                dia_texto = m_data.group(1).strip().lower()
                dia = NUMEROS_EXTENSO.get(dia_texto, dia_texto)
                mes = MESES_PT.get(m_data.group(2).lower(), "??")
                ano = m_data.group(3)
                meta["data"] = f"{str(dia).zfill(2)}/{mes}/{ano}"
            else:
                # Fallback: data no título "4 de março de 2026"
                m_data2 = re.search(
                    r'(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})',
                    conteudo_limpo
                )
                if m_data2:
                    mes = MESES_PT.get(m_data2.group(2).lower(), "??")
                    meta["data"] = f"{m_data2.group(1).zfill(2)}/{mes}/{m_data2.group(3)}"

            # Horário de início: "às 10 horas" ou "às 10:30"
            m_hora = re.search(r'às\s+(\d{1,2})\s+horas', conteudo_limpo)
            if m_hora:
                meta["horario_inicio"] = f"{m_hora.group(1).zfill(2)}:00"
            else:
                m_hora2 = re.search(r'às\s+(\d{1,2}):(\d{2})', conteudo_limpo)
                if m_hora2:
                    meta["horario_inicio"] = f"{m_hora2.group(1).zfill(2)}:{m_hora2.group(2)}"

            # Tipo de reunião
            if re.search(r'extraordin[áa]ria', conteudo_limpo, re.IGNORECASE):
                meta["tipo_reuniao"] = "Extraordinária"
            elif re.search(r'Grupo Operacional', conteudo_limpo, re.IGNORECASE):
                meta["tipo_reuniao"] = "Grupo Operacional"

        else:
            # Formato legado
            m = re.search(
                r"Data[:\s]+(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})",
                conteudo_limpo, re.IGNORECASE
            )
            if m:
                dia, mes_txt, ano = m.group(1), m.group(2).lower(), m.group(3)
                mes = MESES_PT.get(mes_txt, "??")
                meta["data"] = f"{dia.zfill(2)}/{mes}/{ano}"
            else:
                m2 = re.search(r"Data[:\s]+(\d{2}/\d{2}/\d{4})", conteudo_limpo, re.IGNORECASE)
                if m2:
                    meta["data"] = m2.group(1)

            m = re.search(
                r"Hor[aá]rio[:\s]+(\d{1,2}:\d{2})\s*[-–àas]+\s*(\d{1,2}:\d{2})",
                conteudo_limpo, re.IGNORECASE
            )
            if m:
                meta["horario_inicio"] = m.group(1)
                meta["horario_fim"] = m.group(2)

            m = re.search(r"Local[:\s]+(.+?)(?:\n|$)", conteudo_limpo, re.IGNORECASE)
            if m:
                meta["local"] = m.group(1).strip()

            if re.search(r"extraordin[áa]ria", conteudo_limpo, re.IGNORECASE):
                meta["tipo_reuniao"] = "Extraordinária"

        return meta

    # ------------------------------------------------------------------
    # COORDENADOR
    # ------------------------------------------------------------------

    def extrair_coordenador(self) -> dict:
        coordenador = {"nome": "", "cargo": "", "tribunal": ""}
        conteudo_limpo = self._limpar_paginacao(self.conteudo)

        if self._formato_sei:
            # Busca coordenadora nas DELIBERAÇÃO blocks ou abertura
            # Padrão: "A Juíza Auxiliar da Corregedoria-Geral e Coordenadora ... Vânila Cardoso André de Moraes"
            m = re.search(
                r'(?:Juíza(?:\s+Federal)?\s+Auxiliar da Corregedoria-Geral\s+e\s+Coordenadora'
                r'[^,\.]*,\s+)([\w\s]+?(?:Moraes|André|Cardoso)[^\n,\.]*)',
                conteudo_limpo
            )
            if m:
                coordenador["nome"] = m.group(1).strip()
                coordenador["cargo"] = "Juíza Federal Auxiliar da Corregedoria-Geral e Coordenadora do Grupo Operacional"
                coordenador["tribunal"] = "TRF6"
            else:
                # Fallback: busca pelo nome padrão
                m2 = re.search(r'(V[âa]nila\s+Cardoso[^\n,;(]*)', conteudo_limpo)
                if m2:
                    coordenador["nome"] = m2.group(1).strip()
                    coordenador["cargo"] = "Juíza Federal Auxiliar da Corregedoria-Geral e Coordenadora do Grupo Operacional"
                    coordenador["tribunal"] = "TRF6"
        else:
            m = re.search(
                r"[Cc]oordenador[ae]?\w*[:\s]+([^\n(]+?)\s*\(([^)]+)\)\s*[-–]\s*(TRF\d+|TJ\w+|STJ|STF|CJF)",
                conteudo_limpo
            )
            if m:
                coordenador["nome"] = m.group(1).strip()
                coordenador["cargo"] = m.group(2).strip()
                coordenador["tribunal"] = m.group(3).strip()
            else:
                m2 = re.search(r"(V[âa]nila\s+Cardoso[^\n,]*)", conteudo_limpo)
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
        if self._formato_sei:
            return self._extrair_participantes_sei()
        return self._extrair_participantes_legado()

    def _extrair_participantes_sei(self) -> list:
        """Extrai participantes do parágrafo narrativo do formato SEI."""
        participantes = []
        conteudo_limpo = self._limpar_paginacao(self.conteudo)

        # Participantes efetivos: "Estavam presentes..."
        m = re.search(
            r'Estavam presentes[^:]*:\s*(.*?)(?=\n(?:Participaram como convidados|Quadro-resumo:))',
            conteudo_limpo, re.DOTALL | re.IGNORECASE
        )
        if m:
            bloco = m.group(1).strip()
            participantes.extend(self._parse_lista_participantes_sei(bloco, status="Presente"))

        # Convidados: "Participaram como convidados..."
        m_conv = re.search(
            r'Participaram como convidados[^:]*:\s*(.*?)(?=\nQuadro-resumo:|$)',
            conteudo_limpo, re.DOTALL | re.IGNORECASE
        )
        if m_conv:
            bloco = m_conv.group(1).strip()
            participantes.extend(self._parse_lista_participantes_sei(bloco, status="Convidado"))

        return participantes

    def _parse_lista_participantes_sei(self, bloco: str, status: str = "Presente") -> list:
        """Divide por ponto-e-vírgula e parseia cada entrada de participante."""
        participantes = []
        # Divide por ";" mas preserva entradas que terminam com ". " (último item termina com ".")
        entries = [e.strip() for e in bloco.split(';')]

        for entry in entries:
            entry = entry.strip().rstrip('.')
            if not entry or len(entry) < 5:
                continue
            p = self._parse_participante_sei(entry, status)
            if p["nome"]:
                participantes.append(p)

        return participantes

    def _parse_participante_sei(self, entry: str, status: str = "Presente") -> dict:
        """
        Parseia uma entrada de participante no formato SEI.
        Exemplos:
          "a Juíza Federal Vânila Cardoso André de Moraes (vanila.moraes@trf6.jus.br), do TRF6"
          "o servidor Cristino Alves Brandão (cbrandao@trf3.jus.br), do TRF3"
          "a Assessora da Secretaria do Centro de Estudos Judiciário Celeni Rocha Lopes da Silva (celeni@cjf.jus.br), CEJ"
        """
        entry = entry.strip()

        # Extrai email (em parênteses, contendo @)
        email = ""
        m_email = re.search(r'\(([^)]*@[^)]*)\)', entry)
        if m_email:
            email = m_email.group(1).strip()
            texto_antes_email = entry[:m_email.start()].strip()
            texto_depois_email = entry[m_email.end():].strip()
        else:
            texto_antes_email = entry
            texto_depois_email = ""

        # Extrai tribunal (após ", do/da" ou diretamente após a vírgula)
        tribunal = ""
        texto_sem_trib = texto_depois_email
        m_trib = re.search(
            r',?\s*d[ao]s?\s+([\w\d/\-]+(?:\s*[/\-]\s*[\w\d]+)*)\s*$',
            texto_depois_email
        )
        if m_trib:
            tribunal = m_trib.group(1).strip()
        else:
            # Formatos especiais como "CEJ", "TNU/CJF", "SCE/CJF"
            m_trib2 = re.search(
                r',?\s*((?:TRF|STJ|STF|CNJ|CJF|CEJ|TNU|TJMG|TJSP|CLIRS|RIJEF|SCE|REINT|NUGEP)\S*)\s*$',
                texto_depois_email, re.IGNORECASE
            )
            if m_trib2:
                tribunal = m_trib2.group(1).strip()

        # Remove artigo inicial (a, o, e)
        texto_cargo_nome = re.sub(r'^[aoe]\s+', '', texto_antes_email, flags=re.IGNORECASE).strip()

        # Tenta extrair cargo usando prefixos conhecidos
        cargo = ""
        nome = texto_cargo_nome
        for prefixo in CARGOS_PREFIXOS:
            if texto_cargo_nome.lower().startswith(prefixo.lower()):
                cargo = prefixo
                # Capitaliza conforme versão original no texto
                idx = len(prefixo)
                nome = texto_cargo_nome[idx:].strip()
                break

        if not cargo:
            # Fallback: tudo antes do primeiro nome próprio (sequência de maiúsculas)
            # Tenta separar na primeira letra maiúscula que segue um minúsculo
            m_nome = re.search(r'\s([A-ZÁÉÍÓÚÂÊÔÃÕÜ][a-záéíóú]+(?:\s+[A-ZÁÉÍÓÚÂÊÔÃÕÜ][a-záéíóú]+)+)', texto_cargo_nome)
            if m_nome and m_nome.start() > 0:
                cargo = texto_cargo_nome[:m_nome.start()].strip()
                nome = m_nome.group(1).strip()
            else:
                nome = texto_cargo_nome

        return {
            "nome": nome,
            "cargo": cargo,
            "tribunal": tribunal,
            "regiao": "",
            "status": status,
        }

    def _extrair_participantes_legado(self) -> list:
        """Extrai participantes no formato legado (bloco PARTICIPANTES: com linhas '- ')."""
        participantes = []
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

            m_trib = re.search(r"(TRF\d+|TJ\w+|STJ|STF|CNJ|CJF)", linha, re.IGNORECASE)
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
        if self._formato_sei:
            return self._extrair_temas_sei()
        return self._extrair_temas_legado()

    def _extrair_temas_sei(self) -> list:
        """
        Extrai temas das seções numeradas da Memória de Reunião SEI.
        Seções 3-9 correspondem a análises de Notas Técnicas.
        Seções 1 (Objetivo), 2 (Abertura) e 10 (Encerramento) são ignoradas.
        O quadro-resumo (que também tem entradas numeradas) é ignorado buscando
        apenas a parte do texto a partir de '1. Objetivo' ou '2. Abertura'.
        """
        temas = []
        conteudo_limpo = self._limpar_paginacao(self.conteudo)

        # Encontra onde começam as seções reais de discussão (após o quadro-resumo)
        # "1. Objetivo" aparece logo após o quadro-resumo no formato SEI
        m_inicio = re.search(r'\n1\.\s+Objetivo|\n2\.\s+Abertura', conteudo_limpo)
        texto_discussoes = conteudo_limpo[m_inicio.start():] if m_inicio else conteudo_limpo

        # Divide pelo início de seções numeradas: "N. Análise" ou "N. Encerramento"
        # Usa lookahead para preservar o header da seção
        partes = re.split(r'(?=\n\d{1,2}\.\s+(?:Análise|Encerramento|Abertura|Objetivo))', texto_discussoes)

        for parte in partes:
            parte = parte.strip()
            if not parte:
                continue

            # Testa se começa com uma seção de análise
            m_header = re.match(r'^(\d{1,2})\.\s+(.+?)(?:\n|$)', parte)
            if not m_header:
                continue

            num_secao = int(m_header.group(1))
            titulo = m_header.group(2).strip()

            # Ignora seções que não são análises de notas técnicas
            if num_secao < 3:
                continue
            if re.search(r'Encerramento', titulo, re.IGNORECASE):
                continue

            # Corpo da seção (sem o header)
            corpo = parte[m_header.end():].strip()

            # Extrai relator: primeiro parágrafo que menciona o(a) relator(a) da NT
            relator = self._extrair_relator_secao(corpo)

            # Extrai bloco DELIBERAÇÃO (com cabeçalho explícito "DELIBERAÇÃO:")
            m_delib = re.search(r'DELIBERAÇÃO:\s*(.*?)(?=\n\d{1,2}\.\s+(?:Análise|Encerramento)|$)', parte, re.DOTALL)
            deliberacao = ""
            if m_delib:
                deliberacao = " ".join(m_delib.group(1).split())
            else:
                # Fallback: seção pode ter "Deliberou-se" inline sem cabeçalho explícito
                # O hífen pode ser ASCII (-) ou Unicode (‑ U+2011) dependendo do export SEI
                m_inline = re.search(r'(Deliberou.{0,2}se[^.]+\.)', parte)
                if m_inline:
                    deliberacao = " ".join(m_inline.group(1).split())

            # Votação
            votacao = "Unânime"
            if re.search(r'por unanimidade', parte, re.IGNORECASE):
                votacao = "Unânime"
            elif re.search(r'maioria', parte, re.IGNORECASE):
                votacao = "Por maioria"
            elif re.search(r'mantid[ao].*afeta', parte, re.IGNORECASE):
                votacao = "Adiado"

            # Discussões (parágrafos antes da DELIBERAÇÃO)
            texto_discussao = parte[:m_delib.start()].strip() if m_delib else parte
            texto_discussao = texto_discussao[m_header.end():]  # Remove header
            discussoes = self._extrair_discussoes_sei(texto_discussao)

            # Resumo da apresentação = primeiro parágrafo com o relator
            resumo_apres = discussoes[0]["resumo"] if discussoes else ""

            temas.append({
                "numero": f"NT-{num_secao - 2}",
                "titulo": titulo,
                "relator": relator,
                "resumo_apresentacao": resumo_apres[:400],
                "discussoes": discussoes[:5],
                "deliberacao": deliberacao[:600],
                "votacao": votacao,
            })

        return temas

    def _extrair_relator_secao(self, corpo: str) -> str:
        """Extrai o nome do relator do primeiro parágrafo da seção."""
        paragrafos = [p.strip() for p in corpo.split('\n') if p.strip()]
        if not paragrafos:
            return ""

        primeiro = paragrafos[0]
        # Padrão: "A/O [título] Nome (tribunal) iniciou..."
        m = re.match(
            r'^(?:A|O)\s+(?:Juíz[ao] Federal|Desembargador[ao] Federal|'
            r'(?:Juíz[ao]|Juiz[ao])\s+de\s+Direito|servidor[ao]|Assessor[ao])?'
            r'\s*([A-ZÁÉÍÓÚÂÊÔÃÕÜ][^\(\n,]{5,50}?)\s*\([^\)]+\)',
            primeiro
        )
        if m:
            return m.group(1).strip()

        # Fallback: busca pelo primeiro nome próprio com parênteses de tribunal
        m2 = re.search(r'([A-ZÁÉÍÓÚÂÊÔÃÕÜ][a-záéíóúâêôãõü]+(?:\s+[A-ZÁÉÍÓÚÂÊÔÃÕÜ][a-záéíóú]+)+)\s*\(([^)]+)\)', primeiro)
        if m2:
            return m2.group(1).strip()

        return ""

    def _extrair_discussoes_sei(self, texto: str) -> list:
        """
        Extrai contribuições de cada participante de uma seção de discussão.
        Cada parágrafo é uma contribuição (linha longa = um parágrafo no formato SEI).
        """
        discussoes = []
        paragrafos = [p.strip() for p in texto.split('\n') if p.strip() and len(p.strip()) > 30]

        for para in paragrafos:
            # Padrão: "A/O [título] Nome (tribunal), [cargo], [verbo]..."
            m = re.match(
                r'^(?:A|O)\s+(?:(?:Juíz[ao](?:\s+Federal)?|Desembargador[ao](?:\s+Federal)?|'
                r'servidor[ao]|Assessor[ao](?:-Chefe)?)\s+)?'
                r'([A-ZÁÉÍÓÚÂÊÔÃÕÜ][^\(\n,]{4,60}?)\s*\(([^)]{2,30})\)',
                para
            )
            if m:
                nome = m.group(1).strip()
                tribunal = m.group(2).strip()
                resumo = " ".join(para.split())[:280]
                discussoes.append({
                    "participante": nome,
                    "cargo": "",
                    "tribunal": tribunal,
                    "resumo": resumo,
                })
            else:
                # Inclui parágrafo sem identificação explícita (DELIBERAÇÃO já foi removida)
                resumo = " ".join(para.split())[:280]
                if resumo and not resumo.upper().startswith("DELIBERAÇÃO"):
                    discussoes.append({
                        "participante": "",
                        "cargo": "",
                        "tribunal": "",
                        "resumo": resumo,
                    })

        return discussoes

    def _extrair_temas_legado(self) -> list:
        """Extrai temas do formato legado (bloco DISCUSSÕES com seções numeradas)."""
        temas = []
        bloco_discussoes = re.search(
            r"DISCUSS[OÕ]ES?[:\s]*(.*?)(?=ENCAMINHAMENTOS|$)",
            self.conteudo, re.DOTALL | re.IGNORECASE
        )
        if not bloco_discussoes:
            return temas

        texto = bloco_discussoes.group(1)
        secoes = re.split(r"\n(?=\d+\.\s+[A-ZÁÉÍÓÚ])", texto)

        for secao in secoes:
            if not secao.strip():
                continue
            m_titulo = re.match(r"(\d+)\.\s+(.+?)(?:\n|$)", secao)
            titulo = m_titulo.group(2).strip() if m_titulo else secao.split("\n")[0].strip()
            m_nt = re.search(r"(Nota T[ée]cnica(?:\s+n\.?\s*[\d/]+)?(?:\s*\([^)]*\))?)", secao, re.IGNORECASE)
            numero_nt = m_nt.group(1).strip() if m_nt else titulo[:60]
            m_rel = re.search(r"(?:Relator[ae]?|Apresenta[çc][aã]o)[:\s]+([^\n]+)", secao, re.IGNORECASE)
            relator = m_rel.group(1).strip() if m_rel else ""
            m_delib = re.search(r"Delibera[çc][aã]o[:\s]*(.*?)(?=\n\n|\nVota[çc][aã]o|$)", secao, re.DOTALL | re.IGNORECASE)
            deliberacao = " ".join(m_delib.group(1).split()) if m_delib else ""
            votacao = "Unânime"
            if re.search(r"maioria", secao, re.IGNORECASE):
                votacao = "Por maioria"
            temas.append({
                "numero": numero_nt,
                "titulo": titulo,
                "relator": relator,
                "resumo_apresentacao": "",
                "discussoes": self._extrair_discussoes_legado(secao),
                "deliberacao": deliberacao,
                "votacao": votacao,
            })

        return temas

    def _extrair_discussoes_legado(self, secao: str) -> list:
        discussoes = []
        bloco = re.search(r"Discuss[aã]o[:\s]*(.*?)(?=Delibera[çc][aã]o|$)", secao, re.DOTALL | re.IGNORECASE)
        if not bloco:
            return discussoes
        texto = bloco.group(1)
        linhas = [l.strip() for l in texto.splitlines() if l.strip().startswith("-")]
        for linha in linhas:
            linha = linha.lstrip("-").strip()
            m = re.match(r"([^(:\n]+?)\s*\(([^)]+)\)\s*(?:\(([^)]+)\))?[:\s]+(.*)", linha)
            if m:
                resumo = m.group(4).strip()[:240]
                discussoes.append({
                    "participante": m.group(1).strip(),
                    "cargo": m.group(2).strip(),
                    "tribunal": m.group(3).strip() if m.group(3) else "",
                    "resumo": resumo + ("..." if len(m.group(4)) > 240 else ""),
                })
        return discussoes

    # ------------------------------------------------------------------
    # ENCAMINHAMENTOS
    # ------------------------------------------------------------------

    def extrair_encaminhamentos(self) -> list:
        if self._formato_sei:
            return self._extrair_encaminhamentos_sei()
        return self._extrair_encaminhamentos_legado()

    def _extrair_encaminhamentos_sei(self) -> list:
        """
        Extrai encaminhamentos dos blocos DELIBERAÇÃO da Memória de Reunião SEI.
        Cada deliberação gera pelo menos um encaminhamento.
        """
        encaminhamentos = []
        conteudo_limpo = self._limpar_paginacao(self.conteudo)
        num = 1

        # Encontra todos os blocos DELIBERAÇÃO junto com o título da seção antecedente
        padrao_secao = re.compile(
            r'(\d{1,2})\.\s+(Análise[^\n]+)\n.*?DELIBERAÇÃO:\s*(.*?)(?=\n\d{1,2}\.\s+(?:Análise|Encerramento)|\Z)',
            re.DOTALL
        )

        for m in padrao_secao.finditer(conteudo_limpo):
            num_secao = int(m.group(1))
            titulo_secao = m.group(2).strip()
            texto_delib = " ".join(m.group(3).split())

            # 1. Encaminhamento principal: aprovação/nacionalização/manutenção
            if re.search(r'Deliberou-se pela aprovação', texto_delib, re.IGNORECASE):
                encaminhamentos.append({
                    "numero": num,
                    "descricao": f"Nota Técnica aprovada: {titulo_secao[:100]}. Submeter ao Grupo Decisório.",
                    "responsavel": "Grupo Operacional CIn",
                    "cargo": "Coordenadora",
                    "prazo": "Próxima reunião do Grupo Decisório",
                    "prioridade": "Alta",
                    "status": "Pendente",
                })
                num += 1

            elif re.search(r'Deliberou-se pela nacionalização conjunta|nationalização conjunta', texto_delib, re.IGNORECASE):
                encaminhamentos.append({
                    "numero": num,
                    "descricao": f"Nacionalização conjunta das notas técnicas referentes a: {titulo_secao[:80]}. Submeter ao Grupo Decisório.",
                    "responsavel": "Grupo Operacional CIn",
                    "cargo": "Coordenadora",
                    "prazo": "A definir",
                    "prioridade": "Alta",
                    "status": "Pendente",
                })
                num += 1

            elif re.search(r'Deliberou-se pela nacionalização', texto_delib, re.IGNORECASE):
                encaminhamentos.append({
                    "numero": num,
                    "descricao": f"Nacionalização da nota técnica: {titulo_secao[:100]}. Submeter ao Grupo Decisório.",
                    "responsavel": "Grupo Operacional CIn",
                    "cargo": "Coordenadora",
                    "prazo": "A definir",
                    "prioridade": "Alta",
                    "status": "Pendente",
                })
                num += 1

            elif re.search(r'manutenção da afetação|continuidade dos estudos|continuar estudando', texto_delib, re.IGNORECASE):
                encaminhamentos.append({
                    "numero": num,
                    "descricao": f"Manter afetação do tema em análise: {titulo_secao[:100]}. Aprofundar estudos antes de deliberar.",
                    "responsavel": "Grupo Operacional CIn",
                    "cargo": "Coordenadora",
                    "prazo": "Próxima reunião",
                    "prioridade": "Média",
                    "status": "Em análise",
                })
                num += 1

            # 2. Encaminhamentos específicos: "Solicitou que X verificasse/fizesse/elaborasse..."
            for m_sol in re.finditer(
                r'[Ss]olicitou(?:\s+(?:também|ainda))?\s+que\s+([^,\.]{5,80}?)\s+'
                r'(verificasse|fizesse|elaborasse|encaminhasse|levantasse|sistematizasse'
                r'|apresentasse|checasse|confirmasse|providenciasse)([^\.]+)\.',
                texto_delib
            ):
                responsavel_raw = m_sol.group(1).strip()
                verbo = m_sol.group(2)
                complemento = m_sol.group(3).strip()
                # Monta descrição preservando o texto original da deliberação
                descricao_raw = m_sol.group(0).strip()
                descricao = descricao_raw[0].upper() + descricao_raw[1:] if descricao_raw else ""

                # Simplifica responsável
                responsavel = responsavel_raw
                m_nome = re.search(r'([A-ZÁÉÍÓÚ][a-záéíóú]+(?:\s+[A-ZÁÉÍÓÚ][a-záéíóú]+)+)', responsavel_raw)
                if m_nome:
                    responsavel = m_nome.group(1)

                encaminhamentos.append({
                    "numero": num,
                    "descricao": descricao[:300],
                    "responsavel": responsavel,
                    "cargo": "",
                    "prazo": "A definir",
                    "prioridade": "Alta",
                    "status": "Pendente",
                })
                num += 1

            # 3. Recomendações gerais: "Recomendou que todos verificassem..."
            for m_rec in re.finditer(
                r'[Rr]ecomendou\s+(?:ainda\s+)?que\s+(todos[^\,\.]{0,80}?)\s+'
                r'(verificassem|levantassem|identificassem)([^\.]+)\.',
                texto_delib
            ):
                encaminhamentos.append({
                    "numero": num,
                    "descricao": f"{m_rec.group(1).capitalize()} {m_rec.group(2)}{m_rec.group(3)}.",
                    "responsavel": "Todos os participantes do Grupo Operacional",
                    "cargo": "",
                    "prazo": "Próxima reunião",
                    "prioridade": "Média",
                    "status": "Pendente",
                })
                num += 1

        return encaminhamentos

    def _extrair_encaminhamentos_legado(self) -> list:
        """Extrai encaminhamentos do formato legado (seção ENCAMINHAMENTOS:)."""
        encaminhamentos = []
        m = re.search(
            r"ENCAMINHAMENTOS[:\s]*(.*?)(?=\n[A-ZÁÉÍÓÚ]{4,}|\Z)",
            self.conteudo, re.DOTALL | re.IGNORECASE
        )
        if not m:
            return encaminhamentos

        bloco = m.group(1)
        itens = re.split(r"\n(?=\d+\.|[-•])", bloco)

        for i, item in enumerate(itens, start=1):
            item = item.strip()
            if not item:
                continue
            item = re.sub(r"^\d+\.\s*", "", item).strip()
            item = item.lstrip("-•").strip()

            m_resp = re.search(r"[Rr]espons[aá]vel[:\s]+([^\.\n-]+)", item)
            responsavel = m_resp.group(1).strip() if m_resp else ""
            m_prazo = re.search(r"[Pp]razo[:\s]+([^\.\n]+)", item)
            prazo = m_prazo.group(1).strip() if m_prazo else "A definir"

            descricao = item
            descricao = re.sub(r"[Rr]espons[aá]vel[:\s]+[^\.\n-]+", "", descricao)
            descricao = re.sub(r"[Pp]razo[:\s]+[^\.\n]+", "", descricao)
            descricao = re.sub(r"\s{2,}", " ", descricao).strip().rstrip(".-")

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
        conteudo_limpo = self._limpar_paginacao(self.conteudo)

        # Resoluções, Leis, Decretos, Portarias, Recomendações
        for m in re.finditer(
            r"(Resolu[çc][aã]o|Lei|Decreto|Portaria|Instrução Normativa|Recomendação|Medida Provisória)"
            r"\s+([A-Z]*\s*(?:n[º°\.]?\s*)?[\d.]+(?:/\d+)?)",
            conteudo_limpo, re.IGNORECASE
        ):
            tipo = m.group(1).strip().capitalize()
            numero = m.group(2).strip()
            # Adiciona órgão ao tipo se presente (ex: "Resolução CNJ")
            contexto = conteudo_limpo[max(0, m.start()-10):m.start()].strip()
            orgao_match = re.search(r'(CNJ|CJF|STJ|STF)\s*$', contexto)
            if orgao_match:
                tipo = f"{tipo} {orgao_match.group(1)}"
            entrada = {"tipo": tipo, "numero": numero, "descricao": ""}
            if entrada not in normas:
                normas.append(entrada)

        # Temas Repetitivos e julgados
        for m in re.finditer(
            r"(Tema(?:\s+Repetitivo)?|REsp|AREsp|RE|AI|IRDR|Controvérsia)\s*"
            r"[n°º\.]?\s*([\d.,]+(?:/\d+)?)"
            r"(?:\s+(?:do|da)\s+(STJ|STF|TRF\d+|CNJ))?",
            conteudo_limpo, re.IGNORECASE
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
