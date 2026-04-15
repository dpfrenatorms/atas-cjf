"""
MAIN — Orquestrador do Pipeline CJF
Uso:
    python scripts/main.py --transcricao <arquivo.txt> [--participantes <lista.xlsx>]

Saídas em Saida/:
    ATA_YYYYMMDD.docx
    ATA_YYYYMMDD.json
    OFICIO_1_YYYYMMDD.docx, OFICIO_2_YYYYMMDD.docx, ...
    VALIDACAO_YYYYMMDD.json
"""

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# Garante saída UTF-8 no terminal Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf-8-sig"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

# Garante que o diretório scripts/ está no path
sys.path.insert(0, str(Path(__file__).parent))

from transcription_extractor import TranscriptionExtractor
from quality_validator import QualityValidator
from docx_generator import DocxGenerator

# ------------------------------------------------------------------
# CONFIGURAÇÃO DE PASTAS
# ------------------------------------------------------------------

BASE_DIR = Path(__file__).parent.parent  # raiz do projeto
SAIDA_DIR = BASE_DIR / "Saida"
ENTRADA_DIR = BASE_DIR / "Entrada"
BACKUP_DIR = SAIDA_DIR / "Backup"

SAIDA_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)


def _sufixo_data(data_str: str) -> str:
    """Converte 'DD/MM/YYYY' em 'YYYYMMDD' para nomes de arquivo."""
    try:
        partes = data_str.split("/")
        if len(partes) == 3:
            return f"{partes[2]}{partes[1].zfill(2)}{partes[0].zfill(2)}"
    except Exception:
        pass
    return datetime.now().strftime("%Y%m%d")


# ------------------------------------------------------------------
# PIPELINE PRINCIPAL
# ------------------------------------------------------------------

def processar_reuniao(arquivo_transcricao: str, arquivo_participantes: str = None) -> dict:
    print(f"\n{'='*60}")
    print("  PIPELINE DE AUTOMAÇÃO CJF — INICIANDO")
    print(f"{'='*60}")
    print(f"  Transcrição : {arquivo_transcricao}")
    if arquivo_participantes:
        print(f"  Participantes: {arquivo_participantes}")
    print()

    # ------------------------------------------------------------------
    # ETAPA 1 — Extração de dados
    # ------------------------------------------------------------------
    print("[1/4] Extraindo dados da transcrição...")
    extractor = TranscriptionExtractor(arquivo_transcricao, arquivo_participantes)
    json_str = extractor.gerar_json_estruturado()
    dados = json.loads(json_str)

    metadata = dados.get("ata", {}).get("metadata", {})
    data_reuniao = metadata.get("data", "")
    sufixo = _sufixo_data(data_reuniao)

    # Salva JSON estruturado
    caminho_json = SAIDA_DIR / f"ATA_{sufixo}.json"
    caminho_json.write_text(json_str, encoding="utf-8")
    print(f"    ✓ JSON salvo: {caminho_json.name}")

    # ------------------------------------------------------------------
    # ETAPA 2 — Geração da Ata DOCX
    # ------------------------------------------------------------------
    print("[2/4] Gerando ATA.DOCX...")
    try:
        gen_ata = DocxGenerator(tipo_documento="ata")
        gen_ata.adicionar_cabecalho_ata(metadata)
        gen_ata.adicionar_secao_participantes(dados["ata"].get("participantes", []))
        gen_ata.adicionar_secao_deliberacoes(dados["ata"].get("temas", []))
        gen_ata.adicionar_secao_encaminhamentos(dados["ata"].get("encaminhamentos", []))
        gen_ata.adicionar_secao_normas(
            dados["ata"].get("normas_citadas", []),
            dados["ata"].get("julgados_citados", [])
        )
        gen_ata.adicionar_assinatura(metadata.get("coordenador", {}))

        caminho_ata = SAIDA_DIR / f"ATA_{sufixo}.docx"
        gen_ata.salvar(str(caminho_ata))
        print(f"    ✓ ATA salva: {caminho_ata.name}")
    except Exception as e:
        print(f"    [ERRO] Falha ao gerar ATA.DOCX: {e}")
        caminho_ata = None

    # ------------------------------------------------------------------
    # ETAPA 3 — Validação de qualidade
    # ------------------------------------------------------------------
    print("[3/4] Validando qualidade...")
    arquivos_para_validar = []
    if caminho_ata and caminho_ata.exists():
        arquivos_para_validar.append(str(caminho_ata))

    relatorio_validacao = {}
    if arquivos_para_validar:
        validator = QualityValidator(arquivos_para_validar[0])
        relatorio_validacao = validator.imprimir_relatorio(dados)

        caminho_val = SAIDA_DIR / f"VALIDACAO_{sufixo}.json"
        caminho_val.write_text(
            json.dumps(relatorio_validacao, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        print(f"    ✓ Relatório de validação salvo: {caminho_val.name}")

        if relatorio_validacao.get("status") == "ERROS CRÍTICOS":
            print("\n  [ATENÇÃO] Erros críticos encontrados. Revise antes de enviar à Amanda.")

    # ------------------------------------------------------------------
    # ETAPA 4 — Geração dos Ofícios
    # ------------------------------------------------------------------
    encaminhamentos = dados["ata"].get("encaminhamentos", [])
    print(f"[4/4] Gerando {len(encaminhamentos)} ofício(s)...")

    arquivos_oficios = []
    for enc in encaminhamentos:
        num = enc.get("numero", 1)
        try:
            gen_oficio = DocxGenerator(tipo_documento="oficio")
            gen_oficio.gerar_oficio(enc, dados)
            caminho_oficio = SAIDA_DIR / f"OFICIO_{num}_{sufixo}.docx"
            gen_oficio.salvar(str(caminho_oficio))
            arquivos_oficios.append(caminho_oficio.name)
            print(f"    ✓ Ofício {num} salvo: {caminho_oficio.name}")
        except Exception as e:
            print(f"    [ERRO] Falha ao gerar ofício {num}: {e}")

    # ------------------------------------------------------------------
    # RESUMO FINAL
    # ------------------------------------------------------------------
    print(f"\n{'='*60}")
    print("  PROCESSAMENTO CONCLUÍDO")
    print(f"{'='*60}")
    print(f"  Pasta de saída: {SAIDA_DIR}")
    print(f"  Arquivos gerados:")
    if caminho_ata:
        print(f"    – {caminho_ata.name}")
    print(f"    – ATA_{sufixo}.json")
    for arq in arquivos_oficios:
        print(f"    – {arq}")
    print()
    print("  PRÓXIMOS PASSOS PARA AMANDA:")
    print("  1. Revisar os documentos gerados na pasta Saida/")
    print("  2. Verificar o relatório VALIDACAO_*.json")
    print("  3. Importar ATA_*.docx no SEI (processo correspondente)")
    print("  4. Revisar e preencher campos [A PREENCHER] nos ofícios")
    print("  5. Importar OFICIO_*.docx no SEI para assinatura digital")
    print(f"{'='*60}\n")

    return {
        "status": "concluído",
        "arquivos": [str(caminho_ata)] + [
            str(SAIDA_DIR / a) for a in arquivos_oficios
        ],
        "json": str(caminho_json),
        "validacao": relatorio_validacao,
    }


def consolidar_relatorio(pasta_jsons: str = None, periodo: str = ""):
    """Consolida múltiplas atas JSON em relatório executivo."""
    pasta = Path(pasta_jsons) if pasta_jsons else SAIDA_DIR
    jsons = list(pasta.glob("ATA_*.json"))

    if not jsons:
        print("[AVISO] Nenhum arquivo ATA_*.json encontrado para consolidar.")
        return

    print(f"\nConsolidando {len(jsons)} ata(s) em relatório executivo...")
    lista_atas = []
    for j in jsons:
        try:
            lista_atas.append(json.loads(j.read_text(encoding="utf-8")))
        except Exception as e:
            print(f"  [AVISO] Não foi possível ler {j.name}: {e}")

    try:
        gen = DocxGenerator(tipo_documento="relatorio")
        gen.gerar_relatorio_executivo(lista_atas, periodo)
        sufixo = datetime.now().strftime("%Y%m")
        caminho = SAIDA_DIR / f"RELATORIO_EXECUTIVO_{sufixo}.docx"
        gen.salvar(str(caminho))
        print(f"✓ Relatório executivo salvo: {caminho.name}")
    except Exception as e:
        print(f"[ERRO] Falha ao gerar relatório executivo: {e}")


# ------------------------------------------------------------------
# ENTRADA DE LINHA DE COMANDO
# ------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Pipeline CJF: Atas e Ofícios",
        formatter_class=argparse.RawTextHelpFormatter
    )
    subparsers = parser.add_subparsers(dest="comando")

    # Comando padrão: processar reunião
    p_proc = subparsers.add_parser("processar", help="Processar transcrição de reunião")
    p_proc.add_argument("--transcricao", "-t", required=True, help="Caminho do arquivo TXT")
    p_proc.add_argument("--participantes", "-p", help="Caminho do arquivo XLSX (opcional)")

    # Comando: consolidar relatório
    p_rel = subparsers.add_parser("relatorio", help="Gerar relatório executivo consolidado")
    p_rel.add_argument("--pasta", help="Pasta com os JSONs das atas (padrão: Saida/)")
    p_rel.add_argument("--periodo", default="", help="Período de análise (ex: 'Março 2026')")

    # Suporte a uso direto sem subcomando: python main.py --transcricao ...
    parser.add_argument("--transcricao", "-t", help="Caminho do arquivo TXT")
    parser.add_argument("--participantes", "-p", help="Caminho do arquivo XLSX (opcional)")

    args = parser.parse_args()

    if args.comando == "relatorio":
        consolidar_relatorio(args.pasta, args.periodo)
    elif args.comando == "processar" or args.transcricao:
        transcricao = getattr(args, "transcricao", None)
        participantes = getattr(args, "participantes", None)
        if not transcricao:
            parser.print_help()
            sys.exit(1)
        processar_reuniao(transcricao, participantes)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
