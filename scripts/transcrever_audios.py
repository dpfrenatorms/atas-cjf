"""
TRANSCREVER ÁUDIOS — Minuta de transcrição com interlocutores numerados
Uso:
    python scripts/transcrever_audios.py [arquivos ou links...] [--titulo "..."] [--nome "..."]
        [--pistas pistas.json] [--formatos docx txt srt vtt] [--limpar-audio]

Sem arquivos, usa todos os áudios e vídeos de Entrada/ (mp4, m4a, mp3, ogg, opus, wav, aac, webm,
mov, mkv, flac, wma). A ordem dos arquivos na linha de comando é a ordem das partes na minuta —
ordene pelo conteúdo da reunião, não pela ordem de envio no WhatsApp.

Um link (YouTube e demais sites do yt-dlp) no lugar de um arquivo baixa só o áudio para Entrada/.

Etapas:
    1. Descarta duplicados (compara o MD5 do áudio decodificado, não do arquivo).
    2. Transcreve com ElevenLabs Scribe, com separação de vozes (diarize).
       Cache em Saida/transcricoes/<arquivo>.json — rodar de novo não reenvia o áudio.
       Com --limpar-audio (ruído de fundo, volume desigual), o cache é <arquivo>.limpo.json.
    3. Gera as saídas pedidas em --formatos (padrão: docx):
       docx  Saida/MINUTA_TRANSCRICAO_<slug>_YYYYMMDD.docx, para a redação da ata
       txt   o mesmo nome com .txt, só a transcrição
       srt, vtt  Saida/legendas/<arquivo>.srt|.vtt, uma legenda por áudio

Requer ffmpeg e ffprobe; links requerem também o yt-dlp (pip install yt-dlp).
Chave: ELEVENLABS_API_KEY no ambiente ou no .env do projeto video-vox.
Atenção: o áudio é enviado aos servidores da ElevenLabs.

Arquivo de pistas (opcional), preenchido depois de ler a transcrição:
    {"pistas": {"1": "Juíza federal da SJMS...", "2": "..."},
     "termos": ["\"TRE\" (Parte 1, 24:29): provavelmente TRF", "..."],
     "observacoes": "Ordem das partes definida pelo conteúdo..."}
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unicodedata
from datetime import datetime
from pathlib import Path

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

# Garante saída UTF-8 no terminal Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf-8-sig"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

# ------------------------------------------------------------------
# CONFIGURAÇÃO
# ------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent
ENTRADA_DIR = BASE_DIR / "Entrada"
SAIDA_DIR = BASE_DIR / "Saida"
CACHE_DIR = SAIDA_DIR / "transcricoes"
LEGENDAS_DIR = SAIDA_DIR / "legendas"
ENV_ELEVENLABS = Path(r"C:\Workspace\Projetos\video-vox\.env")

EXTENSOES_AUDIO = {".mp4", ".m4a", ".mp3", ".ogg", ".opus", ".wav", ".aac", ".webm", ".mov", ".mkv", ".flac",
                   ".wma"}
SCRIBE_URL = "https://api.elevenlabs.io/v1/speech-to-text"
SCRIBE_MODELO = "scribe_v1"

# corta o grave sem fala, reduz o ruído constante e nivela o volume entre quem fala perto e longe
FILTRO_LIMPEZA = "highpass=f=80,afftdn=nf=-25,loudnorm=I=-16:TP=-1.5:LRA=11"

FORMATOS = ["docx", "txt", "srt", "vtt"]
LEGENDA_LARGURA = 42      # caracteres por linha
LEGENDA_LINHAS = 2
LEGENDA_SEGUNDOS = 6.0
LEGENDA_PAUSA = 1.0       # silêncio que encerra a legenda
LEGENDA_MINIMO = 1.0      # tempo na tela quando há folga até a legenda seguinte
LEGENDA_PISO = 0.5        # tempo na tela garantido, mesmo sem folga

# O ffmpeg do winget só entra no PATH de terminais abertos depois da instalação
if not shutil.which("ffmpeg"):
    for _bin in (Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "WinGet" / "Packages").glob(
            "Gyan.FFmpeg*/*/bin"):
        os.environ["PATH"] += os.pathsep + str(_bin)


# ------------------------------------------------------------------
# ÁUDIO
# ------------------------------------------------------------------

def baixar_link(url: str) -> Path:
    """Baixa só o áudio de um link para Entrada/ (não baixa de novo se o arquivo já existe)."""
    try:
        from yt_dlp import YoutubeDL
    except ImportError:
        sys.exit("Para transcrever links, instale o yt-dlp: pip install yt-dlp")
    opcoes = {"format": "bestaudio/best", "outtmpl": str(ENTRADA_DIR / "%(title).80s [%(id)s].%(ext)s"),
              "windowsfilenames": True, "noplaylist": True, "quiet": True, "no_warnings": True}
    with YoutubeDL(opcoes) as ydl:
        info = ydl.extract_info(url, download=True)
    return Path(info["requested_downloads"][0]["filepath"])


def duracao(caminho: Path) -> float:
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", str(caminho)],
        capture_output=True, text=True, check=True,
    )
    return float(out.stdout.strip())


def md5_audio(caminho: Path) -> str:
    """MD5 do áudio decodificado — detecta cópias mesmo com metadados diferentes."""
    out = subprocess.run(
        ["ffmpeg", "-v", "error", "-i", str(caminho), "-map", "0:a", "-f", "md5", "-"],
        capture_output=True, text=True, check=True,
    )
    return out.stdout.strip().split("=")[-1]


def separar_duplicados(arquivos):
    unicos, duplicados, vistos = [], [], {}
    for a in arquivos:
        h = md5_audio(a)
        if h in vistos:
            duplicados.append((a, vistos[h]))
        else:
            vistos[h] = a
            unicos.append(a)
    return unicos, duplicados


# ------------------------------------------------------------------
# TRANSCRIÇÃO (ElevenLabs Scribe)
# ------------------------------------------------------------------

def chave_elevenlabs() -> str:
    chave = os.environ.get("ELEVENLABS_API_KEY")
    if not chave and ENV_ELEVENLABS.exists():
        for linha in ENV_ELEVENLABS.read_text(encoding="utf-8").splitlines():
            if linha.startswith("ELEVENLABS_API_KEY="):
                chave = linha.split("=", 1)[1].strip().strip('"')
    if not chave:
        sys.exit(f"ELEVENLABS_API_KEY não encontrada (ambiente ou {ENV_ELEVENLABS}).")
    return chave


def caminho_cache(caminho: Path, limpar: bool) -> Path:
    return CACHE_DIR / f"{caminho.stem}{'.limpo' if limpar else ''}.json"


def transcrever(caminho: Path, chave: str, limpar: bool = False) -> dict:
    cache = caminho_cache(caminho, limpar)
    if cache.exists():
        print(f"  cache: {cache.name}")
        return json.loads(cache.read_text(encoding="utf-8"))

    with tempfile.TemporaryDirectory() as tmp:
        mp3 = Path(tmp) / "audio.mp3"
        # mono 16 kHz reduz o upload (~27 MB por hora) sem perder qualidade de fala
        subprocess.run(["ffmpeg", "-v", "error", "-y", "-i", str(caminho), "-vn",
                        *(["-af", FILTRO_LIMPEZA] if limpar else []),
                        "-ac", "1", "-ar", "16000", "-b:a", "48k", str(mp3)], check=True)
        print(f"  enviando {mp3.stat().st_size / 1e6:.1f} MB à ElevenLabs...")
        with open(mp3, "rb") as f:
            resp = requests.post(
                SCRIBE_URL,
                headers={"xi-api-key": chave},
                data={"model_id": SCRIBE_MODELO, "diarize": "true", "language_code": "por",
                      "tag_audio_events": "false"},
                files={"file": ("audio.mp3", f, "audio/mpeg")},
                timeout=1800,
            )
    if resp.status_code != 200:
        sys.exit(f"ElevenLabs respondeu {resp.status_code}: {resp.text[:300]}")
    dados = resp.json()
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
    return dados


def agrupar_falas(dados: dict):
    """Junta palavras consecutivas do mesmo interlocutor em uma fala."""
    falas, atual = [], None
    for w in dados.get("words", []):
        if w.get("type") == "audio_event":
            continue
        spk = w.get("speaker_id")
        if atual and atual["spk"] == spk:
            atual["texto"] += w["text"]
        else:
            if atual:
                falas.append(atual)
            atual = {"spk": spk, "inicio": w["start"], "texto": w["text"]}
    if atual:
        falas.append(atual)
    for f in falas:
        f["texto"] = f["texto"].strip()
    return [f for f in falas if f["texto"]]


def hms(seg: float) -> str:
    s = int(seg)
    return f"{s // 3600:02d}:{s % 3600 // 60:02d}:{s % 60:02d}"


def duracao_legivel(seg: float) -> str:
    s = int(seg)
    h, m, r = s // 3600, s % 3600 // 60, s % 60
    return f"{h}h{m:02d}min" if h else f"{m}min{r:02d}s"


# ------------------------------------------------------------------
# TXT E LEGENDAS
# ------------------------------------------------------------------

def gerar_txt(partes, titulo, saida: Path):
    linhas = ["MINUTA DE TRANSCRIÇÃO", titulo,
              "Transcrição automática (ElevenLabs Scribe, com separação de vozes), sem revisão humana.", ""]
    for p in partes:
        linhas += [f"{p['nome']} — {p['arquivo'].name} ({duracao_legivel(p['duracao'])})", ""]
        linhas += [f"[{hms(f['inicio'])}] Interlocutor {p['mapa'][f['spk']]}: {f['texto']}" for f in p["falas"]]
        linhas.append("")
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text("\n".join(linhas), encoding="utf-8")


def montar_legendas(dados: dict):
    """Quebra as palavras em legendas curtas, sem misturar interlocutores na mesma legenda."""
    legendas, atual, spk_anterior = [], None, None
    for w in dados.get("words", []):
        if w.get("type") != "word":
            continue
        if atual and (w.get("speaker_id") != atual["spk"]
                      or w["start"] - atual["fim"] > LEGENDA_PAUSA
                      or w["end"] - atual["inicio"] > LEGENDA_SEGUNDOS
                      or len(textwrap.wrap(f"{atual['reserva']}{atual['texto']} {w['text']}",
                                           LEGENDA_LARGURA)) > LEGENDA_LINHAS
                      or (atual["texto"][-1] in ".?!" and len(atual["texto"]) >= LEGENDA_LARGURA)):
            legendas.append(atual)
            spk_anterior, atual = atual["spk"], None
        if atual:
            atual["texto"] += " " + w["text"]
            atual["fim"] = w["end"]
        else:
            spk = w.get("speaker_id")
            # na troca de interlocutor o SRT mostra o rótulo, que ocupa parte da primeira linha
            atual = {"spk": spk, "inicio": w["start"], "fim": w["end"], "texto": w["text"],
                     "reserva": "" if spk == spk_anterior else "Interlocutor 00: "}
    if atual:
        legendas.append(atual)
    # O Scribe dá duração zero a palavras de fala sobreposta: garante o piso na tela, atrasando o início da
    # legenda seguinte quando não há folga, e segura até LEGENDA_MINIMO quando há.
    for anterior, leg in zip([None] + legendas, legendas):
        if anterior:
            leg["inicio"] = max(leg["inicio"], anterior["fim"])
        leg["fim"] = max(leg["fim"], leg["inicio"] + LEGENDA_PISO)
    for leg, prox in zip(legendas, legendas[1:] + [None]):
        alvo = leg["inicio"] + LEGENDA_MINIMO
        leg["fim"] = max(leg["fim"], min(alvo, prox["inicio"]) if prox else alvo)
    return legendas


def tempo_legenda(seg: float, separador: str) -> str:
    ms = round(seg * 1000)
    return f"{ms // 3600000:02d}:{ms % 3600000 // 60000:02d}:{ms % 60000 // 1000:02d}{separador}{ms % 1000:03d}"


def gerar_legenda(parte, formato: str, saida: Path):
    """SRT mostra o interlocutor quando ele muda; VTT marca a voz em todas as legendas (<v>)."""
    sep = "," if formato == "srt" else "."
    blocos, anterior = [], None
    for i, leg in enumerate(montar_legendas(parte["dados"]), 1):
        rotulo = f"Interlocutor {parte['mapa'][leg['spk']]}"
        tempos = f"{tempo_legenda(leg['inicio'], sep)} --> {tempo_legenda(leg['fim'], sep)}"
        if formato == "srt":
            texto = leg["texto"] if leg["spk"] == anterior else f"{rotulo}: {leg['texto']}"
            blocos.append(f"{i}\n{tempos}\n{textwrap.fill(texto, LEGENDA_LARGURA)}")
        else:
            blocos.append(f"{tempos}\n<v {rotulo}>{textwrap.fill(leg['texto'], LEGENDA_LARGURA)}")
        anterior = leg["spk"]
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(("WEBVTT\n\n" if formato == "vtt" else "") + "\n\n".join(blocos) + "\n", encoding="utf-8")


# ------------------------------------------------------------------
# DOCX
# ------------------------------------------------------------------

AZUL = RGBColor(0x1F, 0x3A, 0x5F)
CINZA = RGBColor(0x70, 0x70, 0x70)


def _titulo(doc, texto, nivel):
    p = doc.add_heading(texto, level=nivel)
    for r in p.runs:
        r.font.color.rgb = AZUL


def _tabela(doc, cabecalho, linhas, larguras):
    t = doc.add_table(rows=1, cols=len(cabecalho))
    t.style = "Table Grid"
    for i, c in enumerate(cabecalho):
        r = t.rows[0].cells[i].paragraphs[0].add_run(c)
        r.bold = True
        r.font.size = Pt(9.5)
    for linha in linhas:
        cells = t.add_row().cells
        for i, v in enumerate(linha):
            cells[i].paragraphs[0].add_run(str(v)).font.size = Pt(9.5)
    for row in t.rows:
        for i, w in enumerate(larguras):
            row.cells[i].width = Cm(w)


def _centro(doc, texto, tamanho, bold=False, italic=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(texto)
    r.bold, r.italic = bold, italic
    r.font.size = Pt(tamanho)


def gerar_minuta(partes, duplicados, titulo, pistas, saida: Path):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    st.paragraph_format.space_after = Pt(4)
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.5)
        s.top_margin = s.bottom_margin = Cm(2)

    _centro(doc, "MINUTA DE TRANSCRIÇÃO", 14, bold=True)
    _centro(doc, titulo, 12)
    _centro(doc, "Documento de trabalho para elaboração de ata — não é a ata", 10, italic=True)

    total = max(p["base"] + len(p["mapa"]) - 1 for p in partes)
    _titulo(doc, "1. Como usar esta minuta", 1)
    for t in [
        "A transcrição foi feita automaticamente (ElevenLabs Scribe, com separação de vozes) e não passou por "
        "revisão humana. Nomes próprios, siglas e números devem ser conferidos com o áudio antes de ir para a ata.",
        "Cada voz recebeu um número (Interlocutor 1, 2, 3…). Para identificar alguém, preencha a coluna "
        "\"Nome / cargo\" do quadro do item 3 e depois use Localizar e Substituir no Word (ex.: \"Interlocutor 1\" → "
        f"\"Juíza Federal Fulana\"). Substitua do maior número para o menor (Interlocutor {total} antes de "
        "Interlocutor 1), para não trocar parte de um número maior.",
        "A separação de vozes é automática e pode errar: falas curtas e sobrepostas às vezes aparecem no "
        "interlocutor errado, e a mesma pessoa recebe números diferentes em cada parte. As pistas do quadro "
        "ajudam a conferir.",
        "O horário entre colchetes [hh:mm:ss] é o ponto do respectivo arquivo de áudio onde a fala começa.",
    ]:
        doc.add_paragraph(t, style="List Bullet")

    _titulo(doc, "2. Arquivos de áudio", 1)
    linhas = [[p["arquivo"].name, duracao_legivel(p["duracao"]),
               f"Transcrito — {p['nome']}" + (" (áudio com redução de ruído)" if p["limpo"] else "")]
              for p in partes]
    linhas += [[d.name, duracao_legivel(duracao(d)),
                f"Não transcrito: cópia idêntica de {orig.name} (áudio comparado por inteiro)"]
               for d, orig in duplicados]
    _tabela(doc, ["Arquivo", "Duração", "Situação"], linhas, [7, 2.2, 6.8])
    if pistas.get("observacoes"):
        doc.add_paragraph()
        doc.add_paragraph(pistas["observacoes"])

    _titulo(doc, "3. Quadro de interlocutores", 1)
    linhas = []
    for p in partes:
        for spk, n in p["mapa"].items():
            fs = [f for f in p["falas"] if f["spk"] == spk]
            palavras = sum(len(f["texto"].split()) for f in fs)
            linhas.append([f"Interlocutor {n}", p["nome"], f"{len(fs)} / {palavras}", hms(fs[0]["inicio"]),
                           pistas.get("pistas", {}).get(str(n), ""), ""])
    _tabela(doc, ["Nº", "Parte", "Falas / palavras", "1ª fala", "Pistas para identificação", "Nome / cargo"],
            linhas, [2.2, 1.4, 1.8, 1.6, 6.5, 2.5])

    if pistas.get("termos"):
        _titulo(doc, "4. Termos a conferir no áudio", 1)
        for t in pistas["termos"]:
            doc.add_paragraph(t, style="List Bullet")

    _titulo(doc, f"{5 if pistas.get('termos') else 4}. Transcrição", 1)
    for p in partes:
        _titulo(doc, f"{p['nome']} — {p['arquivo'].name} ({duracao_legivel(p['duracao'])})", 2)
        for f in p["falas"]:
            par = doc.add_paragraph()
            par.paragraph_format.space_after = Pt(5)
            r = par.add_run(f"[{hms(f['inicio'])}] ")
            r.font.size = Pt(9)
            r.font.color.rgb = CINZA
            par.add_run(f"Interlocutor {p['mapa'][f['spk']]}: ").bold = True
            par.add_run(f["texto"])

    saida.parent.mkdir(parents=True, exist_ok=True)
    doc.save(saida)


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Transcreve áudios de reunião e gera minuta para ata.")
    ap.add_argument("arquivos", nargs="*",
                    help="Áudios, vídeos ou links na ordem das partes (padrão: todos de Entrada/)")
    ap.add_argument("--titulo", default="Reunião — Conselho da Justiça Federal")
    ap.add_argument("--pistas", help="JSON com pistas por interlocutor, termos a conferir e observações")
    ap.add_argument("--nome", help="Nome curto para o arquivo de saída (padrão: início do título)")
    ap.add_argument("--data", default=datetime.now().strftime("%Y%m%d"), help="Data no nome do arquivo (YYYYMMDD)")
    ap.add_argument("--formatos", nargs="+", choices=FORMATOS, default=["docx"],
                    help="Saídas a gerar (padrão: docx)")
    ap.add_argument("--limpar-audio", action="store_true",
                    help="Reduz ruído e nivela o volume antes de transcrever (só para gravação ruim)")
    args = ap.parse_args()

    if not shutil.which("ffmpeg") or not shutil.which("ffprobe"):
        sys.exit("ffmpeg/ffprobe não encontrados. Instale com: winget install Gyan.FFmpeg")

    if args.arquivos:
        arquivos = []
        for a in args.arquivos:
            if re.match(r"https?://", a):
                print(f"Baixando áudio: {a}")
                arquivos.append(baixar_link(a))
            else:
                arquivos.append(Path(a) if Path(a).exists() else ENTRADA_DIR / a)
    else:
        arquivos = sorted(a for a in ENTRADA_DIR.iterdir() if a.suffix.lower() in EXTENSOES_AUDIO)
    faltando = [a for a in arquivos if not a.exists()]
    if faltando or not arquivos:
        sys.exit(f"Áudios não encontrados: {faltando or ENTRADA_DIR}")

    print("Verificando duplicados...")
    unicos, duplicados = separar_duplicados(arquivos)
    for d, orig in duplicados:
        print(f"  duplicado: {d.name} = {orig.name}")

    chave = None
    partes, base = [], 1
    for i, arq in enumerate(unicos, 1):
        print(f"Parte {i}: {arq.name}")
        if not caminho_cache(arq, args.limpar_audio).exists() and chave is None:
            chave = chave_elevenlabs()
        dados = transcrever(arq, chave, args.limpar_audio)
        falas = agrupar_falas(dados)
        ids = sorted({f["spk"] for f in falas}, key=lambda s: int(s.split("_")[-1]))
        mapa = {s: base + j for j, s in enumerate(ids)}
        partes.append({"nome": f"Parte {i}", "arquivo": arq, "duracao": duracao(arq), "falas": falas,
                       "mapa": mapa, "base": base, "dados": dados, "limpo": args.limpar_audio})
        print(f"  {len(falas)} falas, interlocutores {base} a {base + len(ids) - 1}")
        base += len(ids)

    pistas = json.loads(Path(args.pistas).read_text(encoding="utf-8")) if args.pistas else {}
    sem_acento = unicodedata.normalize("NFKD", args.nome or args.titulo.split("—")[0]).encode("ascii", "ignore").decode()
    slug = re.sub(r"[^A-Za-z0-9]+", "_", sem_acento).strip("_")[:40] or "Reuniao"
    saida = SAIDA_DIR / f"MINUTA_TRANSCRICAO_{slug}_{args.data}.docx"
    if "docx" in args.formatos:
        gerar_minuta(partes, duplicados, args.titulo, pistas, saida)
        print(f"Minuta salva: {saida}")
    if "txt" in args.formatos:
        gerar_txt(partes, args.titulo, saida.with_suffix(".txt"))
        print(f"Texto salvo: {saida.with_suffix('.txt')}")
    for formato in ("srt", "vtt"):
        if formato in args.formatos:
            for p in partes:
                legenda = LEGENDAS_DIR / f"{p['arquivo'].stem}.{formato}"
                gerar_legenda(p, formato, legenda)
                print(f"Legenda salva: {legenda}")


if __name__ == "__main__":
    main()
