import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import json
import time
import urllib.request
import urllib.parse
from pathlib import Path
from datetime import datetime


# ── API REST do LanguageTool ─────────────────────────────────────────────────
API_URL = "https://api.languagetool.org/v2/check"
IDIOMA  = "pt-BR"
PAUSA_ENTRE_BLOCOS = 1.5
TAMANHO_BLOCO = 1500


def _chamar_api(texto: str) -> list[dict]:
    dados = urllib.parse.urlencode({
        "text":     texto,
        "language": IDIOMA,
    }).encode("utf-8")
    req = urllib.request.Request(
        API_URL, data=dados,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept":       "application/json",
            "User-Agent":   "CorretorPT/1.0",
        }
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        resultado = json.loads(resp.read().decode("utf-8"))
    return resultado.get("matches", [])


def _aplicar_correcoes(texto: str, matches: list[dict]) -> tuple[str, list[dict]]:
    erros = []
    for m in sorted(matches, key=lambda x: x["offset"], reverse=True):
        offset = m["offset"]
        length = m["length"]
        repls  = [r["value"] for r in m.get("replacements", [])]
        trecho = texto[offset: offset + length]
        erros.append({
            "trecho":    trecho,
            "mensagem":  m.get("message", ""),
            "sugestoes": repls[:3],
        })
        if repls:
            texto = texto[:offset] + repls[0] + texto[offset + length:]
    erros.reverse()
    return texto, erros


def _dividir_em_blocos(texto: str, tamanho: int = TAMANHO_BLOCO) -> list[str]:
    paragrafos = texto.split("\n")
    blocos, atual = [], ""
    for p in paragrafos:
        if len(atual) + len(p) + 1 > tamanho and atual:
            blocos.append(atual)
            atual = p + "\n"
        else:
            atual += p + "\n"
    if atual.strip():
        blocos.append(atual)
    return blocos or [texto]


def corrigir_texto_completo(texto: str, callback=None) -> tuple[str, list[dict]]:
    blocos = _dividir_em_blocos(texto)
    resultado, todos_erros = [], []
    for i, bloco in enumerate(blocos):
        if callback:
            callback(i, len(blocos))
        matches = _chamar_api(bloco)
        bloco_corrigido, erros = _aplicar_correcoes(bloco, matches)
        resultado.append(bloco_corrigido)
        todos_erros.extend(erros)
        if i < len(blocos) - 1:
            time.sleep(PAUSA_ENTRE_BLOCOS)
    return "".join(resultado), todos_erros


def processar_txt(caminho_entrada: str, callback_progresso):
    callback_progresso("Lendo arquivo .txt…", 10)
    with open(caminho_entrada, "r", encoding="utf-8", errors="replace") as f:
        texto = f.read()

    def cb(i, total):
        pct = 20 + int((i / max(total, 1)) * 65)
        callback_progresso(f"Corrigindo bloco {i+1}/{total}…", pct)

    texto_corrigido, erros = corrigir_texto_completo(texto, cb)

    caminho_saida = _caminho_saida(caminho_entrada, ".txt")
    callback_progresso("Salvando arquivo corrigido…", 90)
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(texto_corrigido)
    return caminho_saida, erros


def processar_docx(caminho_entrada: str, callback_progresso):
    import docx as _docx
    callback_progresso("Lendo arquivo .docx…", 10)
    doc = _docx.Document(caminho_entrada)
    paragrafos_validos = [p for p in doc.paragraphs if p.text.strip()]
    total = len(paragrafos_validos)
    todos_erros = []

    for i, paragrafo in enumerate(paragrafos_validos):
        pct = 15 + int((i / max(total, 1)) * 70)
        callback_progresso(f"Corrigindo parágrafo {i+1}/{total}…", pct)
        matches = _chamar_api(paragrafo.text)
        texto_corrigido, erros = _aplicar_correcoes(paragrafo.text, matches)
        todos_erros.extend(erros)
        if paragrafo.runs:
            paragrafo.runs[0].text = texto_corrigido
            for run in paragrafo.runs[1:]:
                run.text = ""
        else:
            paragrafo.text = texto_corrigido
        if i < total - 1:
            time.sleep(PAUSA_ENTRE_BLOCOS)

    caminho_saida = _caminho_saida(caminho_entrada, ".docx")
    callback_progresso("Salvando arquivo corrigido…", 92)
    doc.save(caminho_saida)
    return caminho_saida, todos_erros


def _caminho_saida(caminho_entrada: str, extensao: str) -> str:
    p  = Path(caminho_entrada)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return str(p.parent / f"{p.stem}_corrigido_{ts}{extensao}")


def verificar_dependencias() -> list[str]:
    faltando = []
    try:
        import docx  # noqa: F401
    except ImportError:
        faltando.append("python-docx")
    return faltando


# ── Interface gráfica ─────────────────────────────────────────────────────────
class CorretorApp(tk.Tk):
    COR_FUNDO       = "#1e1e2e"
    COR_PAINEL      = "#2a2a3d"
    COR_DESTAQUE    = "#7c3aed"
    COR_DESTAQUE2   = "#a855f7"
    COR_TEXTO       = "#e2e8f0"
    COR_SUBTEXT     = "#94a3b8"
    COR_SUCESSO     = "#22c55e"
    COR_ERRO        = "#ef4444"
    COR_AVISO       = "#f59e0b"
    FONTE_NORMAL    = ("Segoe UI", 10)
    FONTE_TITULO    = ("Segoe UI", 18, "bold")
    FONTE_SUBTITULO = ("Segoe UI", 11)
    FONTE_MONO      = ("Consolas", 9)

    def __init__(self):
        super().__init__()
        self.title("Corretor Ortográfico e Gramatical — Português")
        self.geometry("920x700")
        self.minsize(750, 560)
        self.configure(bg=self.COR_FUNDO)
        self.resizable(True, True)
        self.arquivo_selecionado = tk.StringVar()
        self._pronto = False
        self._construir_ui()
        self._verificar_deps_e_iniciar()

    def _construir_ui(self):
        self._cabecalho()
        self._secao_arquivo()
        self._secao_progresso()
        self._secao_log()
        self._rodape()

    def _cabecalho(self):
        frm = tk.Frame(self, bg=self.COR_PAINEL, pady=18)
        frm.pack(fill="x")
        tk.Label(frm, text="✦  Corretor PT-BR", font=self.FONTE_TITULO,
                 bg=self.COR_PAINEL, fg=self.COR_DESTAQUE2).pack()
        tk.Label(frm,
                 text="Ortografia · Gramática · Acentuação · Pontuação  —  via API LanguageTool",
                 font=self.FONTE_SUBTITULO, bg=self.COR_PAINEL,
                 fg=self.COR_SUBTEXT).pack(pady=(2, 0))

    def _secao_arquivo(self):
        frm = tk.LabelFrame(self, text=" Arquivo ", font=self.FONTE_NORMAL,
                            bg=self.COR_FUNDO, fg=self.COR_SUBTEXT,
                            bd=1, relief="solid")
        frm.pack(fill="x", padx=20, pady=(14, 6))
        linha = tk.Frame(frm, bg=self.COR_FUNDO)
        linha.pack(fill="x", padx=10, pady=8)
        self._entry_arquivo = tk.Entry(
            linha, textvariable=self.arquivo_selecionado,
            font=self.FONTE_NORMAL, bg=self.COR_PAINEL, fg=self.COR_TEXTO,
            insertbackground=self.COR_TEXTO, relief="flat", bd=6, state="readonly")
        self._entry_arquivo.pack(side="left", fill="x", expand=True)
        self._btn_selecionar = tk.Button(
            linha, text="📂  Selecionar", font=self.FONTE_NORMAL,
            bg=self.COR_DESTAQUE, fg="white",
            activebackground=self.COR_DESTAQUE2, activeforeground="white",
            relief="flat", bd=0, padx=12, pady=4, cursor="hand2",
            state="disabled", command=self._selecionar_arquivo)
        self._btn_selecionar.pack(side="left", padx=(8, 0))
        self._btn_corrigir = tk.Button(
            frm, text="⚡  Corrigir Arquivo",
            font=("Segoe UI", 11, "bold"),
            bg=self.COR_SUCESSO, fg="white",
            activebackground="#16a34a", activeforeground="white",
            relief="flat", bd=0, padx=18, pady=7, cursor="hand2",
            state="disabled", command=self._iniciar_correcao)
        self._btn_corrigir.pack(pady=(0, 10))

    def _secao_progresso(self):
        frm = tk.Frame(self, bg=self.COR_FUNDO)
        frm.pack(fill="x", padx=20, pady=4)
        self._lbl_status = tk.Label(
            frm, text="Verificando conexão com a API…",
            font=self.FONTE_NORMAL, bg=self.COR_FUNDO,
            fg=self.COR_SUBTEXT, anchor="w")
        self._lbl_status.pack(fill="x")
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Roxo.Horizontal.TProgressbar",
                         troughcolor=self.COR_PAINEL,
                         background=self.COR_DESTAQUE,
                         bordercolor=self.COR_PAINEL,
                         lightcolor=self.COR_DESTAQUE2,
                         darkcolor=self.COR_DESTAQUE)
        self._barra = ttk.Progressbar(
            frm, style="Roxo.Horizontal.TProgressbar",
            orient="horizontal", length=200, mode="determinate")
        self._barra.pack(fill="x", pady=(4, 0))

    def _secao_log(self):
        frm = tk.LabelFrame(self, text=" Log de Erros Detectados ",
                            font=self.FONTE_NORMAL,
                            bg=self.COR_FUNDO, fg=self.COR_SUBTEXT,
                            bd=1, relief="solid")
        frm.pack(fill="both", expand=True, padx=20, pady=(10, 6))
        self._log = scrolledtext.ScrolledText(
            frm, font=self.FONTE_MONO,
            bg=self.COR_PAINEL, fg=self.COR_TEXTO,
            insertbackground=self.COR_TEXTO,
            selectbackground=self.COR_DESTAQUE,
            relief="flat", bd=0, wrap="word", state="disabled")
        self._log.pack(fill="both", expand=True, padx=6, pady=6)
        self._log.tag_config("info",     foreground=self.COR_SUBTEXT)
        self._log.tag_config("ok",       foreground=self.COR_SUCESSO)
        self._log.tag_config("erro",     foreground=self.COR_ERRO)
        self._log.tag_config("aviso",    foreground=self.COR_AVISO)
        self._log.tag_config("destaque", foreground=self.COR_DESTAQUE2,
                              font=("Consolas", 9, "bold"))

    def _rodape(self):
        frm = tk.Frame(self, bg=self.COR_PAINEL, pady=6)
        frm.pack(fill="x", side="bottom")
        tk.Label(frm,
                 text="Arquivo corrigido salvo na mesma pasta com sufixo _corrigido_YYYYMMDD_HHMMSS  •  Requer internet",
                 font=("Segoe UI", 8), bg=self.COR_PAINEL, fg=self.COR_SUBTEXT).pack()

    def _log_escrever(self, texto: str, tag: str = "info"):
        self._log.configure(state="normal")
        self._log.insert("end", texto + "\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _set_status(self, msg: str, pct: int = None):
        self._lbl_status.configure(text=msg)
        if pct is not None:
            self._barra["value"] = pct
        self.update_idletasks()

    def _verificar_deps_e_iniciar(self):
        faltando = verificar_dependencias()
        if faltando:
            self._log_escrever("❌  Dependências não instaladas:", "erro")
            for lib in faltando:
                self._log_escrever(f"    python -m pip install {lib}", "aviso")
            self._log_escrever("\nInstale e reinicie o aplicativo.", "erro")
            self._set_status("Dependências ausentes — veja o log.", 0)
            return
        threading.Thread(target=self._testar_api, daemon=True).start()

    def _testar_api(self):
        self._set_status("Testando conexão com a API do LanguageTool…", 5)
        self._log_escrever("⏳  Testando API do LanguageTool…", "info")
        try:
            _chamar_api("Testando conexão.")
            self._pronto = True
            self._set_status("API pronta. Selecione um arquivo para começar.", 0)
            self._log_escrever("✅  API do LanguageTool acessível (pt-BR).", "ok")
            self._log_escrever("ℹ️   Sem Java necessário — usa API REST via internet.", "info")
            self._btn_selecionar.configure(state="normal")
        except Exception as e:
            self._log_escrever(f"❌  Erro ao conectar com a API: {e}", "erro")
            self._log_escrever("    Verifique sua conexão com a internet.", "aviso")
            self._set_status("Sem conexão com a API. Veja o log.", 0)

    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo para corrigir",
            filetypes=[("Documentos", "*.txt *.docx"),
                       ("Texto simples", "*.txt"),
                       ("Word", "*.docx"),
                       ("Todos", "*.*")])
        if not caminho:
            return
        self.arquivo_selecionado.set(caminho)
        ext = Path(caminho).suffix.lower()
        if ext in (".txt", ".docx") and self._pronto:
            self._btn_corrigir.configure(state="normal")
            self._log_escrever(
                f"\n📄  Arquivo selecionado: {os.path.basename(caminho)}", "destaque")
        elif ext not in (".txt", ".docx"):
            messagebox.showwarning("Formato inválido",
                                   "Apenas arquivos .txt e .docx são suportados.")
            self._btn_corrigir.configure(state="disabled")

    def _iniciar_correcao(self):
        caminho = self.arquivo_selecionado.get()
        if not caminho:
            return
        self._btn_corrigir.configure(state="disabled")
        self._btn_selecionar.configure(state="disabled")
        self._barra["value"] = 0
        threading.Thread(target=self._executar_correcao,
                         args=(caminho,), daemon=True).start()

    def _executar_correcao(self, caminho: str):
        ext = Path(caminho).suffix.lower()
        try:
            self._log_escrever("\n" + "─" * 54, "info")
            self._log_escrever(
                f"🔍  Iniciando correção: {os.path.basename(caminho)}", "destaque")

            def progresso(msg, pct):
                self._set_status(msg, pct)
                self._log_escrever(f"   {msg}", "info")

            if ext == ".txt":
                caminho_saida, erros = processar_txt(caminho, progresso)
            else:
                caminho_saida, erros = processar_docx(caminho, progresso)

            self._set_status("Concluído!", 100)
            self._relatar_erros(erros)
            self._log_escrever(f"\n✅  Arquivo salvo em:\n   {caminho_saida}", "ok")
            self._log_escrever("─" * 54, "info")
            messagebox.showinfo(
                "Correção concluída",
                f"Arquivo corrigido salvo em:\n{caminho_saida}\n\n"
                f"Total de correções aplicadas: {len(erros)}")
        except Exception as e:
            self._log_escrever(f"\n❌  Erro durante a correção: {e}", "erro")
            self._set_status("Erro durante a correção. Veja o log.", 0)
            messagebox.showerror("Erro", str(e))
        finally:
            self._btn_corrigir.configure(state="normal")
            self._btn_selecionar.configure(state="normal")

    def _relatar_erros(self, erros: list[dict]):
        if not erros:
            self._log_escrever("\n🎉  Nenhum erro encontrado!", "ok")
            return
        self._log_escrever(f"\n⚠️   {len(erros)} correção(ões) aplicada(s):", "aviso")
        for i, e in enumerate(erros[:60], 1):
            sug    = ", ".join(e["sugestoes"]) if e["sugestoes"] else "sem sugestão"
            trecho = repr(e["trecho"])[:40]
            msg    = e["mensagem"][:80]
            self._log_escrever(
                f"  {i:>3}. {trecho}  →  {sug}\n       ({msg})", "aviso")
        if len(erros) > 60:
            self._log_escrever(
                f"  … e mais {len(erros)-60} correção(ões) não exibida(s).", "info")


if __name__ == "__main__":
    app = CorretorApp()
    app.mainloop()