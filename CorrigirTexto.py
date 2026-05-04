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


# ── Interface gráfica Refinada ────────────────────────────────────────────────
class CorretorApp(tk.Tk):
    COR_FUNDO       = "#121214"  # Dark theme limpo e moderno
    COR_PAINEL      = "#202024"
    COR_DESTAQUE    = "#8257e5"  # Roxo vibrante
    COR_DESTAQUE_H  = "#996dff"  # Roxo hover
    COR_TEXTO       = "#e1e1e6"
    COR_SUBTEXT     = "#a8a8b3"
    COR_SUCESSO     = "#04d361"
    COR_SUCESSO_H   = "#05e66a"
    COR_ERRO        = "#f75a68"
    COR_AVISO       = "#eba417"
    FONTE_NORMAL    = ("Segoe UI", 10)
    FONTE_TITULO    = ("Segoe UI", 18, "bold")
    FONTE_SUBTITULO = ("Segoe UI", 10)
    FONTE_MONO      = ("Consolas", 10)

    def __init__(self):
        super().__init__()
        self.title("Corretor Ortográfico e Gramatical")
        self.geometry("920x720")
        self.minsize(750, 560)
        self.configure(bg=self.COR_FUNDO)
        self.resizable(True, True)
        self.arquivo_selecionado = tk.StringVar()
        self._pronto = False
        self._construir_ui()
        self._verificar_deps_e_iniciar()

    def _add_hover(self, widget, cor_normal, cor_hover):
        def on_enter(e):
            if widget['state'] != 'disabled':
                widget.config(bg=cor_hover)
        def on_leave(e):
            if widget['state'] != 'disabled':
                widget.config(bg=cor_normal)
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def _construir_ui(self):
        self._cabecalho()
        self._secao_arquivo()
        self._secao_progresso()
        self._secao_log()
        self._rodape()

    def _cabecalho(self):
        frm = tk.Frame(self, bg=self.COR_PAINEL, pady=25)
        frm.pack(fill="x")
        tk.Label(frm, text="✦  CORRETOR PT-BR", font=self.FONTE_TITULO,
                 bg=self.COR_PAINEL, fg=self.COR_DESTAQUE).pack()
        tk.Label(frm,
                 text="Ortografia · Gramática · Acentuação · Pontuação",
                 font=self.FONTE_SUBTITULO, bg=self.COR_PAINEL,
                 fg=self.COR_SUBTEXT).pack(pady=(5, 0))

    def _secao_arquivo(self):
        frm = tk.Frame(self, bg=self.COR_FUNDO)
        frm.pack(fill="x", padx=40, pady=(30, 10))
        
        tk.Label(frm, text="ARQUIVO DE ENTRADA", font=("Segoe UI", 8, "bold"),
                 bg=self.COR_FUNDO, fg=self.COR_SUBTEXT, anchor="w").pack(fill="x", pady=(0, 5))

        linha = tk.Frame(frm, bg=self.COR_FUNDO)
        linha.pack(fill="x")
        
        self._entry_arquivo = tk.Entry(
            linha, textvariable=self.arquivo_selecionado,
            font=self.FONTE_NORMAL, bg=self.COR_PAINEL, fg=self.COR_TEXTO,
            insertbackground=self.COR_TEXTO, relief="flat", bd=10, state="readonly")
        self._entry_arquivo.pack(side="left", fill="x", expand=True)
        
        self._btn_selecionar = tk.Button(
            linha, text="📂 Selecionar", font=("Segoe UI", 10, "bold"),
            bg=self.COR_PAINEL, fg=self.COR_TEXTO,
            activebackground=self.COR_PAINEL, activeforeground="white",
            relief="flat", bd=0, padx=20, pady=8, cursor="hand2",
            state="disabled", command=self._selecionar_arquivo)
        self._btn_selecionar.pack(side="left", padx=(10, 0))
        self._add_hover(self._btn_selecionar, self.COR_PAINEL, "#323238")

        self._btn_corrigir = tk.Button(
            frm, text="⚡ CORRIGIR ARQUIVO",
            font=("Segoe UI", 11, "bold"),
            bg=self.COR_SUCESSO, fg="white",
            activebackground=self.COR_SUCESSO_H, activeforeground="white",
            relief="flat", bd=0, padx=20, pady=12, cursor="hand2",
            state="disabled", command=self._iniciar_correcao)
        self._btn_corrigir.pack(pady=(15, 0), fill="x")
        self._add_hover(self._btn_corrigir, self.COR_SUCESSO, self.COR_SUCESSO_H)

    def _secao_progresso(self):
        frm = tk.Frame(self, bg=self.COR_FUNDO)
        frm.pack(fill="x", padx=40, pady=10)
        self._lbl_status = tk.Label(
            frm, text="Verificando conexão com a API…",
            font=self.FONTE_NORMAL, bg=self.COR_FUNDO,
            fg=self.COR_SUBTEXT, anchor="w")
        self._lbl_status.pack(fill="x")
        
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Slim.Horizontal.TProgressbar",
                         troughcolor=self.COR_PAINEL,
                         background=self.COR_DESTAQUE,
                         bordercolor=self.COR_PAINEL,
                         lightcolor=self.COR_DESTAQUE,
                         darkcolor=self.COR_DESTAQUE,
                         thickness=4) # Barra bem mais fina e elegante
        
        self._barra = ttk.Progressbar(
            frm, style="Slim.Horizontal.TProgressbar",
            orient="horizontal", length=200, mode="determinate")
        self._barra.pack(fill="x", pady=(8, 0))

    def _secao_log(self):
        frm = tk.Frame(self, bg=self.COR_FUNDO)
        frm.pack(fill="both", expand=True, padx=40, pady=(10, 20))
        
        tk.Label(frm, text="LOG DE PROCESSAMENTO", font=("Segoe UI", 8, "bold"),
                 bg=self.COR_FUNDO, fg=self.COR_SUBTEXT, anchor="w").pack(fill="x", pady=(0, 5))

        self._log = scrolledtext.ScrolledText(
            frm, font=self.FONTE_MONO,
            bg=self.COR_PAINEL, fg=self.COR_TEXTO,
            insertbackground=self.COR_TEXTO,
            selectbackground=self.COR_DESTAQUE,
            relief="flat", bd=12, wrap="word", state="disabled")
        self._log.pack(fill="both", expand=True)
        
        self._log.tag_config("info",     foreground=self.COR_SUBTEXT)
        self._log.tag_config("ok",       foreground=self.COR_SUCESSO)
        self._log.tag_config("erro",     foreground=self.COR_ERRO)
        self._log.tag_config("aviso",    foreground=self.COR_AVISO)
        self._log.tag_config("destaque", foreground=self.COR_DESTAQUE_H,
                              font=("Consolas", 10, "bold"))

    def _rodape(self):
        frm = tk.Frame(self, bg=self.COR_FUNDO, pady=10)
        frm.pack(fill="x", side="bottom")
        tk.Label(frm,
                 text="Salva uma cópia com sufixo _corrigido na pasta original  •  Requer internet",
                 font=("Segoe UI", 8), bg=self.COR_FUNDO, fg=self.COR_SUBTEXT).pack()

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
            self._log_escrever("❌ Dependências não instaladas:", "erro")
            for lib in faltando:
                self._log_escrever(f"   python -m pip install {lib}", "aviso")
            self._log_escrever("\nInstale e reinicie o aplicativo.", "erro")
            self._set_status("Dependências ausentes.", 0)
            return
        threading.Thread(target=self._testar_api, daemon=True).start()

    def _testar_api(self):
        self._set_status("Testando conexão com a API do LanguageTool…", 5)
        self._log_escrever("⏳ Testando API do LanguageTool…", "info")
        try:
            _chamar_api("Testando conexão.")
            self._pronto = True
            self._set_status("Pronto. Selecione um arquivo para começar.", 0)
            self._log_escrever("✅ Conectado com sucesso (pt-BR).", "ok")
            self._btn_selecionar.configure(state="normal")
        except Exception as e:
            self._log_escrever(f"❌ Erro de conexão: {e}", "erro")
            self._set_status("Falha na conexão.", 0)

    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo",
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
            self._log_escrever(f"\n📄 Selecionado: {os.path.basename(caminho)}", "destaque")
        elif ext not in (".txt", ".docx"):
            messagebox.showwarning("Aviso", "Apenas arquivos .txt e .docx são suportados.")
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
            self._log_escrever("\n" + "─" * 50, "info")
            self._log_escrever(f"🔍 Iniciando correção: {os.path.basename(caminho)}", "destaque")

            def progresso(msg, pct):
                self._set_status(msg, pct)
                self._log_escrever(f"   {msg}", "info")

            if ext == ".txt":
                caminho_saida, erros = processar_txt(caminho, progresso)
            else:
                caminho_saida, erros = processar_docx(caminho, progresso)

            self._set_status("Concluído!", 100)
            self._relatar_erros(erros)
            self._log_escrever(f"\n✅ Salvo em:\n   {caminho_saida}", "ok")
            self._log_escrever("─" * 50, "info")
        except Exception as e:
            self._log_escrever(f"\n❌ Erro: {e}", "erro")
            self._set_status("Erro durante a execução.", 0)
        finally:
            self._btn_corrigir.configure(state="normal")
            self._btn_selecionar.configure(state="normal")

    def _relatar_erros(self, erros: list[dict]):
        if not erros:
            self._log_escrever("\n🎉 Nenhum erro encontrado!", "ok")
            return
        self._log_escrever(f"\n⚠️  {len(erros)} correções aplicadas:", "aviso")
        for i, e in enumerate(erros[:60], 1):
            sug    = ", ".join(e["sugestoes"]) if e["sugestoes"] else "sem sugestão"
            trecho = repr(e["trecho"])[:40]
            msg    = e["mensagem"][:80]
            self._log_escrever(f"  {i:>3}. {trecho}  →  {sug}\n       ({msg})", "aviso")
        if len(erros) > 60:
            self._log_escrever(f"  … e mais {len(erros)-60} ocultas.", "info")


if __name__ == "__main__":
    app = CorretorApp()
    app.mainloop()