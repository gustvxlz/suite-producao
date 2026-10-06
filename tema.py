"""Tema visual e utilitários compartilhados pela Suite de Produção."""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox

# Paleta escura leve, pensada para boa leitura em monitores de escritório.
BG = "#111318"
BG_PANEL = "#191d24"
BG_CARD = "#1f242d"
BG_ENTRY = "#171b22"
FG = "#edf1f7"
FG_MUTED = "#98a2b3"
BORDER = "#2b3240"
SELECT = "#264b6a"
ACCENT = "#58a6ff"
ACCENT_HOVER = "#79b8ff"
SUCCESS = "#5cc489"
WARNING = "#e4b85b"
DANGER = "#ef7a7a"
# Cores de ação derivadas da própria paleta: destaque sem poluir a interface.
ACTION = ACCENT
ACTION_HOVER = ACCENT_HOVER
COMMIT = SUCCESS
COMMIT_HOVER = "#76d89f"
OUTPUT = BG_PANEL
OUTPUT_HOVER = "#27303c"

TAG_CORES = {
    "ok": ("#173325", "#b9f4d0"),
    "revisar": ("#403617", "#ffe29a"),
    "pendente": ("#3f2024", "#ffc1c5"),
    "excluido": ("#252931", "#8d96a6"),
}


def aplicar_tema_escuro(root):
    root.configure(background=BG)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    fonte = ("Segoe UI", 10)
    style.configure(".", background=BG, foreground=FG, fieldbackground=BG_ENTRY,
                    bordercolor=BORDER, focuscolor=ACCENT, lightcolor=BG, darkcolor=BG,
                    font=fonte)
    style.configure("TFrame", background=BG)
    style.configure("Panel.TFrame", background=BG_PANEL)
    style.configure("Toolbar.TFrame", background=BG_PANEL)
    style.configure("Card.TFrame", background=BG_CARD, relief="flat")
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("Panel.TLabel", background=BG_PANEL, foreground=FG)
    style.configure("Muted.TLabel", background=BG, foreground=FG_MUTED)
    style.configure("PanelMuted.TLabel", background=BG_PANEL, foreground=FG_MUTED)
    style.configure("Card.TLabel", background=BG_CARD, foreground=FG)
    style.configure("CardTitle.TLabel", background=BG_CARD, foreground=FG, font=("Segoe UI Semibold", 13))
    style.configure("CardDesc.TLabel", background=BG_CARD, foreground=FG_MUTED, font=("Segoe UI", 9))
    style.configure("Header.TLabel", background=BG, foreground=FG, font=("Segoe UI Semibold", 19))
    style.configure("Section.TLabel", background=BG, foreground=FG, font=("Segoe UI Semibold", 11))
    style.configure("SubHeader.TLabel", background=BG, foreground=FG_MUTED, font=("Segoe UI", 10))
    style.configure("Status.TLabel", background=BG_PANEL, foreground=FG_MUTED, padding=(10, 6))

    style.configure("TLabelframe", background=BG, foreground=FG, bordercolor=BORDER, relief="solid")
    style.configure("TLabelframe.Label", background=BG, foreground=FG, font=("Segoe UI Semibold", 10))

    style.configure("TButton", background=BG_PANEL, foreground=FG, bordercolor=BORDER, padding=(10, 7), relief="flat")
    style.map("TButton", background=[("active", "#27303c"), ("pressed", "#303a48"), ("disabled", BG_PANEL)],
              foreground=[("disabled", "#677080")])
    style.configure("Small.TButton", padding=(8, 5), font=("Segoe UI", 9))
    style.configure("Accent.TButton", background=ACCENT, foreground="#09111c", bordercolor=ACCENT,
                    padding=(11, 8), font=("Segoe UI Semibold", 10))
    style.map("Accent.TButton", background=[("active", ACCENT_HOVER), ("pressed", "#91c5ff"), ("disabled", "#334252")],
              foreground=[("disabled", "#788391")])
    style.configure("Success.TButton", background=SUCCESS, foreground="#07140d", bordercolor=SUCCESS,
                    padding=(10, 7), font=("Segoe UI Semibold", 10))
    style.map("Success.TButton", background=[("active", "#76d89f")])

    # Ações principais compactas: azul = processar/gerar, verde = gravar/lançar.
    # As variantes desabilitadas continuam levemente coloridas para serem fáceis de localizar.
    style.configure("PrimaryAction.TButton", background=ACCENT, foreground="#09111c", bordercolor=ACCENT,
                    padding=(10, 6), font=("Segoe UI Semibold", 9))
    style.map("PrimaryAction.TButton",
              background=[("active", ACCENT_HOVER), ("pressed", "#91c5ff"), ("disabled", "#273c54")],
              foreground=[("disabled", "#8da4bd")])
    style.configure("LaunchAction.TButton", background=SUCCESS, foreground="#07140d", bordercolor=SUCCESS,
                    padding=(10, 6), font=("Segoe UI Semibold", 9))
    style.map("LaunchAction.TButton",
              background=[("active", "#76d89f"), ("pressed", "#8ee4ae"), ("disabled", "#294536")],
              foreground=[("disabled", "#9cb9a8")])
    style.configure("Danger.TButton", background=BG_PANEL, foreground=DANGER, bordercolor="#604147", padding=(9, 7))
    style.map("Danger.TButton", background=[("active", "#512d33")], foreground=[("active", "#ffd7da")])

    style.configure("TCheckbutton", background=BG, foreground=FG)
    style.map("TCheckbutton", background=[("active", BG)])
    style.configure("TEntry", fieldbackground=BG_ENTRY, foreground=FG, insertcolor=FG, bordercolor=BORDER, padding=6)
    style.configure("TSpinbox", fieldbackground=BG_ENTRY, foreground=FG, insertcolor=FG, background=BG_ENTRY,
                    arrowcolor=FG, bordercolor=BORDER, padding=5)
    style.map("TSpinbox", fieldbackground=[("readonly", BG_ENTRY)])
    style.configure("TCombobox", fieldbackground=BG_ENTRY, foreground=FG, background=BG_ENTRY,
                    arrowcolor=FG, bordercolor=BORDER, padding=5)
    style.map("TCombobox", fieldbackground=[("readonly", BG_ENTRY)], foreground=[("readonly", FG)])

    style.configure("Accent.Horizontal.TProgressbar", background=ACCENT, troughcolor=BG_PANEL,
                    bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT, thickness=7)
    style.configure("Treeview", background=BG_ENTRY, fieldbackground=BG_ENTRY, foreground=FG,
                    bordercolor=BORDER, rowheight=27, font=("Segoe UI", 9))
    style.configure("Treeview.Heading", background=BG_PANEL, foreground=FG, bordercolor=BORDER,
                    relief="flat", font=("Segoe UI Semibold", 9), padding=(5, 6))
    style.map("Treeview.Heading", background=[("active", "#252c37")])
    style.map("Treeview", background=[("selected", SELECT)], foreground=[("selected", "#ffffff")])

    style.configure("Vertical.TScrollbar", background=BG_PANEL, troughcolor=BG, bordercolor=BORDER, arrowcolor=FG)
    style.configure("Horizontal.TScrollbar", background=BG_PANEL, troughcolor=BG, bordercolor=BORDER, arrowcolor=FG)
    style.configure("TNotebook", background=BG, bordercolor=BORDER, tabmargins=(0, 4, 0, 0))
    style.configure("TNotebook.Tab", background=BG_PANEL, foreground=FG_MUTED, padding=(15, 9), font=("Segoe UI Semibold", 10))
    style.map("TNotebook.Tab", background=[("selected", BG), ("active", "#232934")], foreground=[("selected", FG), ("active", FG)])
    style.configure("TPanedwindow", background=BG, sashwidth=6)
    style.configure("TSeparator", background=BORDER)


def estilizar_texto(widget_texto):
    widget_texto.configure(background=BG_ENTRY, foreground=FG, insertbackground=FG,
                           selectbackground=SELECT, selectforeground="#ffffff", relief="flat",
                           highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT)


def estilizar_listbox(widget_listbox):
    widget_listbox.configure(background=BG_ENTRY, foreground=FG, selectbackground=SELECT,
                             selectforeground="#ffffff", relief="flat", highlightthickness=1,
                             highlightbackground=BORDER, activestyle="none")


class ScrollableArea(ttk.Frame):
    """Área vertical rolável para telas que não cabem em monitores menores.

    A roda do mouse rola a página quando o cursor não está sobre um widget que
    já possui rolagem própria (texto, lista ou tabela).
    """

    _WIDGETS_COM_SCROLL_PROPRIO = (tk.Text, tk.Listbox, ttk.Treeview)

    def __init__(self, parent, *, padding=(0, 0, 0, 0)):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, background=BG, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.content = ttk.Frame(self.canvas, padding=padding)
        self._window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.content.bind("<Configure>", self._ajustar_regiao, add="+")
        self.canvas.bind("<Configure>", self._ajustar_largura, add="+")
        self.bind_all("<MouseWheel>", self._roda_mouse, add="+")
        self.bind_all("<Button-4>", self._roda_linux, add="+")
        self.bind_all("<Button-5>", self._roda_linux, add="+")

    def _ajustar_regiao(self, _evento=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _ajustar_largura(self, evento):
        self.canvas.itemconfigure(self._window, width=evento.width)

    def _ponteiro_dentro(self, evento):
        if not self.winfo_ismapped():
            return False
        x0, y0 = self.winfo_rootx(), self.winfo_rooty()
        return x0 <= evento.x_root < x0 + self.winfo_width() and y0 <= evento.y_root < y0 + self.winfo_height()

    def _widget_tem_scroll_proprio(self, widget):
        atual = widget
        while atual is not None:
            if isinstance(atual, self._WIDGETS_COM_SCROLL_PROPRIO):
                return True
            if atual == self:
                break
            try:
                atual = atual.master
            except Exception:
                break
        return False

    def _roda_mouse(self, evento):
        if not self._ponteiro_dentro(evento) or self._widget_tem_scroll_proprio(evento.widget):
            return
        if evento.delta:
            passos = -1 if evento.delta > 0 else 1
            self.canvas.yview_scroll(passos * 3, "units")

    def _roda_linux(self, evento):
        if not self._ponteiro_dentro(evento) or self._widget_tem_scroll_proprio(evento.widget):
            return
        self.canvas.yview_scroll(-3 if evento.num == 4 else 3, "units")


class Tooltip:
    def __init__(self, widget, texto, atraso_ms=450):
        self.widget = widget
        self.texto = texto
        self.atraso_ms = atraso_ms
        self._id_agendado = None
        self._janela = None
        widget.bind("<Enter>", self._agendar, add="+")
        widget.bind("<Leave>", self._cancelar, add="+")
        widget.bind("<ButtonPress>", self._cancelar, add="+")

    def _agendar(self, _evento=None):
        self._cancelar()
        self._id_agendado = self.widget.after(self.atraso_ms, self._mostrar)

    def _cancelar(self, _evento=None):
        if self._id_agendado is not None:
            try:
                self.widget.after_cancel(self._id_agendado)
            except tk.TclError:
                pass
            self._id_agendado = None
        self._esconder()

    def _mostrar(self):
        if self._janela is not None or not self.texto or not self.widget.winfo_exists():
            return
        x = self.widget.winfo_rootx() + 4
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self._janela = tk.Toplevel(self.widget)
        self._janela.wm_overrideredirect(True)
        try:
            self._janela.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        self._janela.wm_geometry(f"+{x}+{y}")
        tk.Label(self._janela, text=self.texto, background="#090b0f", foreground="#f3f6fa",
                 font=("Segoe UI", 9), padx=9, pady=6, wraplength=340, justify="left").pack()

    def _esconder(self):
        if self._janela is not None:
            try:
                self._janela.destroy()
            except tk.TclError:
                pass
            self._janela = None


def dica(widget, texto):
    return Tooltip(widget, texto)


EXTENSOES_EXCEL_VALIDAS = (".xlsx", ".xlsm")


def validar_caminho_planilha(caminho):
    caminho = (caminho or "").strip()
    if not caminho:
        return False, "Nenhum arquivo foi selecionado."
    if not os.path.isfile(caminho):
        return False, f"Não encontrei esse arquivo no computador:\n{caminho}"
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".xls":
        return False, "O formato .xls antigo não é suportado nesta versão. Abra no Excel e salve como .xlsx."
    if ext not in EXTENSOES_EXCEL_VALIDAS:
        return False, f"Selecione uma planilha .xlsx ou .xlsm.\n\nArquivo recebido:\n{caminho}"
    return True, ""


def confirmar_sobrescrita(caminho):
    if not os.path.exists(caminho):
        return True
    return messagebox.askyesno("Arquivo já existe", f"Já existe um arquivo chamado \"{os.path.basename(caminho)}\" nessa pasta.\n\nDeseja substituí-lo pelo novo?", icon="warning")


def executar_assincrono(widget, trabalho, ao_sucesso, ao_erro=None, ao_progresso=None, intervalo_ms=60):
    """Executa trabalho pesado sem travar o Tkinter.

    `trabalho` recebe uma função `report(valor)` thread-safe. Nenhum widget é
    acessado fora da thread principal; a UI é atualizada pelo polling da fila.
    """
    fila = queue.Queue()

    def report(valor):
        fila.put(("progress", valor))

    def worker():
        try:
            fila.put(("result", trabalho(report)))
        except BaseException as exc:  # mantém o erro para exibir na UI
            fila.put(("error", exc))

    thread = threading.Thread(target=worker, daemon=True, name="SuiteProducaoWorker")
    thread.start()

    def poll():
        terminou = False
        while True:
            try:
                tipo, valor = fila.get_nowait()
            except queue.Empty:
                break
            if tipo == "progress" and ao_progresso:
                ao_progresso(valor)
            elif tipo == "result":
                terminou = True
                ao_sucesso(valor)
            elif tipo == "error":
                terminou = True
                if ao_erro:
                    ao_erro(valor)
                else:
                    messagebox.showerror("Erro", str(valor))
        if not terminou and thread.is_alive():
            try:
                widget.after(intervalo_ms, poll)
            except tk.TclError:
                pass
        elif not terminou:
            # A thread pode ter terminado entre a checagem e a fila ser drenada.
            try:
                widget.after(1, poll)
            except tk.TclError:
                pass

    widget.after(intervalo_ms, poll)
    return thread


def caminho_saida_sugerido(caminho_origem, pasta_saida=None):
    base, ext = os.path.splitext(caminho_origem)
    destino_dir = pasta_saida or os.path.dirname(caminho_origem)
    return os.path.join(destino_dir, os.path.basename(base) + "_ATUALIZADO" + ext)
