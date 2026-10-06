"""Suite de Produção — aplicativo unificado e otimizado."""

import tkinter as tk
from tkinter import ttk

import tema
from extrator_producao import ExtratorFrame
from lancador_producao import LancadorProducaoFrame
from lancador_individual import LancadorIndividualFrame
from configuracao import ConfiguracoesFrame


class TelaInicial(ttk.Frame):
    FERRAMENTAS = [
        ("01  Extrator de Produção", "Gere o relatório que será copiado/postado. A ação principal fica em LARANJA.", 1),
        ("02  Cadeiras / Mesas", "Analise e depois lance nas planilhas. LANÇAR fica sempre em VERDE.", 2),
        ("03  Mensal por Funcionário", "Analise e grave a produção mensal. LANÇAR fica sempre em VERDE.", 3),
        ("Configurações", "Edite grupos, exclusões, regras automáticas e preferências sem mexer no código.", 4),
    ]

    def __init__(self, parent, ao_abrir_aba):
        super().__init__(parent)
        self.ao_abrir_aba = ao_abrir_aba
        self._build()

    def _build(self):
        wrapper = ttk.Frame(self, padding=(28, 26))
        wrapper.pack(fill="both", expand=True)
        ttk.Label(wrapper, text="Suite de Produção", style="Header.TLabel", font=("Segoe UI Semibold", 25)).pack(anchor="w")
        ttk.Label(wrapper, text="Ferramentas rápidas para produção, com configurações persistentes e processamento sem travar a tela.", style="SubHeader.TLabel").pack(anchor="w", pady=(3, 22))

        grid = ttk.Frame(wrapper)
        grid.pack(fill="both", expand=True)
        grid.columnconfigure(0, weight=1, uniform="cards")
        grid.columnconfigure(1, weight=1, uniform="cards")
        grid.rowconfigure(0, weight=1, uniform="rows")
        grid.rowconfigure(1, weight=1, uniform="rows")
        for i, (titulo, desc, aba) in enumerate(self.FERRAMENTAS):
            card = ttk.Frame(grid, style="Card.TFrame", padding=20)
            card.grid(row=i // 2, column=i % 2, sticky="nsew", padx=(0 if i % 2 == 0 else 8, 8 if i % 2 == 0 else 0), pady=(0 if i < 2 else 8, 8 if i < 2 else 0))
            ttk.Label(card, text=titulo, style="CardTitle.TLabel").pack(anchor="w")
            ttk.Label(card, text=desc, style="CardDesc.TLabel", wraplength=430, justify="left").pack(anchor="w", fill="x", expand=True, pady=(7, 16))
            ttk.Button(card, text="ABRIR FERRAMENTA  →", style="Accent.TButton", command=lambda a=aba: self.ao_abrir_aba(a)).pack(anchor="e")

        foot = ttk.Frame(wrapper)
        foot.pack(fill="x", pady=(18, 0))
        ttk.Separator(foot).pack(fill="x", pady=(0, 9))
        ttk.Label(foot, text="Os arquivos originais continuam protegidos: os lançadores salvam cópias com sufixo _ATUALIZADO.", style="Muted.TLabel").pack(anchor="w")


class SuiteApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Suite de Produção")
        self.geometry("1280x840")
        self.minsize(980, 650)
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        tema.aplicar_tema_escuro(self)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.tela_inicial = TelaInicial(self.notebook, self._abrir_aba)
        self.aba_extrator = ExtratorFrame(self.notebook)
        self.aba_cadeiras_mesas = LancadorProducaoFrame(self.notebook)
        self.aba_individual = LancadorIndividualFrame(self.notebook)
        self.aba_config = ConfiguracoesFrame(self.notebook)

        self.notebook.add(self.tela_inicial, text="Início")
        self.notebook.add(self.aba_extrator, text="01  Extrator")
        self.notebook.add(self.aba_cadeiras_mesas, text="02  Cadeiras / Mesas")
        self.notebook.add(self.aba_individual, text="03  Individual")
        self.notebook.add(self.aba_config, text="Configurações")

    def _abrir_aba(self, indice):
        self.notebook.select(indice)


if __name__ == "__main__":
    SuiteApp().mainloop()
