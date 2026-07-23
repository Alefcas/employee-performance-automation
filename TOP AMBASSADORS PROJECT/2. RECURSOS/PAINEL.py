# -*- coding: utf-8 -*-
# ============================================================
#   PAINEL.py  -  Tela de boas-vindas do TOP AMBASSADORS
#   Painel de apresentacao (paleta Amazon). Mostra o titulo,
#   uma breve descricao do projeto e um botao para iniciar o
#   processo completo. Retorna True se o usuario clicar em
#   "Iniciar" e False se apenas fechar a janela.
# ============================================================
import tkinter as tk
from tkinter import font as tkfont

# ============================================
# PALETA (mesma dos sites "Site of Alefcas")
# ============================================
C_NAVY     = "#131A22"   # fundo topo (Amazon ink)
C_DARK     = "#232F3E"   # azul escuro Amazon
C_MEDIUM   = "#37475A"   # azul medio
C_ORANGE   = "#FF9900"   # laranja Amazon
C_ORANGE_D = "#E88B00"   # laranja hover
C_GOLD     = "#FEBD69"   # dourado
C_LIGHT    = "#EAEDED"   # cinza claro de fundo
C_WHITE    = "#FFFFFF"
C_TEXT     = "#0F1111"
C_SUBTEXT  = "#565959"


DESCRICAO = (
    "O TOP AMBASSADORS reconhece os Embaixadores \u2014 os treinadores que "
    "capacitam os novos associados no in\u00edcio da sua jornada.\n\n"
    "Este programa automatiza todo o ciclo de reconhecimento: re\u00fane os dados "
    "de qualidade e produtividade, aplica as regras de neg\u00f3cio e gera o "
    "ranking, o e-mail de reconhecimento e as cartas dos destaques.\n\n"
    "Um projeto que abriu portas importantes na minha trajet\u00f3ria dentro da "
    "Amazon. Espero que gostem."
)


class PainelBoasVindas:
    def __init__(self, root):
        self.root = root
        self.iniciar = False

        self.root.title("TOP AMBASSADORS")
        self.root.configure(bg=C_LIGHT)
        self.root.geometry("780x620")
        self.root.minsize(680, 560)

        # Fontes
        self.f_titulo = tkfont.Font(family="Segoe UI", size=44, weight="bold")
        self.f_tag    = tkfont.Font(family="Segoe UI", size=13)
        self.f_desc   = tkfont.Font(family="Segoe UI", size=12)
        self.f_btn    = tkfont.Font(family="Segoe UI", size=14, weight="bold")
        self.f_foot   = tkfont.Font(family="Segoe UI", size=9)

        self._construir()

    # ----------------------------------------
    def _construir(self):
        # Faixa laranja no topo
        tk.Frame(self.root, bg=C_ORANGE, height=5).pack(fill="x")

        # ===== HEADER =====
        header = tk.Frame(self.root, bg=C_NAVY, height=210)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(header, text="TOP AMBASSADORS", bg=C_NAVY, fg=C_WHITE,
                 font=self.f_titulo).pack(pady=(46, 0))

        tk.Label(header,
                 text="Reconhecimento de Embaixadores  \u00b7  Fulfillment Center (FC locale)",
                 bg=C_NAVY, fg=C_GOLD, font=self.f_tag).pack(pady=(6, 0))

        # ===== CORPO =====
        corpo = tk.Frame(self.root, bg=C_LIGHT)
        corpo.pack(fill="both", expand=True, padx=44, pady=(28, 10))

        # Cartao com a descricao
        cartao = tk.Frame(corpo, bg=C_WHITE, highlightbackground="#D5D9D9",
                          highlightthickness=1)
        cartao.pack(fill="both", expand=True)

        barra = tk.Frame(cartao, bg=C_ORANGE, width=5)
        barra.pack(side="left", fill="y")

        tk.Label(cartao, text=DESCRICAO, bg=C_WHITE, fg=C_TEXT,
                 font=self.f_desc, justify="left", wraplength=600,
                 anchor="nw").pack(side="left", fill="both", expand=True,
                                   padx=26, pady=24)

        # ===== BOTAO =====
        area_btn = tk.Frame(self.root, bg=C_LIGHT)
        area_btn.pack(fill="x", padx=44, pady=(0, 18))

        self.btn = tk.Button(area_btn, text="Iniciar processo completo",
                             font=self.f_btn, bg=C_ORANGE, fg=C_WHITE,
                             activebackground=C_ORANGE_D, activeforeground=C_WHITE,
                             relief="flat", cursor="hand2", bd=0,
                             padx=28, pady=14, command=self._on_iniciar)
        self.btn.pack(fill="x")
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg=C_ORANGE_D))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=C_ORANGE))

        # ===== RODAPE =====
        rodape = tk.Frame(self.root, bg=C_DARK, height=40)
        rodape.pack(fill="x", side="bottom")
        rodape.pack_propagate(False)
        tk.Label(rodape,
                 text="FC locale Operations  \u00b7  Work hard. Have fun. Make history.",
                 bg=C_DARK, fg=C_GOLD, font=self.f_foot).pack(pady=10)

        # ENTER tambem inicia
        self.root.bind("<Return>", lambda e: self._on_iniciar())

    # ----------------------------------------
    def _on_iniciar(self):
        self.iniciar = True
        self.root.destroy()


def mostrar_painel():
    """Abre o painel de boas-vindas. Retorna True se o usuario clicou
    em 'Iniciar', ou False se apenas fechou a janela."""
    root = tk.Tk()
    painel = PainelBoasVindas(root)
    # Centraliza a janela na tela
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    x = (root.winfo_screenwidth()  - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry("+" + str(max(x, 0)) + "+" + str(max(y, 0)))
    root.mainloop()
    return painel.iniciar


if __name__ == "__main__":
    print("Iniciar?" , mostrar_painel())
