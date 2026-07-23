# -*- coding: utf-8 -*-
# ============================================================
#   TOP AMBASSADORS - FC locale
#   >>> DE UM PLAY NESTE ARQUIVO <<<
#
#   Este e o unico arquivo que voce precisa executar.
#   Todo o resto do programa fica na pasta "2. RECURSOS".
# ============================================================
import os
import sys
import importlib.util
from pathlib import Path

RAIZ        = Path(__file__).resolve().parent
# Se este arquivo estiver dentro de "2. RECURSOS", a raiz e a pasta-pai
if RAIZ.name.strip().upper() == "2. RECURSOS":
    RAIZ = RAIZ.parent
RECURSOS    = RAIZ / "2. RECURSOS"
AMBASSADORS = RECURSOS / "AMBASSADORS.py"
PAINEL      = RECURSOS / "PAINEL.py"


def _carregar(nome, caminho):
    spec = importlib.util.spec_from_file_location(nome, str(caminho))
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    # Garante que a pasta de recursos esteja no path (para imports internos)
    if str(RECURSOS) not in sys.path:
        sys.path.insert(0, str(RECURSOS))

    # 1) Painel de boas-vindas (titulo grande + descricao do projeto).
    #    Se o usuario fechar a janela sem clicar em "Iniciar", encerra.
    if PAINEL.exists():
        try:
            painel = _carregar("PAINEL", PAINEL)
            if not painel.mostrar_painel():
                return   # usuario fechou sem iniciar
        except Exception as e:
            # Sem interface grafica disponivel: segue direto para o processo.
            print("Aviso: nao foi possivel abrir o painel grafico: " + str(e))

    # 2) Inicia direto o PROCESSO COMPLETO desde o comeco
    #    (Employee Roster -> associados -> produtividade -> ATLAS
    #     -> ranking -> email -> cartas).
    if AMBASSADORS.exists():
        mod = _carregar("AMBASSADORS", AMBASSADORS)
        mod.main(modo_forcado="completo")
    else:
        print("ERRO: arquivos do programa nao encontrados em '2. RECURSOS'.")
        input("Pressione ENTER para fechar...")


if __name__ == "__main__":
    main()
