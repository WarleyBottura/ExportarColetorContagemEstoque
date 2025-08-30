#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exportador TXT – concatenação dinâmica a partir do XLS/XLSX/CSV
-----------------------------------------------------------------
Requisitos (v2 – ajustado):
- Exportar SEM cabeçalho e no formato **COD_INTERNO|COD_EAN|DES_PRODUTO**.
- `DES_PRODUTO` = coluna **DES_PRODUTO** (fixa) + bloco formatado com colunas marcadas.
  Ex.:  ADSTRIGENTE 387 FACE BEAUTIFUL (DEP: GERAL / QTD: 3)
- **Base** da descrição é sempre `DES_PRODUTO` (não configurável na UI).
- **COD_INTERNO** e **COD_EAN** são sempre exportadas (não participam da lista de concatenação).
- Parâmetros de formato: abertura/fechamento (padrão: `(`, `)`), sep. pares (padrão: ` / `), sep. rótulo/valor (padrão: `: `).
- Prévia (10 linhas) mostra o formato final.

Notas técnicas:
- Suporte a XLS/XLSX/CSV. CSV: detecta `, ; \t |` e faz fallback para vírgula.
- Valores vazios/NaN não entram no bloco adicional.
- Floats inteiros (ex.: 3.0) são emitidos como `3`.
- Inclui **self-tests** acionados por `--selftest` (sem abrir a janela) para validar funções-chave.

"""

from __future__ import annotations
import os
import sys
import csv
import math
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any

try:
    import pandas as pd
except Exception:
    raise SystemExit("Erro: pandas não está instalado. Instale com: pip install pandas openpyxl xlrd==1.2.0")

# ------------------------------ Constantes ------------------------------
BASE_COL = "DES_PRODUTO"  # base fixa
MANDATORY_PREFIX = ("COD_INTERNO", "COD_EAN")  # sempre exportadas

# ------------------------------ Utilidades ------------------------------

def smart_str(x: Any) -> str:
    """Converte valores em string de forma amigável.
    - None/NaN -> ""
    - float inteiro -> sem .0
    - strip em strings
    """
    if x is None:
        return ""
    if isinstance(x, float):
        if math.isnan(x):
            return ""
        if x.is_integer():
            return str(int(x))
        return str(x)
    s = str(x)
    return s.strip()

def abbrev_label(colname: str) -> str:
    """Rótulo curto padrão a partir do nome da coluna (1ª palavra, até 4 letras)."""
    if not colname:
        return "VAL"
    seg = colname.replace("-", "_").replace(" ", "_").split("_")[0]
    seg = seg.strip().upper()
    return seg if len(seg) <= 4 else seg[:4]

# ------------------------------ App ------------------------------

class ExportadorTXTApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Exportador Relatorio Custo Estoque.xls para importar no APP Contagem de Estoque (Por Bottura)")
        self.geometry("1080x720")

        # Estado
        self.df: pd.DataFrame | None = None
        self.file_path: str | None = None
        self.sheet_names: List[str] = []
        self.col_vars: Dict[str, tk.BooleanVar] = {}
        self.label_vars: Dict[str, tk.StringVar] = {}

        # Config de formatação
        self.opening_var = tk.StringVar(value="(")
        self.closing_var = tk.StringVar(value=")")
        self.pair_sep_var = tk.StringVar(value=" / ")
        self.label_sep_var = tk.StringVar(value=": ")

        # Widgets
        self._build_ui()

    # -------------------------- UI --------------------------
    def _build_ui(self):
        # Top: Botões de arquivo e escolha de planilha/aba
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=8)
        
        tk.Button(
            top,
            text='Link para exportar "Relatorio Custo Estoque"',
            command=self.on_open_export_link,
            cursor="hand2",
            relief="flat",
            fg="blue",
            activeforeground="blue"
        ).pack(side=tk.LEFT, padx=(0,8))


        ttk.Button(top, text="Abrir Planilha", command=self.on_open_file).pack(side=tk.LEFT)
        self.file_lbl = ttk.Label(top, text="(nenhum arquivo)")
        self.file_lbl.pack(side=tk.LEFT, padx=10)

        ttk.Label(top, text="Aba/Sheet:").pack(side=tk.LEFT, padx=(20, 4))
        self.sheet_cbx = ttk.Combobox(top, state="disabled", width=28)
        self.sheet_cbx.pack(side=tk.LEFT)
        self.sheet_cbx.bind("<<ComboboxSelected>>", self.on_select_sheet)

        # Linha: formato (base é fixa: DES_PRODUTO)
        fmt = ttk.Frame(self)
        fmt.pack(side=tk.TOP, fill=tk.X, padx=10, pady=6)

        ttk.Label(fmt, text="Base da descrição:", foreground="#0a0").pack(side=tk.LEFT)
        ttk.Label(fmt, text=BASE_COL, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 18))

        ttk.Label(fmt, text="Abertura").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.opening_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Fechamento").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.closing_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. pares").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.pair_sep_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. rótulo/valor").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.label_sep_var).pack(side=tk.LEFT)

        # Centro: lista de colunas com checkboxes + rótulo editável (exclui base e mandatórias)
        mid = ttk.Frame(self)
        mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=6)

        head = ttk.Frame(mid)
        head.pack(fill=tk.X)
        ttk.Label(head, text=(
            "Marque as colunas a concatenar em DES_PRODUTO (COD_INTERNO, COD_EAN e a base 'DES_PRODUTO' já são fixos no arquivo de saida):"
        ), font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)

        columns_frame = ttk.Frame(mid)
        columns_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(columns_frame, borderwidth=0)
        self.scroll_y = ttk.Scrollbar(columns_frame, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.configure(yscrollcommand=self.scroll_y.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Rodapé de ações
        actions = ttk.Frame(self)
        actions.pack(side=tk.TOP, fill=tk.X, padx=10, pady=8)
        ttk.Button(actions, text="Pré-visualizar (10 linhas)", command=self.on_preview).pack(side=tk.LEFT)
        ttk.Button(actions, text="Salvar TXT…", command=self.on_save).pack(side=tk.LEFT, padx=8)

        # Área de prévia
        preview_wrap = ttk.LabelFrame(self, text="Prévia do resultado – COD_INTERNO|COD_EAN|DES_PRODUTO")
        preview_wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0,10))
        self.preview_txt = tk.Text(preview_wrap, height=12, wrap="none")
        self.preview_txt.pack(fill=tk.BOTH, expand=True)

    def _populate_columns_ui(self):
        # Limpa UI antiga
        for w in self.inner.winfo_children():
            w.destroy()
        self.col_vars.clear()
        self.label_vars.clear()

        if self.df is None:
            return

        # Valida base e cria avisos se necessário
        cols = list(self.df.columns)
        missing = []
        if BASE_COL not in cols:
            missing.append(BASE_COL)
        warn_opt = [c for c in ("COD_INTERNO", "COD_EAN") if c not in cols]

        if missing:
            msg = (
                f"A coluna base '{BASE_COL}' não foi encontrada na planilha.\n"
                "Colunas disponíveis:\n- " + "\n- ".join(cols)
            )
            messagebox.showerror("Coluna base ausente", msg)
            return
        if warn_opt:
            msg = (
                "As seguintes colunas opcionais de cabeçalho fixo não foram encontradas:\n"
                + "\n".join("- " + c for c in warn_opt)
                + "\nElas sairão vazias no TXT."
            )
            messagebox.showwarning("Aviso", msg)

        # Cabeçalhos
        header = ttk.Frame(self.inner)
        header.grid(row=0, column=0, sticky="ew", padx=(4,4), pady=(4,2))
        ttk.Label(header, text="Usar", width=6).grid(row=0, column=0, padx=4)
        ttk.Label(header, text="Coluna", width=48).grid(row=0, column=1, padx=4)
        ttk.Label(header, text="Rótulo (editável)", width=24).grid(row=0, column=2, padx=4)

        # Linhas de colunas – exclui base e mandatórias do prefixo
        excluded = set([BASE_COL, *MANDATORY_PREFIX])
        r = 1
        for col in self.df.columns:
            if col in excluded:
                continue
            rowf = ttk.Frame(self.inner)
            rowf.grid(row=r, column=0, sticky="ew", padx=4, pady=2)
            r += 1

            v = tk.BooleanVar(value=False)
            self.col_vars[col] = v
            ttk.Checkbutton(rowf, variable=v).grid(row=0, column=0, padx=(0,8))

            ttk.Label(rowf, text=col).grid(row=0, column=1, sticky="w")

            sv = tk.StringVar(value=f"{abbrev_label(col)}")
            self.label_vars[col] = sv
            ttk.Entry(rowf, textvariable=sv, width=16).grid(row=0, column=2, padx=(8,0))

    # -------------------------- Fluxo --------------------------
    def on_open_export_link(self):
        # Abre no navegador padrão (new=1 tenta nova aba)
        webbrowser.open(
            "https://app.mentorasolucoes.com.br/Voti-1.0.7/relatorios_base/frm_rel_custo_estoque.xhtml",
            new=1
        )

    
    def on_open_file(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[
                ("Planilhas", "*.xlsx *.xls *.csv"),
                ("Excel moderno", "*.xlsx"),
                ("Excel antigo", "*.xls"),
                ("CSV", "*.csv"),
                ("Todos", "*.*"),
            ],
        )
        if not path:
            return
        try:
            self.load_dataframe(path)
            self.file_path = path
            self.file_lbl.config(text=os.path.basename(path))
        except Exception as e:
            messagebox.showerror("Erro ao abrir arquivo", str(e))
            return

    def load_dataframe(self, path: str):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            try:
                with open(path, "r", newline="", encoding="utf-8", errors="replace") as f:
                    sample = f.read(4096)
                    f.seek(0)
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                    df = pd.read_csv(f, sep=dialect.delimiter)
            except Exception:
                # Fallback para vírgula
                df = pd.read_csv(path, sep=",", encoding="utf-8", engine="python", errors="replace")
        elif ext in (".xlsx", ".xls"):
            xl = pd.ExcelFile(path)
            self.sheet_names = xl.sheet_names
            self.sheet_cbx.config(state="readonly", values=self.sheet_names)
            self.sheet_cbx.set(self.sheet_names[0])
            df = xl.parse(self.sheet_names[0])
        else:
            raise ValueError("Formato não suportado. Use XLSX, XLS ou CSV.")

        self.df = df
        self._populate_columns_ui()

    def on_select_sheet(self, event=None):
        if not self.file_path:
            return
        try:
            xl = pd.ExcelFile(self.file_path)
            sheet = self.sheet_cbx.get()
            self.df = xl.parse(sheet)
            self._populate_columns_ui()
        except Exception as e:
            messagebox.showerror("Erro ao trocar de aba", str(e))

    # ----- Construção da linha final -----
    def _build_extra_block(self, row: pd.Series) -> str:
        opening = self.opening_var.get()
        closing = self.closing_var.get()
        pair_sep = self.pair_sep_var.get()
        label_sep = self.label_sep_var.get()

        parts = []
        for col, var in self.col_vars.items():
            if var.get():
                label = self.label_vars[col].get().strip()
                val = smart_str(row.get(col, ""))
                if val != "":
                    parts.append(f"{label}{label_sep}{val}" if label else f"{val}")
        return f"{opening}{pair_sep.join(parts)}{closing}" if parts else ""

    def _build_line(self, row: pd.Series) -> str:
        # Prefixos (podem estar ausentes na planilha)
        cod_interno = smart_str(row.get("COD_INTERNO", ""))
        cod_ean = smart_str(row.get("COD_EAN", ""))

        # Base obrigatória
        base_text = smart_str(row.get(BASE_COL, ""))

        extra = self._build_extra_block(row)
        desc_final = f"{base_text} {extra}".strip() if extra else base_text

        return f"{cod_interno}|{cod_ean}|{desc_final}"

    def on_preview(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Abra uma planilha primeiro.")
            return
        if BASE_COL not in self.df.columns:
            messagebox.showerror("Erro", f"Coluna base '{BASE_COL}' não encontrada.")
            return
        lines = []
        for _, row in self.df.head(10).iterrows():
            lines.append(self._build_line(row))
        self.preview_txt.delete("1.0", tk.END)
        self.preview_txt.insert(tk.END, "\n".join(lines))

    def on_save(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Abra uma planilha primeiro.")
            return
        if BASE_COL not in self.df.columns:
            messagebox.showerror("Erro", f"Coluna base '{BASE_COL}' não encontrada.")
            return
        out_path = filedialog.asksaveasfilename(
            title="Salvar como TXT",
            defaultextension=".txt",
            filetypes=[("TXT", "*.txt"), ("Todos", "*.*")],
        )
        if not out_path:
            return
        try:
            with open(out_path, "w", encoding="utf-8", newline="\n") as f:
                for _, row in self.df.iterrows():
                    f.write(self._build_line(row) + "\n")
            messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))


# ------------------------------ Self-tests ------------------------------

def _run_self_tests() -> int:
    """Executa testes rápidos de funções e montagem de linha.
    Retorna 0 em sucesso, 1 em falha. Não abre GUI visível.
    """
    import math as _math

    # smart_str
    assert smart_str(None) == ""
    assert smart_str(float("nan")) == ""
    assert smart_str(3.0) == "3"
    assert smart_str(3.5) == "3.5"
    assert smart_str("  abc  ") == "abc"

    # abbrev_label
    assert abbrev_label("QTD_ESTOQUE_ATUAL") in ("QTD", "QTD_"[:3])  # deve começar por QTD
    assert abbrev_label("Departamento") in ("DEPA"[:4], "DEPA"[:4])  # 1ª palavra

    # Montagem de linha sem abrir janela
    app = ExportadorTXTApp()
    app.withdraw()  # evita janela

    data = {
        "COD_INTERNO": ["123", "456"],
        "COD_EAN": ["789", ""],
        "DES_PRODUTO": ["ADSTRIGENTE 387 FACE BEAUTIFUL", "CREME XYZ"],
        "DEPARTAMENTO": ["GERAL", "BELEZA"],
        "QTD_ESTOQUE_ATUAL": [3.0, float("nan")],
    }
    app.df = pd.DataFrame(data)

    # Simula seleção de colunas (excluindo base e mandatórias)
    app.col_vars = {"DEPARTAMENTO": tk.BooleanVar(value=True),
                    "QTD_ESTOQUE_ATUAL": tk.BooleanVar(value=True)}
    app.label_vars = {"DEPARTAMENTO": tk.StringVar(value="DEP"),
                      "QTD_ESTOQUE_ATUAL": tk.StringVar(value="QTD")}

    # Caso 1: ambos presentes
    r0 = app.df.iloc[0]
    out0 = app._build_line(r0)
    exp0 = "123|789|ADSTRIGENTE 387 FACE BEAUTIFUL (DEP: GERAL / QTD: 3)"
    assert out0 == exp0, f"Esperado: {exp0} — Obtido: {out0}"

    # Caso 2: QTD vazio, sem par redundante
    r1 = app.df.iloc[1]
    out1 = app._build_line(r1)
    exp1 = "456||CREME XYZ (DEP: BELEZA)"
    assert out1 == exp1, f"Esperado: {exp1} — Obtido: {out1}"

    app.destroy()
    return 0


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        try:
            rc = _run_self_tests()
            print("Self-tests OK")
            sys.exit(rc)
        except AssertionError as e:
            print("Self-tests FAILED:", e)
            sys.exit(1)
    else:
        app = ExportadorTXTApp()
        app.mainloop()
