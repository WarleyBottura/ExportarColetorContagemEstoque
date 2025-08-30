#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exportador TXT – concatenação dinâmica a partir do XLS/XLSX
-----------------------------------------------------------------
Objetivo
- Ler uma planilha (XLS/XLSX/CSV)
- Detectar dinamicamente os nomes das colunas
- Permitir escolher:
    * a coluna base da descrição (ex.: DESC_PRODUTO)
    * múltiplas colunas adicionais para concatenar
    * rótulo que antecede o valor de cada coluna adicional (ex.: DEP:, QTD:)
    * formato de concatenação (parênteses, separador entre pares, separador rótulo/valor)
- Gerar linhas como:  
  "ADSTRIGENTE 387 FACE BEAUTIFUL (DEP: GERAL / QTD: 3)"
- Mostrar prévia e salvar em TXT

Dependências
- Somente bibliotecas padrão + pandas
  * pandas (>= 1.5) – já usado no seu ambiente
  * openpyxl (para .xlsx) / xlrd==1.2.0 (para .xls) – opcionais, mas recomendados

Observações
- Suporta CSV (delimitador detectado automaticamente) como alternativa rápida.
- Converte valores NaN/None para vazio.
- Números inteiros como 3.0 são formatados como 3 (sem .0).

Autor: ChatGPT (KDE) – 2025-08-25
"""

import os
import sys
import csv
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any

# Tentativa de importar pandas, com mensagem amigável se faltar
try:
    import pandas as pd
except Exception as e:
    raise SystemExit("Erro: pandas não está instalado. Instale com: pip install pandas openpyxl xlrd==1.2.0")

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
    """Gera um rótulo padrão curto a partir do nome da coluna.
    Ex.: 'QTD_ESTOQUE_ATUAL' -> 'QTD'
         'Departamento'      -> 'DEP'
    Regras simples: pega a primeira palavra/segmento e limita a 3-4 letras em maiúsculas.
    """
    if not colname:
        return "VAL"
    seg = colname.replace("-", "_").replace(" ", "_").split("_")[0]
    seg = seg.strip().upper()
    if len(seg) <= 4:
        return seg
    return seg[:4]

# ------------------------------ App ------------------------------

class ExportadorTXTApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Exportador TXT – Concatenação dinâmica a partir do XLS/XLSX/CSV")
        self.geometry("1080x720")

        # Estado
        self.df: pd.DataFrame | None = None
        self.file_path: str | None = None
        self.sheet_names: List[str] = []
        self.col_vars: Dict[str, tk.BooleanVar] = {}
        self.label_vars: Dict[str, tk.StringVar] = {}

        # Config de formatação
        self.base_col_var = tk.StringVar(value="")
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

        ttk.Button(top, text="Abrir XLS/XLSX/CSV…", command=self.on_open_file).pack(side=tk.LEFT)
        self.file_lbl = ttk.Label(top, text="(nenhum arquivo)")
        self.file_lbl.pack(side=tk.LEFT, padx=10)

        ttk.Label(top, text="Aba/Sheet:").pack(side=tk.LEFT, padx=(20, 4))
        self.sheet_cbx = ttk.Combobox(top, state="disabled", width=28)
        self.sheet_cbx.pack(side=tk.LEFT)
        self.sheet_cbx.bind("<<ComboboxSelected>>", self.on_select_sheet)

        # Linha: campo base e formato
        fmt = ttk.Frame(self)
        fmt.pack(side=tk.TOP, fill=tk.X, padx=10, pady=6)

        ttk.Label(fmt, text="Coluna base da descrição:").pack(side=tk.LEFT)
        self.base_cbx = ttk.Combobox(fmt, state="disabled", width=38, textvariable=self.base_col_var)
        self.base_cbx.pack(side=tk.LEFT, padx=6)

        ttk.Label(fmt, text="Abertura").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.opening_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Fechamento").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.closing_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. pares").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.pair_sep_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. rótulo/valor").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.label_sep_var).pack(side=tk.LEFT)

        # Centro: lista de colunas com checkboxes + rótulo editável
        mid = ttk.Frame(self)
        mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=6)

        # Cabeçalho
        head = ttk.Frame(mid)
        head.pack(fill=tk.X)
        ttk.Label(head, text="Marque as colunas a concatenar e ajuste o rótulo:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)

        # Rolável
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
        preview_wrap = ttk.LabelFrame(self, text="Prévia do resultado")
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

        # Cabeçalhos
        header = ttk.Frame(self.inner)
        header.grid(row=0, column=0, sticky="ew", padx=(4,4), pady=(4,2))
        ttk.Label(header, text="Usar", width=6).grid(row=0, column=0, padx=4)
        ttk.Label(header, text="Coluna", width=48).grid(row=0, column=1, padx=4)
        ttk.Label(header, text="Rótulo (editável)", width=24).grid(row=0, column=2, padx=4)

        # Linhas de colunas
        for i, col in enumerate(self.df.columns, start=1):
            rowf = ttk.Frame(self.inner)
            rowf.grid(row=i, column=0, sticky="ew", padx=4, pady=2)

            v = tk.BooleanVar(value=False)
            self.col_vars[col] = v
            ttk.Checkbutton(rowf, variable=v).grid(row=0, column=0, padx=(0,8))

            ttk.Label(rowf, text=col).grid(row=0, column=1, sticky="w")

            sv = tk.StringVar(value=f"{abbrev_label(col)}")
            self.label_vars[col] = sv
            ttk.Entry(rowf, textvariable=sv, width=16).grid(row=0, column=2, padx=(8,0))

        # Habilita combobox da base
        self.base_cbx.config(state="readonly", values=list(self.df.columns))
        if not self.base_col_var.get() and len(self.df.columns) > 0:
            # Heurística: tenta achar uma coluna de descrição
            guess = None
            for cand in ["DESC_PRODUTO", "DESCRICAO", "DESCRICAO_PRODUTO", "NOME", "PRODUTO", "DESCRICAO_COMPLETA"]:
                if cand in self.df.columns:
                    guess = cand
                    break
            self.base_col_var.set(guess or self.df.columns[0])

    # -------------------------- Fluxo --------------------------
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
            # Tenta detectar delimitador
            with open(path, "r", newline="", encoding="utf-8", errors="replace") as f:
                sample = f.read(4096)
                f.seek(0)
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                df = pd.read_csv(f, sep=dialect.delimiter)
        elif ext in (".xlsx", ".xls"):
            # Coleta os sheets primeiro
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

    def _build_line(self, row: pd.Series) -> str:
        base_col = self.base_col_var.get().strip()
        base_text = smart_str(row.get(base_col, "")) if base_col else ""

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
                    if label:
                        parts.append(f"{label}{label_sep}{val}")
                    else:
                        parts.append(f"{val}")
        extra = f"{opening}{pair_sep.join(parts)}{closing}" if parts else ""

        # Espaçamento inteligente: adiciona espaço antes do bloco extra se houver base
        if base_text and extra:
            return f"{base_text} {extra}"
        return base_text or extra

    def on_preview(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Abra uma planilha primeiro.")
            return
        if not self.base_col_var.get():
            messagebox.showwarning("Atenção", "Selecione a coluna base da descrição.")
            return
        # Gera até 10 linhas
        lines = []
        for _, row in self.df.head(10).iterrows():
            lines.append(self._build_line(row))
        self.preview_txt.delete("1.0", tk.END)
        self.preview_txt.insert(tk.END, "\n".join(lines))

    def on_save(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Abra uma planilha primeiro.")
            return
        if not self.base_col_var.get():
            messagebox.showwarning("Atenção", "Selecione a coluna base da descrição.")
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


if __name__ == "__main__":
    app = ExportadorTXTApp()
    app.mainloop()
