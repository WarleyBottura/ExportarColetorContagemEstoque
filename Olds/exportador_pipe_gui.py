#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exportador Pipe GUI
-------------------
Lê um arquivo XLS/XLSX/CSV com as colunas:
  COD_INTERNO, COD_EAN, DES_PRODUTO, (demais colunas ignoradas)
e exporta apenas essas 3 em TXT delimitado por "|", no formato:
  CODIGO_SISTEMA|CODIGO_BARRAS|NOME

Requisitos (instale com pip se necessário):
  pip install pandas openpyxl xlrd

Observações:
- Suporta .xls (via xlrd), .xlsx (via openpyxl) e .csv.
- Converte todos os valores para texto, remove quebras de linha,
  substitui pipes '|' por '/', e recorta espaços nas extremidades.
- O EAN é lido como texto para evitar notação científica ou ".0".
- Permite concatenar um sufixo/prefixo personalizado à descrição.
- Opção para incluir cabeçalho "CODIGO_SISTEMA|CODIGO_BARRAS|NOME".
- Opção de encoding e quebra de linha (Windows CRLF ou Unix LF).
"""
import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, List
import pandas as pd

APP_TITLE = "Exportador Pipe (XLS→TXT)"
HEADER_DEFAULT = "CODIGO_SISTEMA|CODIGO_BARRAS|NOME"

# ------------------ Utilidades ------------------
def log_append(widget: tk.Text, msg: str) -> None:
    widget.configure(state="normal")
    widget.insert("end", msg + "\n")
    widget.see("end")
    widget.configure(state="disabled")

def sanitize_text(val) -> str:
    """Converte qualquer valor em string 'limpa' para o TXT."""
    if pd.isna(val):
        s = ""
    else:
        s = str(val)
    # remove CR/LF e substitui '|' para não quebrar o layout
    s = s.replace("\r", " ").replace("\n", " ").replace("|", "/")
    # tira espaços extras nas bordas
    s = s.strip()
    return s

def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xls", ".xlsx"]:
        # dtype=str garante que EAN não vire número
        df = pd.read_excel(path, dtype=str)
    elif ext == ".csv":
        # tenta ler com separador ; ou , automaticamente
        try:
            df = pd.read_csv(path, dtype=str)
        except Exception:
            df = pd.read_csv(path, dtype=str, sep=";")
    else:
        raise ValueError("Formato não suportado: use .xls, .xlsx ou .csv")
    return df

def pick_columns(df: pd.DataFrame) -> pd.DataFrame:
    # normaliza nomes para comparação case-insensitive
    lower_map = {c.lower().strip(): c for c in df.columns}
    needed = {
        "cod_interno": None,
        "cod_ean": None,
        "des_produto": None,
    }
    for key in list(needed.keys()):
        if key in lower_map:
            needed[key] = lower_map[key]
    missing = [k for k, v in needed.items() if v is None]
    if missing:
        raise KeyError(
            "Colunas obrigatórias não encontradas: "
            + ", ".join(missing)
            + f".\nEncontradas: {list(df.columns)}"
        )
    cols = [needed["cod_interno"], needed["cod_ean"], needed["des_produto"]]
    return df[cols].copy()

def build_lines(df: pd.DataFrame, extra_text: str, extra_pos: str) -> List[str]:
    lines = []
    # renomeia para destino
    df = df.rename(columns=lambda c: c.strip())
    col_sistema, col_barras, col_nome = df.columns.tolist()
    for _, row in df.iterrows():
        cod_sistema = sanitize_text(row[col_sistema])
        cod_barras = sanitize_text(row[col_barras])
        nome = sanitize_text(row[col_nome])

        if extra_text:
            if extra_pos == "prefixo":
                # "EXTRA " + nome
                nome = f"{extra_text} {nome}".strip()
            else:
                # nome + " " + "EXTRA"
                if nome:
                    nome = f"{nome} {extra_text}".strip()
                else:
                    nome = extra_text.strip()

        lines.append(f"{cod_sistema}|{cod_barras}|{nome}")
    return lines

def export_txt(src_path: str, dst_path: str, include_header: bool,
               extra_text: str, extra_pos: str, encoding: str, newline_win: bool,
               log: Optional[tk.Text] = None) -> int:
    if log: log_append(log, f"Lendo: {src_path}")
    df = read_table(src_path)
    df = pick_columns(df)

    if log: log_append(log, f"Registros carregados: {len(df)}")
    lines = build_lines(df, extra_text=extra_text, extra_pos=extra_pos)

    # monta conteúdo final
    final_lines: List[str] = []
    if include_header:
        final_lines.append(HEADER_DEFAULT)
    final_lines.extend(lines)

    nl = "\r\n" if newline_win else "\n"
    text = nl.join(final_lines) + nl

    # garante pasta destino
    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
    with open(dst_path, "w", encoding=encoding, newline="") as f:
        f.write(text)

    if log:
        log_append(log, f"Arquivo salvo: {dst_path}")
        log_append(log, "Concluído.")
    return len(lines)

# ------------------ GUI ------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("760x540")
        self.minsize(720, 520)

        self.src_path = tk.StringVar()
        self.dst_path = tk.StringVar()
        self.extra_text = tk.StringVar()
        self.extra_pos = tk.StringVar(value="sufixo")  # sufixo por padrão
        self.include_header = tk.BooleanVar(value=True)
        self.encoding = tk.StringVar(value="utf-8")
        self.newline_win = tk.BooleanVar(value=True)

        self.build_ui()

    def build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm_paths = ttk.LabelFrame(self, text="Arquivos")
        frm_paths.pack(fill="x", **pad)

        # origem
        ttk.Label(frm_paths, text="Planilha de origem (.xls/.xlsx/.csv):").grid(row=0, column=0, sticky="w", **pad)
        ent_src = ttk.Entry(frm_paths, textvariable=self.src_path, width=70)
        ent_src.grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(frm_paths, text="Escolher...", command=self.pick_src).grid(row=0, column=2, **pad)

        # destino
        ttk.Label(frm_paths, text="Arquivo TXT de destino:").grid(row=1, column=0, sticky="w", **pad)
        ent_dst = ttk.Entry(frm_paths, textvariable=self.dst_path, width=70)
        ent_dst.grid(row=1, column=1, sticky="we", **pad)
        ttk.Button(frm_paths, text="Salvar como...", command=self.pick_dst).grid(row=1, column=2, **pad)

        frm_opts = ttk.LabelFrame(self, text="Opções")
        frm_opts.pack(fill="x", **pad)

        # extra
        ttk.Label(frm_opts, text="Texto adicional para DES_PRODUTO:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm_opts, textvariable=self.extra_text, width=50).grid(row=0, column=1, sticky="w", **pad)

        pos_frame = ttk.Frame(frm_opts)
        pos_frame.grid(row=0, column=2, sticky="w", **pad)
        ttk.Radiobutton(pos_frame, text="Adicionar no final (sufixo)", value="sufixo", variable=self.extra_pos).pack(side="left")
        ttk.Radiobutton(pos_frame, text="Adicionar no início (prefixo)", value="prefixo", variable=self.extra_pos).pack(side="left")

        # header + encoding
        ttk.Checkbutton(frm_opts, text="Incluir cabeçalho (CODIGO_SISTEMA|CODIGO_BARRAS|NOME)", variable=self.include_header).grid(row=1, column=0, columnspan=3, sticky="w", **pad)

        enc_frame = ttk.Frame(frm_opts)
        enc_frame.grid(row=2, column=0, sticky="w", **pad)
        ttk.Label(enc_frame, text="Encoding:").pack(side="left")
        cb_enc = ttk.Combobox(enc_frame, textvariable=self.encoding, state="readonly",
                              values=["utf-8", "utf-8-sig", "latin-1", "cp1252"])
        cb_enc.pack(side="left", padx=6)

        nl_frame = ttk.Frame(frm_opts)
        nl_frame.grid(row=2, column=1, sticky="w", **pad)
        ttk.Checkbutton(nl_frame, text="Quebra de linha Windows (CRLF)", variable=self.newline_win).pack(side="left")

        # ações
        frm_actions = ttk.Frame(self)
        frm_actions.pack(fill="x", **pad)
        ttk.Button(frm_actions, text="Exportar", command=self.on_export).pack(side="left", padx=6)
        ttk.Button(frm_actions, text="Sair", command=self.destroy).pack(side="right", padx=6)

        # log
        frm_log = ttk.LabelFrame(self, text="Log")
        frm_log.pack(fill="both", expand=True, **pad)
        self.txt_log = tk.Text(frm_log, height=14, state="disabled")
        self.txt_log.pack(fill="both", expand=True, padx=8, pady=8)

        # dicas
        tip = ("Dica: se o EAN aparecer estranho (ex.: 7.89214E+12 ou com '.0'), "
               "certifique-se de que o arquivo de origem está com essa coluna formatada como TEXTO.\n"
               "Este aplicativo já tenta ler como texto, mas planilhas muito antigas podem precisar de ajuste.")
        lbl_tip = ttk.Label(self, text=tip, wraplength=700, foreground="#555")
        lbl_tip.pack(fill="x", padx=12, pady=(0, 10))

    def pick_src(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[
                ("Planilhas Excel", "*.xls *.xlsx"),
                ("CSV", "*.csv"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if path:
            self.src_path.set(path)
            # sugere nome de saída
            base, _ = os.path.splitext(path)
            self.dst_path.set(base + "_exportado.txt")

    def pick_dst(self):
        path = filedialog.asksaveasfilename(
            title="Salvar TXT como...",
            defaultextension=".txt",
            filetypes=[("TXT", "*.txt")],
            initialfile="exportado.txt",
        )
        if path:
            self.dst_path.set(path)

    def on_export(self):
        src = self.src_path.get().strip()
        dst = self.dst_path.get().strip()
        if not src:
            messagebox.showwarning(APP_TITLE, "Selecione a planilha de origem.")
            return
        if not dst:
            messagebox.showwarning(APP_TITLE, "Escolha um caminho para salvar o TXT.")
            return

        try:
            n = export_txt(
                src_path=src,
                dst_path=dst,
                include_header=self.include_header.get(),
                extra_text=self.extra_text.get(),
                extra_pos=self.extra_pos.get(),
                encoding=self.encoding.get(),
                newline_win=self.newline_win.get(),
                log=self.txt_log,
            )
            messagebox.showinfo(APP_TITLE, f"Exportação concluída.\nRegistros: {n}\nDestino:\n{dst}")
        except Exception as e:
            log_append(self.txt_log, f"ERRO: {e}")
            messagebox.showerror(APP_TITLE, f"Falha na exportação:\n{e}")

def main():
    # Permite uso sem GUI via argumentos (opcional):
    #   python exportador_pipe_gui.py origem.xls saida.txt --no-header --prefixo "EXTRA"
    # Isso ajuda automações e testes.
    if len(sys.argv) >= 3 and sys.argv[1] != "--gui":
        src = sys.argv[1]
        dst = sys.argv[2]
        include_header = True
        extra_text = ""
        extra_pos = "sufixo"
        encoding = "utf-8"
        newline_win = True
        # parse flags simples
        args = sys.argv[3:]
        i = 0
        while i < len(args):
            a = args[i]
            if a == "--no-header":
                include_header = False
            elif a == "--prefixo":
                extra_pos = "prefixo"
                if i + 1 < len(args):
                    extra_text = args[i+1]; i += 1
            elif a == "--sufixo":
                extra_pos = "sufixo"
                if i + 1 < len(args):
                    extra_text = args[i+1]; i += 1
            elif a == "--encoding" and i + 1 < len(args):
                encoding = args[i+1]; i += 1
            elif a == "--lf":
                newline_win = False
            i += 1
        n = export_txt(src, dst, include_header, extra_text, extra_pos, encoding, newline_win)
        print(f"OK - Registros exportados: {n} -> {dst}")
        return

    # GUI padrão
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
