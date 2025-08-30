
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exportador TXT – XLS/CSV **ou** consulta PostgreSQL (PDV)
---------------------------------------------------------
Novidades desta versão:
- Fonte de dados selecionável: **Planilha** (XLS/XLSX/CSV) **ou** **Consultar BD PDV (PostgreSQL)**.
- Conexão padrão: host=localhost, porta=5432, banco=PDV. Usuário e senha informados na UI.
- SQL padrão (editável na UI) já preenchida com a consulta solicitada.
- Após carregar (arquivo ou consulta), interface de seleção de colunas, prévia e salvamento continuam iguais.

Formato de saída (inalterado):
- **COD_INTERNO|COD_EAN|DES_PRODUTO**
- A base de descrição é **sempre** a coluna **DES_PRODUTO**.
- `DES_PRODUTO` = valor original + bloco formatado com colunas marcadas.
- Colunas vazias não entram no bloco adicional.
- Floats inteiros saem sem ".0".
- Prévia mostra 10 linhas.

Requisitos de instalação:
- pandas, openpyxl, xlrd==1.2.0  (para ler planilhas)
- psycopg2 (ou psycopg2-binary)  (para conectar no PostgreSQL)
  Ex.:  pip install pandas openpyxl xlrd==1.2.0 psycopg2-binary

"""

from __future__ import annotations
import os
import sys
import csv
import math
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any, Optional

try:
    import pandas as pd
except Exception:
    raise SystemExit("Erro: pandas não está instalado. Instale com: pip install pandas openpyxl xlrd==1.2.0")

# Tentativa de importar psycopg2 (necessário apenas se usar "Consultar BD PDV")
_psycopg2_err: Optional[str] = None
try:
    import psycopg2  # type: ignore
except Exception as e:  # adia o erro para quando o usuário optar por BD
    _psycopg2_err = str(e)
    psycopg2 = None  # type: ignore

# ------------------------------ Constantes ------------------------------
BASE_COL = "DES_PRODUTO"  # base fixa
MANDATORY_PREFIX = ("COD_INTERNO", "COD_EAN")  # sempre exportadas

DEFAULT_SQL = """\
SELECT 
  cod_ean,
  dta_alteracao,
  dta_cadastro,
  des_produto, 
  flg_status, 
  qtd_estoque_atual,
  val_custo as custo, 
  val_venda as VR1,
  des_marca, 
  dta_vencimento, 
  cod_interno, 
  codpai, 
  des_cor, 
  flg_pai, 
  des_tamanho, 
  val_venda_dois AS VR2, 
  flg_envia_balanca, 
  cod_imposto AS TRI,  
  obs_produto, 
  ncm, 
  unidade, 
  val_venda_promocao AS PRO, 
  des_secao AS DEP
FROM public.tb_produto;
"""

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
        self.title("Exportador para APP de Contagem – Planilha ou BD PDV (By Bottura)")
        self.geometry("1180x820")

        # Estado
        self.df: pd.DataFrame | None = None
        self.file_path: str | None = None
        self.sheet_names: List[str] = []
        self.col_vars: Dict[str, tk.BooleanVar] = {}
        self.label_vars: Dict[str, tk.StringVar] = {}

        # Fonte de dados: "file" | "db"
        self.source_var = tk.StringVar(value="file")

        # Config de formatação
        self.opening_var = tk.StringVar(value="(")
        self.closing_var = tk.StringVar(value=")")
        self.pair_sep_var = tk.StringVar(value=" / ")
        self.label_sep_var = tk.StringVar(value=": ")

        # Config de DB (padrões solicitados)
        self.db_host = tk.StringVar(value="localhost")
        self.db_port = tk.StringVar(value="5432")
        self.db_name = tk.StringVar(value="PDV")
        self.db_user = tk.StringVar(value="postgres")
        self.db_pass = tk.StringVar(value="")
        self.sql_text = tk.StringVar(value=DEFAULT_SQL.strip())

        # Widgets
        self._build_ui()

    # -------------------------- UI --------------------------
    def _build_ui(self):
        # Top bar: link + seleção de fonte de dados
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=8)

        tk.Button(
            top,
            text='Link: "Relatório Custo Estoque" (web)',
            command=self.on_open_export_link,
            cursor="hand2",
            relief="flat",
            fg="blue",
            activeforeground="blue"
        ).pack(side=tk.LEFT, padx=(0,12))

        # Fonte de dados
        ttk.Label(top, text="Fonte de dados:").pack(side=tk.LEFT)
        ttk.Radiobutton(top, text="Planilha (XLS/XLSX/CSV)", value="file", variable=self.source_var,
                        command=self.on_change_source).pack(side=tk.LEFT, padx=(6,2))
        ttk.Radiobutton(top, text="Consultar BD PDV (PostgreSQL)", value="db", variable=self.source_var,
                        command=self.on_change_source).pack(side=tk.LEFT, padx=(6,2))

        # --- FILE frame ---
        self.file_frame = ttk.Frame(self)
        self.file_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0,6))
        ttk.Button(self.file_frame, text="Abrir Planilha…", command=self.on_open_file).pack(side=tk.LEFT)
        self.file_lbl = ttk.Label(self.file_frame, text="(nenhum arquivo)")
        self.file_lbl.pack(side=tk.LEFT, padx=10)
        ttk.Label(self.file_frame, text="Aba/Sheet:").pack(side=tk.LEFT, padx=(20, 4))
        self.sheet_cbx = ttk.Combobox(self.file_frame, state="disabled", width=28)
        self.sheet_cbx.pack(side=tk.LEFT)
        self.sheet_cbx.bind("<<ComboboxSelected>>", self.on_select_sheet)

        # --- DB frame ---
        self.db_frame = ttk.LabelFrame(self, text="Conexão BD PDV (PostgreSQL)")
        # criado mas só exibido quando source_var == "db"
        self._build_db_frame_contents()

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
            "Marque as colunas a concatenar em DES_PRODUTO "
            "(COD_INTERNO, COD_EAN e a base 'DES_PRODUTO' já são fixos no arquivo de saída):"
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

        # Sincroniza visibilidade inicial
        self.on_change_source()

    def _build_db_frame_contents(self):
        f = self.db_frame

        grid = ttk.Frame(f)
        grid.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        # Linha 1: host/porta/banco
        ttk.Label(grid, text="Host:").grid(row=0, column=0, sticky="e")
        ttk.Entry(grid, textvariable=self.db_host, width=18).grid(row=0, column=1, padx=6, pady=2, sticky="w")

        ttk.Label(grid, text="Porta:").grid(row=0, column=2, sticky="e")
        ttk.Entry(grid, textvariable=self.db_port, width=8).grid(row=0, column=3, padx=6, pady=2, sticky="w")

        ttk.Label(grid, text="Banco:").grid(row=0, column=4, sticky="e")
        ttk.Entry(grid, textvariable=self.db_name, width=16).grid(row=0, column=5, padx=6, pady=2, sticky="w")

        # Linha 2: usuário/senha
        ttk.Label(grid, text="Usuário:").grid(row=1, column=0, sticky="e")
        ttk.Entry(grid, textvariable=self.db_user, width=18).grid(row=1, column=1, padx=6, pady=2, sticky="w")

        ttk.Label(grid, text="Senha:").grid(row=1, column=2, sticky="e")
        ttk.Entry(grid, textvariable=self.db_pass, width=18, show="*").grid(row=1, column=3, padx=6, pady=2, sticky="w")

        ttk.Button(grid, text="Executar consulta", command=self.on_run_query).grid(row=1, column=5, padx=6, pady=2, sticky="e")

        # Caixa de SQL
        sql_box = ttk.Frame(f)
        sql_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        ttk.Label(sql_box, text="SQL (editável):").pack(anchor="w")
        self.sql_textbox = tk.Text(sql_box, height=8, wrap="none")
        self.sql_textbox.pack(fill=tk.BOTH, expand=True)
        self.sql_textbox.insert("1.0", self.sql_text.get())

    # -------------------------- Fluxo --------------------------
    def on_open_export_link(self):
        webbrowser.open(
            "https://app.mentorasolucoes.com.br/Voti-1.0.7/relatorios_base/frm_rel_custo_estoque.xhtml",
            new=1
        )

    def on_change_source(self):
        """Mostra/oculta áreas conforme opção de fonte de dados."""
        src = self.source_var.get()
        if src == "file":
            self.file_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0,6))
            self.db_frame.pack_forget()
        else:
            self.file_frame.pack_forget()
            self.db_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=10, pady=(0,6))

    # -------------------------- Planilha --------------------------
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
            self.load_dataframe_from_file(path)
            self.file_path = path
            self.file_lbl.config(text=os.path.basename(path))
        except Exception as e:
            messagebox.showerror("Erro ao abrir arquivo", str(e))
            return

    def load_dataframe_from_file(self, path: str):
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

    # -------------------------- Banco de Dados --------------------------
    def on_run_query(self):
        """Executa a SQL no BD e carrega o DataFrame."""
        if psycopg2 is None:
            messagebox.showerror(
                "psycopg2 não disponível",
                f"Não foi possível importar psycopg2.\n"
                f"Detalhes: {_psycopg2_err or 'biblioteca ausente'}\n\n"
                f"Instale com:  pip install psycopg2-binary"
            )
            return

        host = self.db_host.get().strip() or "localhost"
        port = self.db_port.get().strip() or "5432"
        db   = self.db_name.get().strip() or "PDV"
        user = self.db_user.get().strip()
        pwd  = self.db_pass.get()

        sql = self.sql_textbox.get("1.0", "end").strip()
        if not sql:
            messagebox.showwarning("Atenção", "Informe uma SQL para executar.")
            return

        try:
            conn = psycopg2.connect(
                host=host, port=port, dbname=db, user=user, password=pwd
            )
        except Exception as e:
            messagebox.showerror("Erro de conexão", str(e))
            return

        try:
            df = pd.read_sql_query(sql, conn)
        except Exception as e:
            conn.close()
            messagebox.showerror("Erro ao executar SQL", str(e))
            return
        finally:
            try:
                conn.close()
            except Exception:
                pass

        # Normaliza nomes de colunas para o padrão usado no app (maiúsculas)
        df.columns = [c.strip().upper() for c in df.columns]

        self.df = df
        self._populate_columns_ui()
        messagebox.showinfo("Consulta concluída", f"Linhas retornadas: {len(df)}")

    # ----- UI de colunas -----
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
                f"A coluna base '{BASE_COL}' não foi encontrada na origem de dados.\n"
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
        # Prefixos (podem estar ausentes na origem)
        cod_interno = smart_str(row.get("COD_INTERNO", ""))
        cod_ean = smart_str(row.get("COD_EAN", ""))

        # Base obrigatória
        base_text = smart_str(row.get(BASE_COL, ""))

        extra = self._build_extra_block(row)
        desc_final = f"{base_text} {extra}".strip() if extra else base_text

        return f"{cod_interno}|{cod_ean}|{desc_final}"

    def on_preview(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Carregue dados primeiro (Planilha ou Consulta).")
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
            messagebox.showwarning("Atenção", "Carregue dados primeiro (Planilha ou Consulta).")
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


if __name__ == "__main__":
    app = ExportadorTXTApp()
    app.mainloop()
