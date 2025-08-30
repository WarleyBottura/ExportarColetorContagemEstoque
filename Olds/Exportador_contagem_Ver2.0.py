#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exportador TXT – XLS/CSV **ou** consulta PostgreSQL (PDV) — (COM INI, EAN13 opcional)
-----------------------------------------------------------------------------
- Layout: Planilha e BD ocupam a **mesma área** (dentro de `self.flex`), alternando com `grid()`
- Prévia fixa (~10 linhas)
- Seleções/labels preservados quando a lista é redesenhada
- **Validação EAN13** (checkbox):
  - < 13 dígitos → completa com zeros à esquerda
  - > 13 dígitos → ignora na saída e registra em `ean13_invalid.log`
- **Sem literais "
" no código**: usamos `os.linesep` (NL) e strings de aspas triplas

Requisitos: `pandas`, `openpyxl`, `xlrd==1.2.0`, `psycopg2-binary`
"""
from __future__ import annotations

import os
import sys
import csv
import math
import webbrowser
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime

# Nova convenção para quebras de linha
NL = os.linesep

try:
    import pandas as pd
except Exception:
    raise SystemExit(
        "Erro: pandas não está instalado. Instale com: pip install pandas openpyxl xlrd==1.2.0"
    )

_psycopg2_err: Optional[str] = None
try:
    import psycopg2  # type: ignore
except Exception as e:  # adia o erro para quando o usuário optar por BD
    _psycopg2_err = str(e)
    psycopg2 = None  # type: ignore

# ------------------------------ INI / Config ------------------------------
import configparser
import pathlib
import subprocess

APP_NAME = (
    "Exportador para APP de Contagem – Planilha ou BD PostgreSQL (By Bottura)"
)
INI_BASENAME = "exportador_contagem.ini"

# SQL padrão em aspas triplas (sem literais)
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


def _default_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser(interpolation=None)
    cfg["database"] = {
        "host": "localhost",
        "port": "5432",
        "name": "PDV",
        "user": "postgres",
        "password": "",
    }
    cfg["query"] = {"sql": DEFAULT_SQL.strip()}
    cfg["format"] = {
        "BASE_COL": "DES_PRODUTO",
        "MANDATORY_PREFIX": "COD_INTERNO,COD_EAN",
    }
    return cfg


def ini_path() -> pathlib.Path:
    try:
        base_dir = pathlib.Path(
            getattr(sys, "_MEIPASS", pathlib.Path(__file__).resolve().parent)
        )
    except Exception:
        base_dir = pathlib.Path(".").resolve()
    return base_dir / INI_BASENAME


def ensure_ini_and_prompt(cfg_path: pathlib.Path) -> configparser.ConfigParser:
    if cfg_path.exists():
        cfg = configparser.ConfigParser(interpolation=None)
        try:
            cfg.read(cfg_path, encoding="utf-8")
            _ = cfg.sections()
            return cfg
        except Exception:
            # Se deu erro de parsing, faz backup e recria
            try:
                bak = cfg_path.with_suffix(cfg_path.suffix + ".bak")
                i = 1
                while bak.exists():
                    bak = cfg_path.with_suffix(cfg_path.suffix + f".bak{i}")
                    i += 1
                cfg_path.replace(bak)
            except Exception:
                pass
            cfg = _default_config()
            with cfg_path.open("w", encoding="utf-8") as f:
                cfg.write(f)
            tk.Tk().withdraw()
            messagebox.showwarning(
                "INI regenerado",
                "Seu .ini estava inválido. Fizemos backup (.bak) e criamos um novo padrão.",
            )
            return cfg

    # Não existe: criar e perguntar se deseja editar
    cfg = _default_config()
    with cfg_path.open("w", encoding="utf-8") as f:
        cfg.write(f)

    root = tk.Tk()
    root.withdraw()
    resp = messagebox.askyesno(
        title="Configuração inicial",
        message=(
            "Um arquivo de configuração foi criado em:" + NL +
            f"{cfg_path}" + 2*NL +
            "Você quer EDITAR o .ini agora?" + 2*NL +
            "Sim = Editar .ini e encerrar"
        ),
    )
    root.destroy()

    if resp:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(cfg_path))  # type: ignore
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(cfg_path)])
            else:
                subprocess.Popen(["xdg-open", str(cfg_path)])
        except Exception as e:
            tk.Tk().withdraw()
            messagebox.showerror("Erro ao abrir .ini", f"{e}")
        raise SystemExit(0)

    return cfg


def parse_mandatory_prefix(value: str) -> Tuple[str, ...]:
    parts = [p.strip().upper() for p in (value or "").split(",")]
    return tuple([p for p in parts if p])


def load_config_values(cfg: configparser.ConfigParser):
    global BASE_COL, MANDATORY_PREFIX
    BASE_COL = cfg.get("format", "BASE_COL", fallback="DES_PRODUTO").strip().upper()
    MANDATORY_PREFIX = parse_mandatory_prefix(
        cfg.get("format", "MANDATORY_PREFIX", fallback="COD_INTERNO,COD_EAN")
    )
    db = {
        "host": cfg.get("database", "host", fallback="localhost"),
        "port": cfg.get("database", "port", fallback="5432"),
        "name": cfg.get("database", "name", fallback="PDV"),
        "user": cfg.get("database", "user", fallback="postgres"),
        "password": cfg.get("database", "password", fallback=""),
    }
    sql = cfg.get("query", "sql", fallback=DEFAULT_SQL).strip()
    return db, sql


def save_config_from_state(
    cfg_path: pathlib.Path, db: dict, sql: str, base_col: str, mandatory_prefix: Tuple[str, ...]
):
    cfg = configparser.ConfigParser(interpolation=None)
    cfg["database"] = {
        "host": db.get("host", "localhost"),
        "port": db.get("port", "5432"),
        "name": db.get("name", "PDV"),
        "user": db.get("user", "postgres"),
        "password": db.get("password", ""),
    }
    cfg["query"] = {"sql": sql}
    cfg["format"] = {
        "BASE_COL": (base_col or "DES_PRODUTO").upper(),
        "MANDATORY_PREFIX": ",".join(
            [p.strip().upper() for p in mandatory_prefix]
        )
        if mandatory_prefix
        else "COD_INTERNO,COD_EAN",
    }
    with cfg_path.open("w", encoding="utf-8") as f:
        cfg.write(f)


# ------------------------------ Constantes (podem ser sobrescritas via INI) ------------------------------
BASE_COL = "DES_PRODUTO"  # base fixa
MANDATORY_PREFIX: Tuple[str, ...] = ("COD_INTERNO", "COD_EAN")


# ------------------------------ Utilidades ------------------------------

def smart_str(x: Any) -> str:
    if x is None:
        return ""
    if isinstance(x, float):
        if math.isnan(x):
            return ""
        if x.is_integer():
            return str(int(x))
    return str(x).strip()


def abbrev_label(colname: str) -> str:
    if not colname:
        return "VAL"
    seg = colname.replace("-", "_").replace(" ", "_").split("_")[0]
    seg = seg.strip().upper()
    return seg if len(seg) <= 4 else seg[:4]


def one_line(text: Any) -> str:
    """Converte texto multi-linha em uma única linha, sem usar 
 literais."""
    return " ".join(str(text).splitlines())


# ------------------------------ App ------------------------------


class ExportadorTXTApp(tk.Tk):
    def __init__(self, cfg_path: pathlib.Path, db_defaults: dict, sql_default: str):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1180x820")

        self.ini_path = cfg_path

        # Estado
        self.df: Optional[pd.DataFrame] = None
        self.file_path: Optional[str] = None
        self.sheet_names: List[str] = []
        self.col_vars: Dict[str, tk.BooleanVar] = {}
        self.label_vars: Dict[str, tk.StringVar] = {}

        # Fonte de dados
        self.source_var = tk.StringVar(value="file")

        # Config de formatação
        self.opening_var = tk.StringVar(value="(")
        self.closing_var = tk.StringVar(value=")")
        self.pair_sep_var = tk.StringVar(value=" / ")
        self.label_sep_var = tk.StringVar(value=": ")

        # Config de DB (do INI)
        self.db_host = tk.StringVar(value=db_defaults.get("host", "localhost"))
        self.db_port = tk.StringVar(value=db_defaults.get("port", "5432"))
        self.db_name = tk.StringVar(value=db_defaults.get("name", "PDV"))
        self.db_user = tk.StringVar(value=db_defaults.get("user", "postgres"))
        self.db_pass = tk.StringVar(value=db_defaults.get("password", ""))
        self.sql_text = tk.StringVar(value=sql_default.strip())

        # Validação EAN-13
        self.validate_ean13 = tk.BooleanVar(value=False)
        self._ean13_oversized: List[Dict[str, Any]] = []

        self._build_ui()

    # -------------------------- UI --------------------------
    def _build_ui(self):
        # Top bar
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=8)

        tk.Button(
            top,
            text='Link: "Relatório Custo Estoque" (web)',
            command=self.on_open_export_link,
            cursor="hand2",
            relief="flat",
            fg="blue",
            activeforeground="blue",
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(top, text="Fonte de dados:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            top,
            text="Planilha (XLS/XLSX/CSV)",
            value="file",
            variable=self.source_var,
            command=self.on_change_source,
        ).pack(side=tk.LEFT, padx=(6, 2))
        ttk.Radiobutton(
            top,
            text="Consultar BD (PostgreSQL)",
            value="db",
            variable=self.source_var,
            command=self.on_change_source,
        ).pack(side=tk.LEFT, padx=(6, 2))

        # Linha formato
        fmt = ttk.Frame(self)
        fmt.pack(side=tk.TOP, fill=tk.X, padx=10, pady=6)
        ttk.Label(fmt, text="Base da descrição:", foreground="#0a0").pack(
            side=tk.LEFT
        )
        ttk.Label(fmt, text=BASE_COL, font=("Segoe UI", 10, "bold")).pack(
            side=tk.LEFT, padx=(4, 18)
        )
        ttk.Label(fmt, text="Abertura").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.opening_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Fechamento").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.closing_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. pares").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.pair_sep_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. rótulo/valor").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.label_sep_var).pack(side=tk.LEFT)

        # Ações principais
        actions = ttk.Frame(self)
        actions.pack(side=tk.TOP, fill=tk.X, padx=10, pady=8)
        ttk.Button(actions, text="Pré-visualizar (10 linhas)", command=self.on_preview).pack(
            side=tk.LEFT
        )
        ttk.Button(actions, text="Salvar TXT…", command=self.on_save).pack(
            side=tk.LEFT, padx=8
        )
        ttk.Checkbutton(
            actions, text="Validar EAN no formato EAN13", variable=self.validate_ean13
        ).pack(side=tk.LEFT, padx=(16, 0))

        # Prévia fixa
        preview_wrap = ttk.LabelFrame(
            self, text="Prévia do resultado – COD_INTERNO|COD_EAN|DES_PRODUTO"
        )
        preview_wrap.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0, 10))
        self.preview_txt = tk.Text(preview_wrap, height=10, wrap="none")
        self.preview_txt.pack(fill=tk.BOTH, expand=True)
        font = tkfont.nametofont(self.preview_txt["font"])
        line_px = font.metrics("linespace")
        preview_wrap.configure(height=line_px * 10 + 24)
        preview_wrap.pack_propagate(False)

        # Flex container (mid + fonte de dados)
        self.flex = ttk.Frame(self)
        self.flex.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.flex.rowconfigure(0, weight=1)  # mid
        self.flex.rowconfigure(1, weight=7)  # fonte (file/db)
        self.flex.columnconfigure(0, weight=1)

        # MID: lista de colunas
        mid = ttk.Frame(self.flex)
        mid.grid(row=0, column=0, sticky="nsew", padx=0, pady=6)

        head = ttk.Frame(mid)
        self.cols_per_block = tk.IntVar(value=3)
        ttk.Label(head, text="Cols:").pack(side=tk.RIGHT)
        tk.Spinbox(
            head,
            from_=1,
            to=4,
            width=3,
            textvariable=self.cols_per_block,
            command=self._populate_columns_ui,
        ).pack(side=tk.RIGHT, padx=(6, 0))
        head.pack(fill=tk.X)
        ttk.Label(
            head,
            text=(
                "Marque as colunas a concatenar em DES_PRODUTO "
                "(COD_INTERNO, COD_EAN e a base 'DES_PRODUTO' já são fixos no arquivo de saída):"
            ),
            font=("Segoe UI", 10, "bold"),
        ).pack(side=tk.LEFT)

        columns_frame = ttk.Frame(mid)
        columns_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(columns_frame, borderwidth=0)
        self.scroll_y = ttk.Scrollbar(columns_frame, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll_y.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # File frame (dentro do flex)
        self.file_frame = ttk.LabelFrame(self.flex, text="Planilha (XLS/XLSX/CSV)")
        row1 = ttk.Frame(self.file_frame)
        row1.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(row1, text="Abrir Planilha…", command=self.on_open_file).pack(side=tk.LEFT)
        self.file_lbl = ttk.Label(row1, text="(nenhum arquivo)")
        self.file_lbl.pack(side=tk.LEFT, padx=10)
        ttk.Label(row1, text="Aba/Sheet:").pack(side=tk.LEFT, padx=(20, 4))
        self.sheet_cbx = ttk.Combobox(row1, state="disabled", width=28)
        self.sheet_cbx.pack(side=tk.LEFT)
        self.sheet_cbx.bind("<<ComboboxSelected>>", self.on_select_sheet)
        self.file_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 6))

        # DB frame (dentro do flex)
        self.db_frame = ttk.LabelFrame(self.flex, text="Conexão BD (PostgreSQL)")
        self._build_db_frame_contents()
        self.db_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 6))
        self.db_frame.grid_remove()  # começa oculto

        # Visibilidade inicial
        self.on_change_source()

    def _build_db_frame_contents(self):
        f = self.db_frame
        grid = ttk.Frame(f)
        grid.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        ttk.Label(grid, text="Host:").grid(row=0, column=0, sticky="e")
        ttk.Entry(grid, textvariable=self.db_host, width=18).grid(
            row=0, column=1, padx=6, pady=2, sticky="w"
        )
        ttk.Label(grid, text="Porta:").grid(row=0, column=2, sticky="e")
        ttk.Entry(grid, textvariable=self.db_port, width=8).grid(
            row=0, column=3, padx=6, pady=2, sticky="w"
        )
        ttk.Label(grid, text="Banco:").grid(row=0, column=4, sticky="e")
        ttk.Entry(grid, textvariable=self.db_name, width=16).grid(
            row=0, column=5, padx=6, pady=2, sticky="w"
        )
        ttk.Label(grid, text="Usuário:").grid(row=1, column=0, sticky="e")
        ttk.Entry(grid, textvariable=self.db_user, width=18).grid(
            row=1, column=1, padx=6, pady=2, sticky="w"
        )
        ttk.Label(grid, text="Senha:").grid(row=1, column=2, sticky="e")
        ttk.Entry(grid, textvariable=self.db_pass, width=18, show="*").grid(
            row=1, column=3, padx=6, pady=2, sticky="w"
        )
        ttk.Button(grid, text="Executar consulta", command=self.on_run_query).grid(
            row=1, column=5, padx=6, pady=2, sticky="e"
        )
        ttk.Button(grid, text="Salvar .ini", command=self.on_save_ini).grid(
            row=1, column=6, padx=6, pady=2, sticky="e"
        )

        sql_box = ttk.Frame(f)
        sql_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        ttk.Label(sql_box, text="SQL (editável):").pack(anchor="w")
        self.sql_textbox = tk.Text(sql_box, height=8, wrap="none")
        self.sql_textbox.pack(fill=tk.BOTH, expand=True)
        self.sql_textbox.insert("1.0", self.sql_text.get())

    # -------------------------- Fluxo --------------------------
    def on_open_export_link(self):
        webbrowser.open(
            "https://app.mentorasolucoes.com.br/Voti-1.0.7/relatorios_base/frm_rel_custo_estoque.xhtml",
            new=1,
        )

    def on_change_source(self):
        """Alterna Planilha/BD dentro do flex."""
        src = self.source_var.get()
        if src == "file":
            self.db_frame.grid_remove()
            self.file_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 6))
        else:
            self.file_frame.grid_remove()
            self.db_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 6))

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
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    sample = f.read(4096)
                    f.seek(0)
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;	|")
                    df = pd.read_csv(f, sep=dialect.delimiter)
            except Exception:
                df = pd.read_csv(path, sep=",", encoding="utf-8", engine="python")
        elif ext in (".xlsx", ".xls"):
            xl = pd.ExcelFile(path)
            self.sheet_names = xl.sheet_names
            self.sheet_cbx.config(state="readonly", values=self.sheet_names)
            self.sheet_cbx.set(self.sheet_names[0])
            df = xl.parse(self.sheet_names[0])
        else:
            raise ValueError("Formato não suportado. Use XLSX, XLS ou CSV.")

        df.columns = [c.strip().upper() for c in df.columns]
        self.df = df
        self._populate_columns_ui()

    def on_select_sheet(self, event=None):
        if not self.file_path:
            return
        try:
            xl = pd.ExcelFile(self.file_path)
            sheet = self.sheet_cbx.get()
            df = xl.parse(sheet)
            df.columns = [c.strip().upper() for c in df.columns]
            self.df = df
            self._populate_columns_ui()
        except Exception as e:
            messagebox.showerror("Erro ao trocar de aba", str(e))

    # -------------------------- Banco de Dados --------------------------
    def on_run_query(self):
        if psycopg2 is None:
            details = _psycopg2_err or "biblioteca ausente"
            messagebox.showerror(
                "psycopg2 não disponível",
                "Não foi possível importar psycopg2." + NL +
                f"Detalhes: {details}" + 2*NL +
                "Instale com:  pip install psycopg2-binary",
            )
            return

        host = self.db_host.get().strip() or "localhost"
        port = self.db_port.get().strip() or "5432"
        db = self.db_name.get().strip() or "PDV"
        user = self.db_user.get().strip()
        pwd = self.db_pass.get()

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
            df.columns = [c.strip().upper() for c in df.columns]
        except Exception as e:
            try:
                conn.close()
            except Exception:
                pass
            messagebox.showerror("Erro ao executar SQL", str(e))
            return
        finally:
            try:
                conn.close()
            except Exception:
                pass

        self.df = df
        self._populate_columns_ui()
        self.on_preview()
        messagebox.showinfo("Consulta concluída", f"Linhas retornadas: {len(df)}")

    def on_save_ini(self):
        try:
            base_col = BASE_COL
            mandatory = MANDATORY_PREFIX
            db = {
                "host": self.db_host.get().strip() or "localhost",
                "port": self.db_port.get().strip() or "5432",
                "name": self.db_name.get().strip() or "PDV",
                "user": self.db_user.get().strip(),
                "password": self.db_pass.get(),
            }
            sql = self.sql_textbox.get("1.0", "end").strip()
            save_config_from_state(self.ini_path, db, sql, base_col, mandatory)
            messagebox.showinfo("Configuração salva", "Arquivo atualizado:" + NL + f"{self.ini_path}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar .ini", str(e))

    # -------------------------- EAN helpers & logging --------------------------
    def _ean13_fix_and_validate(self, value: Any) -> Tuple[bool, str]:
        s = smart_str(value)
        digits = "".join(ch for ch in s if ch.isdigit())
        if len(digits) > 13:
            return False, ""
        return True, digits.rjust(13, "0")

    def _write_ean13_log(self, context: str):
        if not self._ean13_oversized:
            return
        try:
            base = ini_path().parent
            log_path = base / "ean13_invalid.log"
            with open(log_path, "a", encoding="utf-8") as fh:
                fh.write(
                    "[" + datetime.now().strftime('%Y-%m-%d %H:%M:%S') + "] "
                    + f"{context} - Itens com EAN > 13 dígitos: {len(self._ean13_oversized)}" + NL
                )
                fh.write("COD_INTERNO	COD_EAN	DES_PRODUTO	LEN" + NL)
                for rec in self._ean13_oversized:
                    cod_in = str(rec.get("COD_IN", ""))
                    cod_ean = str(rec.get("COD_EAN", ""))
                    descr = one_line(rec.get("DES_PRODUTO", ""))
                    length = str(rec.get("LEN", ""))
                    fh.write(f"{cod_in}	{cod_ean}	{descr}	{length}" + NL)
                fh.write(NL)
        except Exception:
            pass

    # -------------------------- UI de colunas --------------------------
    def _populate_columns_ui(self):
        prev_selected = {col for col, var in self.col_vars.items() if var.get()}
        prev_labels = {col: sv.get() for col, sv in self.label_vars.items()}

        for w in self.inner.winfo_children():
            w.destroy()
        self.col_vars.clear()
        self.label_vars.clear()

        if self.df is None:
            return

        cols = [c.strip().upper() for c in self.df.columns]
        if BASE_COL not in cols:
            messagebox.showerror(
                "Coluna base ausente",
                (
                    f"A coluna base '{BASE_COL}' não foi encontrada." + NL +
                    "Colunas disponíveis:" + NL + "- " + ("- ".join(cols))
                ),
            )
            return

        warn_opt = [c for c in ("COD_INTERNO", "COD_EAN") if c not in cols]
        if warn_opt:
            messagebox.showwarning(
                "Aviso",
                (
                    "As seguintes colunas opcionais não foram encontradas:" + NL +
                    (NL.join("- " + c for c in warn_opt)) + NL +
                    "Elas sairão vazias no TXT."
                ),
            )

        excluded = set([BASE_COL, *MANDATORY_PREFIX])
        cols_all = [c for c in cols if c not in excluded]

        ncols = max(1, int(getattr(self, "cols_per_block", tk.IntVar(value=2)).get()))
        total = len(cols_all)
        if total == 0:
            return
        rows = math.ceil(total / ncols)

        for j in range(ncols):
            self.inner.columnconfigure(j, weight=1)

        for j in range(ncols):
            header = ttk.Frame(self.inner)
            header.grid(row=0, column=j, sticky="ew", padx=(4, 4), pady=(4, 2))
            ttk.Label(header, text="Usar", width=6).grid(row=0, column=0, padx=4)
            ttk.Label(header, text="Coluna", width=24).grid(row=0, column=1, padx=4)
            ttk.Label(header, text="Rótulo (editável)", width=16).grid(row=0, column=2, padx=4)
            header.columnconfigure(1, weight=1)
            header.columnconfigure(2, weight=1)

        for idx, col in enumerate(cols_all):
            r = 1 + (idx % rows)
            c = idx // rows
            rowf = ttk.Frame(self.inner)
            rowf.grid(row=r, column=c, sticky="ew", padx=4, pady=2)
            v = tk.BooleanVar(value=(col in prev_selected))
            self.col_vars[col] = v
            ttk.Checkbutton(rowf, variable=v).grid(row=0, column=0, padx=(0, 8))
            ttk.Label(rowf, text=col).grid(row=0, column=1, sticky="w")
            sv = tk.StringVar(value=prev_labels.get(col, abbrev_label(col)))
            self.label_vars[col] = sv
            ttk.Entry(rowf, textvariable=sv, width=14).grid(
                row=0, column=2, padx=(8, 0), sticky="ew"
            )
            rowf.columnconfigure(1, weight=1)
            rowf.columnconfigure(2, weight=1)

    # -------------------------- Montagem de linha / prévia / salvar --------------------------
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

    def _build_line(self, row: pd.Series) -> Optional[str]:
        cod_interno = smart_str(row.get("COD_INTERNO", ""))
        raw_ean = row.get("COD_EAN", "")
        if self.validate_ean13.get():
            ok, ean = self._ean13_fix_and_validate(raw_ean)
            if not ok:
                self._ean13_oversized.append(
                    {
                        "COD_IN": cod_interno,
                        "COD_EAN": smart_str(raw_ean),
                        "DES_PRODUTO": smart_str(row.get(BASE_COL, "")),
                        "LEN": len("".join(ch for ch in smart_str(raw_ean) if ch.isdigit())),
                    }
                )
                return None
            cod_ean = ean
        else:
            cod_ean = smart_str(raw_ean)
        base_text = smart_str(row.get(BASE_COL, ""))
        extra = self._build_extra_block(row)
        desc_final = (base_text + " " + extra).strip() if extra else base_text
        return f"{cod_interno}|{cod_ean}|{desc_final}"

    def on_preview(self):
        if self.df is None:
            messagebox.showwarning(
                "Atenção", "Carregue dados primeiro (Planilha ou Consulta)."
            )
            return
        if BASE_COL not in self.df.columns:
            messagebox.showerror("Erro", f"Coluna base '{BASE_COL}' não encontrada.")
            return
        self._ean13_oversized = []
        lines: List[str] = []
        for _, row in self.df.head(10).iterrows():
            line = self._build_line(row)
            if line is not None:
                lines.append(line)
        self.preview_txt.delete("1.0", tk.END)
        self.preview_txt.insert(tk.END, NL.join(lines))
        if self.validate_ean13.get() and self._ean13_oversized:
            self._write_ean13_log("Pré-visualização")
            messagebox.showwarning(
                "EAN > 13 dígitos",
                f"{len(self._ean13_oversized)} item(ns) com EAN > 13 dígitos foram ignorados na prévia." + NL +
                "Geramos o log: ean13_invalid.log",
            )

    def on_save(self):
        if self.df is None:
            messagebox.showwarning(
                "Atenção", "Carregue dados primeiro (Planilha ou Consulta)."
            )
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
        self._ean13_oversized = []
        try:
            with open(out_path, "w", encoding="utf-8") as f:
                for _, row in self.df.iterrows():
                    line = self._build_line(row)
                    if line is not None:
                        f.write(line + NL)
            if self.validate_ean13.get() and self._ean13_oversized:
                self._write_ean13_log("Exportação")
                messagebox.showwarning(
                    "Exportado com avisos",
                    "Arquivo salvo em:" + NL + f"{out_path}" + 2*NL +
                    f"{len(self._ean13_oversized)} item(ns) com EAN > 13 dígitos foram ignorados." + NL +
                    "Geramos o log: ean13_invalid.log",
                )
            else:
                messagebox.showinfo("Concluído", "Arquivo salvo em:" + NL + f"{out_path}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))


# ------------------------------ bootstrap ------------------------------

def main():
    cfg_path = ini_path()
    cfg = ensure_ini_and_prompt(cfg_path)
    db_defaults, sql_default = load_config_values(cfg)
    app = ExportadorTXTApp(cfg_path, db_defaults, sql_default)
    app.mainloop()


if __name__ == "__main__":
    main()
