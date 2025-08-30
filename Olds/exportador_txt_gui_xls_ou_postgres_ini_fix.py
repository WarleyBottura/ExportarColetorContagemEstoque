#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exportador TXT – XLS/CSV **ou** consulta PostgreSQL (PDV) — com INI
-------------------------------------------------------------------
Correções/recursos desta build:
- INI com `interpolation=None` (não quebra com % em SQL: ex. LIKE '%%abc%%').
- Primeiro uso: cria `exportador_contagem.ini` com padrões e pergunta se deseja editar.
- Botão "Salvar .ini" para persistir credenciais/SQL.
- Robustez: se o INI existente estiver corrompido (erro de parsing), é feito backup
  automático como `.bak` e um novo INI é gerado com padrões.

Requisitos (pip): pandas, openpyxl, xlrd==1.2.0, psycopg2-binary
"""
from __future__ import annotations
import os
import sys
import csv
import math
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any, Optional, Tuple

try:
    import pandas as pd
except Exception:
    raise SystemExit("Erro: pandas não está instalado. Instale com: pip install pandas openpyxl xlrd==1.2.0")

# psycopg2 é opcional (usado apenas em "Consultar BD PDV")
_psycopg2_err: Optional[str] = None
try:
    import psycopg2  # type: ignore
except Exception as e:  # adia o erro para quando o usuário optar por BD
    _psycopg2_err = str(e)
    psycopg2 = None  # type: ignore

# ------------------------------ Config / INI ------------------------------
import configparser
import pathlib
import subprocess

INI_BASENAME = "exportador_contagem.ini"

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
        "password": ""
    }
    cfg["query"] = {"sql": DEFAULT_SQL.strip("\n")}
    cfg["format"] = {
        "BASE_COL": "DES_PRODUTO",
        "MANDATORY_PREFIX": "COD_INTERNO,COD_EAN"
    }
    return cfg

def ini_path() -> pathlib.Path:
    try:
        base_dir = pathlib.Path(getattr(sys, "_MEIPASS", pathlib.Path(__file__).resolve().parent))
    except Exception:
        base_dir = pathlib.Path(".").resolve()
    return base_dir / INI_BASENAME

def _backup_bad_ini(cfg_path: pathlib.Path):
    try:
        bak = cfg_path.with_suffix(cfg_path.suffix + ".bak")
        i = 1
        while bak.exists():
            bak = cfg_path.with_suffix(cfg_path.suffix + f".bak{i}")
            i += 1
        cfg_path.replace(bak)
        return bak
    except Exception:
        return None

def ensure_ini_and_prompt(cfg_path: pathlib.Path) -> configparser.ConfigParser:
    if cfg_path.exists():
        cfg = configparser.ConfigParser(interpolation=None)
        try:
            cfg.read(cfg_path, encoding="utf-8")
            # Força leitura de uma chave para disparar ParsingError cedo, se houver
            _ = cfg.sections()
            return cfg
        except Exception as e:
            # Faz backup e recria com defaults
            _backup_bad_ini(cfg_path)
            cfg = _default_config()
            with cfg_path.open("w", encoding="utf-8") as f:
                cfg.write(f)
            tk.Tk().withdraw()
            messagebox.showwarning(
                "INI regenerado",
                "Seu arquivo INI estava corrompido e foi salvo como .bak.\n"
                "Geramos um novo com valores padrão."
            )
            return cfg

    # Não existe: criar com defaults e perguntar se deseja editar
    cfg = _default_config()
    with cfg_path.open("w", encoding="utf-8") as f:
        cfg.write(f)

    root = tk.Tk(); root.withdraw()
    resp = messagebox.askyesno(
        title="Configuração inicial",
        message=(
            "Um arquivo de configuração foi criado em:\n"
            f"{cfg_path}\n\n"
            "Você quer EDITAR o .ini agora?\n\n"
            "Sim = Editar .ini e encerrar\n"
        )
    )
    root.destroy()

    if resp:  # Sim -> editar e encerrar
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(cfg_path))  # type: ignore
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(cfg_path)])
            else:
                subprocess.Popen(["xdg-open", str(cfg_path)])
        except Exception as e:
            tk.Tk().withdraw(); messagebox.showerror("Erro ao abrir .ini", f"{e}")
        raise SystemExit(0)

    return cfg

def parse_mandatory_prefix(value: str) -> Tuple[str, ...]:
    parts = [p.strip().upper() for p in (value or "").split(",")]
    return tuple([p for p in parts if p])

def load_config_values(cfg: configparser.ConfigParser):
    global BASE_COL, MANDATORY_PREFIX
    BASE_COL = cfg.get("format", "BASE_COL", fallback="DES_PRODUTO").strip().upper()
    MANDATORY_PREFIX = parse_mandatory_prefix(cfg.get("format", "MANDATORY_PREFIX", fallback="COD_INTERNO,COD_EAN"))
    db = {
        "host": cfg.get("database", "host", fallback="localhost"),
        "port": cfg.get("database", "port", fallback="5432"),
        "name": cfg.get("database", "name", fallback="PDV"),
        "user": cfg.get("database", "user", fallback="postgres"),
        "password": cfg.get("database", "password", fallback=""),
    }
    sql = cfg.get("query", "sql", fallback=DEFAULT_SQL).strip()
    return db, sql

def save_config_from_state(cfg_path: pathlib.Path, db: dict, sql: str, base_col: str, mandatory_prefix: Tuple[str, ...]):
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
        "MANDATORY_PREFIX": ",".join([p.strip().upper() for p in mandatory_prefix]) if mandatory_prefix else "COD_INTERNO,COD_EAN"
    }
    with cfg_path.open("w", encoding="utf-8") as f:
        cfg.write(f)

# ------------------------------ Constantes (podem ser sobrescritas pelo INI) ------------------------------
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
        return str(x)
    return str(x).strip()

def abbrev_label(colname: str) -> str:
    if not colname:
        return "VAL"
    seg = colname.replace("-", "_").replace(" ", "_").split("_")[0].strip().upper()
    return seg if len(seg) <= 4 else seg[:4]

# ------------------------------ App ------------------------------
class ExportadorTXTApp(tk.Tk):
    def __init__(self, ini_cfg_path: pathlib.Path, db_defaults: dict, sql_default: str):
        super().__init__()
        self.title("Exportador para APP de Contagem – Planilha ou BD PDV (By Bottura)")
        self.geometry("1180x860")
        self.ini_path = ini_cfg_path

        # Estado
        self.df: pd.DataFrame | None = None
        self.file_path: str | None = None
        self.sheet_names: List[str] = []
        self.col_vars: Dict[str, tk.BooleanVar] = {}
        self.label_vars: Dict[str, tk.StringVar] = {}

        # Fonte de dados: "file" | "db"
        self.source_var = tk.StringVar(value="file")

        # Config visual
        self.opening_var = tk.StringVar(value="(")
        self.closing_var = tk.StringVar(value=")")
        self.pair_sep_var = tk.StringVar(value=" / ")
        self.label_sep_var = tk.StringVar(value=": ")

        # Config de DB (carregadas do INI)
        self.db_host = tk.StringVar(value=db_defaults.get("host", "localhost"))
        self.db_port = tk.StringVar(value=db_defaults.get("port", "5432"))
        self.db_name = tk.StringVar(value=db_defaults.get("name", "PDV"))
        self.db_user = tk.StringVar(value=db_defaults.get("user", "postgres"))
        self.db_pass = tk.StringVar(value=db_defaults.get("password", ""))
        self.sql_text = tk.StringVar(value=sql_default.strip())

        # Monta UI
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
            activeforeground="blue"
        ).pack(side=tk.LEFT, padx=(0,12))

        ttk.Label(top, text="Fonte de dados:").pack(side=tk.LEFT)
        ttk.Radiobutton(top, text="Planilha (XLS/XLSX/CSV)", value="file", variable=self.source_var,
                        command=self.on_change_source).pack(side=tk.LEFT, padx=(6,2))
        ttk.Radiobutton(top, text="Consultar BD PDV (PostgreSQL)", value="db", variable=self.source_var,
                        command=self.on_change_source).pack(side=tk.LEFT, padx=(6,2))

        # FILE frame
        self.file_frame = ttk.Frame(self)
        self.file_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0,6))
        ttk.Button(self.file_frame, text="Abrir Planilha…", command=self.on_open_file).pack(side=tk.LEFT)
        self.file_lbl = ttk.Label(self.file_frame, text="(nenhum arquivo)")
        self.file_lbl.pack(side=tk.LEFT, padx=10)
        ttk.Label(self.file_frame, text="Aba/Sheet:").pack(side=tk.LEFT, padx=(20, 4))
        self.sheet_cbx = ttk.Combobox(self.file_frame, state="disabled", width=28)
        self.sheet_cbx.pack(side=tk.LEFT)
        self.sheet_cbx.bind("<<ComboboxSelected>>", self.on_select_sheet)

        # DB frame
        self.db_frame = ttk.LabelFrame(self, text="Conexão BD PDV (PostgreSQL)")
        self._build_db_frame_contents()

        # Linha de formatação e base
        fmt = ttk.Frame(self)
        fmt.pack(side=tk.TOP, fill=tk.X, padx=10, pady=6)
        ttk.Label(fmt, text="Base da descrição:", foreground="#0a0").pack(side=tk.LEFT)
        self.base_lbl = ttk.Label(fmt, text=BASE_COL, font=("Segoe UI", 10, "bold"))
        self.base_lbl.pack(side=tk.LEFT, padx=(4, 18))

        ttk.Label(fmt, text="Abertura").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.opening_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Fechamento").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=4, textvariable=self.closing_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. pares").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.pair_sep_var).pack(side=tk.LEFT)
        ttk.Label(fmt, text="Sep. rótulo/valor").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(fmt, width=10, textvariable=self.label_sep_var).pack(side=tk.LEFT)

        # Centro: lista de colunas
        mid = ttk.Frame(self)
        mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=6)

        head = ttk.Frame(mid)
        head.pack(fill=tk.X)
        ttk.Label(head, text=(
            "Marque as colunas a concatenar em DES_PRODUTO "
            "(COD_INTERNO, COD_EAN e a base já são fixos no arquivo de saída):"
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

        ttk.Label(grid, text="Host:").grid(row=0, column=0, sticky="e")
        ttk.Entry(grid, textvariable=self.db_host, width=18).grid(row=0, column=1, padx=6, pady=2, sticky="w")
        ttk.Label(grid, text="Porta:").grid(row=0, column=2, sticky="e")
        ttk.Entry(grid, textvariable=self.db_port, width=8).grid(row=0, column=3, padx=6, pady=2, sticky="w")
        ttk.Label(grid, text="Banco:").grid(row=0, column=4, sticky="e")
        ttk.Entry(grid, textvariable=self.db_name, width=16).grid(row=0, column=5, padx=6, pady=2, sticky="w")

        ttk.Label(grid, text="Usuário:").grid(row=1, column=0, sticky="e")
        ttk.Entry(grid, textvariable=self.db_user, width=18).grid(row=1, column=1, padx=6, pady=2, sticky="w")
        ttk.Label(grid, text="Senha:").grid(row=1, column=2, sticky="e")
        ttk.Entry(grid, textvariable=self.db_pass, width=18, show="*").grid(row=1, column=3, padx=6, pady=2, sticky="w")

        ttk.Button(grid, text="Executar consulta", command=self.on_run_query).grid(row=1, column=5, padx=6, pady=2, sticky="e")
        ttk.Button(grid, text="Salvar .ini", command=self.on_save_ini).grid(row=1, column=6, padx=6, pady=2, sticky="e")

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

        # Normaliza nomes de colunas para o padrão usado no app (maiúsculas)
        df.columns = [c.strip().upper() for c in df.columns]

        self.df = df
        self._populate_columns_ui()
        messagebox.showinfo("Consulta concluída", f"Linhas retornadas: {len(df)}")

    # --- Salvar INI ---
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
            messagebox.showinfo("Configuração salva", f"Arquivo atualizado:\n{self.ini_path}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar .ini", str(e))

    # ----- UI de colunas -----
    def _populate_columns_ui(self):
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


def main():
    cfg_path = ini_path()
    cfg = ensure_ini_and_prompt(cfg_path)
    db_defaults, sql_default = load_config_values(cfg)
    # As globais BASE_COL/MANDATORY_PREFIX já foram atualizadas por load_config_values()

    app = ExportadorTXTApp(cfg_path, db_defaults, sql_default)
    app.mainloop()


if __name__ == "__main__":
    main()
