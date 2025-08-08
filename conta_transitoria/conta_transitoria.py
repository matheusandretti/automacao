# conta_transitoria.py
import sys
import re
import glob
from pathlib import Path
import pandas as pd
from itertools import combinations
from openpyxl import load_workbook  # usado no writer; leitura ignora estilos

# ======= CONFIGURAÇÕES =======
COLUMN_ALIASES = {
    "date": [r"^data(\s+do\s+lan[cç]amento)?$", r"^dt$", r"^emiss[aã]o$", r"^lan[cç]amento$", r"^data$"],
    "hist": [r"^hist[oó]rico$", r"^hist[oó]rico\s+do\s+lan[cç]amento$", r"^historico$", r"^descri[cç][aã]o$", r"^descri[cç][aã]o\s+do\s+lan[cç]amento$"],
    "debit": [r"^d[eé]bito$", r"^valor\s*d[eé]bito$", r"^vlr\s*d[eé]bito$", r"^debito$"],
    "credit":[r"^cr[eé]dito$", r"^valor\s*cr[eé]dito$", r"^vlr\s*cr[eé]dito$", r"^credito$"],
    "batch": [r"^lote$", r"^n[ºo]\s*lote$"],
}

EPS = 0.01
USE_FIRST_NOTE_ONLY = True

import tempfile
import shutil

# ======= LEITURA COM DETECÇÃO DE CABEÇALHO =======
def _detect_header_row(temp_df: pd.DataFrame) -> int:
    header_row = None
    for i, row in temp_df.iterrows():
        row_str = row.astype(str).str.lower()
        has_date = row_str.str.contains(r"\bdata\b|lan[cç]amento|emiss[aã]o", regex=True, na=False).any()
        has_hist = row_str.str.contains(r"hist[oó]rico|descri[cç][aã]o", regex=True, na=False).any()
        has_deb  = row_str.str.contains(r"d[eé]bito|debito", regex=True, na=False).any()
        has_cred = row_str.str.contains(r"cr[eé]dito|credito", regex=True, na=False).any()
        if has_date and (has_hist or has_deb or has_cred) and (has_deb or has_cred):
            header_row = i
            break
    if header_row is None:
        for i, row in temp_df.iterrows():
            row_str = row.astype(str).str.lower()
            score = 0
            score += row_str.str.contains(r"\bdata\b|lan[cç]amento|emiss[aã]o", regex=True, na=False).any()
            score += row_str.str.contains(r"hist[oó]rico|descri[cç][aã]o", regex=True, na=False).any()
            score += row_str.str.contains(r"d[eé]bito|cr[eé]dito|debito|credito", regex=True, na=False).any()
            if score >= 2:
                header_row = i
                break
    if header_row is None:
        raise ValueError("Não consegui localizar a linha de cabeçalho com 'Data/Histórico/Débito/Crédito'.")
    return header_row

def _finalize_from_temp(temp: pd.DataFrame, header_row: int) -> pd.DataFrame:
    print(f"[INFO] Cabeçalho detectado na linha (0-based): {header_row}")
    header_vals = list(temp.iloc[header_row])
    cols = []
    for idx, c in enumerate(header_vals):
        if c is None:
            cols.append(f"Unnamed: {idx}")
        else:
            name = str(c).strip()
            cols.append(name if name and name.lower() != "nan" else f"Unnamed: {idx}")
    data_rows = temp.iloc[header_row + 1 : ].copy()
    data_rows.columns = cols
    df = data_rows.dropna(axis=1, how="all")
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def read_with_header_detection(path: Path) -> pd.DataFrame:
    # 1) tenta engine 'calamine'
    try:
        temp = pd.read_excel(path, header=None, engine="calamine")
        header_row = _detect_header_row(temp)
        return _finalize_from_temp(temp, header_row)
    except Exception:
        print("[WARN] Falha no engine 'calamine' (ou não instalado). Tentando limpar via Excel/COM...")

    # 2) fallback: Excel/COM re-salva um XLSX limpo
    try:
        import win32com.client as win32
        import pythoncom
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        tmpdir = Path(tempfile.mkdtemp(prefix="xls_clean_"))
        cleaned_path = tmpdir / (path.stem + "_clean.xlsx")

        wb = excel.Workbooks.Open(str(path))
        wb.SaveAs(str(cleaned_path), FileFormat=51)  # 51 = .xlsx
        wb.Close(False)
        excel.Quit()

        temp = pd.read_excel(cleaned_path, header=None)
        header_row = _detect_header_row(temp)
        df = _finalize_from_temp(temp, header_row)

        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass

        return df

    except Exception as e2:
        raise RuntimeError(
            "Falha na leitura e limpeza do arquivo Excel.\n"
            "Alternativas: abra e 'Salvar como' .xlsx manualmente ou instale pandas-calamine."
        ) from e2

# ======= AUXILIARES =======
def normalize_cols(df: pd.DataFrame):
    colmap = {}
    for canon, patterns in COLUMN_ALIASES.items():
        for c in df.columns:
            name = str(c).strip().lower()
            if any(re.search(p, name) for p in patterns):
                colmap[canon] = c; break
    missing = [k for k in ("date", "hist", "debit", "credit") if k not in colmap]
    if missing:
        raise ValueError("Não consegui identificar as colunas obrigatórias: " + ", ".join(missing) +
                         f". Colunas disponíveis: {list(df.columns)}.\n→ Ajuste COLUMN_ALIASES.")
    print("[INFO] Colunas mapeadas:", {k: colmap[k] for k in colmap})
    return colmap

def to_number(x):
    if pd.isna(x): return pd.NA
    s = str(x).strip()
    if s == "": return pd.NA
    s = re.sub(r"[^\d,.\-]", "", s)
    if "," in s and "." in s and s.rfind(",") > s.rfind("."):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return pd.NA

def parse_date(col):
    return pd.to_datetime(col, errors="coerce", dayfirst=True)

def consolidate_history(df, col_date, col_hist):
    df = df.copy()
    df["_has_date"] = ~df[col_date].isna()
    df["_hist_add"] = (~df["_has_date"]) & df[col_hist].notna() & (df[col_hist].astype(str).str.strip() != "")
    df["_anchor_idx"] = df.index.to_series().where(df["_has_date"]).ffill()
    cont = df[df["_hist_add"]]
    if not cont.empty:
        add_text = cont.groupby("_anchor_idx")[col_hist].apply(lambda s: " | ".join(s.astype(str)))
        for anchor, extra in add_text.items():
            base = str(df.at[anchor, col_hist]) if pd.notna(df.at[anchor, col_hist]) else ""
            sep = " | " if base and extra else ""
            df.at[anchor, col_hist] = f"{base}{sep}{extra}"
    df = df[df["_has_date"]].copy()
    df.drop(columns=["_has_date", "_hist_add", "_anchor_idx"], inplace=True)
    return df

# ======= EXTRAÇÃO DE NOTA: NF-FIRST + FALLBACK =======
NF_NUMBER = re.compile(r"(?:\bNF(?:E)?\b[^\d]{0,3})(\d{3,}(?:[.\s]\d{3})*(?:[\/\-]\d{1,3})?)", re.I)
GEN_NUMBER = re.compile(r"(?<!\d)(\d{3,}(?:[.\s]\d{3})*(?:[\/\-]\d{1,3})?)(?!\d)", re.I)
DATE_LIKE  = re.compile(r"\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b", re.I)

def _normalize_note(token: str) -> str:
    parts = re.split(r"([\/\-])", token, maxsplit=1)
    if len(parts) == 3:
        base, sep, suf = parts
    else:
        base, sep, suf = parts[0], "", None
    base = re.sub(r"[.\s]", "", base)
    if re.fullmatch(r"\d+", base):
        base = str(int(base))
    return f"{base}{sep}{suf}" if sep and suf is not None else base

def extract_note_ids(text):
    if pd.isna(text): return []
    raw = str(text)

    # 1) prioriza número logo após NF/NFE
    nf_hits = [m.group(1) for m in NF_NUMBER.finditer(raw)]
    nf_hits = [h for h in nf_hits if not DATE_LIKE.search(h)]
    if nf_hits:
        best = _normalize_note(nf_hits[-1])
        return [best]

    # 2) fallback geral
    cands = []
    for m in GEN_NUMBER.finditer(raw):
        token = m.group(1)
        if DATE_LIKE.search(token):
            continue
        cands.append(_normalize_note(token))
    if not cands:
        return []
    cands.sort(key=lambda s: (("/" in s) or ("-" in s), len(re.sub(r"\D", "", s))), reverse=True)
    return [cands[0]] if USE_FIRST_NOTE_ONLY else cands

# ======= SELEÇÃO DE RESPONSÁVEIS (subset até 4 notas) =======
def _close(a, b, eps=EPS):
    return abs(a - b) <= eps

def pick_responsible_sets(por_dia_nota, dias_com_diff):
    selected = {}
    for _, row in dias_com_diff.iterrows():
        dia = row["Dia"]
        target = float(row["Diferenca"])
        sample = por_dia_nota[(por_dia_nota["Dia"] == dia) & (por_dia_nota["Diferenca"].abs() > EPS)].copy()
        diffs = list(zip(sample["NotaID"], sample["Diferenca"].astype(float)))
        found = set()

        # tenta tamanhos de 1..4
        for k in range(1, min(4, len(diffs)) + 1):
            solved = False
            for combo in combinations(diffs, k):
                s = sum(d for _, d in combo)
                if _close(s, target):
                    found = {n for n, _ in combo}
                    solved = True
                    break
            if solved:
                break

        # se não achar combinação exata, aproxima com as maiores diferenças
        if not found and diffs:
            diffs_sorted = sorted(diffs, key=lambda t: abs(t[1]), reverse=True)
            best = None; best_gap = float("inf")
            for k in range(1, min(4, len(diffs_sorted)) + 1):
                s = sum(d for _, d in diffs_sorted[:k])
                gap = abs(s - target)
                if gap < best_gap:
                    best_gap, best = gap, {n for n, _ in diffs_sorted[:k]}
            found = best or set()

        selected[dia] = found
    return selected

# ======= PIPELINE PRINCIPAL =======
def process_file(xlsx_path: Path):
    df_raw = read_with_header_detection(xlsx_path)
    if df_raw.empty:
        raise ValueError("Planilha vazia após detecção de cabeçalho.")

    colmap = normalize_cols(df_raw)
    col_date  = colmap["date"]
    col_hist  = colmap["hist"]
    col_deb   = colmap["debit"]
    col_cred  = colmap["credit"]
    col_batch = colmap.get("batch", None)

    df = consolidate_history(df_raw, col_date, col_hist)
    df[col_date] = parse_date(df[col_date])
    df["Debito"] = df[col_deb].apply(to_number).astype("Float64")
    df["Credito"] = df[col_cred].apply(to_number).astype("Float64")
    df["Valor"] = (df["Debito"].fillna(0) - df["Credito"].fillna(0)).astype("Float64")

    df = df[df[col_date].notna()].copy()
    df = df[~(df["Debito"].fillna(0).eq(0) & df["Credito"].fillna(0).eq(0))].copy()

    # Nota e dia
    df["NotaIDs"] = df[col_hist].apply(extract_note_ids)
    df["NotaID"] = df["NotaIDs"].apply(lambda l: l[0] if l else "SEM_NOTA")
    df["Dia"] = df[col_date].dt.date

    # Resumo mensal
    resumo_mensal = df.agg({"Debito": "sum", "Credito": "sum", "Valor": "sum"}).to_frame(name="Total").T
    resumo_mensal["Fechou"] = (resumo_mensal["Valor"].abs() <= EPS)

    # Totais por dia
    dias = df.groupby("Dia", as_index=False).agg(Debito=("Debito","sum"), Credito=("Credito","sum"))
    dias["Diferenca"] = (dias["Debito"] - dias["Credito"]).astype(float)
    dias["Fechou"] = dias["Diferenca"].abs() <= EPS
    dias_com_diff = dias[~dias["Fechou"]].sort_values("Dia")

    # Por nota no mês (informativo)
    por_nota_mes = df.groupby("NotaID", as_index=False).agg(Debito=("Debito","sum"), Credito=("Credito","sum"))
    por_nota_mes["Diferenca"] = (por_nota_mes["Debito"] - por_nota_mes["Credito"]).astype(float)
    por_nota_mes = por_nota_mes.sort_values(["Diferenca","NotaID"], ascending=[False, True])

    # Por dia + nota
    por_dia_nota = df.groupby(["Dia","NotaID"], as_index=False).agg(Debito=("Debito","sum"), Credito=("Credito","sum"))
    por_dia_nota["Diferenca"] = (por_dia_nota["Debito"] - por_dia_nota["Credito"]).astype(float)

    # Só dias problemáticos
    dias_problematicos = set(dias_com_diff["Dia"])
    diffs_por_nota = (
        por_dia_nota[
            (por_dia_nota["Dia"].isin(dias_problematicos)) &
            (por_dia_nota["Diferenca"].abs() > EPS)
        ].sort_values(["Dia","Diferenca"], ascending=[True, False])
    )

    # Seleciona notas responsáveis por dia
    selected_by_day = pick_responsible_sets(diffs_por_nota, dias_com_diff)

    # Flag “Provavel_Responsavel” (fica no Excel)
    diffs_por_nota["Provavel_Responsavel"] = diffs_por_nota.apply(
        lambda r: r["NotaID"] in selected_by_day.get(r["Dia"], set()), axis=1
    )

    # Lançamentos responsáveis: somente as linhas das notas selecionadas
    chaves_dia_nota = {(d, n) for d, notas in selected_by_day.items() for n in notas}
    responsaveis = df[df.apply(lambda r: (r["Dia"], r["NotaID"]) in chaves_dia_nota, axis=1)].copy()

    # Sem contrapartida
    side_by_note = df.groupby(["Dia","NotaID"]).agg(
        deb=("Debito", lambda s: (s.fillna(0) > 0).sum()),
        cred=("Credito", lambda s: (s.fillna(0) > 0).sum())
    ).reset_index()
    side_by_note["SemContrapartida"] = (side_by_note["deb"].eq(0) | side_by_note["cred"].eq(0))
    responsaveis = responsaveis.merge(side_by_note[["Dia","NotaID","SemContrapartida"]], on=["Dia","NotaID"], how="left")

    # >>> NOVO: valor da diferença da nota no dia (para a aba 'Lancamentos_Responsaveis')
    diffs_key = diffs_por_nota[["Dia", "NotaID", "Diferenca"]].rename(columns={"Diferenca": "DiferencaNotaDia"})
    responsaveis = responsaveis.merge(diffs_key, on=["Dia", "NotaID"], how="left")

    # Organiza colunas
    show_cols = ["Dia", col_date, col_hist, "NotaID", "Debito", "Credito", "Valor", "DiferencaNotaDia", "SemContrapartida"]
    if col_batch and col_batch in df.columns:
        show_cols.insert(1, col_batch)
    responsaveis = responsaveis[show_cols].sort_values(["Dia","NotaID"])

    # Saída Excel
    out_path = xlsx_path.with_name(xlsx_path.stem + "_relatorio.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as xlw:
        resumo_mensal.to_excel(xlw, sheet_name="Resumo_Mensal", index=False)
        dias.to_excel(xlw, sheet_name="Totais_por_Dia", index=False)
        dias_com_diff.to_excel(xlw, sheet_name="Dias_com_Diferenca", index=False)
        por_nota_mes.to_excel(xlw, sheet_name="Notas_Mes", index=False)
        diffs_por_nota.to_excel(xlw, sheet_name="Diferencas_por_Nota", index=False)
        responsaveis.to_excel(xlw, sheet_name="Lancamentos_Responsaveis", index=False)

    # ======= Console =======
    print("\n== RESUMO MENSAL ==")
    print(resumo_mensal.to_string(index=False))
    if not dias_com_diff.empty:
        print("\n== DIAS COM DIFERENÇA ==")
        print(dias_com_diff.to_string(index=False))

        # (Oculto) Diferenças por Nota no console

        # helper BRL
        def _fmt_brl(v):
            try:
                return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return str(v)

        print("\n== NOTAS SELECIONADAS COMO RESPONSÁVEIS ==")
        for d, notas in selected_by_day.items():
            if not notas:
                continue
            itens = []
            for n in sorted(notas):
                val = diffs_por_nota.loc[
                    (diffs_por_nota["Dia"] == d) & (diffs_por_nota["NotaID"] == n),
                    "Diferenca"
                ].sum()
                itens.append(f"{n} (R$ {_fmt_brl(val)})")
            print(f"{d} -> " + ", ".join(itens))
    else:
        print("\nTodos os dias fecharam em R$ 0,00.")
    print(f"\nRelatório salvo em: {out_path}")

# ======= ENTRADA =======
def _pick_file_dialog():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        path = filedialog.askopenfilename(title="Selecione a planilha do mês",
            filetypes=[("Planilhas Excel","*.xlsx *.xlsm *.xls")])
        return Path(path) if path else None
    except Exception:
        return None

def _fallback_latest_xlsx():
    here = Path(__file__).resolve().parent
    files = sorted((Path(p) for p in glob.glob(str(here / "*.xls*"))),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None

def main():
    xlsx = Path(sys.argv[1]) if len(sys.argv) >= 2 else None
    if xlsx is None: xlsx = _pick_file_dialog()
    if xlsx is None: xlsx = _fallback_latest_xlsx()
    if xlsx is None or not xlsx.exists():
        print("Não encontrei a planilha.\n→ Rode: python conta_transitoria.py \"CAMINHO\\arquivo.xlsx\"\n"
              "ou selecione pelo diálogo ao dar duplo-clique no .py.")
        sys.exit(3)
    print(f"[INFO] Processando: {xlsx}")
    process_file(xlsx)

if __name__ == "__main__":
    main()
