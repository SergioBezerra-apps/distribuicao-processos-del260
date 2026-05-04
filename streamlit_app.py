import os
import io
import re
import zipfile
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from xlsxwriter.utility import xl_col_to_name

# =============================================================================
# E-mail (usa st.secrets / env; se ausentes, não envia)
# =============================================================================

def _get_mail_creds():
    user = st.secrets.get("SMTP_USERNAME", os.getenv("SMTP_USERNAME", ""))
    pwd  = st.secrets.get("SMTP_PASSWORD", os.getenv("SMTP_PASSWORD", ""))
    host = st.secrets.get("SMTP_SERVER",   os.getenv("SMTP_SERVER",   "smtp.gmail.com"))
    port = int(st.secrets.get("SMTP_PORT", os.getenv("SMTP_PORT", 465)))
    return host, port, user, pwd

def send_email_with_multiple_attachments(to_emails, subject, body, attachments):
    smtp_server, smtp_port, smtp_username, smtp_password = _get_mail_creds()
    if not smtp_username or not smtp_password:
        st.warning("Credenciais de e-mail ausentes (SMTP_USERNAME/SMTP_PASSWORD). E-mail não enviado.")
        return
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_username
    msg["To"] = ", ".join(to_emails)
    msg.set_content(body)
    for attachment_bytes, filename in attachments:
        msg.add_attachment(
            attachment_bytes,
            maintype="application",
            subtype="octet-stream",
            filename=filename
        )
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10) as server:
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            st.info(f"E-mail enviado para: {to_emails}")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail para {to_emails}: {e}")

def send_email_with_two_attachments(to_email, subject, body, attachment_pre, filename_pre, attachment_res, filename_res):
    smtp_server, smtp_port, smtp_username, smtp_password = _get_mail_creds()
    if not smtp_username or not smtp_password:
        st.warning("Credenciais de e-mail ausentes (SMTP_USERNAME/SMTP_PASSWORD). E-mail não enviado.")
        return
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_username
    msg["To"] = to_email
    msg.set_content(body)
    if attachment_pre is not None:
        msg.add_attachment(attachment_pre, maintype="application", subtype="octet-stream", filename=filename_pre)
    if attachment_res is not None:
        msg.add_attachment(attachment_res, maintype="application", subtype="octet-stream", filename=filename_res)
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10) as server:
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            st.info(f"E-mail enviado para: {to_email}")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail para {to_email}: {e}")

# =============================================================================
# Excel helpers
# =============================================================================

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha")
        workbook = writer.book
        worksheet = writer.sheets["Planilha"]
        if "Grupo Natureza" in df.columns:
            col_index = df.columns.get_loc("Grupo Natureza")
            col_letter = xl_col_to_name(col_index)
            last_row = len(df) + 1
            cell_range = f"{col_letter}2:{col_letter}{last_row}"
            color_list = [
                "#990000", "#006600", "#996600", "#003366",
                "#660066", "#663300", "#003300", "#000066"
            ]
            unique_values = df["Grupo Natureza"].dropna().unique()
            color_mapping = {val: color_list[i % len(color_list)] for i, val in enumerate(unique_values)}
            for value, color in color_mapping.items():
                fmt = workbook.add_format({"font_color": color, "bold": True})
                worksheet.conditional_format(cell_range, {
                    "type": "cell", "criteria": "==", "value": f'"{value}"', "format": fmt
                })
    return output.getvalue()

def create_zip_from_dict(file_dict: dict) -> bytes:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for filename, file_bytes in file_dict.items():
            zip_file.writestr(filename, file_bytes)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# =============================================================================
# Validações de cadastro
# =============================================================================

def _is_valid_email(x: str) -> bool:
    if not isinstance(x, str):
        return False
    x = x.strip()
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", x))

# =============================================================================
# Critério / Prioridade
# =============================================================================

DIAS_5_ANOS = 1825
DIAS_4_ANOS = 1460

def calcula_criterio(row):
    try:
        tempo = float(row["Tempo TCERJ"])
    except Exception:
        tempo = -1
    try:
        dias_orgao = float(row["Dias no Orgão"])
    except Exception:
        dias_orgao = -1

    if tempo >= DIAS_5_ANOS:
        return "01 Mais de cinco anos de autuado"
    elif DIAS_4_ANOS <= tempo < DIAS_5_ANOS:
        return "02 A completar 5 anos de autuado"
    elif dias_orgao >= 150:
        return "03 Mais de 5 meses na 3CAP"
    else:
        return "04 Data da carga"

priority_map = {
    "01 Mais de cinco anos de autuado": 0,
    "02 A completar 5 anos de autuado": 1,
    "03 Mais de 5 meses na 3CAP": 2,
    "04 Data da carga": 3
}

# =============================================================================
# Whitelist — novo modelo tabular
#
# whitelist_pairs[inf] = lista de tuplas (natureza_filtro, orgao_filtro)
#   string vazia = curinga (qualquer valor)
#   lista vazia  = informante sem restrição (aceita tudo)
#
# Exemplo:
#   ANA → [("APOSENTADORIA", "ÓRGÃO A"),
#           ("PENSÃO",        "ÓRGÃO A"),
#           ("APOSENTADORIA", "ÓRGÃO B"),
#           ("PENSÃO",        "ÓRGÃO D")]
#
# Um processo é aceito se QUALQUER linha da lista fizer match.
# =============================================================================

WhitelistPairs = Dict[str, List[Tuple[str, str]]]

def build_whitelist_pairs(df_wl: pd.DataFrame) -> WhitelistPairs:
    """
    Constrói o dicionário whitelist_pairs a partir do data_editor.
    Colunas esperadas: Informante | Grupo Natureza | Orgão Origem
    Valor vazio em Natureza ou Órgão = curinga.
    """
    pairs: WhitelistPairs = {}
    if df_wl is None or df_wl.empty:
        return pairs
    for _, row in df_wl.iterrows():
        inf = str(row.get("Informante", "")).strip().upper()
        if not inf:
            continue
        nat = str(row.get("Grupo Natureza", "")).strip().upper()
        org = str(row.get("Orgão Origem",   "")).strip().upper()
        # "(QUALQUER)" ou vazio → curinga
        if nat in ("(QUALQUER)", "NAN", "NONE", ""):
            nat = ""
        if org in ("(QUALQUER)", "NAN", "NONE", ""):
            org = ""
        pairs.setdefault(inf, []).append((nat, org))
    return pairs

def _accepts(inf: str, orgao: str, natureza: str,
             whitelist_pairs: WhitelistPairs) -> bool:
    """
    Verifica se o informante aceita o processo com base na whitelist tabular.

    Regras:
      - Informante sem entradas na tabela → aceita tudo (sem restrição).
      - Informante com entradas → aceita se ALGUMA linha fizer match.
      - Curinga (string vazia) → qualquer valor casa.
    """
    pairs = whitelist_pairs.get(inf, [])
    if not pairs:
        return True   # sem restrição configurada
    for nat_f, org_f in pairs:
        ok_nat = (not nat_f) or (nat_f == natureza)
        ok_org = (not org_f) or (org_f == orgao)
        if ok_nat and ok_org:
            return True
    return False

# =============================================================================
# Helpers internos
# =============================================================================

def _locked_mask(df: pd.DataFrame) -> pd.Series:
    if "Locked" in df.columns:
        return df["Locked"].astype(bool).reindex(df.index).fillna(False)
    return pd.Series(False, index=df.index)

# =============================================================================
# Regras de roteamento
# =============================================================================

def _apply_routing_rules(df_pool: pd.DataFrame,
                         rules: list,
                         whitelist_pairs: WhitelistPairs,
                         only_locked_map: Optional[Dict[str, bool]] = None):
    if only_locked_map is None:
        only_locked_map = {}

    if df_pool.empty:
        df_empty = df_pool.copy()
        if "Locked" not in df_empty.columns:
            df_empty["Locked"] = False
        return df_empty.head(0), df_empty

    df_pool = df_pool.copy()
    if "Locked" not in df_pool.columns:
        df_pool["Locked"] = False

    assigned_rows, remaining_rows = [], []

    for _, row in df_pool.iterrows():
        natureza = str(row["Grupo Natureza"])
        orgao    = str(row["Orgão Origem"])

        best = None
        for idx, r in enumerate(rules):
            inf      = r["Informante"]
            nat_rule = r["Grupo Natureza"]
            org_rule = r["Orgão Origem"]

            ok_nat = (nat_rule == "(QUALQUER)") or (nat_rule == natureza)
            ok_org = (org_rule == "(QUALQUER)") or (org_rule == orgao)
            if not (ok_nat and ok_org):
                continue

            score = (0 if nat_rule == "(QUALQUER)" else 1) + (0 if org_rule == "(QUALQUER)" else 1)
            cand  = (score, idx, r)
            if (best is None) or (cand[0] > best[0]) or (cand[0] == best[0] and cand[1] < best[1]):
                best = cand

        if best is None:
            remaining_rows.append(row)
        else:
            _, _, rbest = best
            inf_best  = rbest["Informante"]
            exclusiva = bool(rbest.get("Exclusiva?", False))

            if (not exclusiva) and bool(only_locked_map.get(inf_best, False)):
                remaining_rows.append(row)
                continue

            if exclusiva or _accepts(inf_best, orgao, natureza, whitelist_pairs):
                new_row = row.copy()
                new_row["Informante"] = inf_best
                new_row["Locked"]     = exclusiva
                assigned_rows.append(new_row)
            else:
                remaining_rows.append(row)

    df_assigned  = pd.DataFrame(assigned_rows)  if assigned_rows  else pd.DataFrame(columns=df_pool.columns)
    df_remaining = pd.DataFrame(remaining_rows) if remaining_rows else pd.DataFrame(columns=df_pool.columns)
    for d in (df_assigned, df_remaining):
        if "Locked" not in d.columns:
            d["Locked"] = False
    return df_assigned, df_remaining

# =============================================================================
# Redistribuição com reserva por whitelist tabular
# =============================================================================

def _redistribute(df_unassigned: pd.DataFrame,
                  informantes_grupo_a: List[str],
                  informantes_grupo_b: List[str],
                  origens_especiais: List[str],
                  whitelist_pairs: WhitelistPairs,
                  only_locked_map: Dict[str, bool]) -> pd.DataFrame:
    """
    Round-robin por natureza garantindo 100% de destino.

    RESERVA POR WHITELIST TABULAR:
      Se algum informante do pool possui entradas na tabela de whitelist
      que cobrem o processo em questão, o processo é reservado para esse
      subgrupo. Informantes sem whitelist não concorrem enquanto houver
      candidato reservado disponível.

    Cascata de fallback:
      T0: pool reservado (ou completo se não há reservado), not-only-locked
      T1: pool reservado, not-only-locked, ignora whitelist
      T2: pool completo, only-locked liberado, mantém whitelist
      T3: pool completo, tudo liberado (fallback final)
    """
    if df_unassigned.empty:
        return df_unassigned

    df_unassigned = df_unassigned.copy()
    if "Fallback Tier"   not in df_unassigned.columns:
        df_unassigned["Fallback Tier"]   = ""
    if "Fallback Motivo" not in df_unassigned.columns:
        df_unassigned["Fallback Motivo"] = ""

    rr_indices = {gn: 0 for gn in df_unassigned["Grupo Natureza"].unique()}

    def _candidates(pool: List[str], orgao: str, natureza: str,
                    allow_only_locked: bool, ignore_whitelist: bool) -> List[str]:
        cand = []
        for inf in pool:
            if (not allow_only_locked) and bool(only_locked_map.get(inf, False)):
                continue
            if ignore_whitelist or _accepts(inf, orgao, natureza, whitelist_pairs):
                cand.append(inf)
        return cand

    out_rows = []
    for _, row in df_unassigned.iterrows():
        natureza = row["Grupo Natureza"]
        orgao    = row["Orgão Origem"]

        pool = informantes_grupo_a if (origens_especiais and orgao in origens_especiais) else informantes_grupo_b

        # Pool reservado: informantes que têm whitelist configurada
        # E essa whitelist cobre este processo específico.
        pool_reservado = [
            inf for inf in pool
            if whitelist_pairs.get(inf)   # tem ao menos uma linha configurada
            and _accepts(inf, orgao, natureza, whitelist_pairs)
        ]
        pool_primario = pool_reservado if pool_reservado else pool

        candidatos = _candidates(pool_primario, orgao, natureza, allow_only_locked=False, ignore_whitelist=False)
        tier   = "T0"
        motivo = ""

        if not candidatos:
            candidatos = _candidates(pool_primario, orgao, natureza, allow_only_locked=False, ignore_whitelist=True)
            tier   = "T1"
            motivo = "Whitelist ignorada (sem candidato disponível no pool reservado)"

        if not candidatos:
            candidatos = _candidates(pool, orgao, natureza, allow_only_locked=True, ignore_whitelist=False)
            tier   = "T2"
            motivo = "Somente candidatos only-locked disponíveis; manteve whitelist"

        if not candidatos:
            candidatos = _candidates(pool, orgao, natureza, allow_only_locked=True, ignore_whitelist=True)
            tier   = "T3"
            motivo = "Fallback final: only-locked e whitelist ignorados"

        if not candidatos:
            row["Informante"]    = ""
            row["Fallback Tier"] = "SEM_CANDIDATO"
            row["Fallback Motivo"] = "Nenhum informante disponível no pool"
            out_rows.append(row)
            continue

        idx = rr_indices.get(natureza, 0) % len(candidatos)
        row["Informante"] = candidatos[idx]
        rr_indices[natureza] = rr_indices.get(natureza, 0) + 1

        if tier != "T0":
            row["Fallback Tier"]   = tier
            row["Fallback Motivo"] = motivo

        out_rows.append(row)

    return pd.DataFrame(out_rows)

# =============================================================================
# PREVENÇÃO — Top-200
# =============================================================================

def _apply_prevention_top200(res_final: pd.DataFrame,
                             df_prev_map: Optional[pd.DataFrame],
                             whitelist_pairs: WhitelistPairs) -> dict:
    out  = {}
    base = res_final.copy()
    base["Critério"]       = base.apply(calcula_criterio, axis=1)
    base["CustomPriority"] = base["Critério"].apply(lambda x: priority_map.get(x, 4))
    base = base.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])

    prev_map = {}
    if df_prev_map is not None and not df_prev_map.empty:
        prev = df_prev_map.copy()
        prev.columns = [c.strip() for c in prev.columns]
        if {"Processo", "Informante"}.issubset(set(prev.columns)):
            prev["Processo"]   = prev["Processo"].astype(str).str.strip()
            prev["Informante"] = prev["Informante"].astype(str).str.strip().str.upper()
            prev_map = dict(zip(prev["Processo"], prev["Informante"]))

    informantes = base["Informante"].dropna().astype(str).str.strip().str.upper().unique()

    for inf in informantes:
        if not inf:
            continue

        df_inf = base[base["Informante"].astype(str).str.upper() == inf].copy()
        if df_inf.empty:
            continue

        locked_part = df_inf[_locked_mask(df_inf)].copy()
        nonlocked   = df_inf[~_locked_mask(df_inf)].copy()

        locked_part["ForaWhitelist"] = False
        nonlocked["ForaWhitelist"]   = nonlocked.apply(
            lambda r: not _accepts(inf, r["Orgão Origem"], r["Grupo Natureza"], whitelist_pairs),
            axis=1
        )

        df_inf = pd.concat([locked_part, nonlocked], ignore_index=True)

        if prev_map:
            df_inf["preferido"] = df_inf["Processo"].astype(str).map(lambda p: 1 if prev_map.get(p) == inf else 0)
        else:
            df_inf["preferido"] = 0

        preferidos = df_inf[df_inf["preferido"] == 1].sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
        nao_pref   = df_inf[df_inf["preferido"] == 0].sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])

        df_top = pd.concat([preferidos, nao_pref], ignore_index=True).head(200)
        df_top = df_top.drop(columns=["CustomPriority", "preferido"], errors="ignore")
        out[inf] = df_top

    return out

# =============================================================================
# STICK ESTRITO — Top-200
# =============================================================================

def _stick_previous_topN_STRICT(res_df: pd.DataFrame,
                                df_prev_map: pd.DataFrame,
                                whitelist_pairs: WhitelistPairs,
                                only_locked_map: Dict[str, bool],
                                top_n_per_inf: int = 200):
    if res_df.empty or df_prev_map is None or df_prev_map.empty:
        return res_df, pd.DataFrame(columns=["Informante","colados","bloq_only_locked","bloq_whitelist"])

    prev = df_prev_map.copy()
    prev.columns = [c.strip() for c in prev.columns]
    if not {"Processo", "Informante"}.issubset(set(prev.columns)):
        return res_df, pd.DataFrame(columns=["Informante","colados","bloq_only_locked","bloq_whitelist"])

    prev["Processo"]   = prev["Processo"].astype(str).str.strip()
    prev["Informante"] = prev["Informante"].astype(str).str.strip().str.upper()

    df = res_df.copy()
    if "Locked" not in df.columns:
        df["Locked"] = False

    df["Critério"]       = df.apply(calcula_criterio, axis=1)
    df["CustomPriority"] = df["Critério"].apply(lambda x: priority_map.get(x, 4))

    diag_rows = []
    for inf in sorted(df["Informante"].astype(str).str.upper().dropna().unique()):
        if str(inf).strip() == "":
            continue

        prev_set = set(prev[prev["Informante"] == inf]["Processo"].astype(str))
        bloco = df[df["Processo"].astype(str).isin(prev_set)].copy()
        if bloco.empty:
            diag_rows.append({"Informante": inf, "colados": 0, "bloq_only_locked": 0, "bloq_whitelist": 0})
            continue

        bloco["ok_whitelist"] = bloco.apply(
            lambda r: True if bool(r.get("Locked", False))
                      else _accepts(inf, r["Orgão Origem"], r["Grupo Natureza"], whitelist_pairs),
            axis=1
        )

        bloco = bloco[bloco["ok_whitelist"]].copy()
        bloco = bloco.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False]).head(top_n_per_inf)

        if bloco.empty:
            diag_rows.append({"Informante": inf, "colados": 0, "bloq_only_locked": 0, "bloq_whitelist": int(len(prev_set))})
            continue

        sl_only = bool(only_locked_map.get(inf, False))
        bloco["bloq_only_locked"] = bloco.apply(lambda r: bool(sl_only and not r.get("Locked", False)), axis=1)

        colaveis = bloco[~bloco["bloq_only_locked"]].copy()
        if not colaveis.empty:
            df.loc[df["Processo"].astype(str).isin(colaveis["Processo"].astype(str)), "Informante"] = inf

        diag_rows.append({
            "Informante":       inf,
            "colados":          int(len(colaveis)),
            "bloq_only_locked": int(bloco["bloq_only_locked"].sum()),
            "bloq_whitelist":   0
        })

    df = df.drop(columns=["CustomPriority"], errors="ignore")
    diag_df = pd.DataFrame(diag_rows, columns=["Informante","colados","bloq_only_locked","bloq_whitelist"])
    return df, diag_df

# =============================================================================
# Auto-sync pools
# =============================================================================

def _sync_pool_selections(informantes_disponiveis: List[str]) -> None:
    avail_set = set(informantes_disponiveis)

    if "pool_b_sel" not in st.session_state:
        st.session_state["pool_b_sel"] = informantes_disponiveis[:]
    if "pool_a_sel" not in st.session_state:
        st.session_state["pool_a_sel"] = []
    if "pool_unico_sel" not in st.session_state:
        st.session_state["pool_unico_sel"] = informantes_disponiveis[:]

    roster_key = "roster_sig_v2"
    roster_sig = "|".join(sorted(informantes_disponiveis))
    prev_sig   = st.session_state.get(roster_key)

    if prev_sig != roster_sig:
        a_set = set(st.session_state.get("pool_a_sel", []))
        b_set = set(st.session_state.get("pool_b_sel", []))
        p_set = set(st.session_state.get("pool_unico_sel", []))

        a_set &= avail_set
        b_set &= avail_set
        p_set &= avail_set

        novos  = avail_set.difference(a_set.union(b_set).union(p_set))
        b_set |= novos
        p_set |= novos

        st.session_state["pool_a_sel"]     = sorted(a_set)
        st.session_state["pool_b_sel"]     = sorted(b_set)
        st.session_state["pool_unico_sel"] = sorted(p_set)
        st.session_state[roster_key]       = roster_sig

# =============================================================================
# UI
# =============================================================================

st.title("Distribuição de Processos da Del. 260")

if "numero" not in st.session_state:
    st.session_state.numero = "184"

uploaded_files = st.file_uploader(
    "Carregar: processos.xlsx, processosmanter.xlsx, observacoes.xlsx, disponibilidade_equipe.xlsx",
    type=["xlsx"],
    accept_multiple_files=True
)
prev_file = st.file_uploader(
    "Opcional: carregar a PLANILHA GERAL PRINCIPAL da semana anterior (para prevenção Top-200). Deve conter colunas 'Processo' e 'Informante'.",
    type=["xlsx"],
    accept_multiple_files=False
)

files_dict = {}
for file in uploaded_files or []:
    fname = file.name.lower().strip()
    if fname == "processos.xlsx":
        files_dict["processos"] = file
    elif fname == "processosmanter.xlsx":
        files_dict["processosmanter"] = file
    elif fname in ["observacoes.xlsx", "obervacoes.xlsx"]:
        files_dict["observacoes"] = file
    elif fname == "disponibilidade_equipe.xlsx":
        files_dict["disponibilidade"] = file

numero = st.text_input("Numeração desta planilha de distribuição:", value=st.session_state.numero)

modo = st.radio("Selecione o modo:", options=["Teste", "Produção"], horizontal=True)
test_mode = (modo == "Teste")
if test_mode:
    st.info("Modo Teste: não envia e-mails; planilhas para download.")
else:
    st.info("Modo Produção: e-mails serão enviados conforme seleção.")
st.markdown(f"**Modo selecionado:** {modo}")

modo_envio = None
if not test_mode:
    modo_envio = st.radio(
        "Modo de envio:",
        options=["Produção - Gestores e Informantes", "Produção - Apenas Gestores"],
        horizontal=False
    )

managers_emails = st.text_input(
    "E-mails dos gestores/revisores (separados por vírgula):",
    value="annapc@tcerj.tc.br, fabiovf@tcerj.tc.br, sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com"
)

# -----------------------------------------------------------------------------
# Somente após carregar os 4 arquivos principais
# -----------------------------------------------------------------------------

if all(k in files_dict for k in ["processos", "processosmanter", "observacoes", "disponibilidade"]):

    df_proc = pd.read_excel(files_dict["processos"])
    df_proc.columns = df_proc.columns.str.strip()
    if "Grupo Natureza" in df_proc.columns:
        df_proc["Grupo Natureza"] = df_proc["Grupo Natureza"].astype(str).str.strip().str.upper()
    if "Orgão Origem" in df_proc.columns:
        df_proc["Orgão Origem"] = df_proc["Orgão Origem"].astype(str).str.strip().str.upper()

    df_manter = pd.read_excel(files_dict["processosmanter"])
    df_manter.columns = df_manter.columns.str.strip()
    processos_validos = df_manter["Processo"].dropna().astype(str).unique()
    df_proc["Processo"] = df_proc["Processo"].astype(str)
    df_proc = df_proc[df_proc["Processo"].isin(processos_validos)]
    if "Tipo Processo" in df_proc.columns:
        df_proc = df_proc[df_proc["Tipo Processo"].astype(str).str.upper().str.strip() == "PRINCIPAL"]

    grupo_natureza_options = sorted(df_proc["Grupo Natureza"].dropna().unique())
    orgaos_origem_options  = sorted(df_proc["Orgão Origem"].dropna().unique())

    # =========================
    # CADASTRO EDITÁVEL
    # =========================
    st.markdown("## Cadastro de Informantes (válido para a execução)")
    st.caption("Você pode incluir/editar informantes e e-mails na hora. A execução usa este cadastro.")

    df_disp_base = pd.read_excel(files_dict["disponibilidade"])
    df_disp_base.columns = df_disp_base.columns.str.strip()

    for col, default in [("informantes", ""), ("email", ""), ("disponibilidade", "sim")]:
        if col not in df_disp_base.columns:
            df_disp_base[col] = default

    df_disp_base["informantes"]    = df_disp_base["informantes"].astype(str).str.strip().str.upper()
    df_disp_base["email"]          = df_disp_base["email"].astype(str).str.strip()
    df_disp_base["disponibilidade"]= df_disp_base["disponibilidade"].astype(str).str.strip().str.lower()

    if "df_disp_editor" not in st.session_state:
        st.session_state["df_disp_editor"] = df_disp_base.copy()

    df_disp_editor = st.data_editor(
        st.session_state["df_disp_editor"],
        num_rows="dynamic",
        use_container_width=True,
        key="disp_editor",
        column_config={
            "informantes":     st.column_config.TextColumn("Informante (NOME)", required=True),
            "email":           st.column_config.TextColumn("E-mail", required=False),
            "disponibilidade": st.column_config.SelectboxColumn("Disponível?", options=["sim", "nao"], required=True),
        }
    )

    df_disp_editor = df_disp_editor.copy()
    df_disp_editor["informantes"]    = df_disp_editor["informantes"].astype(str).str.strip().str.upper()
    df_disp_editor["email"]          = df_disp_editor["email"].astype(str).str.strip()
    df_disp_editor["disponibilidade"]= df_disp_editor["disponibilidade"].astype(str).str.strip().str.lower()
    df_disp_editor = df_disp_editor[df_disp_editor["informantes"].astype(str).str.strip() != ""].copy()

    invalid_emails = df_disp_editor[
        (df_disp_editor["email"].astype(str).str.strip() != "") &
        (~df_disp_editor["email"].apply(_is_valid_email))
    ]
    if not invalid_emails.empty:
        st.warning("Há e-mails com formato inválido no cadastro (o envio pode falhar).")

    st.session_state["df_disp_editor"] = df_disp_editor

    st.markdown("### Exportar cadastro atualizado")
    st.caption("Baixe um XLSX com o cadastro acima para reaproveitar na próxima semana.")
    cadastro_export   = df_disp_editor[["informantes", "email", "disponibilidade"]].copy()
    cadastro_filename = f"disponibilidade_equipe_atualizada_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        "Baixar disponibilidade atualizada (XLSX)",
        data=to_excel_bytes(cadastro_export),
        file_name=cadastro_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =========================
    # Disponíveis + auto-sync
    # =========================
    df_disp_disponiveis  = df_disp_editor[df_disp_editor["disponibilidade"] == "sim"].copy()
    informantes_disponiveis = sorted(df_disp_disponiveis["informantes"].dropna().unique().tolist())

    if not informantes_disponiveis:
        st.error("Não há informantes com disponibilidade = 'sim'. Ajuste o cadastro acima para prosseguir.")
        st.stop()

    _sync_pool_selections(informantes_disponiveis)

    # =========================
    # Modelo de distribuição
    # =========================
    st.markdown("## Modelo de Distribuição (PRINCIPAIS)")
    modelo_dist = st.radio(
        "Selecione o modelo:",
        options=["A/B por órgão (Opção 1: A adicional + B base)", "Pool único (todos para todos)"],
        horizontal=True
    )
    st.session_state["modelo_dist"] = modelo_dist

    st.markdown("## Seleção do Pool de Informantes (PRINCIPAIS)")
    st.caption("Opção 1: B é o pool geral (base). A é um pool adicional usado apenas para órgãos especiais.")

    if modelo_dist.startswith("A/B"):
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Selecionar TODOS no pool B (geral)"):
                st.session_state["pool_b_sel"] = informantes_disponiveis[:]
        with col2:
            if st.button("Limpar pool A (órgãos especiais)"):
                st.session_state["pool_a_sel"] = []

        pool_b_sel = st.multiselect(
            "Pool B (geral/base) — recebe processos de órgãos NÃO especiais",
            options=informantes_disponiveis,
            default=[x for x in st.session_state.get("pool_b_sel", []) if x in informantes_disponiveis] or informantes_disponiveis[:],
            key="pool_b_multiselect"
        )
        pool_a_sel = st.multiselect(
            "Pool A (adicional) — usado APENAS para órgãos especiais (pode sobrepor B)",
            options=informantes_disponiveis,
            default=[x for x in st.session_state.get("pool_a_sel", []) if x in informantes_disponiveis],
            key="pool_a_multiselect"
        )
        st.session_state["pool_b_sel"] = pool_b_sel
        st.session_state["pool_a_sel"] = pool_a_sel

        if not pool_b_sel:
            st.warning("Pool B (geral) está vazio.")

        st.markdown("### Órgãos especiais (disparam uso do pool A)")
        origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]
        st.write(origens_especiais)

    else:
        if st.button("Selecionar TODOS no Pool Único"):
            st.session_state["pool_unico_sel"] = informantes_disponiveis[:]

        pool_unico_sel = st.multiselect(
            "Pool Único (todos para todos)",
            options=informantes_disponiveis,
            default=[x for x in st.session_state.get("pool_unico_sel", []) if x in informantes_disponiveis] or informantes_disponiveis[:],
            key="pool_unico_multiselect"
        )
        st.session_state["pool_unico_sel"] = pool_unico_sel

        if not pool_unico_sel:
            st.warning("Pool único vazio.")

    # =========================
    # WHITELIST TABULAR (novo)
    # =========================
    st.markdown("## Filtros (whitelist) por informante")
    st.caption(
        "Cada linha define uma combinação aceita por um informante. "
        "Deixe Natureza ou Órgão em branco para funcionar como curinga (qualquer valor). "
        "Informante sem nenhuma linha = aceita tudo. "
        "Um processo é aceito se qualquer linha fizer match."
    )

    natureza_opts_wl = [""] + grupo_natureza_options   # "" = qualquer
    orgao_opts_wl    = [""] + orgaos_origem_options

    wl_cols = ["Informante", "Grupo Natureza", "Orgão Origem"]

    if "whitelist_tabela" not in st.session_state:
        st.session_state["whitelist_tabela"] = pd.DataFrame(columns=wl_cols)

    # Garante que colunas existam mesmo após edições parciais
    wl_df_state = st.session_state["whitelist_tabela"].copy()
    for c in wl_cols:
        if c not in wl_df_state.columns:
            wl_df_state[c] = ""

    whitelist_editor = st.data_editor(
        wl_df_state[wl_cols],
        num_rows="dynamic",
        use_container_width=True,
        key="whitelist_editor",
        column_config={
            "Informante":    st.column_config.SelectboxColumn(
                                 "Informante", options=informantes_disponiveis, required=True),
            "Grupo Natureza": st.column_config.SelectboxColumn(
                                 "Natureza (vazio = qualquer)", options=natureza_opts_wl, required=False),
            "Orgão Origem":   st.column_config.SelectboxColumn(
                                 "Órgão (vazio = qualquer)",   options=orgao_opts_wl,    required=False),
        }
    )
    st.session_state["whitelist_tabela"] = whitelist_editor

    # Preview por informante
    if not whitelist_editor.empty:
        with st.expander("Ver resumo da whitelist por informante"):
            for inf in informantes_disponiveis:
                linhas = whitelist_editor[
                    whitelist_editor["Informante"].astype(str).str.upper() == inf
                ]
                if linhas.empty:
                    st.write(f"**{inf}** — sem restrição (aceita tudo)")
                else:
                    st.write(f"**{inf}** — {len(linhas)} combinação(ões) configurada(s):")
                    st.dataframe(linhas[["Grupo Natureza", "Orgão Origem"]].reset_index(drop=True),
                                 use_container_width=True, hide_index=True)

    # =========================
    # Preferências: only-locked
    # =========================
    st.markdown("## Preferências por informante")
    st.caption("Marque para que o informante receba apenas itens de regras EXCLUSIVAS (Locked).")
    only_locked_map: Dict[str, bool] = st.session_state.get("only_locked_map", {})
    for inf in informantes_disponiveis:
        key = f"only_locked_{inf.replace(' ', '_')}"
        val = st.checkbox(
            f"{inf}: receber apenas itens exclusivos (Locked)?",
            value=only_locked_map.get(inf, False),
            key=key
        )
        only_locked_map[inf] = val
    st.session_state["only_locked_map"] = only_locked_map

    # =========================
    # Regras de roteamento
    # =========================
    st.markdown("## Regras de roteamento por (Natureza, Órgão) → Informante")
    st.caption("Use '(QUALQUER)' como curinga. Marque 'Exclusiva?' para reservar o par ao informante indicado.")
    natureza_opts = ["(QUALQUER)"] + grupo_natureza_options
    orgao_opts    = ["(QUALQUER)"] + orgaos_origem_options
    inf_opts      = list(informantes_disponiveis)

    if "rules_state_v2" not in st.session_state:
        st.session_state["rules_state_v2"] = pd.DataFrame(columns=["Informante", "Grupo Natureza", "Orgão Origem", "Exclusiva?"])

    rules_state = st.data_editor(
        st.session_state["rules_state_v2"],
        num_rows="dynamic",
        use_container_width=True,
        key="rules_editor_v2",
        column_config={
            "Informante":    st.column_config.SelectboxColumn("Informante",    options=inf_opts,      required=True),
            "Grupo Natureza":st.column_config.SelectboxColumn("Grupo Natureza",options=natureza_opts, required=True),
            "Orgão Origem":  st.column_config.SelectboxColumn("Orgão Origem",  options=orgao_opts,    required=True),
            "Exclusiva?":    st.column_config.CheckboxColumn("Exclusiva?")
        }
    )
    st.session_state["rules_state_v2"] = rules_state

    # =========================
    # Botão principal
    # =========================
    if st.button("Executar Distribuição"):

        def run_distribution(processos_file, processosmanter_file, obs_file, numero,
                             whitelist_pairs: WhitelistPairs,
                             rules_df=None, prev_file=None):

            # 1) Leitura base
            df = pd.read_excel(processos_file)
            df.columns = df.columns.str.strip()
            for col in ["Grupo Natureza", "Orgão Origem"]:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.upper()
            df["Processo"] = df["Processo"].astype(str)

            df_manter_local = pd.read_excel(processosmanter_file)
            df_manter_local.columns = df_manter_local.columns.str.strip()
            processos_validos_local = df_manter_local["Processo"].dropna().astype(str).unique()
            df = df[df["Processo"].isin(processos_validos_local)]
            if "Tipo Processo" in df.columns:
                df = df[df["Tipo Processo"].astype(str).str.upper().str.strip() == "PRINCIPAL"]

            required_cols = [
                "Processo", "Grupo Natureza", "Orgão Origem", "Dias no Orgão",
                "Tempo TCERJ", "Data Última Carga", "Descrição Informação", "Funcionário Informação"
            ]
            df = df[required_cols]
            df["Descrição Informação"]  = df["Descrição Informação"].astype(str).str.strip().str.lower()
            df["Funcionário Informação"]= df["Funcionário Informação"].astype(str).str.strip().str.upper()

            df_obs = pd.read_excel(obs_file)
            df_obs.columns = df_obs.columns.str.strip()
            df = pd.merge(df, df_obs[["Processo", "Obs", "Data Obs"]], on="Processo", how="left")
            if "Obs" in df.columns:
                mask_suspensa = df["Obs"].astype(str).str.lower().str.contains("análise suspensa")
                df = df[~mask_suspensa].copy()

            df["Data Última Carga"] = pd.to_datetime(df["Data Última Carga"], errors="coerce")
            df["Data Obs"]          = pd.to_datetime(df["Data Obs"],          errors="coerce")

            def update_obs(row):
                if pd.notna(row["Data Obs"]) and pd.notna(row["Data Última Carga"]) and row["Data Obs"] > row["Data Última Carga"]:
                    return pd.Series([row["Obs"], row["Data Obs"]])
                return pd.Series(["", ""])

            df[["Obs", "Data Obs"]] = df.apply(update_obs, axis=1)
            df = df.drop(columns=["Data Última Carga"])

            # 2) Pré-atribuídos e principais
            mask_pre = df["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df["Funcionário Informação"] != "")
            pre_df = df[mask_pre].copy()
            pre_df["Informante"] = pre_df["Funcionário Informação"]
            res_df = df[~mask_pre].copy()

            if "Locked" not in pre_df.columns:
                pre_df["Locked"] = False

            # 3) Disponibilidade
            df_disp_local = st.session_state["df_disp_editor"].copy()
            df_disp_local["informantes"]    = df_disp_local["informantes"].astype(str).str.strip().str.upper()
            df_disp_local["email"]          = df_disp_local["email"].astype(str).str.strip()
            df_disp_local["disponibilidade"]= df_disp_local["disponibilidade"].astype(str).str.strip().str.lower()

            df_disp_ok = df_disp_local[df_disp_local["disponibilidade"] == "sim"].copy()
            available  = sorted(df_disp_ok["informantes"].dropna().unique().tolist())
            informantes_emails = dict(zip(df_disp_local["informantes"], df_disp_local["email"]))

            if not available:
                raise ValueError("Nenhum informante disponível.")

            modelo = st.session_state.get("modelo_dist", "A/B por órgão (Opção 1: A adicional + B base)")
            only_locked_map_local: Dict[str, bool] = st.session_state.get("only_locked_map", {})

            # 4) Pools
            if modelo.startswith("A/B"):
                pool_b = [x for x in st.session_state.get("pool_b_sel", []) if x in available]
                pool_a = [x for x in st.session_state.get("pool_a_sel", []) if x in available]
                if not pool_b:
                    pool_b = available[:]
                informantes_grupo_b  = [s.upper() for s in pool_b]
                informantes_grupo_a  = [s.upper() for s in pool_a]
                origens_especiais    = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]
                if not informantes_grupo_a:
                    informantes_grupo_a = informantes_grupo_b[:]
            else:
                pool = [x for x in st.session_state.get("pool_unico_sel", []) if x in available]
                if not pool:
                    pool = available[:]
                informantes_grupo_a = []
                informantes_grupo_b = [s.upper() for s in pool]
                origens_especiais   = []

            # 5) Regras
            rules_list = []
            if rules_df is not None and not rules_df.empty:
                tmp = rules_df.copy()
                tmp["Informante"]      = tmp["Informante"].map(lambda v: str(v).strip().upper() if pd.notna(v) else "")
                tmp["Grupo Natureza"]  = tmp["Grupo Natureza"].map(lambda v: "(QUALQUER)" if str(v).strip() == "" else str(v).strip().upper())
                tmp["Orgão Origem"]    = tmp["Orgão Origem"].map(lambda v: "(QUALQUER)" if str(v).strip() == "" else str(v).strip().upper())
                tmp = tmp[tmp["Informante"] != ""]
                for _, r in tmp.iterrows():
                    rules_list.append({
                        "Informante":    r["Informante"],
                        "Grupo Natureza":r["Grupo Natureza"],
                        "Orgão Origem":  r["Orgão Origem"],
                        "Exclusiva?":    bool(r.get("Exclusiva?", False))
                    })

            # 6) Distribuição principal
            df_pool = res_df.copy().sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            df_pool["Informante"] = ""
            if "Locked" not in df_pool.columns:
                df_pool["Locked"] = False

            assigned_by_rules, rem_after_rules = _apply_routing_rules(
                df_pool, rules_list, whitelist_pairs,
                only_locked_map=only_locked_map_local
            )

            distributed_df = _redistribute(
                rem_after_rules,
                informantes_grupo_a, informantes_grupo_b, origens_especiais,
                whitelist_pairs,
                only_locked_map=only_locked_map_local
            )

            res_final = pd.concat([assigned_by_rules, distributed_df], ignore_index=True)
            res_final["Informante"] = res_final["Informante"].astype(str).str.strip().str.upper()

            if (res_final["Informante"].astype(str).str.strip() == "").any():
                st.warning("Há processos sem informante (SEM_CANDIDATO). Verifique pools e disponibilidade.")

            res_final["Critério"]       = res_final.apply(calcula_criterio, axis=1)
            res_final["CustomPriority"] = res_final["Critério"].apply(lambda x: priority_map.get(x, 4))
            res_final = res_final.sort_values(
                by=["Informante", "CustomPriority", "Dias no Orgão"],
                ascending=[True, True, False]
            ).reset_index(drop=True)
            res_final = res_final.drop(columns=["CustomPriority"], errors="ignore")

            # 7) Prevenção (Stick estrito)
            df_prev_map = None
            if prev_file is not None:
                try:
                    df_prev_raw = pd.read_excel(prev_file)
                    df_prev_raw.columns = [c.strip() for c in df_prev_raw.columns]
                    if {"Processo", "Informante"}.issubset(set(df_prev_raw.columns)):
                        df_prev_map = df_prev_raw[["Processo", "Informante"]].copy()
                    else:
                        st.warning("Planilha anterior sem colunas 'Processo' e 'Informante'. Prevenção não aplicada.")
                except Exception as e:
                    st.warning(f"Falha ao ler planilha anterior: {e}")

            if "Locked" not in res_final.columns:
                res_final["Locked"] = False

            diag_df = None
            if df_prev_map is not None:
                res_final, diag_df = _stick_previous_topN_STRICT(
                    res_final, df_prev_map,
                    whitelist_pairs,
                    only_locked_map=only_locked_map_local,
                    top_n_per_inf=200
                )
                res_final["Critério"]       = res_final.apply(calcula_criterio, axis=1)
                res_final["CustomPriority"] = res_final["Critério"].apply(lambda x: priority_map.get(x, 4))
                res_final = res_final.sort_values(
                    by=["Informante", "CustomPriority", "Dias no Orgão"],
                    ascending=[True, True, False]
                ).reset_index(drop=True)
                res_final = res_final.drop(columns=["CustomPriority"], errors="ignore")

            # 8) Planilhas gerais
            pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            pre_geral_bytes    = to_excel_bytes(pre_df)
            res_geral_filename = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            res_geral_bytes    = to_excel_bytes(
                res_final.drop(columns=["Descrição Informação", "Funcionário Informação"], errors="ignore")
            )

            # 9) Planilhas individuais — Pré
            def build_pre_individuals():
                pre_individual_files = {}
                pre_df_local = pre_df.copy()
                pre_df_local["Informante"]    = pre_df_local["Informante"].astype(str).str.strip().str.upper()
                pre_df_local["Critério"]      = pre_df_local.apply(calcula_criterio, axis=1)
                pre_df_local["CustomPriority"]= pre_df_local["Critério"].apply(lambda x: priority_map.get(x, 4))

                for inf in pre_df_local["Informante"].dropna().astype(str).str.upper().unique():
                    if not inf:
                        continue
                    df_inf = pre_df_local[pre_df_local["Informante"].astype(str).str.upper() == inf].copy()
                    if df_inf.empty:
                        continue

                    locked_part = df_inf[_locked_mask(df_inf)].copy()
                    nonlocked   = df_inf[~_locked_mask(df_inf)].copy()

                    locked_part["ForaWhitelist"] = False
                    nonlocked["ForaWhitelist"]   = nonlocked.apply(
                        lambda r: not _accepts(inf, r["Orgão Origem"], r["Grupo Natureza"], whitelist_pairs),
                        axis=1
                    )

                    df_inf2 = pd.concat([locked_part, nonlocked], ignore_index=True)
                    df_inf2 = df_inf2.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
                    df_inf2 = df_inf2.drop(columns=["CustomPriority"], errors="ignore")
                    df_inf2 = df_inf2.head(200)
                    pre_individual_files[inf] = to_excel_bytes(df_inf2)

                return pre_individual_files

            pre_individual_files = build_pre_individuals()

            # 10) Planilhas individuais — Principal (Top-200)
            res_individual_files = {}
            top200_dict = _apply_prevention_top200(res_final, df_prev_map, whitelist_pairs)
            for inf, df_inf in top200_dict.items():
                res_individual_files[inf] = to_excel_bytes(df_inf)

            return (pre_geral_filename, pre_geral_bytes,
                    res_geral_filename, res_geral_bytes,
                    pre_individual_files, res_individual_files,
                    informantes_emails, diag_df)

        # ---- Constrói whitelist_pairs a partir do editor ----
        whitelist_pairs_exec = build_whitelist_pairs(st.session_state.get("whitelist_tabela"))

        try:
            (pre_geral_filename, pre_geral_bytes,
             res_geral_filename, res_geral_bytes,
             pre_individual_files, res_individual_files,
             informantes_emails, diag_df) = run_distribution(
                files_dict["processos"], files_dict["processosmanter"],
                files_dict["observacoes"], numero,
                whitelist_pairs_exec,
                rules_df=st.session_state.get("rules_state_v2"),
                prev_file=prev_file
            )
        except Exception as e:
            st.error(f"Falha na distribuição: {e}")
            st.stop()

        st.success("Distribuição executada com sucesso!")

        if isinstance(diag_df, pd.DataFrame) and not diag_df.empty:
            st.markdown("### Diagnóstico da Prevenção (Stick Estrito Top-200)")
            st.dataframe(diag_df.sort_values(by=["Informante"]).reset_index(drop=True), use_container_width=True)

        st.download_button("Baixar Planilha Geral PRE-ATRIBUÍDA",
                           data=pre_geral_bytes, file_name=pre_geral_filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Baixar Planilha Geral PRINCIPAL",
                           data=res_geral_bytes, file_name=res_geral_filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### Planilhas Individuais — Pré-Atribuídos")
        for inf, b in pre_individual_files.items():
            filename_inf = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(f"Baixar para {inf} (Pré-Atribuído)", data=b, file_name=filename_inf,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### Planilhas Individuais — Principal (com prevenção Top-200)")
        for inf, b in res_individual_files.items():
            filename_inf = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(f"Baixar para {inf} (Principal)", data=b, file_name=filename_inf,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if not test_mode:
            managers_list = [e.strip() for e in managers_emails.split(",") if e.strip()]
            if managers_list:
                all_individual_files = {}
                for inf, b in pre_individual_files.items():
                    all_individual_files[f"{inf.replace(' ','_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"] = b
                for inf, b in res_individual_files.items():
                    all_individual_files[f"{inf.replace(' ','_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"] = b

                zip_individual_bytes = create_zip_from_dict(all_individual_files)
                zip_filename = f"{numero}_planilhas_individuais_{datetime.now().strftime('%Y%m%d')}.zip"

                attachments = [
                    (pre_geral_bytes, pre_geral_filename),
                    (res_geral_bytes, res_geral_filename),
                    (zip_individual_bytes, zip_filename)
                ]
                send_email_with_multiple_attachments(
                    managers_list,
                    "Planilhas Gerais e Individuais de Processos",
                    "Prezado(a) Gestor(a),\n\nSeguem anexas as planilhas:\n- Geral de Processos Pré-Atribuídos\n- Geral de Processos Principais\n- ZIP com todas as planilhas individuais\n\nAtenciosamente,\nGestão da 3ª CAP",
                    attachments
                )

            if modo_envio == "Produção - Gestores e Informantes":
                for inf in set(list(pre_individual_files.keys()) + list(res_individual_files.keys())):
                    email_destino = informantes_emails.get(inf.upper(), "")
                    if email_destino:
                        attachment_pre = pre_individual_files.get(inf)
                        attachment_res = res_individual_files.get(inf)
                        filename_pre = f"{inf.replace(' ','_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx" if attachment_pre else None
                        filename_res = f"{inf.replace(' ','_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"    if attachment_res else None
                        send_email_with_two_attachments(
                            email_destino,
                            f"Distribuição de Processos - {inf}",
                            "Prezado(a) Informante,\n\nSeguem anexas as planilhas referentes à distribuição de processos:\n\n• Pré-Atribuídos: vinculados a você no sistema (andamento/conclusão).\n• Principais: novos processos distribuídos.\n\nAtenciosamente,\nGestão da 3ª CAP",
                            attachment_pre, filename_pre,
                            attachment_res, filename_res
                        )

        st.session_state.numero = numero

else:
    st.info("Carregue os quatro arquivos exigidos para habilitar a execução.")
