import os
import io
import re
import zipfile
from datetime import datetime
from typing import Dict, List, Optional

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

DIAS_5_ANOS = 1825   # >= 5 anos
DIAS_4_ANOS = 1460   # >= 4 anos

def calcula_criterio(row):
    """
    Ordem:
      01 >= 5 anos de autuado
      02 Entre 4 e <5 anos de autuado
      03 >= 150 dias no órgão
      04 Demais (data da carga)
    """
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
# Core helpers
# =============================================================================

def _accepts(inf: str, orgao: str, natureza: str,
             filtros_grupo_natureza: Dict[str, List[str]],
             filtros_orgao_origem: Dict[str, List[str]]) -> bool:
    """Whitelist: vazio = aceita tudo. Se marcado, aceita somente o que foi marcado (E para natureza/órgão)."""
    grupos_ok = filtros_grupo_natureza.get(inf, [])
    orgaos_ok = filtros_orgao_origem.get(inf, [])
    if grupos_ok and natureza not in grupos_ok:
        return False
    if orgaos_ok and orgao not in orgaos_ok:
        return False
    return True

def _locked_mask(df: pd.DataFrame) -> pd.Series:
    if "Locked" in df.columns:
        return df["Locked"].astype(bool).reindex(df.index).fillna(False)
    return pd.Series(False, index=df.index)

def _apply_routing_rules(df_pool: pd.DataFrame,
                         rules: list,
                         filtros_grupo_natureza: Dict[str, List[str]],
                         filtros_orgao_origem: Dict[str, List[str]],
                         only_locked_map: Optional[Dict[str, bool]] = None):
    """
    Regras (Natureza, Órgão) → Informante, com 'Exclusiva?' opcional.
    Escolhe a mais específica; empate: a primeira. Exclusiva? => Locked=True.

    only_locked_map:
      - Se destino exige "somente Locked", regras NÃO-exclusivas não atribuem para ele.
      - Regras exclusivas (Locked) sempre prevalecem.
    """
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

        best = None   # (score, idx, rule_dict)
        for idx, r in enumerate(rules):
            inf       = r["Informante"]
            nat_rule  = r["Grupo Natureza"]
            org_rule  = r["Orgão Origem"]

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
            inf_best    = rbest["Informante"]
            exclusiva   = bool(rbest.get("Exclusiva?", False))

            # destino quer "somente Locked": regra NÃO-exclusiva não atribui
            if (not exclusiva) and bool(only_locked_map.get(inf_best, False)):
                remaining_rows.append(row)
                continue

            if exclusiva or _accepts(inf_best, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
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

def _redistribute(df_unassigned: pd.DataFrame,
                  informantes_grupo_a: List[str],
                  informantes_grupo_b: List[str],
                  origens_especiais: List[str],
                  filtros_grupo_natureza: Dict[str, List[str]],
                  filtros_orgao_origem: Dict[str, List[str]],
                  only_locked_map: Dict[str, bool]) -> pd.DataFrame:
    """
    Round-robin por natureza, garantindo DESTINO para 100% dos itens
    (desde que exista ao menos 1 informante no pool aplicável).

    Cascata:
      T0: não-only-locked + respeita whitelist
      T1: não-only-locked + ignora whitelist
      T2: only-locked + respeita whitelist
      T3: only-locked + ignora whitelist (final)

    Se cair em T1/T2/T3: marca 'Fallback Tier' e 'Fallback Motivo'.
    """
    if df_unassigned.empty:
        return df_unassigned

    df_unassigned = df_unassigned.copy()
    if "Fallback Tier" not in df_unassigned.columns:
        df_unassigned["Fallback Tier"] = ""
    if "Fallback Motivo" not in df_unassigned.columns:
        df_unassigned["Fallback Motivo"] = ""

    rr_indices = {gn: 0 for gn in df_unassigned["Grupo Natureza"].unique()}

    def _candidates(pool: List[str], orgao: str, natureza: str,
                    allow_only_locked: bool, ignore_whitelist: bool) -> List[str]:
        cand = []
        for inf in pool:
            if (not allow_only_locked) and bool(only_locked_map.get(inf, False)):
                continue
            if ignore_whitelist or _accepts(inf, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
                cand.append(inf)
        return cand

    out_rows = []
    for _, row in df_unassigned.iterrows():
        natureza = row["Grupo Natureza"]
        orgao = row["Orgão Origem"]

        pool = informantes_grupo_a if (origens_especiais and orgao in origens_especiais) else informantes_grupo_b

        candidatos = _candidates(pool, orgao, natureza, allow_only_locked=False, ignore_whitelist=False)
        tier = "T0"
        motivo = ""

        if not candidatos:
            candidatos = _candidates(pool, orgao, natureza, allow_only_locked=False, ignore_whitelist=True)
            tier = "T1"
            motivo = "Sem candidato na whitelist; whitelist ignorada"

        if not candidatos:
            candidatos = _candidates(pool, orgao, natureza, allow_only_locked=True, ignore_whitelist=False)
            tier = "T2"
            motivo = "Somente candidates only-locked disponíveis; manteve whitelist"

        if not candidatos:
            candidatos = _candidates(pool, orgao, natureza, allow_only_locked=True, ignore_whitelist=True)
            tier = "T3"
            motivo = "Fallback final: only-locked e whitelist ignorados"

        if not candidatos:
            row["Informante"] = ""
            row["Fallback Tier"] = "SEM_CANDIDATO"
            row["Fallback Motivo"] = "Nenhum informante disponível no pool"
            out_rows.append(row)
            continue

        idx = rr_indices.get(natureza, 0) % len(candidatos)
        row["Informante"] = candidatos[idx]
        rr_indices[natureza] = rr_indices.get(natureza, 0) + 1

        if tier != "T0":
            row["Fallback Tier"] = tier
            row["Fallback Motivo"] = motivo

        out_rows.append(row)

    return pd.DataFrame(out_rows)

# =============================================================================
# PREVENÇÃO — Top-200 (não "some" processos: marca ForaWhitelist)
# =============================================================================

def _apply_prevention_top200(res_final: pd.DataFrame, df_prev_map: Optional[pd.DataFrame],
                             filtros_grupo_natureza: Dict[str, List[str]],
                             filtros_orgao_origem: Dict[str, List[str]]) -> dict:
    out = {}
    base = res_final.copy()
    base["Critério"] = base.apply(calcula_criterio, axis=1)
    base["CustomPriority"] = base["Critério"].apply(lambda x: priority_map.get(x, 4))
    base = base.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])

    prev_map = {}
    if df_prev_map is not None and not df_prev_map.empty:
        prev = df_prev_map.copy()
        prev.columns = [c.strip() for c in prev.columns]
        if {"Processo", "Informante"}.issubset(set(prev.columns)):
            prev["Processo"] = prev["Processo"].astype(str).str.strip()
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
        nonlocked["ForaWhitelist"] = nonlocked.apply(
            lambda r: not _accepts(inf, r["Orgão Origem"], r["Grupo Natureza"], filtros_grupo_natureza, filtros_orgao_origem),
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
# STICK ESTRITO — Top-200 (respeita only-locked)
# =============================================================================

def _stick_previous_topN_STRICT(res_df: pd.DataFrame, df_prev_map: pd.DataFrame,
                                filtros_grupo_natureza: Dict[str, List[str]],
                                filtros_orgao_origem: Dict[str, List[str]],
                                only_locked_map: Dict[str, bool], top_n_per_inf: int = 200):
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

    df["Critério"] = df.apply(calcula_criterio, axis=1)
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
            lambda r: True if bool(r.get("Locked", False)) else _accepts(
                inf, r["Orgão Origem"], r["Grupo Natureza"],
                filtros_grupo_natureza, filtros_orgao_origem
            ),
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
            "Informante": inf,
            "colados": int(len(colaveis)),
            "bloq_only_locked": int(bloco["bloq_only_locked"].sum()),
            "bloq_whitelist": 0
        })

    df = df.drop(columns=["CustomPriority"], errors="ignore")
    diag_df = pd.DataFrame(diag_rows, columns=["Informante","colados","bloq_only_locked","bloq_whitelist"])
    return df, diag_df

# =============================================================================
# Auto-sync pools
#   - Pool Único: auto-inclui novos disponíveis
#   - Pool Geral (B): auto-inclui novos disponíveis (recomendado no modelo A/B)
#   - Pool Especial (A adicional): NÃO auto-inclui
# =============================================================================

def _sync_pool_selections(informantes_disponiveis: List[str]) -> None:
    """
    Modelo A/B (Opção 1):
      - B é o pool geral (base).
      - A é um pool adicional para órgãos especiais (pode sobrepor B).
      - Alguns podem estar só em B; alguns em A e B.

    Sincronização:
      - novos informantes "sim" entram automaticamente em B e no Pool Único.
      - A não é preenchido automaticamente.
      - remove das seleções quem deixou de estar disponível.
    """
    avail_set = set(informantes_disponiveis)

    if "pool_b_sel" not in st.session_state:
        st.session_state["pool_b_sel"] = informantes_disponiveis[:]
    if "pool_a_sel" not in st.session_state:
        st.session_state["pool_a_sel"] = []
    if "pool_unico_sel" not in st.session_state:
        st.session_state["pool_unico_sel"] = informantes_disponiveis[:]

    roster_key = "roster_sig_v2"
    roster_sig = "|".join(sorted(informantes_disponiveis))
    prev_sig = st.session_state.get(roster_key)

    if prev_sig != roster_sig:
        a_set = set(st.session_state.get("pool_a_sel", []))
        b_set = set(st.session_state.get("pool_b_sel", []))
        p_set = set(st.session_state.get("pool_unico_sel", []))

        # remove indisponíveis
        a_set &= avail_set
        b_set &= avail_set
        p_set &= avail_set

        # novos disponíveis
        novos = avail_set.difference(a_set.union(b_set).union(p_set))

        # auto em B e no pool único
        b_set |= novos
        p_set |= novos

        st.session_state["pool_a_sel"] = sorted(a_set)  # A NÃO recebe novos automaticamente
        st.session_state["pool_b_sel"] = sorted(b_set)
        st.session_state["pool_unico_sel"] = sorted(p_set)

        st.session_state[roster_key] = roster_sig

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
    # ====== Leitura e normalização para opções de filtro ======
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
    # CADASTRO EDITÁVEL (UI)
    # =========================
    st.markdown("## Cadastro de Informantes (válido para a execução)")
    st.caption("Você pode incluir/editar informantes e e-mails na hora. A execução usa este cadastro.")

    df_disp_base = pd.read_excel(files_dict["disponibilidade"])
    df_disp_base.columns = df_disp_base.columns.str.strip()

    # garante colunas mínimas
    if "informantes" not in df_disp_base.columns:
        df_disp_base["informantes"] = ""
    if "email" not in df_disp_base.columns:
        df_disp_base["email"] = ""
    if "disponibilidade" not in df_disp_base.columns:
        df_disp_base["disponibilidade"] = "sim"

    df_disp_base["informantes"] = df_disp_base["informantes"].astype(str).str.strip().str.upper()
    df_disp_base["email"] = df_disp_base["email"].astype(str).str.strip()
    df_disp_base["disponibilidade"] = df_disp_base["disponibilidade"].astype(str).str.strip().str.lower()

    if "df_disp_editor" not in st.session_state:
        st.session_state["df_disp_editor"] = df_disp_base.copy()

    df_disp_editor = st.data_editor(
        st.session_state["df_disp_editor"],
        num_rows="dynamic",
        use_container_width=True,
        key="disp_editor",
        column_config={
            "informantes": st.column_config.TextColumn("Informante (NOME)", required=True),
            "email": st.column_config.TextColumn("E-mail", required=False),
            "disponibilidade": st.column_config.SelectboxColumn("Disponível?", options=["sim", "nao"], required=True),
        }
    )

    df_disp_editor = df_disp_editor.copy()
    df_disp_editor["informantes"] = df_disp_editor["informantes"].astype(str).str.strip().str.upper()
    df_disp_editor["email"] = df_disp_editor["email"].astype(str).str.strip()
    df_disp_editor["disponibilidade"] = df_disp_editor["disponibilidade"].astype(str).str.strip().str.lower()
    df_disp_editor = df_disp_editor[df_disp_editor["informantes"].astype(str).str.strip() != ""].copy()

    invalid_emails = df_disp_editor[
        (df_disp_editor["email"].astype(str).str.strip() != "") &
        (~df_disp_editor["email"].apply(_is_valid_email))
    ]
    if not invalid_emails.empty:
        st.warning("Há e-mails com formato inválido no cadastro (o envio pode falhar).")

    st.session_state["df_disp_editor"] = df_disp_editor

    # =========================
    # BOTÃO: baixar cadastro atualizado
    # =========================
    st.markdown("### Exportar cadastro atualizado")
    st.caption("Baixe um XLSX com o cadastro acima para reaproveitar na próxima semana (renomeie para 'disponibilidade_equipe.xlsx').")

    cadastro_export = df_disp_editor.copy()
    cols = ["informantes", "email", "disponibilidade"]
    for c in cols:
        if c not in cadastro_export.columns:
            cadastro_export[c] = ""
    cadastro_export = cadastro_export[cols]

    cadastro_filename = f"disponibilidade_equipe_atualizada_{datetime.now().strftime('%Y%m%d')}.xlsx"
    cadastro_bytes = to_excel_bytes(cadastro_export)

    st.download_button(
        "Baixar disponibilidade atualizada (XLSX)",
        data=cadastro_bytes,
        file_name=cadastro_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =========================
    # Disponíveis + auto-sync pools
    # =========================
    df_disp_disponiveis = df_disp_editor[df_disp_editor["disponibilidade"] == "sim"].copy()
    informantes_disponiveis = sorted(df_disp_disponiveis["informantes"].dropna().unique().tolist())
    informantes_emails_ui = dict(zip(df_disp_editor["informantes"], df_disp_editor["email"]))

    if not informantes_disponiveis:
        st.error("Não há informantes com disponibilidade = 'sim'. Ajuste o cadastro acima para prosseguir.")
        st.stop()

    # >>> auto-inclusão:
    #     - novos "sim" entram automaticamente em B e no Pool Único
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

    # =========================
    # Seletores: Opção 1 (A adicional + B base) OU Pool Único
    # =========================
    st.markdown("## Seleção do Pool de Informantes (PRINCIPAIS)")
    st.caption("Opção 1: B é o pool geral (base). A é um pool adicional usado apenas para órgãos especiais — e pode sobrepor B.")

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
            st.warning("Pool B (geral) está vazio. A distribuição poderá ficar sem candidatos para órgãos não especiais.")

        st.markdown("### Órgãos especiais (disparam uso do pool A)")
        st.caption("Por enquanto, a lista é fixa. Se quiser, eu transformo em multiselect alimentado pelo 'Orgão Origem' real do dia.")
        origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]
        st.write(origens_especiais)

    else:
        if st.button("Selecionar TODOS no Pool Único"):
            st.session_state["pool_unico_sel"] = informantes_disponiveis[:]

        pool_unico_sel = st.multiselect(
            "Pool Único (todos para todos) — selecione quem participa",
            options=informantes_disponiveis,
            default=[x for x in st.session_state.get("pool_unico_sel", []) if x in informantes_disponiveis] or informantes_disponiveis[:],
            key="pool_unico_multiselect"
        )
        st.session_state["pool_unico_sel"] = pool_unico_sel

        if not pool_unico_sel:
            st.warning("Pool único vazio. A distribuição não terá candidatos.")

    # =========================
    # Filtros (whitelist) por informante
    # =========================
    st.markdown("## Filtros (whitelist) por informante")
    st.caption("Vazio = aceita tudo. A whitelist é preferencial: em falta de candidato, o sistema faz fallback para garantir destino.")

    filtros_grupo_natureza: Dict[str, List[str]] = {}
    filtros_orgao_origem: Dict[str, List[str]] = {}
    for inf in informantes_disponiveis:
        filtros_grupo_natureza[inf] = st.multiselect(
            f"Naturezas aceitas — {inf}",
            options=grupo_natureza_options,
            key=f"gn_{inf.replace(' ','_')}"
        )
        filtros_orgao_origem[inf] = st.multiselect(
            f"Órgãos aceitos — {inf}",
            options=orgaos_origem_options,
            key=f"org_{inf.replace(' ','_')}"
        )

    # =========================
    # Preferências: only-locked
    # =========================
    st.markdown("## Preferências por informante")
    st.caption("Marque para que o informante receba apenas itens de regras EXCLUSIVAS (Locked). Em fallback extremo, ele pode receber outros itens para garantir destino (marcado em Fallback Tier).")
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
            "Informante": st.column_config.SelectboxColumn("Informante", options=inf_opts, required=True),
            "Grupo Natureza": st.column_config.SelectboxColumn("Grupo Natureza", options=natureza_opts, required=True),
            "Orgão Origem": st.column_config.SelectboxColumn("Orgão Origem", options=orgao_opts, required=True),
            "Exclusiva?": st.column_config.CheckboxColumn("Exclusiva?")
        }
    )
    st.session_state["rules_state_v2"] = rules_state

    # =========================
    # Botão principal
    # =========================
    if st.button("Executar Distribuição"):
        def run_distribution(processos_file, processosmanter_file, obs_file, numero,
                             filtros_grupo_natureza, filtros_orgao_origem,
                             rules_df=None, prev_file=None):

            # 1) Leitura base
            df = pd.read_excel(processos_file); df.columns = df.columns.str.strip()
            for col in ["Grupo Natureza", "Orgão Origem"]:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.upper()
            df["Processo"] = df["Processo"].astype(str)

            df_manter_local = pd.read_excel(processosmanter_file); df_manter_local.columns = df_manter_local.columns.str.strip()
            processos_validos_local = df_manter_local["Processo"].dropna().astype(str).unique()
            df = df[df["Processo"].isin(processos_validos_local)]
            if "Tipo Processo" in df.columns:
                df = df[df["Tipo Processo"].astype(str).str.upper().str.strip() == "PRINCIPAL"]

            required_cols = [
                "Processo", "Grupo Natureza", "Orgão Origem", "Dias no Orgão",
                "Tempo TCERJ", "Data Última Carga", "Descrição Informação", "Funcionário Informação"
            ]
            df = df[required_cols]
            df["Descrição Informação"] = df["Descrição Informação"].astype(str).str.strip().str.lower()
            df["Funcionário Informação"] = df["Funcionário Informação"].astype(str).str.strip().str.upper()

            df_obs = pd.read_excel(obs_file); df_obs.columns = df_obs.columns.str.strip()
            df = pd.merge(df, df_obs[["Processo", "Obs", "Data Obs"]], on="Processo", how="left")
            if "Obs" in df.columns:
                mask_suspensa = df["Obs"].astype(str).str.lower().str.contains("análise suspensa")
                df = df[~mask_suspensa].copy()

            df["Data Última Carga"] = pd.to_datetime(df["Data Última Carga"], errors="coerce")
            df["Data Obs"] = pd.to_datetime(df["Data Obs"], errors="coerce")

            def update_obs(row):
                if pd.notna(row["Data Obs"]) and pd.notna(row["Data Última Carga"]) and row["Data Obs"] > row["Data Última Carga"]:
                    return pd.Series([row["Obs"], row["Data Obs"]])
                return pd.Series(["", ""])

            df[["Obs", "Data Obs"]] = df.apply(update_obs, axis=1)
            df = df.drop(columns=["Data Última Carga"])

            # 2) Pré-atribuídos e principais
            mask_pre = df["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df["Funcionário Informação"] != "")
            pre_df = df[mask_pre].copy(); pre_df["Informante"] = pre_df["Funcionário Informação"]
            res_df = df[~mask_pre].copy()

            if "Locked" not in pre_df.columns:
                pre_df["Locked"] = False

            # 3) Cadastro/Disponibilidade em memória (UI)
            df_disp_local = st.session_state["df_disp_editor"].copy()
            df_disp_local["informantes"] = df_disp_local["informantes"].astype(str).str.strip().str.upper()
            df_disp_local["email"] = df_disp_local["email"].astype(str).str.strip()
            df_disp_local["disponibilidade"] = df_disp_local["disponibilidade"].astype(str).str.strip().str.lower()

            df_disp_ok = df_disp_local[df_disp_local["disponibilidade"] == "sim"].copy()
            available = sorted(df_disp_ok["informantes"].dropna().unique().tolist())
            informantes_emails = dict(zip(df_disp_local["informantes"], df_disp_local["email"]))

            if not available:
                raise ValueError("Nenhum informante disponível (cadastro/planilha).")

            modelo = st.session_state.get("modelo_dist", "A/B por órgão (Opção 1: A adicional + B base)")
            only_locked_map_local: Dict[str, bool] = st.session_state.get("only_locked_map", {})

            # 4) Monta pools conforme modelo
            if modelo.startswith("A/B"):
                pool_b = [x for x in st.session_state.get("pool_b_sel", []) if x in available]
                pool_a = [x for x in st.session_state.get("pool_a_sel", []) if x in available]

                # fallback defensivo: se B vazio, usa todos disponíveis
                if not pool_b:
                    pool_b = available[:]

                informantes_grupo_b = [s.upper() for s in pool_b]  # BASE
                informantes_grupo_a = [s.upper() for s in pool_a]  # ADICIONAL (pode sobrepor B)

                origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]

                # Se A estiver vazio, órgãos especiais caem no B (por desenho)
                if not informantes_grupo_a:
                    informantes_grupo_a = informantes_grupo_b[:]

            else:
                pool = [x for x in st.session_state.get("pool_unico_sel", []) if x in available]
                if not pool:
                    pool = available[:]
                informantes_grupo_a = []                  # não usado
                informantes_grupo_b = [s.upper() for s in pool]
                origens_especiais = []                    # tudo cai no B

            # 5) Regras
            rules_list = []
            if rules_df is not None and not rules_df.empty:
                tmp = rules_df.copy()
                tmp["Informante"]     = tmp["Informante"].map(lambda v: str(v).strip().upper() if pd.notna(v) else "")
                tmp["Grupo Natureza"] = tmp["Grupo Natureza"].map(lambda v: "(QUALQUER)" if str(v).strip()=="" else str(v).strip().upper())
                tmp["Orgão Origem"]   = tmp["Orgão Origem"].map(lambda v: "(QUALQUER)" if str(v).strip()=="" else str(v).strip().upper())
                tmp = tmp[tmp["Informante"] != ""]
                for _, r in tmp.iterrows():
                    rules_list.append({
                        "Informante": r["Informante"],
                        "Grupo Natureza": r["Grupo Natureza"],
                        "Orgão Origem": r["Orgão Origem"],
                        "Exclusiva?": bool(r.get("Exclusiva?", False))
                    })

            # 6) PRINCIPAIS: pool -> regras -> redistribuição (100% destino)
            df_pool = res_df.copy().sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            df_pool["Informante"] = ""
            if "Locked" not in df_pool.columns:
                df_pool["Locked"] = False

            assigned_by_rules, rem_after_rules = _apply_routing_rules(
                df_pool, rules_list,
                filtros_grupo_natureza, filtros_orgao_origem,
                only_locked_map=only_locked_map_local
            )

            distributed_df = _redistribute(
                rem_after_rules,
                informantes_grupo_a, informantes_grupo_b, origens_especiais,
                filtros_grupo_natureza, filtros_orgao_origem,
                only_locked_map=only_locked_map_local
            )

            res_final = pd.concat([assigned_by_rules, distributed_df], ignore_index=True)
            res_final["Informante"] = res_final["Informante"].astype(str).str.strip().str.upper()

            # sanity: não deveria ocorrer se houver pool válido
            if (res_final["Informante"].astype(str).str.strip() == "").any():
                st.warning("Há processos sem informante (SEM_CANDIDATO). Verifique pools e disponibilidade.")

            # Ordenação final por critério
            res_final["Critério"] = res_final.apply(calcula_criterio, axis=1)
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
                    filtros_grupo_natureza, filtros_orgao_origem,
                    only_locked_map=only_locked_map_local, top_n_per_inf=200
                )
                # Reordena após stick
                res_final["Critério"] = res_final.apply(calcula_criterio, axis=1)
                res_final["CustomPriority"] = res_final["Critério"].apply(lambda x: priority_map.get(x, 4))
                res_final = res_final.sort_values(
                    by=["Informante", "CustomPriority", "Dias no Orgão"],
                    ascending=[True, True, False]
                ).reset_index(drop=True)
                res_final = res_final.drop(columns=["CustomPriority"], errors="ignore")

            # 8) Planilhas gerais
            pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            pre_geral_bytes = to_excel_bytes(pre_df)

            res_geral_filename = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            res_geral_bytes = to_excel_bytes(
                res_final.drop(columns=["Descrição Informação", "Funcionário Informação"], errors="ignore")
            )

            # 9) Planilhas individuais — Pré
            def build_pre_individuals():
                pre_individual_files = {}
                pre_df_local = pre_df.copy()
                pre_df_local["Informante"] = pre_df_local["Informante"].astype(str).str.strip().str.upper()
                pre_df_local["Critério"] = pre_df_local.apply(calcula_criterio, axis=1)
                pre_df_local["CustomPriority"] = pre_df_local["Critério"].apply(lambda x: priority_map.get(x, 4))

                for inf in pre_df_local["Informante"].dropna().astype(str).str.upper().unique():
                    if not inf:
                        continue
                    df_inf = pre_df_local[pre_df_local["Informante"].astype(str).str.upper() == inf].copy()
                    if df_inf.empty:
                        continue

                    locked_part = df_inf[_locked_mask(df_inf)].copy()
                    nonlocked   = df_inf[~_locked_mask(df_inf)].copy()

                    locked_part["ForaWhitelist"] = False
                    nonlocked["ForaWhitelist"] = nonlocked.apply(
                        lambda r: not _accepts(inf, r["Orgão Origem"], r["Grupo Natureza"], filtros_grupo_natureza, filtros_orgao_origem),
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
            top200_dict = _apply_prevention_top200(
                res_final, df_prev_map, filtros_grupo_natureza, filtros_orgao_origem
            )
            for inf, df_inf in top200_dict.items():
                res_individual_files[inf] = to_excel_bytes(df_inf)

            return (pre_geral_filename, pre_geral_bytes,
                    res_geral_filename, res_geral_bytes,
                    pre_individual_files, res_individual_files,
                    informantes_emails, diag_df)

        # ---- Executa ----
        try:
            (pre_geral_filename, pre_geral_bytes,
             res_geral_filename, res_geral_bytes,
             pre_individual_files, res_individual_files,
             informantes_emails, diag_df) = run_distribution(
                files_dict["processos"], files_dict["processosmanter"],
                files_dict["observacoes"], numero,
                filtros_grupo_natureza, filtros_orgao_origem,
                rules_df=st.session_state.get("rules_state_v2"), prev_file=prev_file
            )
        except Exception as e:
            st.error(f"Falha na distribuição: {e}")
            st.stop()

        st.success("Distribuição executada com sucesso!")

        if isinstance(diag_df, pd.DataFrame) and not diag_df.empty:
            st.markdown("### Diagnóstico da Prevenção (Stick Estrito Top-200)")
            st.dataframe(diag_df.sort_values(by=["Informante"]).reset_index(drop=True), use_container_width=True)

        # Downloads
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

        # Envio de e-mails
        if not test_mode:
            managers_list = [e.strip() for e in managers_emails.split(",") if e.strip()]
            if managers_list:
                all_individual_files = {}
                for inf, b in pre_individual_files.items():
                    filename_ind = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    all_individual_files[filename_ind] = b
                for inf, b in res_individual_files.items():
                    filename_ind = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    all_individual_files[filename_ind] = b

                zip_individual_bytes = create_zip_from_dict(all_individual_files)
                zip_filename = f"{numero}_planilhas_individuais_{datetime.now().strftime('%Y%m%d')}.zip"

                attachments = [
                    (pre_geral_bytes, pre_geral_filename),
                    (res_geral_bytes, res_geral_filename),
                    (zip_individual_bytes, zip_filename)
                ]
                subject_managers = "Planilhas Gerais e Individuais de Processos"
                body_managers = (
                    "Prezado(a) Gestor(a),\n\n"
                    "Seguem anexas as planilhas:\n"
                    "- Geral de Processos Pré-Atribuídos\n"
                    "- Geral de Processos Principais\n"
                    "- ZIP com todas as planilhas individuais\n\n"
                    "Atenciosamente,\nGestão da 3ª CAP"
                )
                send_email_with_multiple_attachments(managers_list, subject_managers, body_managers, attachments)

            if modo_envio == "Produção - Gestores e Informantes":
                for inf in set(list(pre_individual_files.keys()) + list(res_individual_files.keys())):
                    email_destino = informantes_emails.get(inf.upper(), "")
                    if email_destino:
                        attachment_pre = pre_individual_files.get(inf)
                        attachment_res = res_individual_files.get(inf)

                        filename_pre = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx" if attachment_pre else None
                        filename_res = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx" if attachment_res else None

                        subject_inf = f"Distribuição de Processos - {inf}"
                        body_inf = (
                            "Prezado(a) Informante,\n\n"
                            "Seguem anexas as planilhas referentes à distribuição de processos:\n\n"
                            "• Pré-Atribuídos: vinculados a você no sistema (andamento/conclusão).\n"
                            "• Principais: novos processos distribuídos.\n\n"
                            "Atenciosamente,\nGestão da 3ª CAP"
                        )
                        send_email_with_two_attachments(
                            email_destino, subject_inf, body_inf,
                            attachment_pre, filename_pre,
                            attachment_res, filename_res
                        )

        st.session_state.numero = numero

else:
    st.info("Carregue os quatro arquivos exigidos para habilitar a execução.")
