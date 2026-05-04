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
# E-mail
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
        st.warning("Credenciais de e-mail ausentes. E-mail não enviado.")
        return
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"]    = smtp_username
    msg["To"]      = ", ".join(to_emails)
    msg.set_content(body)
    for attachment_bytes, filename in attachments:
        msg.add_attachment(attachment_bytes, maintype="application",
                           subtype="octet-stream", filename=filename)
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10) as server:
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            st.info(f"E-mail enviado para: {to_emails}")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail para {to_emails}: {e}")

def send_email_with_two_attachments(to_email, subject, body,
                                    attachment_pre, filename_pre,
                                    attachment_res, filename_res):
    smtp_server, smtp_port, smtp_username, smtp_password = _get_mail_creds()
    if not smtp_username or not smtp_password:
        st.warning("Credenciais de e-mail ausentes. E-mail não enviado.")
        return
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"]    = smtp_username
    msg["To"]      = to_email
    msg.set_content(body)
    if attachment_pre is not None:
        msg.add_attachment(attachment_pre, maintype="application",
                           subtype="octet-stream", filename=filename_pre)
    if attachment_res is not None:
        msg.add_attachment(attachment_res, maintype="application",
                           subtype="octet-stream", filename=filename_res)
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
        workbook  = writer.book
        worksheet = writer.sheets["Planilha"]
        if "Grupo Natureza" in df.columns:
            col_index  = df.columns.get_loc("Grupo Natureza")
            col_letter = xl_col_to_name(col_index)
            last_row   = len(df) + 1
            cell_range = f"{col_letter}2:{col_letter}{last_row}"
            color_list = ["#990000","#006600","#996600","#003366",
                          "#660066","#663300","#003300","#000066"]
            unique_values = df["Grupo Natureza"].dropna().unique()
            color_mapping = {val: color_list[i % len(color_list)]
                             for i, val in enumerate(unique_values)}
            for value, color in color_mapping.items():
                fmt = workbook.add_format({"font_color": color, "bold": True})
                worksheet.conditional_format(cell_range, {
                    "type": "cell", "criteria": "==",
                    "value": f'"{value}"', "format": fmt
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
# Validações
# =============================================================================

def _is_valid_email(x: str) -> bool:
    if not isinstance(x, str):
        return False
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", x.strip()))

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
    "03 Mais de 5 meses na 3CAP":       2,
    "04 Data da carga":                  3,
}

# =============================================================================
# Whitelist tabular — leitura do session_state
#
# Cada linha da whitelist é armazenada como três widgets independentes:
#   st.session_state[f"wl_inf_{id}"]  → informante
#   st.session_state[f"wl_nat_{id}"]  → natureza  (vazio = qualquer)
#   st.session_state[f"wl_org_{id}"]  → órgão     (vazio = qualquer)
#
# A lista de IDs ativos fica em st.session_state["wl_rows"].
# Como os widgets têm keys estáveis, o Streamlit nunca os apaga em reruns.
# =============================================================================

WhitelistPairs = Dict[str, List[Tuple[str, str]]]

def _read_whitelist_pairs() -> WhitelistPairs:
    """Lê os valores dos widgets de whitelist diretamente do session_state."""
    pairs: WhitelistPairs = {}
    for row_id in st.session_state.get("wl_rows", []):
        inf = st.session_state.get(f"wl_inf_{row_id}", "")
        nat = st.session_state.get(f"wl_nat_{row_id}", "")
        org = st.session_state.get(f"wl_org_{row_id}", "")
        if inf:
            pairs.setdefault(inf, []).append((nat, org))
    return pairs

def _accepts(inf: str, orgao: str, natureza: str,
             whitelist_pairs: WhitelistPairs) -> bool:
    """
    Whitelist tabular: lista vazia = aceita tudo.
    Aceita se QUALQUER linha da lista fizer match.
    String vazia no filtro = curinga.
    """
    pairs = whitelist_pairs.get(inf, [])
    if not pairs:
        return True
    for nat_f, org_f in pairs:
        ok_nat = (not nat_f) or (nat_f == natureza)
        ok_org = (not org_f) or (org_f == orgao)
        if ok_nat and ok_org:
            return True
    return False

# =============================================================================
# Helpers
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

            score = (0 if nat_rule == "(QUALQUER)" else 1) + \
                    (0 if org_rule == "(QUALQUER)" else 1)
            cand  = (score, idx, r)
            if (best is None) or (cand[0] > best[0]) or \
               (cand[0] == best[0] and cand[1] < best[1]):
                best = cand

        if best is None:
            remaining_rows.append(row)
        else:
            _, _, rbest  = best
            inf_best     = rbest["Informante"]
            exclusiva    = bool(rbest.get("Exclusiva?", False))

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
# Redistribuição
# =============================================================================

def _redistribute(df_unassigned: pd.DataFrame,
                  informantes_grupo_a: List[str],
                  informantes_grupo_b: List[str],
                  origens_especiais: List[str],
                  whitelist_pairs: WhitelistPairs,
                  only_locked_map: Dict[str, bool]) -> pd.DataFrame:
    if df_unassigned.empty:
        return df_unassigned

    df_unassigned = df_unassigned.copy()
    if "Fallback Tier"   not in df_unassigned.columns:
        df_unassigned["Fallback Tier"]   = ""
    if "Fallback Motivo" not in df_unassigned.columns:
        df_unassigned["Fallback Motivo"] = ""

    rr_indices = {gn: 0 for gn in df_unassigned["Grupo Natureza"].unique()}

    def _candidates(pool, orgao, natureza, allow_only_locked, ignore_whitelist):
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

        pool = (informantes_grupo_a
                if (origens_especiais and orgao in origens_especiais)
                else informantes_grupo_b)

        # Reserva: informantes com whitelist que cobre este processo
        pool_reservado = [
            inf for inf in pool
            if whitelist_pairs.get(inf)
            and _accepts(inf, orgao, natureza, whitelist_pairs)
        ]
        pool_primario = pool_reservado if pool_reservado else pool

        candidatos = _candidates(pool_primario, orgao, natureza,
                                 allow_only_locked=False, ignore_whitelist=False)
        tier, motivo = "T0", ""

        if not candidatos:
            candidatos = _candidates(pool_primario, orgao, natureza,
                                     allow_only_locked=False, ignore_whitelist=True)
            tier   = "T1"
            motivo = "Whitelist ignorada (sem candidato no pool reservado)"

        if not candidatos:
            candidatos = _candidates(pool, orgao, natureza,
                                     allow_only_locked=True, ignore_whitelist=False)
            tier   = "T2"
            motivo = "Somente candidatos only-locked; manteve whitelist"

        if not candidatos:
            candidatos = _candidates(pool, orgao, natureza,
                                     allow_only_locked=True, ignore_whitelist=True)
            tier   = "T3"
            motivo = "Fallback final: only-locked e whitelist ignorados"

        if not candidatos:
            row["Informante"]     = ""
            row["Fallback Tier"]  = "SEM_CANDIDATO"
            row["Fallback Motivo"]= "Nenhum informante disponível no pool"
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
# Prevenção Top-200
# =============================================================================

def _apply_prevention_top200(res_final: pd.DataFrame,
                             df_prev_map: Optional[pd.DataFrame],
                             whitelist_pairs: WhitelistPairs) -> dict:
    out  = {}
    base = res_final.copy()
    base["Critério"]       = base.apply(calcula_criterio, axis=1)
    base["CustomPriority"] = base["Critério"].apply(lambda x: priority_map.get(x, 4))
    base = base.sort_values(["CustomPriority", "Dias no Orgão"], ascending=[True, False])

    prev_map = {}
    if df_prev_map is not None and not df_prev_map.empty:
        prev = df_prev_map.copy()
        prev.columns = [c.strip() for c in prev.columns]
        if {"Processo", "Informante"}.issubset(prev.columns):
            prev["Processo"]   = prev["Processo"].astype(str).str.strip()
            prev["Informante"] = prev["Informante"].astype(str).str.strip().str.upper()
            prev_map = dict(zip(prev["Processo"], prev["Informante"]))

    for inf in base["Informante"].dropna().astype(str).str.strip().str.upper().unique():
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
        df_inf["preferido"] = df_inf["Processo"].astype(str).map(
            lambda p: 1 if prev_map.get(p) == inf else 0
        ) if prev_map else 0

        preferidos = df_inf[df_inf["preferido"] == 1].sort_values(
            ["CustomPriority", "Dias no Orgão"], ascending=[True, False])
        nao_pref   = df_inf[df_inf["preferido"] == 0].sort_values(
            ["CustomPriority", "Dias no Orgão"], ascending=[True, False])

        df_top = pd.concat([preferidos, nao_pref], ignore_index=True).head(200)
        df_top = df_top.drop(columns=["CustomPriority", "preferido"], errors="ignore")
        out[inf] = df_top

    return out

# =============================================================================
# Stick estrito Top-200
# =============================================================================

def _stick_previous_topN_STRICT(res_df: pd.DataFrame,
                                df_prev_map: pd.DataFrame,
                                whitelist_pairs: WhitelistPairs,
                                only_locked_map: Dict[str, bool],
                                top_n_per_inf: int = 200):
    empty_diag = pd.DataFrame(columns=["Informante","colados","bloq_only_locked","bloq_whitelist"])
    if res_df.empty or df_prev_map is None or df_prev_map.empty:
        return res_df, empty_diag

    prev = df_prev_map.copy()
    prev.columns = [c.strip() for c in prev.columns]
    if not {"Processo", "Informante"}.issubset(prev.columns):
        return res_df, empty_diag

    prev["Processo"]   = prev["Processo"].astype(str).str.strip()
    prev["Informante"] = prev["Informante"].astype(str).str.strip().str.upper()

    df = res_df.copy()
    if "Locked" not in df.columns:
        df["Locked"] = False

    df["Critério"]       = df.apply(calcula_criterio, axis=1)
    df["CustomPriority"] = df["Critério"].apply(lambda x: priority_map.get(x, 4))

    diag_rows = []
    for inf in sorted(df["Informante"].astype(str).str.upper().dropna().unique()):
        if not inf.strip():
            continue

        prev_set = set(prev[prev["Informante"] == inf]["Processo"].astype(str))
        bloco    = df[df["Processo"].astype(str).isin(prev_set)].copy()
        if bloco.empty:
            diag_rows.append({"Informante": inf, "colados": 0,
                              "bloq_only_locked": 0, "bloq_whitelist": 0})
            continue

        bloco["ok_whitelist"] = bloco.apply(
            lambda r: True if bool(r.get("Locked", False))
                      else _accepts(inf, r["Orgão Origem"], r["Grupo Natureza"], whitelist_pairs),
            axis=1
        )
        bloco = bloco[bloco["ok_whitelist"]].copy()
        bloco = bloco.sort_values(["CustomPriority","Dias no Orgão"],
                                  ascending=[True,False]).head(top_n_per_inf)

        if bloco.empty:
            diag_rows.append({"Informante": inf, "colados": 0,
                              "bloq_only_locked": 0, "bloq_whitelist": int(len(prev_set))})
            continue

        sl_only = bool(only_locked_map.get(inf, False))
        bloco["bloq_only_locked"] = bloco.apply(
            lambda r: bool(sl_only and not r.get("Locked", False)), axis=1)

        colaveis = bloco[~bloco["bloq_only_locked"]].copy()
        if not colaveis.empty:
            df.loc[df["Processo"].astype(str).isin(
                colaveis["Processo"].astype(str)), "Informante"] = inf

        diag_rows.append({
            "Informante":       inf,
            "colados":          int(len(colaveis)),
            "bloq_only_locked": int(bloco["bloq_only_locked"].sum()),
            "bloq_whitelist":   0
        })

    df = df.drop(columns=["CustomPriority"], errors="ignore")
    return df, pd.DataFrame(diag_rows,
                            columns=["Informante","colados","bloq_only_locked","bloq_whitelist"])

# =============================================================================
# Auto-sync pools
# =============================================================================

def _sync_pool_selections(informantes_disponiveis: List[str]) -> None:
    avail_set = set(informantes_disponiveis)

    for key, default in [("pool_b_sel", informantes_disponiveis[:]),
                         ("pool_a_sel",  []),
                         ("pool_unico_sel", informantes_disponiveis[:])]:
        if key not in st.session_state:
            st.session_state[key] = default

    roster_sig = "|".join(sorted(informantes_disponiveis))
    if st.session_state.get("roster_sig_v2") != roster_sig:
        a_set = set(st.session_state.get("pool_a_sel", []))   & avail_set
        b_set = set(st.session_state.get("pool_b_sel", []))   & avail_set
        p_set = set(st.session_state.get("pool_unico_sel",[])) & avail_set
        novos = avail_set - (a_set | b_set | p_set)
        b_set |= novos
        p_set |= novos
        st.session_state["pool_a_sel"]     = sorted(a_set)
        st.session_state["pool_b_sel"]     = sorted(b_set)
        st.session_state["pool_unico_sel"] = sorted(p_set)
        st.session_state["roster_sig_v2"]  = roster_sig

# =============================================================================
# UI — whitelist estável com selectboxes individuais
# =============================================================================

def _render_whitelist_editor(informantes_disponiveis, grupo_natureza_options, orgaos_origem_options):
    """
    Renderiza o editor de whitelist usando selectboxes individuais com keys
    estáveis. Não usa data_editor, evitando perda de estado em reruns.
    """
    # Inicializa estruturas de controle
    if "wl_rows" not in st.session_state:
        st.session_state["wl_rows"] = []        # lista de IDs ativos
    if "wl_next_id" not in st.session_state:
        st.session_state["wl_next_id"] = 0

    nat_options = ["(qualquer)"] + grupo_natureza_options
    org_options = ["(qualquer)"] + orgaos_origem_options

    # Cabeçalho das colunas
    if st.session_state["wl_rows"]:
        h1, h2, h3, h4 = st.columns([2, 2, 2, 0.5])
        h1.markdown("**Informante**")
        h2.markdown("**Natureza** *(vazio = qualquer)*")
        h3.markdown("**Órgão** *(vazio = qualquer)*")
        h4.markdown("** **")

    # Linhas ativas
    ids_para_remover = []
    for row_id in list(st.session_state["wl_rows"]):
        c1, c2, c3, c4 = st.columns([2, 2, 2, 0.5])

        # Garante valor padrão na 1ª vez que o widget aparece
        if f"wl_inf_{row_id}" not in st.session_state:
            st.session_state[f"wl_inf_{row_id}"] = informantes_disponiveis[0] if informantes_disponiveis else ""
        if f"wl_nat_{row_id}" not in st.session_state:
            st.session_state[f"wl_nat_{row_id}"] = "(qualquer)"
        if f"wl_org_{row_id}" not in st.session_state:
            st.session_state[f"wl_org_{row_id}"] = "(qualquer)"

        # Selectbox de informante — garante que o valor salvo ainda existe nas opções
        inf_saved = st.session_state[f"wl_inf_{row_id}"]
        inf_idx   = informantes_disponiveis.index(inf_saved) if inf_saved in informantes_disponiveis else 0
        c1.selectbox("inf", informantes_disponiveis,
                     index=inf_idx,
                     key=f"wl_inf_{row_id}",
                     label_visibility="collapsed")

        nat_saved = st.session_state[f"wl_nat_{row_id}"]
        nat_idx   = nat_options.index(nat_saved) if nat_saved in nat_options else 0
        c2.selectbox("nat", nat_options,
                     index=nat_idx,
                     key=f"wl_nat_{row_id}",
                     label_visibility="collapsed")

        org_saved = st.session_state[f"wl_org_{row_id}"]
        org_idx   = org_options.index(org_saved) if org_saved in org_options else 0
        c3.selectbox("org", org_options,
                     index=org_idx,
                     key=f"wl_org_{row_id}",
                     label_visibility="collapsed")

        if c4.button("✕", key=f"wl_del_{row_id}", help="Remover esta linha"):
            ids_para_remover.append(row_id)

    # Processa remoções
    if ids_para_remover:
        for rid in ids_para_remover:
            st.session_state["wl_rows"].remove(rid)
        st.rerun()

    # Botão adicionar linha
    if st.button("＋ Adicionar linha de whitelist"):
        new_id = st.session_state["wl_next_id"]
        st.session_state["wl_rows"].append(new_id)
        st.session_state["wl_next_id"] += 1
        # Pré-inicializa os valores para evitar flash de widget vazio
        st.session_state[f"wl_inf_{new_id}"] = informantes_disponiveis[0] if informantes_disponiveis else ""
        st.session_state[f"wl_nat_{new_id}"] = "(qualquer)"
        st.session_state[f"wl_org_{new_id}"] = "(qualquer)"
        st.rerun()

    # Resumo compacto
    if st.session_state["wl_rows"]:
        with st.expander("Ver resumo por informante"):
            resumo: Dict[str, list] = {}
            for row_id in st.session_state["wl_rows"]:
                inf = st.session_state.get(f"wl_inf_{row_id}", "")
                nat = st.session_state.get(f"wl_nat_{row_id}", "(qualquer)")
                org = st.session_state.get(f"wl_org_{row_id}", "(qualquer)")
                if inf:
                    resumo.setdefault(inf, []).append(
                        f"Natureza={nat or '(qualquer)'} | Órgão={org or '(qualquer)'}"
                    )
            for inf in informantes_disponiveis:
                if inf in resumo:
                    st.write(f"**{inf}** — {len(resumo[inf])} regra(s):")
                    for linha in resumo[inf]:
                        st.write(f"  • {linha}")
                else:
                    st.write(f"**{inf}** — sem restrição (aceita tudo)")
    else:
        st.info("Nenhuma linha configurada — todos os informantes aceitam qualquer processo.")

# =============================================================================
# UI principal
# =============================================================================

st.title("Distribuição de Processos da Del. 260")

if "numero" not in st.session_state:
    st.session_state.numero = "184"

uploaded_files = st.file_uploader(
    "Carregar: processos.xlsx, processosmanter.xlsx, observacoes.xlsx, disponibilidade_equipe.xlsx",
    type=["xlsx"], accept_multiple_files=True
)
prev_file = st.file_uploader(
    "Opcional: PLANILHA GERAL PRINCIPAL da semana anterior (prevenção Top-200). "
    "Deve conter colunas 'Processo' e 'Informante'.",
    type=["xlsx"], accept_multiple_files=False
)

files_dict = {}
for file in uploaded_files or []:
    fname = file.name.lower().strip()
    if   fname == "processos.xlsx":            files_dict["processos"]     = file
    elif fname == "processosmanter.xlsx":      files_dict["processosmanter"] = file
    elif fname in ["observacoes.xlsx","obervacoes.xlsx"]: files_dict["observacoes"] = file
    elif fname == "disponibilidade_equipe.xlsx": files_dict["disponibilidade"] = file

numero = st.text_input("Numeração desta planilha de distribuição:",
                       value=st.session_state.numero)

modo      = st.radio("Selecione o modo:", ["Teste","Produção"], horizontal=True)
test_mode = (modo == "Teste")
st.info("Modo Teste: não envia e-mails." if test_mode else "Modo Produção: e-mails serão enviados.")
st.markdown(f"**Modo selecionado:** {modo}")

modo_envio = None
if not test_mode:
    modo_envio = st.radio("Modo de envio:",
                          ["Produção - Gestores e Informantes",
                           "Produção - Apenas Gestores"])

managers_emails = st.text_input(
    "E-mails dos gestores/revisores (separados por vírgula):",
    value="annapc@tcerj.tc.br, fabiovf@tcerj.tc.br, sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com"
)

# -----------------------------------------------------------------------

if all(k in files_dict for k in ["processos","processosmanter","observacoes","disponibilidade"]):

    df_proc = pd.read_excel(files_dict["processos"])
    df_proc.columns = df_proc.columns.str.strip()
    for col in ["Grupo Natureza","Orgão Origem"]:
        if col in df_proc.columns:
            df_proc[col] = df_proc[col].astype(str).str.strip().str.upper()

    df_manter = pd.read_excel(files_dict["processosmanter"])
    df_manter.columns = df_manter.columns.str.strip()
    processos_validos = df_manter["Processo"].dropna().astype(str).unique()
    df_proc["Processo"] = df_proc["Processo"].astype(str)
    df_proc = df_proc[df_proc["Processo"].isin(processos_validos)]
    if "Tipo Processo" in df_proc.columns:
        df_proc = df_proc[df_proc["Tipo Processo"].astype(str).str.upper().str.strip() == "PRINCIPAL"]

    grupo_natureza_options = sorted(df_proc["Grupo Natureza"].dropna().unique())
    orgaos_origem_options  = sorted(df_proc["Orgão Origem"].dropna().unique())

    # ---- Cadastro de informantes ----
    st.markdown("## Cadastro de Informantes")
    st.caption("Edite na hora. A execução usa este cadastro.")

    df_disp_base = pd.read_excel(files_dict["disponibilidade"])
    df_disp_base.columns = df_disp_base.columns.str.strip()
    for col, default in [("informantes",""),("email",""),("disponibilidade","sim")]:
        if col not in df_disp_base.columns:
            df_disp_base[col] = default
    df_disp_base["informantes"]    = df_disp_base["informantes"].astype(str).str.strip().str.upper()
    df_disp_base["email"]          = df_disp_base["email"].astype(str).str.strip()
    df_disp_base["disponibilidade"]= df_disp_base["disponibilidade"].astype(str).str.strip().str.lower()

    if "df_disp_editor" not in st.session_state:
        st.session_state["df_disp_editor"] = df_disp_base.copy()

    df_disp_editor = st.data_editor(
        st.session_state["df_disp_editor"],
        num_rows="dynamic", use_container_width=True, key="disp_editor",
        column_config={
            "informantes":     st.column_config.TextColumn("Informante (NOME)", required=True),
            "email":           st.column_config.TextColumn("E-mail"),
            "disponibilidade": st.column_config.SelectboxColumn("Disponível?",
                                   options=["sim","nao"], required=True),
        }
    )
    df_disp_editor = df_disp_editor.copy()
    df_disp_editor["informantes"]    = df_disp_editor["informantes"].astype(str).str.strip().str.upper()
    df_disp_editor["email"]          = df_disp_editor["email"].astype(str).str.strip()
    df_disp_editor["disponibilidade"]= df_disp_editor["disponibilidade"].astype(str).str.strip().str.lower()
    df_disp_editor = df_disp_editor[df_disp_editor["informantes"].str.strip() != ""].copy()

    invalid_emails = df_disp_editor[
        (df_disp_editor["email"].str.strip() != "") &
        (~df_disp_editor["email"].apply(_is_valid_email))
    ]
    if not invalid_emails.empty:
        st.warning("Há e-mails com formato inválido no cadastro.")
    st.session_state["df_disp_editor"] = df_disp_editor

    st.markdown("### Exportar cadastro atualizado")
    cadastro_bytes = to_excel_bytes(df_disp_editor[["informantes","email","disponibilidade"]])
    st.download_button(
        "Baixar disponibilidade atualizada (XLSX)", data=cadastro_bytes,
        file_name=f"disponibilidade_equipe_atualizada_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---- Disponíveis ----
    df_disp_ok = df_disp_editor[df_disp_editor["disponibilidade"] == "sim"].copy()
    informantes_disponiveis = sorted(df_disp_ok["informantes"].dropna().unique().tolist())

    if not informantes_disponiveis:
        st.error("Não há informantes disponíveis. Ajuste o cadastro.")
        st.stop()

    _sync_pool_selections(informantes_disponiveis)

    # ---- Modelo de distribuição ----
    st.markdown("## Modelo de Distribuição")
    modelo_dist = st.radio(
        "Selecione o modelo:",
        ["A/B por órgão (Opção 1: A adicional + B base)", "Pool único (todos para todos)"],
        horizontal=True
    )
    st.session_state["modelo_dist"] = modelo_dist

    st.markdown("## Seleção do Pool de Informantes")

    if modelo_dist.startswith("A/B"):
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Selecionar TODOS no pool B"):
                st.session_state["pool_b_sel"] = informantes_disponiveis[:]
        with c2:
            if st.button("Limpar pool A"):
                st.session_state["pool_a_sel"] = []

        pool_b_sel = st.multiselect(
            "Pool B (geral/base)", options=informantes_disponiveis,
            default=[x for x in st.session_state.get("pool_b_sel",[]) if x in informantes_disponiveis]
                    or informantes_disponiveis[:],
            key="pool_b_multiselect"
        )
        pool_a_sel = st.multiselect(
            "Pool A (adicional — órgãos especiais)", options=informantes_disponiveis,
            default=[x for x in st.session_state.get("pool_a_sel",[]) if x in informantes_disponiveis],
            key="pool_a_multiselect"
        )
        st.session_state["pool_b_sel"] = pool_b_sel
        st.session_state["pool_a_sel"] = pool_a_sel
        if not pool_b_sel:
            st.warning("Pool B vazio.")

        st.markdown("### Órgãos especiais")
        origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]
        st.write(origens_especiais)
    else:
        if st.button("Selecionar TODOS no Pool Único"):
            st.session_state["pool_unico_sel"] = informantes_disponiveis[:]
        pool_unico_sel = st.multiselect(
            "Pool Único", options=informantes_disponiveis,
            default=[x for x in st.session_state.get("pool_unico_sel",[]) if x in informantes_disponiveis]
                    or informantes_disponiveis[:],
            key="pool_unico_multiselect"
        )
        st.session_state["pool_unico_sel"] = pool_unico_sel
        if not pool_unico_sel:
            st.warning("Pool único vazio.")

    # ---- WHITELIST — selectboxes estáveis ----
    st.markdown("## Filtros (whitelist) por informante")
    st.caption(
        "Cada linha reserva um par (Natureza, Órgão) para um informante. "
        "Use '(qualquer)' como curinga. "
        "Informante sem nenhuma linha aceita qualquer processo."
    )
    _render_whitelist_editor(informantes_disponiveis,
                             grupo_natureza_options,
                             orgaos_origem_options)

    # ---- Preferências only-locked ----
    st.markdown("## Preferências por informante")
    st.caption("Marque para que o informante receba apenas itens de regras EXCLUSIVAS (Locked).")
    only_locked_map: Dict[str, bool] = st.session_state.get("only_locked_map", {})
    for inf in informantes_disponiveis:
        key = f"only_locked_{inf.replace(' ','_')}"
        only_locked_map[inf] = st.checkbox(
            f"{inf}: receber apenas itens exclusivos (Locked)?",
            value=only_locked_map.get(inf, False), key=key
        )
    st.session_state["only_locked_map"] = only_locked_map

    # ---- Regras de roteamento ----
    st.markdown("## Regras de roteamento por (Natureza, Órgão) → Informante")
    st.caption("Use '(QUALQUER)' como curinga. Marque 'Exclusiva?' para reservar o par.")
    natureza_opts = ["(QUALQUER)"] + grupo_natureza_options
    orgao_opts    = ["(QUALQUER)"] + orgaos_origem_options

    if "rules_state_v2" not in st.session_state:
        st.session_state["rules_state_v2"] = pd.DataFrame(
            columns=["Informante","Grupo Natureza","Orgão Origem","Exclusiva?"])

    rules_state = st.data_editor(
        st.session_state["rules_state_v2"],
        num_rows="dynamic", use_container_width=True, key="rules_editor_v2",
        column_config={
            "Informante":     st.column_config.SelectboxColumn("Informante",
                                  options=informantes_disponiveis, required=True),
            "Grupo Natureza": st.column_config.SelectboxColumn("Grupo Natureza",
                                  options=natureza_opts, required=True),
            "Orgão Origem":   st.column_config.SelectboxColumn("Orgão Origem",
                                  options=orgao_opts, required=True),
            "Exclusiva?":     st.column_config.CheckboxColumn("Exclusiva?"),
        }
    )
    st.session_state["rules_state_v2"] = rules_state

    # ---- Execução ----
    if st.button("Executar Distribuição"):

        def run_distribution(processos_file, processosmanter_file, obs_file,
                             numero, whitelist_pairs, rules_df=None, prev_file=None):

            df = pd.read_excel(processos_file)
            df.columns = df.columns.str.strip()
            for col in ["Grupo Natureza","Orgão Origem"]:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.upper()
            df["Processo"] = df["Processo"].astype(str)

            df_mt = pd.read_excel(processosmanter_file)
            df_mt.columns = df_mt.columns.str.strip()
            pv = df_mt["Processo"].dropna().astype(str).unique()
            df = df[df["Processo"].isin(pv)]
            if "Tipo Processo" in df.columns:
                df = df[df["Tipo Processo"].astype(str).str.upper().str.strip() == "PRINCIPAL"]

            required_cols = ["Processo","Grupo Natureza","Orgão Origem","Dias no Orgão",
                             "Tempo TCERJ","Data Última Carga",
                             "Descrição Informação","Funcionário Informação"]
            df = df[required_cols]
            df["Descrição Informação"]  = df["Descrição Informação"].astype(str).str.strip().str.lower()
            df["Funcionário Informação"]= df["Funcionário Informação"].astype(str).str.strip().str.upper()

            df_obs = pd.read_excel(obs_file)
            df_obs.columns = df_obs.columns.str.strip()
            df = pd.merge(df, df_obs[["Processo","Obs","Data Obs"]], on="Processo", how="left")
            if "Obs" in df.columns:
                df = df[~df["Obs"].astype(str).str.lower().str.contains("análise suspensa")].copy()

            df["Data Última Carga"] = pd.to_datetime(df["Data Última Carga"], errors="coerce")
            df["Data Obs"]          = pd.to_datetime(df["Data Obs"],          errors="coerce")

            def update_obs(row):
                if pd.notna(row["Data Obs"]) and pd.notna(row["Data Última Carga"]) \
                        and row["Data Obs"] > row["Data Última Carga"]:
                    return pd.Series([row["Obs"], row["Data Obs"]])
                return pd.Series(["",""])

            df[["Obs","Data Obs"]] = df.apply(update_obs, axis=1)
            df = df.drop(columns=["Data Última Carga"])

            mask_pre = (df["Descrição Informação"].isin(["em elaboração","concluída"])
                        & (df["Funcionário Informação"] != ""))
            pre_df = df[mask_pre].copy()
            pre_df["Informante"] = pre_df["Funcionário Informação"]
            if "Locked" not in pre_df.columns:
                pre_df["Locked"] = False
            res_df = df[~mask_pre].copy()

            df_disp_local = st.session_state["df_disp_editor"].copy()
            df_disp_local["informantes"]    = df_disp_local["informantes"].astype(str).str.strip().str.upper()
            df_disp_local["disponibilidade"]= df_disp_local["disponibilidade"].astype(str).str.strip().str.lower()
            available = sorted(df_disp_local[df_disp_local["disponibilidade"]=="sim"]
                               ["informantes"].dropna().unique().tolist())
            informantes_emails = dict(zip(df_disp_local["informantes"], df_disp_local["email"]))

            if not available:
                raise ValueError("Nenhum informante disponível.")

            modelo = st.session_state.get("modelo_dist","A/B por órgão (Opção 1: A adicional + B base)")
            only_locked_map_local = st.session_state.get("only_locked_map", {})

            if modelo.startswith("A/B"):
                pool_b = [x for x in st.session_state.get("pool_b_sel",[]) if x in available] or available[:]
                pool_a = [x for x in st.session_state.get("pool_a_sel",[]) if x in available]
                informantes_grupo_b = [s.upper() for s in pool_b]
                informantes_grupo_a = [s.upper() for s in pool_a] or informantes_grupo_b[:]
                origens_especiais_  = ["SEC EST POLICIA MILITAR","SEC EST DEFESA CIVIL"]
            else:
                pool = [x for x in st.session_state.get("pool_unico_sel",[]) if x in available] or available[:]
                informantes_grupo_a = []
                informantes_grupo_b = [s.upper() for s in pool]
                origens_especiais_  = []

            rules_list = []
            if rules_df is not None and not rules_df.empty:
                tmp = rules_df.copy()
                tmp["Informante"]     = tmp["Informante"].map(lambda v: str(v).strip().upper() if pd.notna(v) else "")
                tmp["Grupo Natureza"] = tmp["Grupo Natureza"].map(lambda v: "(QUALQUER)" if str(v).strip()=="" else str(v).strip().upper())
                tmp["Orgão Origem"]   = tmp["Orgão Origem"].map(lambda v: "(QUALQUER)" if str(v).strip()=="" else str(v).strip().upper())
                for _, r in tmp[tmp["Informante"] != ""].iterrows():
                    rules_list.append({
                        "Informante":    r["Informante"],
                        "Grupo Natureza":r["Grupo Natureza"],
                        "Orgão Origem":  r["Orgão Origem"],
                        "Exclusiva?":    bool(r.get("Exclusiva?", False))
                    })

            df_pool = res_df.copy().sort_values("Dias no Orgão", ascending=False).reset_index(drop=True)
            df_pool["Informante"] = ""
            if "Locked" not in df_pool.columns:
                df_pool["Locked"] = False

            assigned_by_rules, rem_after_rules = _apply_routing_rules(
                df_pool, rules_list, whitelist_pairs,
                only_locked_map=only_locked_map_local
            )
            distributed_df = _redistribute(
                rem_after_rules,
                informantes_grupo_a, informantes_grupo_b, origens_especiais_,
                whitelist_pairs, only_locked_map=only_locked_map_local
            )

            res_final = pd.concat([assigned_by_rules, distributed_df], ignore_index=True)
            res_final["Informante"] = res_final["Informante"].astype(str).str.strip().str.upper()

            if (res_final["Informante"].str.strip() == "").any():
                st.warning("Há processos sem informante (SEM_CANDIDATO).")

            res_final["Critério"]       = res_final.apply(calcula_criterio, axis=1)
            res_final["CustomPriority"] = res_final["Critério"].apply(lambda x: priority_map.get(x,4))
            res_final = res_final.sort_values(
                ["Informante","CustomPriority","Dias no Orgão"], ascending=[True,True,False]
            ).reset_index(drop=True).drop(columns=["CustomPriority"], errors="ignore")

            df_prev_map = None
            if prev_file is not None:
                try:
                    df_prev_raw = pd.read_excel(prev_file)
                    df_prev_raw.columns = [c.strip() for c in df_prev_raw.columns]
                    if {"Processo","Informante"}.issubset(df_prev_raw.columns):
                        df_prev_map = df_prev_raw[["Processo","Informante"]].copy()
                    else:
                        st.warning("Planilha anterior sem colunas 'Processo' e 'Informante'.")
                except Exception as e:
                    st.warning(f"Falha ao ler planilha anterior: {e}")

            if "Locked" not in res_final.columns:
                res_final["Locked"] = False

            diag_df = None
            if df_prev_map is not None:
                res_final, diag_df = _stick_previous_topN_STRICT(
                    res_final, df_prev_map, whitelist_pairs,
                    only_locked_map=only_locked_map_local, top_n_per_inf=200
                )
                res_final["Critério"]       = res_final.apply(calcula_criterio, axis=1)
                res_final["CustomPriority"] = res_final["Critério"].apply(lambda x: priority_map.get(x,4))
                res_final = res_final.sort_values(
                    ["Informante","CustomPriority","Dias no Orgão"], ascending=[True,True,False]
                ).reset_index(drop=True).drop(columns=["CustomPriority"], errors="ignore")

            pre_geral_fn = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            pre_geral_b  = to_excel_bytes(pre_df)
            res_geral_fn = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            res_geral_b  = to_excel_bytes(
                res_final.drop(columns=["Descrição Informação","Funcionário Informação"], errors="ignore")
            )

            # Individuais — Pré
            pre_individual_files = {}
            pre_local = pre_df.copy()
            pre_local["Informante"]    = pre_local["Informante"].astype(str).str.strip().str.upper()
            pre_local["Critério"]      = pre_local.apply(calcula_criterio, axis=1)
            pre_local["CustomPriority"]= pre_local["Critério"].apply(lambda x: priority_map.get(x,4))
            for inf in pre_local["Informante"].dropna().str.upper().unique():
                if not inf: continue
                di = pre_local[pre_local["Informante"].str.upper() == inf].copy()
                if di.empty: continue
                lk = di[_locked_mask(di)].copy(); nl = di[~_locked_mask(di)].copy()
                lk["ForaWhitelist"] = False
                nl["ForaWhitelist"] = nl.apply(
                    lambda r: not _accepts(inf, r["Orgão Origem"], r["Grupo Natureza"], whitelist_pairs), axis=1)
                di2 = pd.concat([lk, nl], ignore_index=True)
                di2 = di2.sort_values(["CustomPriority","Dias no Orgão"], ascending=[True,False]).head(200)
                di2 = di2.drop(columns=["CustomPriority"], errors="ignore")
                pre_individual_files[inf] = to_excel_bytes(di2)

            # Individuais — Principal Top-200
            res_individual_files = {}
            for inf, di in _apply_prevention_top200(res_final, df_prev_map, whitelist_pairs).items():
                res_individual_files[inf] = to_excel_bytes(di)

            return (pre_geral_fn, pre_geral_b, res_geral_fn, res_geral_b,
                    pre_individual_files, res_individual_files,
                    informantes_emails, diag_df)

        # Lê whitelist do session_state no momento do clique
        whitelist_pairs_exec = _read_whitelist_pairs()
        # Normaliza "(qualquer)" → string vazia (curinga interno)
        whitelist_pairs_exec = {
            inf: [(nat if nat != "(qualquer)" else "",
                   org if org != "(qualquer)" else "")
                  for nat, org in pairs]
            for inf, pairs in whitelist_pairs_exec.items()
        }

        try:
            (pre_geral_fn, pre_geral_b, res_geral_fn, res_geral_b,
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
            st.dataframe(diag_df.sort_values("Informante").reset_index(drop=True),
                         use_container_width=True)

        st.download_button("Baixar Planilha Geral PRE-ATRIBUÍDA",
                           data=pre_geral_b, file_name=pre_geral_fn,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Baixar Planilha Geral PRINCIPAL",
                           data=res_geral_b, file_name=res_geral_fn,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### Planilhas Individuais — Pré-Atribuídos")
        for inf, b in pre_individual_files.items():
            st.download_button(f"Baixar para {inf} (Pré-Atribuído)", data=b,
                file_name=f"{inf.replace(' ','_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### Planilhas Individuais — Principal (com prevenção Top-200)")
        for inf, b in res_individual_files.items():
            st.download_button(f"Baixar para {inf} (Principal)", data=b,
                file_name=f"{inf.replace(' ','_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if not test_mode:
            managers_list = [e.strip() for e in managers_emails.split(",") if e.strip()]
            if managers_list:
                all_ind = {}
                for inf, b in pre_individual_files.items():
                    all_ind[f"{inf.replace(' ','_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"] = b
                for inf, b in res_individual_files.items():
                    all_ind[f"{inf.replace(' ','_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"] = b
                zip_fn = f"{numero}_planilhas_individuais_{datetime.now().strftime('%Y%m%d')}.zip"
                send_email_with_multiple_attachments(
                    managers_list,
                    "Planilhas Gerais e Individuais de Processos",
                    "Prezado(a) Gestor(a),\n\nSeguem anexas as planilhas:\n"
                    "- Geral Pré-Atribuídos\n- Geral Principal\n- ZIP individuais\n\n"
                    "Atenciosamente,\nGestão da 3ª CAP",
                    [(pre_geral_b, pre_geral_fn),
                     (res_geral_b, res_geral_fn),
                     (create_zip_from_dict(all_ind), zip_fn)]
                )

            if modo_envio == "Produção - Gestores e Informantes":
                for inf in set(list(pre_individual_files) + list(res_individual_files)):
                    email_destino = informantes_emails.get(inf.upper(), "")
                    if email_destino:
                        ap = pre_individual_files.get(inf)
                        ar = res_individual_files.get(inf)
                        send_email_with_two_attachments(
                            email_destino,
                            f"Distribuição de Processos - {inf}",
                            "Prezado(a) Informante,\n\nSeguem as planilhas:\n"
                            "• Pré-Atribuídos\n• Principais\n\n"
                            "Atenciosamente,\nGestão da 3ª CAP",
                            ap, f"{inf.replace(' ','_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx" if ap else None,
                            ar, f"{inf.replace(' ','_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"    if ar else None,
                        )

        st.session_state.numero = numero

else:
    st.info("Carregue os quatro arquivos exigidos para habilitar a execução.")
