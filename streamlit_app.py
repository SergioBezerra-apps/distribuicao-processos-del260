import os
import io
import zipfile
from datetime import datetime
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
# Core helpers
# =============================================================================

def _accepts(inf, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem) -> bool:
    """Whitelist: vazio = aceita tudo. Se marcado, aceita somente o que foi marcado (E para natureza/órgão)."""
    grupos_ok = filtros_grupo_natureza.get(inf, [])
    orgaos_ok = filtros_orgao_origem.get(inf, [])
    if grupos_ok and natureza not in grupos_ok:
        return False
    if orgaos_ok and orgao not in orgaos_ok:
        return False
    return True

def _apply_routing_rules(df_pool: pd.DataFrame, rules: list, filtros_grupo_natureza, filtros_orgao_origem):
    """
    Regras (Natureza, Órgão) → Informante, com 'Exclusiva?' opcional.
    Escolhe a mais específica; empate: a primeira. Exclusiva? => Locked=True.
    """
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
                  informantes_grupo_a, informantes_grupo_b, origens_especiais,
                  filtros_grupo_natureza, filtros_orgao_origem,
                  only_locked_map) -> pd.DataFrame:
    """Round-robin por natureza, respeitando grupos A/B, whitelist e ignorando 'somente exclusivos'."""
    if df_unassigned.empty:
        return df_unassigned
    df_unassigned = df_unassigned.copy()
    rr_indices = {gn: 0 for gn in df_unassigned["Grupo Natureza"].unique()}
    out_rows = []
    for _, row in df_unassigned.iterrows():
        natureza = row["Grupo Natureza"]
        orgao = row["Orgão Origem"]
        informantes_do_grupo = informantes_grupo_a if orgao in origens_especiais else informantes_grupo_b

        candidatos = []
        for inf in informantes_do_grupo:
            if bool(only_locked_map.get(inf, False)):
                continue
            if _accepts(inf, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
                candidatos.append(inf)

        if candidatos:
            idx = rr_indices[natureza] % len(candidatos)
            row["Informante"] = candidatos[idx]
            rr_indices[natureza] += 1
        else:
            row["Informante"] = ""
        out_rows.append(row)
    return pd.DataFrame(out_rows)

# ======================
# Critério / Prioridade
# ======================

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

# ============
# Locked util
# ============

def _locked_mask(df: pd.DataFrame) -> pd.Series:
    """Series booleana alinhada indicando Locked."""
    if "Locked" in df.columns:
        return df["Locked"].astype(bool).reindex(df.index).fillna(False)
    return pd.Series(False, index=df.index)

# =============================================================================
# PREVENÇÃO — listas individuais (<=200) com whitelist só p/ não-Locked
# =============================================================================

def _apply_prevention_top200(res_final: pd.DataFrame, df_prev_map: pd.DataFrame,
                             filtros_grupo_natureza, filtros_orgao_origem) -> dict:
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

    for inf in base["Informante"].dropna().unique():
        df_inf = base[base["Informante"] == inf].copy()

        # Whitelist apenas nos não-Locked
        locked_part = df_inf[_locked_mask(df_inf)].copy()
        nonlocked   = df_inf[~_locked_mask(df_inf)].copy()

        grupos_escolhidos = filtros_grupo_natureza.get(inf, [])
        orgaos_escolhidos = filtros_orgao_origem.get(inf, [])
        if grupos_escolhidos:
            nonlocked = nonlocked[nonlocked["Grupo Natureza"].isin(grupos_escolhidos)]
        if orgaos_escolhidos:
            nonlocked = nonlocked[nonlocked["Orgão Origem"].isin(orgaos_escolhidos)]

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
# STICK ESTRITO — cola apenas o que já era do informante (Top-200),
# whitelist NÃO barra itens Locked, e respeita travas/only-locked do destino.
# =============================================================================

def _stick_previous_topN_STRICT(res_df: pd.DataFrame, df_prev_map: pd.DataFrame,
                                filtros_grupo_natureza, filtros_orgao_origem,
                                only_locked_map: dict, top_n_per_inf: int = 200):
    if res_df.empty or df_prev_map is None or df_prev_map.empty:
        return res_df, pd.DataFrame(columns=["Informante","colados","bloq_locked_outro","bloq_only_locked","bloq_whitelist"])

    prev = df_prev_map.copy()
    prev.columns = [c.strip() for c in prev.columns]
    if not {"Processo", "Informante"}.issubset(set(prev.columns)):
        return res_df, pd.DataFrame(columns=["Informante","colados","bloq_locked_outro","bloq_only_locked","bloq_whitelist"])

    prev["Processo"]   = prev["Processo"].astype(str).str.strip()
    prev["Informante"] = prev["Informante"].astype(str).str.strip().str.upper()

    df = res_df.copy()
    if "Locked" not in df.columns:
        df["Locked"] = False

    df["Critério"] = df.apply(calcula_criterio, axis=1)
    df["CustomPriority"] = df["Critério"].apply(lambda x: priority_map.get(x, 4))

    diag_rows = []
    prev_map = dict(zip(prev["Processo"], prev["Informante"]))

    for inf in sorted(df["Informante"].astype(str).str.upper().dropna().unique()):
        prev_set = set(prev[prev["Informante"] == inf]["Processo"].astype(str))

        bloco = df[df["Processo"].astype(str).isin(prev_set)].copy()
        if bloco.empty:
            diag_rows.append({"Informante": inf, "colados": 0, "bloq_locked_outro": 0, "bloq_only_locked": 0, "bloq_whitelist": 0})
            continue

        # Whitelist: itens Locked entram SEM checar whitelist; demais passam por _accepts
        bloco["ok_whitelist"] = bloco.apply(
            lambda r: True if bool(r.get("Locked", False)) else _accepts(
                inf, r["Orgão Origem"], r["Grupo Natureza"],
                filtros_grupo_natureza, filtros_orgao_origem
            ),
            axis=1
        )

        # Ordenar por prioridade e limitar Top-N depois de whitelist
        bloco = bloco[bloco["ok_whitelist"]].copy()
        bloco = bloco.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False]).head(top_n_per_inf)

        if bloco.empty:
            diag_rows.append({"Informante": inf, "colados": 0, "bloq_locked_outro": 0, "bloq_only_locked": 0, "bloq_whitelist": int(len(prev_set))})
            continue

        sl_only = bool(only_locked_map.get(inf, False))

        def blocked_locked_outro(row):
            # Item Locked mas travado para OUTRO informante
            return bool(row.get("Locked", False) and str(row.get("Informante","")).upper() != inf)

        def blocked_only_locked(row):
            # Destino exige only-locked e o item NÃO é Locked
            return bool(sl_only and not row.get("Locked", False))

        bloco["bloq_locked_outro"] = bloco.apply(blocked_locked_outro, axis=1)
        bloco["bloq_only_locked"]  = bloco.apply(blocked_only_locked, axis=1)

        colaveis = bloco[~(bloco["bloq_locked_outro"] | bloco["bloq_only_locked"])].copy()

        if not colaveis.empty:
            df.loc[df["Processo"].astype(str).isin(colaveis["Processo"].astype(str)), "Informante"] = inf

        diag_rows.append({
            "Informante": inf,
            "colados": int(len(colaveis)),
            "bloq_locked_outro": int(bloco["bloq_locked_outro"].sum()),
            "bloq_only_locked": int(bloco["bloq_only_locked"].sum()),
            "bloq_whitelist": 0  # whitelist já aplicada antes do head
        })

    df = df.drop(columns=["CustomPriority"], errors="ignore")
    diag_df = pd.DataFrame(diag_rows, columns=["Informante","colados","bloq_locked_outro","bloq_only_locked","bloq_whitelist"])
    return df, diag_df

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
    modo_envio = st.radio("Modo de envio:", options=["Produção - Gestores e Informantes", "Produção - Apenas Gestores"], horizontal=False)

managers_emails = st.text_input(
    "E-mails dos gestores/revisores (separados por vírgula):",
    value="annapc@tcerj.tc.br, fabiovf@tcerj.tc.br, sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com"
)

# -----------------------------------------------------------------------------
# Somente após carregar os 4 arquivos principais
# -----------------------------------------------------------------------------

if all(k in files_dict for k in ["processos", "processosmanter", "observacoes", "disponibilidade"]):
    # ====== Leitura e normalização UPPER ======
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

    df_disp = pd.read_excel(files_dict["disponibilidade"])
    df_disp.columns = df_disp.columns.str.strip()
    if "disponibilidade" in df_disp.columns:
        df_disp = df_disp[df_disp["disponibilidade"].astype(str).str.lower() == "sim"].copy()
    df_disp["informantes"] = df_disp["informantes"].astype(str).str.strip().str.upper()

    informantes_principais = sorted(df_disp["informantes"].dropna().unique())
    grupo_natureza_options = sorted(df_proc["Grupo Natureza"].dropna().unique())
    orgaos_origem_options  = sorted(df_proc["Orgão Origem"].dropna().unique())

    # ========== Filtros (whitelist) por informante ==========
    st.markdown("### Filtros (whitelist) por informante")
    st.caption("Vazio = aceita tudo. Se marcar Natureza e Órgão, exige interseção (E).")
    filtros_grupo_natureza, filtros_orgao_origem = {}, {}
    for inf in informantes_principais:
        filtros_grupo_natureza[inf] = st.multiselect(
            f"Naturezas aceitas — {inf}",
            options=grupo_natureza_options, key=f"gn_{inf.replace(' ','_')}"
        )
        filtros_orgao_origem[inf] = st.multiselect(
            f"Órgãos aceitos — {inf}",
            options=orgaos_origem_options, key=f"org_{inf.replace(' ','_')}"
        )

    # ========== Preferências: somente exclusivos (Locked) ==========
    st.markdown("### Preferências por informante")
    st.caption("Marque para que o informante receba apenas itens de regras exclusivas (Locked).")
    only_locked_map = st.session_state.get("only_locked_map", {})
    for inf in informantes_principais:
        key = f"only_locked_{inf.replace(' ', '_')}"
        val = st.checkbox(f"{inf}: receber apenas itens exclusivos (Locked)?", value=only_locked_map.get(inf, False), key=key)
        only_locked_map[inf] = val
    st.session_state["only_locked_map"] = only_locked_map

    # ========== Regras ==========
    st.markdown("### Regras de roteamento por (Natureza, Órgão) → Informante")
    st.caption("Use '(QUALQUER)' como curinga. Marque 'Exclusiva?' para reservar o par ao informante indicado.")
    natureza_opts = ["(QUALQUER)"] + grupo_natureza_options
    orgao_opts    = ["(QUALQUER)"] + orgaos_origem_options
    inf_opts      = list(informantes_principais)

    if "rules_state_v2" not in st.session_state:
        st.session_state["rules_state_v2"] = pd.DataFrame(columns=["Informante", "Grupo Natureza", "Orgão Origem", "Exclusiva?"])

    rules_state = st.data_editor(
        st.session_state["rules_state_v2"],
        num_rows="dynamic",
        use_container_width=True,
        key="rules_editor_v2",
        column_config={
            "Informante": st.column_config.SelectboxColumn("Informante", options=inf_opts, required=True),
            # <- corrigido para natureza_opts
            "Grupo Natureza": st.column_config.SelectboxColumn("Grupo Natureza", options=natureza_opts, required=True),
            "Orgão Origem": st.column_config.SelectboxColumn("Orgão Origem", options=orgao_opts, required=True),
            "Exclusiva?": st.column_config.CheckboxColumn("Exclusiva?")
        }
    )
    st.session_state["rules_state_v2"] = rules_state

    # ========== Botão principal ==========
    if st.button("Executar Distribuição"):
        def run_distribution(processos_file, processosmanter_file, obs_file, disp_file, numero,
                             filtros_grupo_natureza, filtros_orgao_origem,
                             rules_df=None, prev_file=None):

            # 1) Leitura base
            df = pd.read_excel(processos_file); df.columns = df.columns.str.strip()
            for col in ["Grupo Natureza", "Orgão Origem"]:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip().str.upper()
            df["Processo"] = df["Processo"].astype(str)

            df_manter = pd.read_excel(processosmanter_file); df_manter.columns = df_manter.columns.str.strip()
            processos_validos = df_manter["Processo"].dropna().astype(str).unique()
            df = df[df["Processo"].isin(processos_validos)]
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
                else:
                    return pd.Series(["", ""])
            df[["Obs", "Data Obs"]] = df.apply(update_obs, axis=1)
            df = df.drop(columns=["Data Última Carga"])

            # 2) Pré-atribuídos e principais
            mask_pre = df["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df["Funcionário Informação"] != "")
            pre_df = df[mask_pre].copy(); pre_df["Informante"] = pre_df["Funcionário Informação"]
            res_df = df[~mask_pre].copy()

            # garante Locked em pre_df
            if "Locked" not in pre_df.columns:
                pre_df["Locked"] = False

            # 3) Disponibilidade / e-mails / A-B
            df_disp_local = pd.read_excel(disp_file); df_disp_local.columns = df_disp_local.columns.str.strip()
            if "disponibilidade" in df_disp_local.columns:
                df_disp_local = df_disp_local[df_disp_local["disponibilidade"].astype(str).str.lower() == "sim"].copy()
            df_disp_local["informantes"] = df_disp_local["informantes"].astype(str).str.strip().str.upper()
            informantes_emails = dict(zip(df_disp_local["informantes"], df_disp_local["email"]))

            informantes_grupo_a = [
                "ALESSANDRO RIBEIRO RIOS", "ANDRE LUIZ BREIA",
                "ROSANE CESAR DE CARVALHO SCHLOSSER", "ANNA PAULA CYMERMAN"
            ]
            informantes_grupo_b = [
                "LUCIA MARIA FELIPE DA SILVA", "MONICA ARANHA GOMES DO NASCIMENTO",
                "RODRIGO SILVEIRA BARRETO", "JOSÉ CARLOS NUNES"
            ]
            available = set(df_disp_local["informantes"])
            informantes_grupo_a = [inf for inf in informantes_grupo_a if inf in available]
            informantes_grupo_b = [inf for inf in informantes_grupo_b if inf in available]
            origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]

            # 4) Round-robin base (A/B)
            df_grupo_a = res_df[res_df["Orgão Origem"].isin(origens_especiais)].copy().sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            df_grupo_b = res_df[~res_df["Orgão Origem"].isin(origens_especiais)].copy().sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            if informantes_grupo_a:
                df_grupo_a["Informante"] = [informantes_grupo_a[i % len(informantes_grupo_a)] for i in range(len(df_grupo_a))]
            if informantes_grupo_b:
                df_grupo_b["Informante"] = [informantes_grupo_b[i % len(informantes_grupo_b)] for i in range(len(df_grupo_b))]
            res_assigned = pd.concat([df_grupo_a, df_grupo_b], ignore_index=True)

            # 5) Critério/Prioridade
            res_assigned["Critério"] = res_assigned.apply(calcula_criterio, axis=1)
            res_assigned["CustomPriority"] = res_assigned["Critério"].apply(lambda x: priority_map.get(x, 4))
            res_assigned = res_assigned.sort_values(by=["Informante", "CustomPriority", "Dias no Orgão"],
                                                    ascending=[True, True, False]).reset_index(drop=True)

            # 6) APLICA REGRAS
            rules_list = []
            rules_df_in = rules_df

            def _norm(x):
                return (str(x).strip().str.upper() if pd.notna(x) else "")

            if rules_df_in is not None and not rules_df_in.empty:
                tmp = rules_df_in.copy()
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

            assigned_by_rules, rem_after_rules = _apply_routing_rules(
                res_assigned, rules_list, filtros_grupo_natureza, filtros_orgao_origem
            )
            res_assigned2 = pd.concat([assigned_by_rules, rem_after_rules], ignore_index=True)
            if "Locked" not in res_assigned2.columns:
                res_assigned2["Locked"] = False

            # 7) FILTROS + only-locked
            only_locked_map_local = st.session_state.get("only_locked_map", {})
            aceitos_parts, rejeitados_parts = [], []
            for inf in res_assigned2["Informante"].dropna().unique():
                df_inf = res_assigned2[res_assigned2["Informante"] == inf].copy()
                if "Locked" not in df_inf.columns:
                    df_inf["Locked"] = False
                only_locked = bool(only_locked_map_local.get(inf, False))
                if only_locked:
                    mask_keep = df_inf["Locked"]
                else:
                    mask_keep = df_inf.apply(
                        lambda r: True if r["Locked"] else _accepts(
                            inf, r["Orgão Origem"], r["Grupo Natureza"], filtros_grupo_natureza, filtros_orgao_origem
                        ),
                        axis=1
                    )
                aceitos_parts.append(df_inf[mask_keep])
                rej = df_inf[~mask_keep].copy()
                rej["Informante"] = ""
                rejeitados_parts.append(rej)

            aceitos_df = pd.concat(aceitos_parts, ignore_index=True) if aceitos_parts else pd.DataFrame(columns=res_assigned2.columns)
            unassigned_df = pd.concat(rejeitados_parts, ignore_index=True) if rejeitados_parts else pd.DataFrame(columns=res_assigned2.columns)

            # 8) Redistribuição do restante (não Locked)
            informantes_grupo_a_upper = [s.upper() for s in informantes_grupo_a]
            informantes_grupo_b_upper = [s.upper() for s in informantes_grupo_b]
            aceitos_df["Informante"] = aceitos_df["Informante"].astype(str).str.upper()
            unassigned_df["Informante"] = unassigned_df["Informante"].astype(str).str.upper()

            redistribuidos_df = _redistribute(
                unassigned_df, informantes_grupo_a_upper, informantes_grupo_b_upper, origens_especiais,
                filtros_grupo_natureza, filtros_orgao_origem,
                only_locked_map=only_locked_map_local
            )
            if not redistribuidos_df.empty:
                mask_bad = redistribuidos_df["Informante"].map(lambda x: bool(only_locked_map_local.get(x, False)))
                if mask_bad.any():
                    redistribuidos_df = redistribuidos_df[~mask_bad]

            redistribuidos_df = redistribuidos_df[redistribuidos_df["Informante"] != ""]
            res_final = pd.concat([aceitos_df, redistribuidos_df], ignore_index=True)

            # 8.1) PREVENÇÃO (STICK ESTRITO) — cola até 200 do que já era de cada informante
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

            # garantir Locked em res_final
            if "Locked" not in res_final.columns:
                res_final["Locked"] = False

            diag_df = None
            if df_prev_map is not None:
                res_final, diag_df = _stick_previous_topN_STRICT(
                    res_final, df_prev_map,
                    filtros_grupo_natureza, filtros_orgao_origem,
                    only_locked_map=only_locked_map_local, top_n_per_inf=200
                )
                # 8.2) REORDENA após o stick para respeitar o Critério (e dias)
                res_final = res_final.copy()
                res_final["Critério"] = res_final.apply(calcula_criterio, axis=1)
                res_final["CustomPriority"] = res_final["Critério"].apply(lambda x: priority_map.get(x, 4))
                res_final = res_final.sort_values(
                    by=["Informante", "CustomPriority", "Dias no Orgão"],
                    ascending=[True, True, False]
                ).reset_index(drop=True)
                res_final = res_final.drop(columns=["CustomPriority"], errors="ignore")

            # 9) Planilhas gerais
            pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            pre_geral_bytes = to_excel_bytes(pre_df)
            res_geral_filename = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            res_geral_bytes = to_excel_bytes(res_final.drop(columns=["Descrição Informação", "Funcionário Informação"], errors="ignore"))

            # 10) Planilhas individuais — Pré (whitelist só para não-Locked)
            def build_pre_individuals():
                pre_individual_files = {}
                pre_df_local = pre_df.copy()
                pre_df_local["Critério"] = pre_df_local.apply(calcula_criterio, axis=1)
                pre_df_local["CustomPriority"] = pre_df_local["Critério"].apply(lambda x: priority_map.get(x, 4))
                for inf in pre_df_local["Informante"].dropna().unique():
                    df_inf = pre_df_local[pre_df_local["Informante"] == inf].copy()

                    locked_part = df_inf[_locked_mask(df_inf)].copy()
                    nonlocked   = df_inf[~_locked_mask(df_inf)].copy()

                    grupos_escolhidos = filtros_grupo_natureza.get(inf, [])
                    orgaos_escolhidos = filtros_orgao_origem.get(inf, [])
                    if grupos_escolhidos:
                        nonlocked = nonlocked[nonlocked["Grupo Natureza"].isin(grupos_escolhidos)]
                    if orgaos_escolhidos:
                        nonlocked = nonlocked[nonlocked["Orgão Origem"].isin(orgaos_escolhidos)]

                    df_inf = pd.concat([locked_part, nonlocked], ignore_index=True)
                    df_inf = df_inf.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
                    df_inf = df_inf.drop(columns=["CustomPriority"])
                    df_inf = df_inf.head(200)
                    pre_individual_files[inf] = to_excel_bytes(df_inf)
                return pre_individual_files

            pre_individual_files = build_pre_individuals()

            # 11) Planilhas individuais — Principal (prevenção Top-200, whitelist só p/ não-Locked)
            res_individual_files = {}
            top200_dict = _apply_prevention_top200(
                res_final, df_prev_map, filtros_grupo_natureza, filtros_orgao_origem
            )
            for inf, df_inf in top200_dict.items():
                filename_inf = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
                res_individual_files[inf] = to_excel_bytes(df_inf)

            return (pre_geral_filename, pre_geral_bytes,
                    res_geral_filename, res_geral_bytes,
                    pre_individual_files, res_individual_files,
                    informantes_emails, diag_df)

        # ---- Executa ----
        (pre_geral_filename, pre_geral_bytes,
         res_geral_filename, res_geral_bytes,
         pre_individual_files, res_individual_files,
         informantes_emails, diag_df) = run_distribution(
            files_dict["processos"], files_dict["processosmanter"],
            files_dict["observacoes"], files_dict["disponibilidade"],
            numero, filtros_grupo_natureza, filtros_orgao_origem,
            rules_df=st.session_state.get("rules_state_v2"), prev_file=prev_file
        )

        st.success("Distribuição executada com sucesso!")

        # Diagnóstico do STICK
        if isinstance(diag_df, pd.DataFrame) and not diag_df.empty:
            st.markdown("#### Diagnóstico da Prevenção (Stick Estrito Top-200)")
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
                        send_email_with_two_attachments(email_destino, subject_inf, body_inf, attachment_pre, filename_pre, attachment_res, filename_res)

        st.session_state.numero = numero
else:
    st.info("Carregue os quatro arquivos exigidos para habilitar a execução.")
