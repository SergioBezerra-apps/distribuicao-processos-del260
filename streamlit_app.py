import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import smtplib
from email.message import EmailMessage
from xlsxwriter.utility import xl_col_to_name

# =============================================================================
# Funções utilitárias
# =============================================================================

def send_email_with_multiple_attachments(to_emails, subject, body, attachments):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    smtp_username = 'sergiolbezerralj@gmail.com'
    smtp_password = 'dimwpnhowxxeqbes'  # SUGESTÃO: mover para st.secrets
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_username
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)
    for attachment_bytes, filename in attachments:
        msg.add_attachment(
            attachment_bytes,
            maintype='application',
            subtype='octet-stream',
            filename=filename
        )
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10) as server:
            server.set_debuglevel(1)
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            st.info(f"E-mail enviado para: {to_emails}")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail para {to_emails}: {e}")

def send_email_with_two_attachments(to_email, subject, body, attachment_pre, filename_pre, attachment_res, filename_res):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    smtp_username = 'sergiolbezerralj@gmail.com'
    smtp_password = 'dimwpnhowxxeqbes'  # SUGESTÃO: mover para st.secrets
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_username
    msg['To'] = to_email
    msg.set_content(body)
    if attachment_pre is not None:
        msg.add_attachment(
            attachment_pre,
            maintype='application',
            subtype='octet-stream',
            filename=filename_pre
        )
    if attachment_res is not None:
        msg.add_attachment(
            attachment_res,
            maintype='application',
            subtype='octet-stream',
            filename=filename_res
        )
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10) as server:
            server.set_debuglevel(1)
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            st.info(f"E-mail enviado para: {to_email}")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail para {to_email}: {e}")

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha")
        workbook = writer.book
        worksheet = writer.sheets["Planilha"]
        if "Grupo Natureza" in df.columns:
            col_index = df.columns.get_loc("Grupo Natureza")
            col_letter = xl_col_to_name(col_index)
            last_row = len(df) + 1
            cell_range = f'{col_letter}2:{col_letter}{last_row}'
            color_list = [
                "#990000", "#006600", "#996600", "#003366",
                "#660066", "#663300", "#003300", "#000066"
            ]
            unique_values = df["Grupo Natureza"].dropna().unique()
            color_mapping = {val: color_list[i % len(color_list)] for i, val in enumerate(unique_values)}
            for value, color in color_mapping.items():
                fmt = workbook.add_format({'font_color': color, 'bold': True})
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '==',
                    'value': f'"{value}"',
                    'format': fmt
                })
    return output.getvalue()

def create_zip_from_dict(file_dict):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for filename, file_bytes in file_dict.items():
            zip_file.writestr(filename, file_bytes)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# =============================================================================
# NOVO: Funções auxiliares para filtros/roteamento/redistribuição/prevenção
# =============================================================================

def _accepts(inf, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
    """True se *inf* aceita (órgão, natureza) conforme filtros da interface."""
    grupos_ok = filtros_grupo_natureza.get(inf, [])
    orgaos_ok = filtros_orgao_origem.get(inf, [])
    if grupos_ok and natureza not in grupos_ok:
        return False
    if orgaos_ok and orgao not in orgaos_ok:
        return False
    return True

def _apply_exclusive_rules(df_pool, rules, filtros_grupo_natureza, filtros_orgao_origem):
    """
    Aplica regras exclusivas (Natureza, Órgão) -> Informante.
    rules: lista de dicts com chaves: Informante, Grupo Natureza, Orgão Origem
           use "(qualquer)" como curinga.
    Retorna: df_assigned, df_remaining
    """
    if df_pool.empty or not rules:
        return pd.DataFrame(columns=df_pool.columns), df_pool

    df_pool = df_pool.copy()
    assigned_rows = []
    remaining_rows = []

    for _, row in df_pool.iterrows():
        natureza = str(row["Grupo Natureza"])
        orgao = str(row["Orgão Origem"])
        matched = False
        for r in rules:
            inf = r["Informante"]
            nat_rule = r["Grupo Natureza"]
            org_rule = r["Orgão Origem"]
            ok_nat = (nat_rule == "(qualquer)") or (nat_rule == natureza)
            ok_org = (org_rule == "(qualquer)") or (org_rule == orgao)
            if ok_nat and ok_org and _accepts(inf, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
                new_row = row.copy()
                new_row["Informante"] = inf
                assigned_rows.append(new_row)
                matched = True
                break
        if not matched:
            remaining_rows.append(row)

    df_assigned = pd.DataFrame(assigned_rows) if assigned_rows else pd.DataFrame(columns=df_pool.columns)
    df_remaining = pd.DataFrame(remaining_rows) if remaining_rows else pd.DataFrame(columns=df_pool.columns)
    return df_assigned, df_remaining

def _redistribute(df_unassigned, informantes_ordem,
                  filtros_grupo_natureza, filtros_orgao_origem,
                  exclusive_mode, exclusive_orgao_map,
                  informantes_grupo_a, informantes_grupo_b, origens_especiais):
    """
    Redispõe processos sem destino, mantendo round-robin por natureza,
    respeitando grupos A/B e exclusividades válidas.
    """
    if df_unassigned.empty:
        return df_unassigned

    rr_indices = {gn: 0 for gn in df_unassigned["Grupo Natureza"].unique()}
    rows = []

    for _, row in df_unassigned.iterrows():
        natureza = row["Grupo Natureza"]
        orgao = row["Orgão Origem"]

        # Escolha do grupo
        informantes_do_grupo = informantes_grupo_a if orgao in origens_especiais else informantes_grupo_b

        candidatos = []
        # Exclusividade por órgão (fallback se regra explícita não capturou)
        if exclusive_mode and orgao in exclusive_orgao_map:
            inf_exc = exclusive_orgao_map[orgao]
            if _accepts(inf_exc, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
                candidatos = [inf_exc]
            else:
                for inf in informantes_do_grupo:
                    if inf != inf_exc and _accepts(inf, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
                        candidatos.append(inf)
        else:
            for inf in informantes_do_grupo:
                if _accepts(inf, orgao, natureza, filtros_grupo_natureza, filtros_orgao_origem):
                    candidatos.append(inf)

        # Round-robin
        if candidatos:
            idx = rr_indices[natureza] % len(candidatos)
            row["Informante"] = candidatos[idx]
            rr_indices[natureza] += 1
        else:
            row["Informante"] = ""  # permanece sem destino
        rows.append(row)

    return pd.DataFrame(rows)

def _apply_prevention_top200(res_final, df_prev_map, calcula_criterio, priority_map, filtros_grupo_natureza, filtros_orgao_origem):
    """
    Garante, por informante, que até 200 processos que *permanecem* e já estavam com o mesmo informante
    sejam priorizados nas listas individuais. Completa com os demais por prioridade.
    df_prev_map: DataFrame com colunas ['Processo','Informante'] da semana anterior.
    Retorna: dict {inf: DataFrame (<=200)}, construído já com a seleção final.
    """
    out = {}
    # Pré-cálculo do critério e ordenação geral para consistência na “complementação”
    base = res_final.copy()
    base["Critério"] = base.apply(calcula_criterio, axis=1)
    base["CustomPriority"] = base["Critério"].apply(lambda x: priority_map.get(x, 4))
    base = base.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])

    prev_map = {}
    if df_prev_map is not None and not df_prev_map.empty:
        # normaliza nomes
        df_prev_map = df_prev_map.copy()
        df_prev_map["Informante"] = df_prev_map["Informante"].astype(str).str.strip()
        df_prev_map["Processo"] = df_prev_map["Processo"].astype(str).str.strip()
        prev_map = dict(zip(df_prev_map["Processo"], df_prev_map["Informante"]))

    for inf in base["Informante"].dropna().unique():
        df_inf = base[base["Informante"] == inf].copy()
        # aplica filtros voluntários do inf na etapa da lista individual
        grupos_escolhidos = filtros_grupo_natureza.get(inf, [])
        orgaos_escolhidos = filtros_orgao_origem.get(inf, [])
        if grupos_escolhidos:
            df_inf = df_inf[df_inf["Grupo Natureza"].isin(grupos_escolhidos)]
        if orgaos_escolhidos:
            df_inf = df_inf[df_inf["Orgão Origem"].isin(orgaos_escolhidos)]

        # preferidos (carry-over): permanecem e eram do mesmo informante
        if prev_map:
            df_inf["preferido"] = df_inf["Processo"].astype(str).map(lambda p: 1 if prev_map.get(str(p)) == inf else 0)
        else:
            df_inf["preferido"] = 0

        preferidos = df_inf[df_inf["preferido"] == 1]
        nao_pref = df_inf[df_inf["preferido"] == 0]

        # mantém ordem por prioridade dentro de cada grupo
        preferidos = preferidos.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
        nao_pref = nao_pref.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])

        # monta até 200
        df_top = pd.concat([preferidos, nao_pref], ignore_index=True).head(200)
        df_top = df_top.drop(columns=["CustomPriority", "preferido"], errors="ignore")
        out[inf] = df_top

    return out

# =============================================================================
# Interface principal
# =============================================================================

st.title("Distribuição de Processos da Del. 260")
st.markdown("### Faça o upload dos arquivos e configure a distribuição.")

if "numero" not in st.session_state:
    st.session_state.numero = "184"

uploaded_files = st.file_uploader(
    "Carregar: processos.xlsx, processosmanter.xlsx, observacoes.xlsx, disponibilidade_equipe.xlsx",
    type=["xlsx"],
    accept_multiple_files=True
)

# NOVO: base preventiva opcional (geral principal da semana anterior, mínima com colunas Processo e Informante)
prev_file = st.file_uploader(
    "Opcional: carregar a PLANILHA GERAL PRINCIPAL da semana anterior (para prevenção Top 200). Deve conter colunas 'Processo' e 'Informante'.",
    type=["xlsx"],
    accept_multiple_files=False
)

files_dict = {}
for file in uploaded_files or []:
    fname = file.name.lower()
    if fname == "processos.xlsx":
        files_dict["processos"] = file
    elif fname == "processosmanter.xlsx":
        files_dict["processosmanter"] = file
    elif fname in ["observacoes.xlsx", "obervacoes.xlsx"]:
        files_dict["observacoes"] = file
    elif fname == "disponibilidade_equipe.xlsx":
        files_dict["disponibilidade"] = file

numero = st.text_input("Qual a numeração dessa planilha de distribuição?", value=st.session_state.numero)

modo = st.radio("Selecione o modo:", options=["Teste", "Produção"])
test_mode = (modo == "Teste")
if test_mode:
    st.info("Modo Teste: Nenhum e-mail será enviado; as planilhas serão disponibilizadas para download.")
else:
    st.info("Modo Produção: Os e-mails serão enviados conforme o modo de envio selecionado.")
st.markdown(f"**Modo selecionado:** {modo}")

modo_envio = None
if not test_mode:
    modo_envio = st.radio("Selecione o modo de envio:", options=["Produção - Gestores e Informantes", "Produção - Apenas Gestores"])
    st.markdown(f"**Modo de envio selecionado:** {modo_envio}")

managers_emails = st.text_input(
    "E-mails dos gestores/revisores (separados por vírgula):",
    value="annapc@tcerj.tc.br, fabiovf@tcerj.tc.br, sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com"
)

# ------------------------------------------
# Filtros e seleção exclusiva
# ------------------------------------------

filtros_grupo_natureza = {}
filtros_orgao_origem = {}

informantes_principais = []
grupo_natureza_options = []
orgaos_origem_options = []
exclusive_mode = False
exclusive_orgao_map = {}

# NOVO: Regras exclusivas por (Natureza, Órgão) → Informante
st.markdown("### Regras exclusivas de roteamento (Natureza, Órgão) → Informante (opcional)")
st.caption("Ex.: APOSENTADORIA do Órgão X para Y, e PENSÃO do Órgão Z para Y. Use '(qualquer)' como curinga.")
rules_df = pd.DataFrame(columns=["Informante", "Grupo Natureza", "Orgão Origem"])
rules_state = st.session_state.get("rules_state", rules_df.copy())
rules_state = st.data_editor(
    rules_state,
    num_rows="dynamic",
    key="rules_editor",
    use_container_width=True,
    column_config={
        "Informante": st.column_config.TextColumn(help="Nome exatamente como em disponibilidade_equipe"),
        "Grupo Natureza": st.column_config.TextColumn(help="Ex.: APOSENTADORIA, PENSÃO, ou '(qualquer)'"),
        "Orgão Origem": st.column_config.TextColumn(help="Ex.: SEC EST ... ou '(qualquer)'")
    }
)
st.session_state["rules_state"] = rules_state

if all(key in files_dict for key in ["processos", "processosmanter", "observacoes", "disponibilidade"]):
    # Carrega DataFrames para montar as opções dos filtros
    df_proc = pd.read_excel(files_dict["processos"])
    df_proc.columns = df_proc.columns.str.strip()
    df_manter = pd.read_excel(files_dict["processosmanter"])
    df_manter.columns = df_manter.columns.str.strip()
    processos_validos = df_manter["Processo"].dropna().unique()
    df_proc = df_proc[df_proc["Processo"].isin(processos_validos)]
    df_proc = df_proc[df_proc["Tipo Processo"] == "Principal"]

    disp_file = files_dict["disponibilidade"]
    df_disp = pd.read_excel(disp_file)
    df_disp.columns = df_disp.columns.str.strip()
    df_disp = df_disp[df_disp["disponibilidade"].str.lower() == "sim"].copy()
    informantes_principais = sorted(df_disp["informantes"].dropna().unique())
    grupo_natureza_options = sorted(df_proc["Grupo Natureza"].dropna().unique())
    orgaos_origem_options = sorted(df_proc["Orgão Origem"].dropna().unique())

    # === Interface de atribuição exclusiva por Órgão (como já existia) ===
    st.markdown("### Atribuição exclusiva de Orgão Origem (opcional)")
    exclusive_mode = st.checkbox("Atribuir cada Orgão Origem a apenas um informante?", value=False)
    exclusive_orgao_map = {}
    if exclusive_mode:
        for orgao in orgaos_origem_options:
            inf_exclusivo = st.selectbox(
                f"Selecione o informante responsável exclusivamente por '{orgao}'",
                options=["(não atribuir exclusivamente)"] + list(informantes_principais),
                key=f"selectbox_exclusivo_{orgao.replace(' ', '_')}"
            )
            if inf_exclusivo != "(não atribuir exclusivamente)":
                exclusive_orgao_map[orgao] = inf_exclusivo

    st.markdown("### Filtros de Grupo Natureza e Orgão Origem para Processos Principais (por informante)")
    for inf in informantes_principais:
        filtros_grupo_natureza[inf] = st.multiselect(
            f"Grupo(s) de Natureza para {inf} (vazio = não filtra):",
            options=grupo_natureza_options,
            key=f"grupo_natureza_{inf.replace(' ', '_')}"
        )
        filtros_orgao_origem[inf] = st.multiselect(
            f"Orgão(s) Origem para {inf} (vazio = não filtra):",
            options=orgaos_origem_options,
            key=f"orgao_origem_{inf.replace(' ', '_')}"
        )

# ------------------------------------------
# Botão principal de execução
# ------------------------------------------

if st.button("Executar Distribuição"):
    required_keys = ["processos", "processosmanter", "observacoes", "disponibilidade"]
    if all(key in files_dict for key in required_keys):

        def run_distribution(
            processos_file,
            processosmanter_file,
            obs_file,
            disp_file,
            numero,
            filtros_grupo_natureza,
            filtros_orgao_origem,
            exclusive_orgao_map=None,
            exclusive_mode=False,
            rules_df=None,          # NOVO
            prev_file=None          # NOVO
        ):
            # 1. Leitura dos dados e pré-processamento
            df = pd.read_excel(processos_file)
            df.columns = df.columns.str.strip()
            df_manter = pd.read_excel(processosmanter_file)
            df_manter.columns = df_manter.columns.str.strip()
            processos_validos = df_manter["Processo"].dropna().unique()
            df = df[df["Processo"].isin(processos_validos)]
            df = df[df["Tipo Processo"] == "Principal"]

            required_cols = [
                "Processo", "Grupo Natureza", "Orgão Origem", "Dias no Orgão",
                "Tempo TCERJ", "Data Última Carga", "Descrição Informação", "Funcionário Informação"
            ]
            df = df[required_cols]
            df["Descrição Informação"] = df["Descrição Informação"].astype(str).str.strip().str.lower()
            df["Funcionário Informação"] = df["Funcionário Informação"].astype(str).str.strip()

            df_obs = pd.read_excel(obs_file)
            df_obs.columns = df_obs.columns.str.strip()
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

            # 2. Separação dos processos pré-atribuídos e dos principais
            mask_preassigned = df["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df["Funcionário Informação"] != "")
            pre_df = df[mask_preassigned].copy()
            pre_df["Informante"] = pre_df["Funcionário Informação"]
            res_df = df[~mask_preassigned].copy()

            # 3. Disponibilidade e grupos A/B
            df_disp = pd.read_excel(disp_file)
            df_disp.columns = df_disp.columns.str.strip()
            df_disp["disponibilidade"] = df_disp["disponibilidade"].str.lower()
            df_disp = df_disp[df_disp["disponibilidade"] == "sim"].copy()
            informantes_emails = dict(zip(df_disp["informantes"].str.upper(), df_disp["email"]))

            informantes_grupo_a = [
                "ALESSANDRO RIBEIRO RIOS", "ANDRE LUIZ BREIA",
                "ROSANE CESAR DE CARVALHO SCHLOSSER", "ANNA PAULA CYMERMAN"
            ]
            informantes_grupo_b = [
                "LUCIA MARIA FELIPE DA SILVA", "MONICA ARANHA GOMES DO NASCIMENTO",
                "RODRIGO SILVEIRA BARRETO", "JOSÉ CARLOS NUNES"
            ]
            informantes_grupo_a = [inf for inf in informantes_grupo_a if inf in informantes_emails]
            informantes_grupo_b = [inf for inf in informantes_grupo_b if inf in informantes_emails]
            origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]

            # 4. Round-robin base por grupo A/B
            df_grupo_a = res_df[res_df["Orgão Origem"].isin(origens_especiais)].copy().sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            df_grupo_b = res_df[~res_df["Orgão Origem"].isin(origens_especiais)].copy().sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            if informantes_grupo_a:
                df_grupo_a["Informante"] = [informantes_grupo_a[i % len(informantes_grupo_a)] for i in range(len(df_grupo_a))]
            if informantes_grupo_b:
                df_grupo_b["Informante"] = [informantes_grupo_b[i % len(informantes_grupo_b)] for i in range(len(df_grupo_b))]
            res_assigned = pd.concat([df_grupo_a, df_grupo_b], ignore_index=True)

            # 5. Exclusividade por Órgão (pós-round-robin)
            if exclusive_mode and exclusive_orgao_map:
                for orgao, inf in exclusive_orgao_map.items():
                    if inf != "(não atribuir exclusivamente)":
                        res_assigned.loc[res_assigned["Orgão Origem"] == orgao, "Informante"] = inf

            # 6. Critério de prioridade
            def calcula_criterio(row):
                if pd.isna(row["Processo"]) or row["Processo"] == "":
                    return ""
                elif row["Tempo TCERJ"] > 1765:
                    return "01 Mais de cinco anos de autuado"
                elif 1220 < row["Tempo TCERJ"] < 1765:
                    return "02 A completar 5 anos de autuado"
                elif row["Dias no Orgão"] >= 150:
                    return "03 Mais de 5 meses na 3CAP"
                else:
                    return "04 Data da carga"
            res_assigned["Critério"] = res_assigned.apply(calcula_criterio, axis=1)
            priority_map = {
                "01 Mais de cinco anos de autuado": 0,
                "02 A completar 5 anos de autuado": 1,
                "03 Mais de 5 meses na 3CAP": 2,
                "04 Data da carga": 3
            }
            res_assigned["CustomPriority"] = res_assigned["Critério"].apply(lambda x: priority_map.get(x, 4))
            res_assigned = res_assigned.sort_values(by=["Informante", "CustomPriority", "Dias no Orgão"], ascending=[True, True, False]).reset_index(drop=True)

            # 7. FILTROS por informante e REDISTRIBUIÇÃO
            aceitos_parts, rejeitados_parts = [], []
            for inf in res_assigned["Informante"].dropna().unique():
                df_inf = res_assigned[res_assigned["Informante"] == inf].copy()
                mask_keep = df_inf.apply(
                    lambda r: _accepts(
                        inf, r["Orgão Origem"], r["Grupo Natureza"],
                        filtros_grupo_natureza, filtros_orgao_origem
                    ), axis=1)
                aceitos_parts.append(df_inf[mask_keep])
                rej = df_inf[~mask_keep].copy()
                rej["Informante"] = ""
                rejeitados_parts.append(rej)
            aceitos_df = pd.concat(aceitos_parts, ignore_index=True) if aceitos_parts else pd.DataFrame(columns=res_assigned.columns)
            unassigned_df = pd.concat(rejeitados_parts, ignore_index=True) if rejeitados_parts else pd.DataFrame(columns=res_assigned.columns)

            informantes_ordem = list(df_disp["informantes"].str.upper())

            # 7.1 NOVO: aplica Regras exclusivas (Natureza, Órgão) -> Informante ANTES da redistribuição final
            rules_list = []
            if rules_df is not None and not rules_df.empty:
                # normaliza e filtra linhas válidas
                tmp = rules_df.fillna("").copy()
                for _, r in tmp.iterrows():
                    if str(r.get("Informante", "")).strip():
                        rules_list.append({
                            "Informante": str(r["Informante"]).strip(),
                            "Grupo Natureza": str(r.get("Grupo Natureza", "(qualquer)")).strip() or "(qualquer)",
                            "Orgão Origem": str(r.get("Orgão Origem", "(qualquer)")).strip() or "(qualquer)"
                        })

            assigned_by_rules, still_unassigned = _apply_exclusive_rules(
                unassigned_df, rules_list, filtros_grupo_natureza, filtros_orgao_origem
            )

            # 7.2 Redistribuição para o restante
            redistribuidos_df = _redistribute(
                still_unassigned, informantes_ordem,
                filtros_grupo_natureza, filtros_orgao_origem,
                exclusive_mode, exclusive_orgao_map,
                informantes_grupo_a, informantes_grupo_b, origens_especiais
            )
            redistribuidos_df = redistribuidos_df[redistribuidos_df["Informante"] != ""]

            # Resultado consolidado PRINCIPAL
            res_final = pd.concat([aceitos_df, assigned_by_rules, redistribuidos_df], ignore_index=True)

            # 8. Planilhas gerais
            pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            pre_geral_bytes = to_excel_bytes(pre_df)
            res_geral_filename = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            res_geral_bytes = to_excel_bytes(
                res_final.drop(columns=["Descrição Informação", "Funcionário Informação"], errors="ignore")
            )

            # 9. Planilhas individuais PRÉ (mantém sua lógica)
            def build_pre_individuals():
                pre_individual_files = {}
                for inf in pre_df["Informante"].dropna().unique():
                    df_inf = pre_df[pre_df["Informante"] == inf].copy()
                    df_inf["Critério"] = df_inf.apply(calcula_criterio, axis=1)
                    df_inf["CustomPriority"] = df_inf["Critério"].apply(lambda x: priority_map.get(x, 4))
                    df_inf = df_inf.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
                    df_inf = df_inf.drop(columns=["CustomPriority"])
                    df_inf = df_inf.head(200)
                    filename_inf = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    pre_individual_files[inf] = to_excel_bytes(df_inf)
                return pre_individual_files

            pre_individual_files = build_pre_individuals()

            # 10. NOVO: prevenção Top 200 por informante usando base anterior (se existir)
            df_prev_map = None
            if prev_file is not None:
                try:
                    df_prev_raw = pd.read_excel(prev_file)
                    # aceita nomes variados/uppercase
                    cols = [c.strip() for c in df_prev_raw.columns]
                    df_prev_raw.columns = cols
                    # tenta resolver nomes comuns
                    col_proc = next((c for c in cols if c.lower() == "processo"), None)
                    col_inf = next((c for c in cols if c.lower() == "informante"), None)
                    if col_proc and col_inf:
                        df_prev_map = df_prev_raw[[col_proc, col_inf]].rename(columns={col_proc: "Processo", col_inf: "Informante"})
                    else:
                        st.warning("Planilha anterior sem colunas 'Processo' e 'Informante'. Prevenção não aplicada.")
                except Exception as e:
                    st.warning(f"Falha ao ler planilha anterior: {e}")

            # Recalcula critérios (já calculados acima em res_final) e aplica prevenção para montar as INDIVIDUAIS PRINCIPAIS
            res_individual_files = {}
            top200_dict = _apply_prevention_top200(
                res_final, df_prev_map, calcula_criterio, priority_map, filtros_grupo_natureza, filtros_orgao_origem
            )
            for inf, df_inf in top200_dict.items():
                filename_inf = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
                res_individual_files[inf] = to_excel_bytes(df_inf)

            return (pre_geral_filename, pre_geral_bytes,
                    res_geral_filename, res_geral_bytes,
                    pre_individual_files, res_individual_files,
                    informantes_emails)

        # ---- Chamada principal ----
        processos_file = files_dict["processos"]
        processosmanter_file = files_dict["processosmanter"]
        obs_file = files_dict["observacoes"]
        disp_file = files_dict["disponibilidade"]

        (pre_geral_filename, pre_geral_bytes,
         res_geral_filename, res_geral_bytes,
         pre_individual_files, res_individual_files,
         informantes_emails) = run_distribution(
            processos_file, processosmanter_file, obs_file, disp_file, numero,
            filtros_grupo_natureza, filtros_orgao_origem,
            exclusive_orgao_map=exclusive_orgao_map, exclusive_mode=exclusive_mode,
            rules_df=st.session_state.get("rules_state"),
            prev_file=prev_file
        )

        st.success("Distribuição executada com sucesso!")

        st.download_button(
            "Baixar Planilha Geral PRE-ATRIBUÍDA",
            data=pre_geral_bytes,
            file_name=pre_geral_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            "Baixar Planilha Geral PRINCIPAL",
            data=res_geral_bytes,
            file_name=res_geral_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("### Planilhas Individuais - Pré-Atribuídos")
        for inf, file_bytes in pre_individual_files.items():
            filename_inf = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                f"Baixar para {inf} (Pré-Atribuído)",
                data=file_bytes,
                file_name=filename_inf,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.markdown("### Planilhas Individuais - Principal (com prevenção Top 200)")
        for inf, file_bytes in res_individual_files.items():
            filename_inf = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                f"Baixar para {inf} (Principal)",
                data=file_bytes,
                file_name=filename_inf,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ----- Envio de e-mails (inalterado) -----
        if not test_mode:
            managers_list = [e.strip() for e in managers_emails.split(",") if e.strip()]
            if managers_list:
                all_individual_files = {}
                for inf, file_bytes in pre_individual_files.items():
                    filename_ind = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    all_individual_files[filename_ind] = file_bytes
                for inf, file_bytes in res_individual_files.items():
                    filename_ind = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    all_individual_files[filename_ind] = file_bytes
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
                    "Segue em anexo as seguintes planilhas:\n\n"
                    "- Planilha Geral de Processos Pré-Atribuídos\n"
                    "- Planilha Geral de Processos Principais\n"
                    "- Arquivo ZIP contendo todas as planilhas individuais (Pré-Atribuídos e Principais)\n\n"
                    "Atenciosamente,\n"
                    "[Equipe de Distribuição de Processos]"
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
                            "Segue em anexo as planilhas referentes à distribuição de processos:\n\n"
                            "Processos Pré-Atribuídos:\n"
                            "Estes são os processos que já estavam vinculados a você antes da distribuição, "
                            "ou seja, os que já estavam com instrução processual em andamento ou concluída no sistema.\n\n"
                            "Processos Principais:\n"
                            "São os novos processos distribuídos entre os informantes disponíveis.\n\n"
                            "Caso tenha dúvidas, entre em contato.\n\n"
                            "Atenciosamente,\n"
                            "Gestão da 3ª CAP"
                        )
                        send_email_with_two_attachments(email_destino, subject_inf, body_inf, attachment_pre, filename_pre, attachment_res, filename_res)

        st.session_state.numero = numero

    else:
        st.error("Por favor, faça o upload dos quatro arquivos necessários: processos.xlsx, processosmanter.xlsx, observacoes.xlsx e disponibilidade_equipe.xlsx.")
