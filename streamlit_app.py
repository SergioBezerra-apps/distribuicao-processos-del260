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
    smtp_password = 'dimwpnhowxxeqbes'
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
    smtp_password = 'dimwpnhowxxeqbes'
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
# Funções auxiliares NOVAS
# =============================================================================
def _accepts(inf, orgao, natureza,
             filtros_grupo_natureza, filtros_orgao_origem):
    """
    True se *inf* aceita (órgão, natureza) conforme filtros da interface.
    """
    grupos_ok = filtros_grupo_natureza.get(inf, [])
    orgaos_ok = filtros_orgao_origem.get(inf, [])
    if grupos_ok and natureza not in grupos_ok:
        return False
    if orgaos_ok and orgao not in orgaos_ok:
        return False
    return True


def _redistribute(df_unassigned, informantes_ordem,
                  filtros_grupo_natureza, filtros_orgao_origem,
                  exclusive_mode, exclusive_orgao_map):
    """
    Redispõe processos sem destino, mantendo round‑robin por natureza e
    respeitando exclusividades válidas.
    """
    if df_unassigned.empty:
        return df_unassigned

    rr_indices = {gn: 0 for gn in df_unassigned["Grupo Natureza"].unique()}
    rows = []

    for _, row in df_unassigned.iterrows():
        natureza = row["Grupo Natureza"]
        orgao = row["Orgão Origem"]

        # 1) Se houver exclusividade válida
        candidatos = []
        if exclusive_mode and exclusive_orgao_map.get(orgao):
            inf_exc = exclusive_orgao_map[orgao]
            if _accepts(inf_exc, orgao, natureza,
                        filtros_grupo_natureza, filtros_orgao_origem):
                candidatos = [inf_exc]
        else:
            # 2) Todos que aceitam
            for inf in informantes_ordem:
                if _accepts(inf, orgao, natureza,
                            filtros_grupo_natureza, filtros_orgao_origem):
                    candidatos.append(inf)

        if candidatos:
            idx = rr_indices[natureza] % len(candidatos)
            row["Informante"] = candidatos[idx]
            rr_indices[natureza] += 1
        else:
            row["Informante"] = ""          # continua sem destino
        rows.append(row)

    return pd.DataFrame(rows)


# =============================================================================
# Interface principal
# =============================================================================

st.title("Distribuição de Processos da Del. 260")
st.markdown("### Faça o upload dos arquivos e configure a distribuição.")

if "numero" not in st.session_state:
    st.session_state.numero = "184"

uploaded_files = st.file_uploader(
    "Carregar os arquivos: processos.xlsx, processosmanter.xlsx, observacoes.xlsx e disponibilidade_equipe.xlsx",
    type=["xlsx"],
    accept_multiple_files=True
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

    # === Interface de atribuição exclusiva ===
    st.markdown("### Atribuição exclusiva de Orgão Origem (opcional)")
    exclusive_mode = st.checkbox("Atribuir cada Orgão Origem a apenas um informante?", value=False)
    exclusive_orgao_map = {}
    if exclusive_mode:
        for orgao in orgaos_origem_options:
            exclusive_orgao_map[orgao] = st.selectbox(
                f"Selecione o informante responsável exclusivamente por '{orgao}'",
                options=["(não atribuir exclusivamente)"] + list(informantes_principais),
                key=f"selectbox_exclusivo_{orgao.replace(' ', '_')}"
            )

    st.markdown("### Filtros de Grupo Natureza e Orgão Origem para Processos Principais (por informante)")
    for inf in informantes_principais:
        filtros_grupo_natureza[inf] = st.multiselect(
            f"Selecione Grupo(s) de Natureza para {inf} (deixe vazio para não filtrar):",
            options=grupo_natureza_options,
            key=f"grupo_natureza_{inf.replace(' ', '_')}"
        )
        filtros_orgao_origem[inf] = st.multiselect(
            f"Selecione Orgão(s) Origem para {inf} (deixe vazio para não filtrar):",
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
            exclusive_mode=False
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
            mask_residual = ~mask_preassigned
            res_df = df[mask_residual].copy()
        
            # 3. Distribuição original por grupo A/B (SEM FILTRO)
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
        
            # Grupo A: órgãos especiais
            df_grupo_a = res_df[res_df["Orgão Origem"].isin(origens_especiais)].copy()
            df_grupo_b = res_df[~res_df["Orgão Origem"].isin(origens_especiais)].copy()
        
            # Round-robin por informante dentro de cada grupo
            df_grupo_a = df_grupo_a.sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            df_grupo_b = df_grupo_b.sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
            if informantes_grupo_a:
                df_grupo_a["Informante"] = [informantes_grupo_a[i % len(informantes_grupo_a)] for i in range(len(df_grupo_a))]
            if informantes_grupo_b:
                df_grupo_b["Informante"] = [informantes_grupo_b[i % len(informantes_grupo_b)] for i in range(len(df_grupo_b))]
        
            res_assigned = pd.concat([df_grupo_a, df_grupo_b], ignore_index=True)
        
            # 4. Atribuição exclusiva de órgão (após round-robin)
            if exclusive_mode and exclusive_orgao_map:
                for orgao, inf in exclusive_orgao_map.items():
                    if inf != "(não atribuir exclusivamente)":
                        res_assigned.loc[res_assigned["Orgão Origem"] == orgao, "Informante"] = inf
        
            # 5. Cálculo de critério e ordenação para prioridade
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
            res_assigned = res_assigned.sort_values(by=["Informante", "CustomPriority", "Dias no Orgão"],
                                                    ascending=[True, True, False]).reset_index(drop=True)
            res_assigned = res_assigned.drop(columns=["CustomPriority"])
                    # ===== FILTROS POR INFORMANTE + REDISTRIBUIÇÃO AUTOMÁTICA =====
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
                    rej["Informante"] = ""  # volta ao pool
                    rejeitados_parts.append(rej)
                    aceitos_df = pd.concat(aceitos_parts, ignore_index=True)
                    unassigned_df = pd.concat(rejeitados_parts, ignore_index=True)
        
                    informantes_ordem = list(df_disp["informantes"].str.upper())
                    redistribuidos_df = _redistribute(
                        unassigned_df, informantes_ordem,
                        filtros_grupo_natureza, filtros_orgao_origem,
                        exclusive_mode, exclusive_orgao_map
                    )
                # ELIMINA QUALQUER UM QUE NÃO FOI DISTRIBUÍDO
                redistribuidos_df = redistribuidos_df[redistribuidos_df["Informante"] != ""]
        
                # Resultado consolidado
                res_final = pd.concat([aceitos_df, redistribuidos_df], ignore_index=True)
        
        
                
            # 6. Geração das planilhas gerais
            pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            pre_geral_bytes = to_excel_bytes(pre_df)
            res_geral_filename = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            # 7. Planilha “principal” usa o dataframe já filtrado e redistribuído
            res_geral_bytes = to_excel_bytes(
                res_final.drop(columns=["Descrição Informação", "Funcionário Informação"], errors="ignore")
            )

        
            # 8. Planilhas individuais por informante
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
        
            res_individual_files = {}
            for inf in res_final["Informante"].dropna().unique():
                df_inf = res_final[res_final["Informante"] == inf].copy()
                grupos_escolhidos = filtros_grupo_natureza.get(inf, [])
                orgaos_escolhidos = filtros_orgao_origem.get(inf, [])
                if grupos_escolhidos:
                    df_inf = df_inf[df_inf["Grupo Natureza"].isin(grupos_escolhidos)]
                if orgaos_escolhidos:
                    df_inf = df_inf[df_inf["Orgão Origem"].isin(orgaos_escolhidos)]
                df_inf["Critério"] = df_inf.apply(calcula_criterio, axis=1)
                df_inf["CustomPriority"] = df_inf["Critério"].apply(lambda x: priority_map.get(x, 4))
                df_inf = df_inf.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
                df_inf = df_inf.drop(columns=["CustomPriority"])
                df_inf = df_inf.head(200)
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
            exclusive_orgao_map=exclusive_orgao_map, exclusive_mode=exclusive_mode
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

        st.markdown("### Planilhas Individuais - Principal")
        for inf, file_bytes in res_individual_files.items():
            filename_inf = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                f"Baixar para {inf} (Principal)", 
                data=file_bytes, 
                file_name=filename_inf,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ----- Envio de e-mails (igual ao seu padrão) -----
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
