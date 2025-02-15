import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import smtplib
from email.message import EmailMessage
from xlsxwriter.utility import xl_col_to_name

# =============================================================================
# Função para envio de e-mail com múltiplos anexos (usada para gestores)
# =============================================================================
def send_email_with_multiple_attachments(to_emails, subject, body, attachments):
    """
    Envia um e-mail com vários anexos.
    :param to_emails: lista de endereços de e-mail
    :param subject: assunto do e-mail
    :param body: corpo do e-mail
    :param attachments: lista de tuplas (attachment_bytes, filename)
    """
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

# =============================================================================
# Função para envio de e-mail com dois anexos (usada para informantes)
# =============================================================================
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

# =============================================================================
# Função para converter um DataFrame em bytes (para download) com formatação condicional
# =============================================================================
def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha")
        workbook = writer.book
        worksheet = writer.sheets["Planilha"]
        # Formatação condicional para a coluna 'Grupo Natureza'
        if "Grupo Natureza" in df.columns:
            col_index = df.columns.get_loc("Grupo Natureza")
            col_letter = xl_col_to_name(col_index)
            last_row = len(df) + 1  # Cabeçalho na linha 1; dados a partir da linha 2
            cell_range = f'{col_letter}2:{col_letter}{last_row}'
            # Lista de cores de fonte mais escuras
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

# =============================================================================
# Função para criar um arquivo ZIP a partir de um dicionário de arquivos
# =============================================================================
def create_zip_from_dict(file_dict):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for filename, file_bytes in file_dict.items():
            zip_file.writestr(filename, file_bytes)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# =============================================================================
# Função que executa a lógica de distribuição
# =============================================================================
def run_distribution(processos_file, obs_file, disp_file, numero):
    # Lê o arquivo de processos e filtra apenas os do tipo "Principal"
    df = pd.read_excel(processos_file)
    df.columns = df.columns.str.strip()
    df = df[df["Tipo Processo"] == "Principal"]
    
    # Seleciona as colunas necessárias
    required_cols = [
        "Processo", "Grupo Natureza", "Orgão Origem", "Dias no Orgão",
        "Tempo TCERJ", "Data Última Carga", "Descrição Informação", "Funcionário Informação"
    ]
    df = df[required_cols]
    
    # Normaliza as colunas de interesse
    df["Descrição Informação"] = df["Descrição Informação"].astype(str).str.strip().str.lower()
    df["Funcionário Informação"] = df["Funcionário Informação"].astype(str).str.strip()
    
    # Merge com o arquivo de observações
    df_obs = pd.read_excel(obs_file)
    df_obs.columns = df_obs.columns.str.strip()
    df = pd.merge(df, df_obs[["Processo", "Obs", "Data Obs"]], on="Processo", how="left")
    
    # Atualiza as colunas "Obs" e "Data Obs" com base na comparação entre datas
    df["Data Última Carga"] = pd.to_datetime(df["Data Última Carga"], errors="coerce")
    df["Data Obs"] = pd.to_datetime(df["Data Obs"], errors="coerce")
    def update_obs(row):
        if pd.notna(row["Data Obs"]) and pd.notna(row["Data Última Carga"]) and row["Data Obs"] > row["Data Última Carga"]:
            return pd.Series([row["Obs"], row["Data Obs"]])
        else:
            return pd.Series(["", ""])
    df[["Obs", "Data Obs"]] = df.apply(update_obs, axis=1)
    df = df.drop(columns=["Data Última Carga"])
    
    # --- Separação dos Processos ---
    # Pré-Atribuídos: processos em que "descrição informação" é "em elaboração" ou "concluída"
    # e "funcionário informação" não está vazio.
    mask_preassigned = df["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df["Funcionário Informação"] != "")
    pre_df = df[mask_preassigned].copy()
    pre_df["Informante"] = pre_df["Funcionário Informação"]
    
    # Processos restantes (Principais)
    mask_residual = ~mask_preassigned
    res_df = df[mask_residual].copy()
    
    # --- Distribuição dos Processos Principais ---
    df_disp = pd.read_excel(disp_file)
    df_disp.columns = df_disp.columns.str.strip()
    df_disp["disponibilidade"] = df_disp["disponibilidade"].str.lower()
    df_disp = df_disp[df_disp["disponibilidade"] == "sim"].copy()
    informantes_emails = dict(zip(df_disp["informantes"].str.upper(), df_disp["email"]))
    
    # Grupos de informantes (em maiúsculas)
    informantes_grupo_a = ["ALESSANDRO RIBEIRO RIOS", "ANDRE LUIZ BREIA", "ROSANE CESAR DE CARVALHO SCHLOSSER", "ANNA PAULA CYMERMAN"]
    informantes_grupo_b = ["LUCIA MARIA FELIPE DA SILVA", "MONICA ARANHA GOMES DO NASCIMENTO", "RODRIGO SILVEIRA BARRETO", "JOSÉ CARLOS NUNES"]
    informantes_grupo_a = [inf for inf in informantes_grupo_a if inf in informantes_emails]
    informantes_grupo_b = [inf for inf in informantes_grupo_b if inf in informantes_emails]
    
    # Distribui os processos principais entre os informantes
    origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]
    df_grupo_a = res_df[res_df["Orgão Origem"].isin(origens_especiais)].copy()
    df_grupo_b = res_df[~res_df["Orgão Origem"].isin(origens_especiais)].copy()
    
    df_grupo_a = df_grupo_a.sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
    df_grupo_b = df_grupo_b.sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
    
    if informantes_grupo_a:
        df_grupo_a["Informante"] = [informantes_grupo_a[i % len(informantes_grupo_a)] for i in range(len(df_grupo_a))]
    if informantes_grupo_b:
        df_grupo_b["Informante"] = [informantes_grupo_b[i % len(informantes_grupo_b)] for i in range(len(df_grupo_b))]
    
    res_assigned = pd.concat([df_grupo_a, df_grupo_b], ignore_index=True)
    
    # --- Cálculo do Critério para ordenação ---
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
    
    # --- Geração das Planilhas Gerais ---
    pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
    pre_geral_bytes = to_excel_bytes(pre_df)
    
    res_geral_filename = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
    # Removendo colunas indesejadas para a planilha geral dos processos principais
    res_assigned = res_assigned.drop(columns=["Descrição Informação", "Funcionário Informação"], errors='ignore')
    cols = res_assigned.columns.tolist()
    if "Critério" in cols:
        cols.remove("Critério")
        cols.insert(2, "Critério")
        res_assigned = res_assigned[cols]
    res_geral_bytes = to_excel_bytes(res_assigned)
    
    # --- Geração das Planilhas Individuais ---
    # Cada informante receberá, em sua planilha, no máximo 200 registros.
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
    for inf in res_assigned["Informante"].dropna().unique():
        df_inf = res_assigned[res_assigned["Informante"] == inf].copy()
        df_inf = df_inf.drop(columns=["Descrição Informação", "Funcionário Informação"], errors='ignore')
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

# =============================================================================
# Configuração de número (mantido em session_state)
# =============================================================================
if "numero" not in st.session_state:
    st.session_state.numero = 184

# =============================================================================
# Interface Gráfica (Streamlit)
# =============================================================================
st.title("Distribuição de Processos da Del. 260")
st.markdown("### Faça o upload dos arquivos e configure a distribuição.")

uploaded_files = st.file_uploader(
    "Carregar os arquivos: processos.xlsx, observacoes.xlsx e disponibilidade_equipe.xlsx",
    type=["xlsx"],
    accept_multiple_files=True
)

files_dict = {}
if uploaded_files:
    for file in uploaded_files:
        fname = file.name.lower()
        if fname == "processos.xlsx":
            files_dict["processos"] = file
        elif fname in ["observacoes.xlsx", "obervacoes.xlsx"]:
            files_dict["observacoes"] = file
        elif fname == "disponibilidade_equipe.xlsx":
            files_dict["disponibilidade"] = file

numero = st.number_input(
    "Qual a numeração dessa planilha de distribuição?",
    value=st.session_state.numero,
    step=1
)

# Seleção do modo geral: Teste ou Produção
modo = st.radio("Selecione o modo:", options=["Teste", "Produção"])
test_mode = (modo == "Teste")
if test_mode:
    st.info("Modo Teste: Nenhum e-mail será enviado; as planilhas serão disponibilizadas para download.")
else:
    st.info("Modo Produção: Os e-mails serão enviados conforme o modo de envio selecionado.")
st.markdown(f"**Modo selecionado:** {modo}")

# Se estiver em modo Produção, pergunta se o envio será para Gestores e Informantes ou Apenas Gestores
modo_envio = None
if not test_mode:
    modo_envio = st.radio("Selecione o modo de envio:", options=["Produção - Gestores e Informantes", "Produção - Apenas Gestores"])
    st.markdown(f"**Modo de envio selecionado:** {modo_envio}")

# Entrada dos e-mails dos gestores
managers_emails = st.text_input(
    "E-mails dos gestores/revisores (separados por vírgula):", 
    value="annapc@tcerj.tc.br, fabiovf@tcerj.tc.br, sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com"
)

if st.button("Executar Distribuição"):
    required_keys = ["processos", "observacoes", "disponibilidade"]
    if all(key in files_dict for key in required_keys):
        processos_file = files_dict["processos"]
        obs_file = files_dict["observacoes"]
        disp_file = files_dict["disponibilidade"]

        (pre_geral_filename, pre_geral_bytes,
         res_geral_filename, res_geral_bytes,
         pre_individual_files, res_individual_files,
         informantes_emails) = run_distribution(processos_file, obs_file, disp_file, numero)

        st.success("Distribuição executada com sucesso!")
        
        # Disponibiliza as planilhas gerais para download (gestores)
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
        
        # Disponibiliza as planilhas individuais para download
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
        
        # Envio de e-mails apenas se não estivermos em modo teste
        if not test_mode:
            managers_list = [e.strip() for e in managers_emails.split(",") if e.strip()]
            if managers_list:
                # Compacta todas as planilhas individuais em um arquivo ZIP
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
            
            # Envio de e-mails individuais para os informantes, somente se o modo de envio permitir
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
                            "[Equipe de Distribuição de Processos]"
                        )
                        send_email_with_two_attachments(email_destino, subject_inf, body_inf, attachment_pre, filename_pre, attachment_res, filename_res)
        
        st.session_state.numero = numero + 1
    else:
        st.error("Por favor, faça o upload dos três arquivos necessários: processos.xlsx, observacoes.xlsx e disponibilidade_equipe.xlsx.")
