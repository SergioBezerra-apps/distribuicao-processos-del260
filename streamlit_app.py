import streamlit as st
import pandas as pd
import io
from datetime import datetime
import smtplib
from email.message import EmailMessage
from xlsxwriter.utility import xl_col_to_name

# =============================================================================
# Função para envio de e-mail com anexos (modo Produção)
# =============================================================================
def send_email_with_attachments(to_emails, subject, body, attachment_bytes, filename):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    smtp_username = 'seuemail@gmail.com'       # <-- Substitua pelo seu e-mail
    smtp_password = 'senhadeaplicativo'          # <-- Substitua pela sua senha de aplicativo

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_username
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)
    msg.add_attachment(
        attachment_bytes,
        maintype='application',
        subtype='octet-stream',
        filename=filename
    )
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10) as server:
            server.set_debuglevel(1)  # ajuste para 0 se preferir menos mensagens
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            st.info(f"E-mail enviado para: {to_emails}")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail para {to_emails}: {e}")

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
            color_list = [
                "#FFC7CE", "#C6EFCE", "#FFEB9C", "#9CC3E6",
                "#D9D2E9", "#FCE4D6", "#D0E0E3", "#E2EFDA"
            ]
            unique_values = df["Grupo Natureza"].dropna().unique()
            color_mapping = {val: color_list[i % len(color_list)] for i, val in enumerate(unique_values)}
            for value, color in color_mapping.items():
                fmt = workbook.add_format({'bg_color': color, 'font_color': '#000000'})
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '==',
                    'value': f'"{value}"',
                    'format': fmt
                })
    return output.getvalue()

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
    
    # Normaliza as colunas de interesse:
    # Converte "Descrição Informação" para minúsculas e remove espaços;
    # "Funcionário Informação" é mantido com espaços removidos.
    df["Descrição Informação"] = df["Descrição Informação"].astype(str).str.strip().str.lower()
    df["Funcionário Informação"] = df["Funcionário Informação"].astype(str).str.strip()
    
    # Merge com o arquivo de observações
    df_obs = pd.read_excel(obs_file)
    df_obs.columns = df_obs.columns.str.strip()
    df = pd.merge(df, df_obs[["Processo", "Obs", "Data Obs"]], on="Processo", how="left")
    
    # Converte datas e atualiza "Obs" e "Data Obs"
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
    # Pré-Atribuídos: processos em que "Descrição Informação" está em ["em elaboração", "concluída"]
    # E "Funcionário Informação" não está vazio.
    mask_preassigned = df["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df["Funcionário Informação"] != "")
    pre_df = df[mask_preassigned].copy()
    pre_df["Informante"] = pre_df["Funcionário Informação"]
    
    # Residual: Complemento dos pré-atribuídos (todos os processos que não foram pré-atribuídos)
    mask_residual = ~mask_preassigned
    res_df = df[mask_residual].copy()
    
    # --- Distribuição dos Processos Residual ---
    # Lê o arquivo de disponibilidade e obtém informantes disponíveis e seus e-mails
    df_disp = pd.read_excel(disp_file)
    df_disp.columns = df_disp.columns.str.strip()
    df_disp["disponibilidade"] = df_disp["disponibilidade"].str.lower()
    df_disp = df_disp[df_disp["disponibilidade"] == "sim"].copy()
    # Converte os nomes dos informantes para maiúsculas para padronização
    informantes_emails = dict(zip(df_disp["informantes"].str.upper(), df_disp["email"]))
    
    # Grupos de informantes (em maiúsculas)
    informantes_grupo_a = ["ALESSANDRO RIBEIRO RIOS", "ANDRE LUIZ BREIA", "ROSANE CESAR DE CARVALHO SCHLOSSER", "ANNA PAULA CYMERMAN"]
    informantes_grupo_b = ["LUCIA MARIA FELIPE DA SILVA", "MONICA ARANHA GOMES DO NASCIMENTO", "RODRIGO SILVEIRA BARRETO", "JOSÉ CARLOS NUNES"]
    informantes_grupo_a = [inf for inf in informantes_grupo_a if inf in informantes_emails]
    informantes_grupo_b = [inf for inf in informantes_grupo_b if inf in informantes_emails]
    
    # Separa os processos residuais em dois grupos com base no "Orgão Origem"
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
    
    # --- Geração das Planilhas Gerais ---
    pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
    pre_geral_bytes = to_excel_bytes(pre_df)
    
    res_geral_filename = f"{numero}_planilha_geral_residual_{datetime.now().strftime('%Y%m%d')}.xlsx"
    res_geral_bytes = to_excel_bytes(res_assigned)
    
    # --- Geração das Planilhas Individuais ---
    # Para pré-atribuídos (por informante)
    pre_individual_files = {}
    for inf in pre_df["Informante"].dropna().unique():
        df_inf = pre_df[pre_df["Informante"] == inf].copy()
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
        df_inf["Critério"] = df_inf.apply(calcula_criterio, axis=1)
        priority_map = {
            "01 Mais de cinco anos de autuado": 0,
            "02 A completar 5 anos de autuado": 1,
            "03 Mais de 5 meses na 3CAP": 2,
            "04 Data da carga": 3
        }
        df_inf["CustomPriority"] = df_inf["Critério"].apply(lambda x: priority_map.get(x, 4))
        df_inf = df_inf.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
        df_inf = df_inf.drop(columns=["CustomPriority"])
        filename_inf = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
        pre_individual_files[inf] = to_excel_bytes(df_inf)
    
    # Para os processos residuais (por informante)
    res_individual_files = {}
    for inf in res_assigned["Informante"].dropna().unique():
        df_inf = res_assigned[res_assigned["Informante"] == inf].copy()
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
        df_inf["Critério"] = df_inf.apply(calcula_criterio, axis=1)
        priority_map = {
            "01 Mais de cinco anos de autuado": 0,
            "02 A completar 5 anos de autuado": 1,
            "03 Mais de 5 meses na 3CAP": 2,
            "04 Data da carga": 3
        }
        df_inf["CustomPriority"] = df_inf["Critério"].apply(lambda x: priority_map.get(x, 4))
        df_inf = df_inf.sort_values(by=["CustomPriority", "Dias no Orgão"], ascending=[True, False])
        df_inf = df_inf.drop(columns=["CustomPriority"])
        filename_inf = f"{inf.replace(' ', '_')}_{numero}_residual_{datetime.now().strftime('%Y%m%d')}.xlsx"
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
st.title("Distribuição de processos da Del. 260 - Interface Gráfica")
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

modo = st.radio("Selecione o modo:", options=["Teste", "Produção"])
test_mode = (modo == "Teste")
if test_mode:
    st.info("Modo Teste: Nenhum e-mail será enviado; as planilhas serão disponibilizadas para download.")
else:
    st.info("Modo Produção: Serão enviados e-mails para gestores (planilhas gerais) e para informantes (planilhas individuais).")
st.markdown(f"**Modo selecionado:** {modo}")

managers_emails = st.text_input(
    "E-mails dos gestores/revisores (separados por vírgula):", 
    value="sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com"
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
        
        st.download_button(
            "Baixar Planilha Geral PRE-ATRIBUÍDA", 
            data=pre_geral_bytes, 
            file_name=pre_geral_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            "Baixar Planilha Geral RESIDUAL", 
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
        
        st.markdown("### Planilhas Individuais - Residual")
        for inf, file_bytes in res_individual_files.items():
            filename_inf = f"{inf.replace(' ', '_')}_{numero}_residual_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                f"Baixar para {inf} (Residual)", 
                data=file_bytes, 
                file_name=filename_inf,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        if test_mode:
            st.info("Modo Teste: Nenhum e-mail foi enviado.")
        else:
            managers_list = [e.strip() for e in managers_emails.split(",") if e.strip()]
            if managers_list:
                subject_pre = "Planilha Geral de Processos Pré-Atribuídos"
                body_pre = "Segue em anexo a planilha geral com os processos pré-atribuídos."
                send_email_with_attachments(managers_list, subject_pre, body_pre, pre_geral_bytes, pre_geral_filename)
                subject_res = "Planilha Geral de Processos Residual"
                body_res = "Segue em anexo a planilha geral com os processos residuais distribuídos."
                send_email_with_attachments(managers_list, subject_res, body_res, res_geral_bytes, res_geral_filename)
            
            for inf in set(list(pre_individual_files.keys()) + list(res_individual_files.keys())):
                # Os nomes dos informantes foram padronizados para maiúsculas no dicionário de e-mails
                email_destino = informantes_emails.get(inf.upper(), "")
                if email_destino:
                    if inf in pre_individual_files:
                        subject_inf = f"Distribuição de Processos - {inf} (Pré-Atribuído)"
                        body_inf = "Segue em anexo a planilha com os processos pré-atribuídos para você."
                        filename_inf = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        send_email_with_attachments([email_destino], subject_inf, body_inf, pre_individual_files[inf], filename_inf)
                    if inf in res_individual_files:
                        subject_inf = f"Distribuição de Processos - {inf} (Residual)"
                        body_inf = "Segue em anexo a planilha com os processos residuais distribuídos para você."
                        filename_inf = f"{inf.replace(' ', '_')}_{numero}_residual_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        send_email_with_attachments([email_destino], subject_inf, body_inf, res_individual_files[inf], filename_inf)
        
        st.session_state.numero = numero + 1
    else:
        st.error("Por favor, faça o upload dos três arquivos necessários: processos.xlsx, observacoes.xlsx e disponibilidade_equipe.xlsx.")
