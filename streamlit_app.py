import streamlit as st
import pandas as pd
import io
from datetime import datetime
import smtplib
from email.message import EmailMessage
from xlsxwriter.utility import xl_col_to_name

# =============================================================================
# Função para envio de e-mail com anexos
# =============================================================================
def send_email_with_attachments(to_emails, subject, body, attachment_bytes, filename):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    smtp_username = 'sergiolbezerralj@gmail.com'
    smtp_password = 'dimwpnhowxxeqbes'  # Verifique se a senha de aplicativo está correta

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_username
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)
    msg.add_attachment(attachment_bytes,
                       maintype='application',
                       subtype='octet-stream',
                       filename=filename)
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=10) as server:
            server.set_debuglevel(1)
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            st.info("E-mail com anexos enviado com sucesso!")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")

# =============================================================================
# Função para converter um DataFrame em bytes (para download) com formatação condicional
# =============================================================================
def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Planilha")
        workbook = writer.book
        worksheet = writer.sheets["Planilha"]
        
        # Aplica formatação condicional para a coluna 'Grupo Natureza'
        if "Grupo Natureza" in df.columns:
            col_index = df.columns.get_loc("Grupo Natureza")
            col_letter = xl_col_to_name(col_index)
            last_row = len(df) + 1  # Cabeçalho na linha 1; dados a partir da linha 2
            cell_range = f'{col_letter}2:{col_letter}{last_row}'
            
            # Lista de cores pré-definidas (será utilizada ciclicamente)
            color_list = ["#FFC7CE", "#C6EFCE", "#FFEB9C", "#9CC3E6", "#D9D2E9", "#FCE4D6", "#D0E0E3", "#E2EFDA"]
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
def run_distribution(processos_file, obs_file, disp_file, numero, test_mode=True):
    # Lê o arquivo de processos
    df = pd.read_excel(processos_file)
    df.columns = df.columns.str.strip()
    
    # Filtra para manter apenas os processos do tipo "Principal"
    df = df[df["Tipo Processo"] == "Principal"]
    
    # Seleciona as colunas necessárias, garantindo que "Descrição Informação" seja incluída
    required_cols = [
        "Processo", "Grupo Natureza", "Orgão Origem", "Dias no Orgão", 
        "Tempo TCERJ", "Data Última Carga", "Descrição Informação", "Funcionário Informação"
    ]
    df = df[required_cols]
    
    # Lê o arquivo de disponibilidade e obtém os informantes disponíveis (disponibilidade == "sim")
    df_disp = pd.read_excel(disp_file)
    df_disp.columns = df_disp.columns.str.strip()
    df_disp["disponibilidade"] = df_disp["disponibilidade"].str.lower()
    available = df_disp[df_disp["disponibilidade"] == "sim"]["informantes"].tolist()
    
    # Pré-atribuição: para processos com "Descrição Informação" igual a "Em Elaboração" ou "Concluída",
    # atribui o informante indicado em "Funcionário Informação" se este estiver disponível.
    mask_preassigned = df["Descrição Informação"].isin(["Em Elaboração", "Concluída"])
    df.loc[mask_preassigned, "Informante"] = df.loc[mask_preassigned, "Funcionário Informação"].where(
        df.loc[mask_preassigned, "Funcionário Informação"].isin(available), ""
    )
    
    # Lê o arquivo de observações
    df_obs = pd.read_excel(obs_file)
    df_obs.columns = df_obs.columns.str.strip()
    
    # Faz o merge trazendo as colunas "Obs" e "Data Obs"
    df = pd.merge(df, df_obs[["Processo", "Obs", "Data Obs"]], on="Processo", how="left")
    
    # Converte as colunas de data
    df["Data Última Carga"] = pd.to_datetime(df["Data Última Carga"], errors="coerce")
    df["Data Obs"] = pd.to_datetime(df["Data Obs"], errors="coerce")
    
    # Atualiza Obs e Data Obs: se "Data Obs" for mais recente que "Data Última Carga", mantém os valores; caso contrário, zera ambos.
    def update_obs(row):
        if pd.notna(row["Data Obs"]) and pd.notna(row["Data Última Carga"]) and row["Data Obs"] > row["Data Última Carga"]:
            return pd.Series([row["Obs"], row["Data Obs"]])
        else:
            return pd.Series(["", ""])
    df[["Obs", "Data Obs"]] = df.apply(update_obs, axis=1)
    
    # Remove a coluna "Data Última Carga" (mantendo "Data Obs")
    df = df.drop(columns=["Data Última Carga"])
    
    # Calcula o Critério com base nos campos "Tempo TCERJ" e "Dias no Orgão"
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
    df["Critério"] = df.apply(calcula_criterio, axis=1)
    
    # Separa os processos que já possuem informante pré-atribuído dos que precisam de distribuição
    df_preassigned = df[df["Informante"].notna() & (df["Informante"] != "")].copy()
    df_remaining = df[df["Informante"].isna() | (df["Informante"] == "")].copy()
    
    # Para os processos restantes, aplica a lógica de distribuição
    origens_especiais = ["SEC EST POLICIA MILITAR", "SEC EST DEFESA CIVIL"]
    df_grupo_a_data = df_remaining[df_remaining["Orgão Origem"].isin(origens_especiais)].copy()
    df_grupo_b_data = df_remaining[~df_remaining["Orgão Origem"].isin(origens_especiais)].copy()
    
    # Listas originais dos informantes
    informantes_grupo_a = ["Alessandro Rios", "André", "Rosane", "Anna Cymerman"]
    informantes_grupo_b = ["Lúcia", "Mônica Aranha", "Rodrigo", "Wellington", "Zezinho"]
    # Filtra os informantes de acordo com a disponibilidade
    informantes_grupo_a = [inf for inf in informantes_grupo_a if inf in available]
    informantes_grupo_b = [inf for inf in informantes_grupo_b if inf in available]
    
    df_grupo_a_data = df_grupo_a_data.sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
    df_grupo_b_data = df_grupo_b_data.sort_values(by="Dias no Orgão", ascending=False).reset_index(drop=True)
    
    if informantes_grupo_a:
        df_grupo_a_data["Informante"] = [informantes_grupo_a[i % len(informantes_grupo_a)] for i in range(len(df_grupo_a_data))]
    if informantes_grupo_b:
        df_grupo_b_data["Informante"] = [informantes_grupo_b[i % len(informantes_grupo_b)] for i in range(len(df_grupo_b_data))]
    
    df_assigned = pd.concat([df_grupo_a_data, df_grupo_b_data], ignore_index=True)
    
    # Combina os processos pré-atribuídos e os que receberam distribuição
    df_final = pd.concat([df_preassigned, df_assigned], ignore_index=True)
    
    # Ordena o DataFrame final por informante, Critério (através de mapeamento) e "Dias no Orgão"
    priority_map = {
        "01 Mais de cinco anos de autuado": 0,
        "02 A completar 5 anos de autuado": 1,
        "03 Mais de 5 meses na 3CAP": 2,
        "04 Data da carga": 3
    }
    df_final["CustomPriority"] = df_final["Critério"].apply(lambda x: priority_map.get(x, 4))
    df_final = df_final.sort_values(by=["Informante", "CustomPriority", "Dias no Orgão"],
                                    ascending=[True, True, False]).reset_index(drop=True)
    df_final = df_final.drop(columns=["CustomPriority"])
    
    # O arquivo final manterá todas as colunas, inclusive "Descrição Informação"
    geral_filename = f"{numero}_planilha_geral_processos_{datetime.now().strftime('%Y%m%d')}.xlsx"
    geral_bytes = to_excel_bytes(df_final)
    
    # Gera as planilhas individuais para cada informante (limite de 200 processos, ordenados por prioridade)
    individual_files = {}
    informantes_list = df_final["Informante"].dropna().unique()
    for informante in informantes_list:
        df_informante = df_final[df_final["Informante"] == informante].copy()
        df_informante["CustomPriority"] = df_informante["Critério"].apply(lambda x: priority_map.get(x, 4))
        df_informante = df_informante.sort_values(by=["CustomPriority", "Dias no Orgão"],
                                                  ascending=[True, False])
        df_informante = df_informante.head(200).drop(columns=["CustomPriority"])
        filename = f"{informante.replace(' ', '_')}_{numero}_processos_{datetime.now().strftime('%Y%m%d')}.xlsx"
        individual_files[informante] = to_excel_bytes(df_informante)
    
    return geral_filename, geral_bytes, individual_files

# =============================================================================
# Configuração de número (mantido em session_state)
# =============================================================================
if "numero" not in st.session_state:
    st.session_state.numero = 184  # valor inicial

# =============================================================================
# Interface Gráfica (Streamlit)
# =============================================================================
st.title("Distribuição de Processos - Interface Gráfica")
st.markdown("### Faça o upload dos arquivos e configure a distribuição.")

# Upload dos arquivos – agora apenas um botão para os três arquivos necessários
uploaded_files = st.file_uploader(
    "Carregar os arquivos: processos.xlsx, observacoes.xlsx e disponibilidade_equipe.xlsx",
    type=["xlsx"],
    accept_multiple_files=True
)

# Mapeia os arquivos enviados com base no nome
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

if st.button("Executar Distribuição"):
    required_keys = ["processos", "observacoes", "disponibilidade"]
    if all(key in files_dict for key in required_keys):
        processos_file = files_dict["processos"]
        obs_file = files_dict["observacoes"]
        disp_file = files_dict["disponibilidade"]

        # Campo para numeração – o valor inicial vem de session_state
        numero = st.number_input("Qual a numeração dessa planilha de distribuição?", value=st.session_state.numero, step=1)

        # Controle de modo: Teste ou Produção
        modo = st.radio("Selecione o modo:", options=["Teste", "Produção"])
        test_mode = (modo == "Teste")
        st.markdown(f"**Modo selecionado:** {modo}")

        # Editor de e-mails dos informantes (permitindo edição)
        default_emails = pd.DataFrame({
            "Informante": ["Alessandro Rios", "André", "Rosane", "Anna Cymerman", 
                            "Lúcia", "Mônica Aranha", "Rodrigo", "Wellington", "Zezinho"],
            "Email": ["alessandrorr@tcerj.tc.br", "andrelbr@tcerj.tc.br", "rosanec@tcerj.tc.br", "anna.cymerman@exemplo.com",
                      "luciamfs@tcerj.tc.br", "monicaag@tcerj.tc.br", "rodrigob@tcerj.tc.br", "wellinsc@tcerj.tc.br", "josecn@tcerj.tc.br"]
        })
        st.markdown("### E-mails dos Informantes (pode editar)")
        edited_emails = st.experimental_data_editor(default_emails, num_rows="dynamic", key="emails_editor")
        emails_informantes = dict(zip(edited_emails["Informante"], edited_emails["Email"]))

        # Campo para e-mails dos gestores/revisores (separados por vírgula)
        managers_emails = st.text_input("E-mails dos gestores/revisores (separados por vírgula):", 
                                        value="sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com")

        # Executa a distribuição
        geral_filename, geral_bytes, individual_files = run_distribution(
            processos_file, obs_file, disp_file, numero, test_mode
        )
        st.success("Distribuição executada com sucesso!")
        
        # Disponibiliza o download da planilha geral
        st.download_button("Baixar Planilha Geral", data=geral_bytes, file_name=geral_filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.markdown("### Planilhas Individuais")
        for informante, file_bytes in individual_files.items():
            filename = f"{informante.replace(' ', '_')}_{numero}_processos_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(f"Baixar para {informante}", data=file_bytes, file_name=filename,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # Envio de e-mails:
        # Em modo Teste, envia apenas para os gestores; em Produção, também para os informantes.
        if test_mode:
            st.info("Modo Teste: E-mails para informantes NÃO serão enviados; apenas os e-mails dos gestores serão enviados.")
        else:
            for informante, file_bytes in individual_files.items():
                arquivo_informante = f"{informante.replace(' ', '_')}_{numero}_processos_{datetime.now().strftime('%Y%m%d')}.xlsx"
                email_destino = emails_informantes.get(informante, "")
                if email_destino:
                    subject = f"Distribuição de Processos - {informante}"
                    body = "Segue em anexo os processos distribuídos para você."
                    send_email_with_attachments([email_destino], subject, body, file_bytes, arquivo_informante)
                else:
                    st.write(f"E-mail não encontrado para {informante}")
        
        # Envia a planilha geral para os gestores/revisores (sempre enviar)
        managers_list = [e.strip() for e in managers_emails.split(",") if e.strip()]
        subject_geral = "Planilha Geral de Processos Distribuídos"
        body_geral = "Segue em anexo a planilha geral com todos os processos distribuídos."
        send_email_with_attachments(managers_list, subject_geral, body_geral, geral_bytes, geral_filename)
        
        # Atualiza a numeração para a próxima execução
        st.session_state.numero = numero + 1
    else:
        st.error("Por favor, faça o upload dos três arquivos necessários: processos.xlsx, observacoes.xlsx e disponibilidade_equipe.xlsx.")
