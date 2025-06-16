import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import smtplib
from email.message import EmailMessage
from xlsxwriter.utility import xl_col_to_name

# =============================================================================
# Funções de envio de e-mail e geração de arquivos
# =============================================================================

def send_email_with_multiple_attachments(to_emails, subject, body, attachments):
    """Envia um único e‑mail com várias planilhas em anexo."""
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    smtp_username = 'sergiolbezerralj@gmail.com'  # TODO: mover para st.secrets
    smtp_password = 'dimwpnhowxxeqbes'            # TODO: mover para st.secrets

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


def send_email_with_two_attachments(to_email, subject, body,
                                    attachment_pre, filename_pre,
                                    attachment_res, filename_res):
    """Envia e‑mail para informante com PRE e PRINCIPAL (se existirem)."""
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


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Converte DataFrame em bytes (xlsx) formatando a coluna 'Grupo Natureza'."""
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


def create_zip_from_dict(file_dict: dict[str, bytes]) -> bytes:
    """Cria ZIP em memória a partir de dict {nome: bytes}."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zf:
        for filename, file_bytes in file_dict.items():
            zf.writestr(filename, file_bytes)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

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

# Mapeia rapidamente pelos nomes (semelhante ao código original)
files_dict: dict[str, io.BytesIO] = {}
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
TEST_MODE = (modo == "Teste")

if TEST_MODE:
    st.info("**Modo Teste**: nenhum e-mail será enviado; as planilhas ficarão disponíveis para download.")
else:
    st.info("**Modo Produção**: e-mails serão enviados conforme seleção abaixo.")

modo_envio = None
if not TEST_MODE:
    modo_envio = st.radio(
        "Selecione o modo de envio:",
        options=["Produção - Gestores e Informantes", "Produção - Apenas Gestores"]
    )

managers_emails = st.text_input(
    "E-mails dos gestores/revisores (separados por vírgula):",
    value="annapc@tcerj.tc.br, fabiovf@tcerj.tc.br, sergiolblj@tcerj.tc.br, sergiollima2@hotmail.com"
)

# ------------------------------------------
# Construção dos filtros por informante
# ------------------------------------------

filtros_grupo_natureza: dict[str, list[str]] = {}
filtros_orgao_origem:   dict[str, list[str]] = {}
informantes_principais: list[str] = []

if all(key in files_dict for key in ("processos", "processosmanter", "observacoes", "disponibilidade")):
    # Carrega e filtra processamentos mínimos para descobrir informantes
    df_proc = pd.read_excel(files_dict["processos"])
    df_proc.columns = df_proc.columns.str.strip()

    df_manter = pd.read_excel(files_dict["processosmanter"])
    df_manter.columns = df_manter.columns.str.strip()
    processos_validos = df_manter["Processo"].dropna().unique()

    df_proc = df_proc[df_proc["Processo"].isin(processos_validos)]
    df_proc = df_proc[df_proc["Tipo Processo"] == "Principal"].copy()

    # Pré‑atribuição removida (manter lógica original)
    df_proc["Descrição Informação"] = df_proc["Descrição Informação"].astype(str).str.strip().str.lower()
    df_proc["Funcionário Informação"] = df_proc["Funcionário Informação"].astype(str).str.strip()
    mask_preassigned = df_proc["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df_proc["Funcionário Informação"] != "")
    res_df = df_proc[~mask_preassigned].copy()

    # Define informantes e grupos A/B
    df_disp = pd.read_excel(files_dict["disponibilidade"])
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
    df_grupo_a = res_df[res_df["Orgão Origem"].isin(origens_especiais)].copy()
    df_grupo_b = res_df[~res_df["Orgão Origem"].isin(origens_especiais)].copy()

    if informantes_grupo_a:
        df_grupo_a["Informante"] = [informantes_grupo_a[i % len(informantes_grupo_a)]
                                      for i in range(len(df_grupo_a))]
    if informantes_grupo_b:
        df_grupo_b["Informante"] = [informantes_grupo_b[i % len(informantes_grupo_b)]
                                      for i in range(len(df_grupo_b))]

    res_assigned = pd.concat([df_grupo_a, df_grupo_b], ignore_index=True)

    informantes_principais = sorted(res_assigned["Informante"].dropna().unique())
    grupo_natureza_options = sorted(res_assigned["Grupo Natureza"].dropna().unique())
    orgao_origem_options   = sorted(res_assigned["Orgão Origem"].dropna().unique())

    st.markdown("### Filtros de Grupo Natureza / Orgão Origem (por informante)")
    for inf in informantes_principais:
        filtros_grupo_natureza[inf] = st.multiselect(
            f"Grupo(s) de Natureza para **{inf}** (opcional)",
            options=grupo_natureza_options,
            key=f"gn_{inf.replace(' ', '_')}"
        )
        filtros_orgao_origem[inf] = st.multiselect(
            f"Orgão(s) de Origem para **{inf}** (opcional)",
            options=orgao_origem_options,
            key=f"oo_{inf.replace(' ', '_')}"
        )

# =============================================================================
# Função principal de distribuição
# =============================================================================

def run_distribution(processos_file, processosmanter_file, obs_file,
                     disp_file, numero,
                     filtros_grupo_natureza: dict[str, list[str]],
                     filtros_orgao_origem: dict[str, list[str]]):
    """Executa toda a lógica de distribuição e devolve bytes/filenames."""

    df = pd.read_excel(processos_file)
    df.columns = df.columns.str.strip()

    df_manter = pd.read_excel(processosmanter_file)
    df_manter.columns = df_manter.columns.str.strip()
    processos_validos = df_manter["Processo"].dropna().unique()

    df = df[df["Processo"].isin(processos_validos)]
    df = df[df["Tipo Processo"] == "Principal"].copy()

    # Campos mínimos que vamos trabalhar
    keep_cols = [
        "Processo", "Grupo Natureza", "Orgão Origem", "Dias no Orgão",
        "Tempo TCERJ", "Data Última Carga", "Descrição Informação", "Funcionário Informação"
    ]
    df = df[keep_cols]

    # Entra a observação + remoção de análise suspensa
    df_obs = pd.read_excel(obs_file)
    df_obs.columns = df_obs.columns.str.strip()

    df = pd.merge(df, df_obs[["Processo", "Obs", "Data Obs"]], on="Processo", how="left")

    mask_suspensa = df["Obs"].astype(str).str.lower().str.contains("análise suspensa", na=False)
    df = df[~mask_suspensa].copy()

    # Limpa datas/obs conforme mais nova
    df["Data Última Carga"] = pd.to_datetime(df["Data Última Carga"], errors="coerce")
    df["Data Obs"] = pd.to_datetime(df["Data Obs"], errors="coerce")

    def _update_obs(row):
        if pd.notna(row["Data Obs"]) and pd.notna(row["Data Última Carga"]) and row["Data Obs"] > row["Data Última Carga"]:
            return pd.Series([row["Obs"], row["Data Obs"]])
        return pd.Series(["", ""])

    df[["Obs", "Data Obs"]] = df.apply(_update_obs, axis=1)
    df = df.drop(columns=["Data Última Carga"])

    # Pré‑atribuídos (já em instrução)
    df["Descrição Informação"] = df["Descrição Informação"].astype(str).str.strip().str.lower()
    df["Funcionário Informação"] = df["Funcionário Informação"].astype(str).str.strip()
    mask_pre = df["Descrição Informação"].isin(["em elaboração", "concluída"]) & (df["Funcionário Informação"] != "")

    pre_df = df[mask_pre].copy()
    pre_df["Informante"] = pre_df["Funcionário Informação"]

    res_df = df[~mask_pre].copy()

    # Disponibilidade
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
    df_grupo_a = res_df[res_df["Orgão Origem"].isin(origens_especiais)].copy()
    df_grupo_b = res_df[~res_df["Orgão Origem"].isin(origens_especiais)].copy()

    df_grupo_a = df_grupo_a.sort_values("Dias no Orgão", ascending=False).reset_index(drop=True)
    df_grupo_b = df_grupo_b.sort_values("Dias no Orgão", ascending=False).reset_index(drop=True)

    if informantes_grupo_a:
        df_grupo_a["Informante"] = [informantes_grupo_a[i % len(informantes_grupo_a)] for i in range(len(df_grupo_a))]
    if informantes_grupo_b:
        df_grupo_b["Informante"] = [informantes_grupo_b[i % len(informantes_grupo_b)] for i in range(len(df_grupo_b))]

    res_assigned = pd.concat([df_grupo_a, df_grupo_b], ignore_index=True)

    # Função critério/prioridade
    def _calc_criterio(row):
        if pd.isna(row["Processo"]) or row["Processo"] == "":
            return ""
        if row["Tempo TCERJ"] > 1765:
            return "01 Mais de cinco anos de autuado"
        if 1220 < row["Tempo TCERJ"] < 1765:
            return "02 A completar 5 anos de autuado"
        if row["Dias no Orgão"] >= 150:
            return "03 Mais de 5 meses na 3CAP"
        return "04 Data da carga"

    priority_map = {
        "01 Mais de cinco anos de autuado": 0,
        "02 A completar 5 anos de autuado": 1,
        "03 Mais de 5 meses na 3CAP":      2,
        "04 Data da carga":                3,
    }

    # ---------------- PLANILHA GERAL PRE ----------------
    pre_geral_filename = f"{numero}_planilha_geral_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
    pre_geral_bytes    = to_excel_bytes(pre_df)

    # ---------------- FILTROS & PLANILHA GERAL PRINCIPAL ----------------
    df_list = []
    for inf in res_assigned["Informante"].dropna().unique():
        df_inf = res_assigned[res_assigned["Informante"] == inf].copy()

        # 1) Grupo Natureza
        grupos = filtros_grupo_natureza.get(inf, [])
        if grupos:
            df_inf = df_inf[df_inf["Grupo Natureza"].isin(grupos)]

        # 2) Orgão Origem
        orgaos = filtros_orgao_origem.get(inf, [])
        if orgaos:
            df_inf = df_inf[df_inf["Orgão Origem"].isin(orgaos)]

        df_list.append(df_inf)

    res_assigned_filtered = pd.concat(df_list, ignore_index=True)

    res_assigned_filtered["Critério"] = res_assigned_filtered.apply(_calc_criterio, axis=1)
    res_assigned_filtered["CustomPriority"] = res_assigned_filtered["Critério"].map(priority_map)
    res_assigned_filtered = res_assigned_filtered.sort_values(
        ["Informante", "CustomPriority", "Dias no Orgão"],
        ascending=[True, True, False]
    ).reset_index(drop=True)
    res_assigned_filtered = res_assigned_filtered.drop(columns=["CustomPriority", "Descrição Informação", "Funcionário Informação"], errors="ignore")

    # Reorganiza ordem colunas para manter Critério na 3ª posição
    cols = res_assigned_filtered.columns.tolist()
    if "Critério" in cols:
        cols.remove("Critério")
        cols.insert(2, "Critério")
        res_assigned_filtered = res_assigned_filtered[cols]

    res_geral_filename = f"{numero}_planilha_geral_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
    res_geral_bytes    = to_excel_bytes(res_assigned_filtered)

    # ---------------- INDIVIDUAIS PRE ----------------
    pre_individual_files = {}
    for inf in pre_df["Informante"].dropna().unique():
        df_inf = pre_df[pre_df["Informante"] == inf].copy()
        df_inf["Critério"] = df_inf.apply(_calc_criterio, axis=1)
        df_inf["CustomPriority"] = df_inf["Critério"].map(priority_map)
        df_inf = df_inf.sort_values(["CustomPriority", "Dias no Orgão"], ascending=[True, False])
        df_inf = df_inf.drop(columns=["CustomPriority"])
        df_inf = df_inf.head(200)
        fname_inf = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
        pre_individual_files[inf] = to_excel_bytes(df_inf)

    # ---------------- INDIVIDUAIS PRINCIPAL ----------------
    res_individual_files = {}
    for inf in res_assigned["Informante"].dropna().unique():
        df_inf = res_assigned[res_assigned["Informante"] == inf].copy()

        # filtros por inf
        grupos = filtros_grupo_natureza.get(inf, [])
        if grupos:
            df_inf = df_inf[df_inf["Grupo Natureza"].isin(grupos)]
        orgaos = filtros_orgao_origem.get(inf, [])
        if orgaos:
            df_inf = df_inf[df_inf["Orgão Origem"].isin(orgaos)]

        df_inf["Critério"] = df_inf.apply(_calc_criterio, axis=1)
        df_inf["CustomPriority"] = df_inf["Critério"].map(priority_map)
        df_inf = df_inf.sort_values(["CustomPriority", "Dias no Orgão"], ascending=[True, False])
        df_inf = df_inf.drop(columns=["CustomPriority"])
        df_inf = df_inf.head(200)
        fname_inf = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
        res_individual_files[inf] = to_excel_bytes(df_inf)

    return (
        pre_geral_filename, pre_geral_bytes,
        res_geral_filename, res_geral_bytes,
        pre_individual_files, res_individual_files,
        informantes_emails
    )

# =============================================================================
# Botão de execução
# =============================================================================

if st.button("Executar Distribuição"):
    required_keys = ["processos", "processosmanter", "observacoes", "disponibilidade"]
    if all(key in files_dict for key in required_keys):
        (
            pre_geral_filename, pre_geral_bytes,
            res_geral_filename, res_geral_bytes,
            pre_individual_files, res_individual_files,
            informantes_emails
        ) = run_distribution(
            files_dict["processos"],
            files_dict["processosmanter"],
            files_dict["observacoes"],
            files_dict["disponibilidade"],
            numero,
            filtros_grupo_natureza,
            filtros_orgao_origem
        )

        st.success("Distribuição executada com sucesso!")

        # Download buttons
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

        # Individuais PRE
        st.markdown("### Planilhas Individuais - Pré-Atribuídos")
        for inf, file_bytes in pre_individual_files.items():
            fname = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                f"Baixar para {inf} (Pré-Atribuído)",
                data=file_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Individuais PRINCIPAL
        st.markdown("### Planilhas Individuais - Principal")
        for inf, file_bytes in res_individual_files.items():
            fname = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                f"Baixar para {inf} (Principal)",
                data=file_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ------------------------------------------------------------------
        # Envio de e-mails (Produção)
        # ------------------------------------------------------------------
        if not TEST_MODE:
            managers_list = [e.strip() for e in managers_emails.split(',') if e.strip()]

            # 1) Gestores sempre recebem ZIP + Geral
            if managers_list:
                all_individual_files = {}
                # junta pré + principal
                for inf, bytes_ in pre_individual_files.items():
                    fname = f"{inf.replace(' ', '_')}_{numero}_pre_atribuida_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    all_individual_files[fname] = bytes_
                for inf, bytes_ in res_individual_files.items():
                    fname = f"{inf.replace(' ', '_')}_{numero}_principal_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    all_individual_files[fname] = bytes_

                zip_bytes = create_zip_from_dict(all_individual_files)
                zip_fname = f"{numero}_planilhas_individuais_{datetime.now().strftime('%Y%m%d')}.zip"

                attachments = [
                    (pre_geral_bytes, pre_geral_filename),
                    (res_geral_bytes, res_geral_filename),
                    (zip_bytes, zip_fname)
                ]

                body_managers = (
                    "Prezados gestores,\n\n"
                    "Seguem anexas:\n"
                    "• Planilha Geral Pré-Atribuída\n"
                    "• Planilha Geral Principal\n"
                    "• ZIP com todas as planilhas individuais\n\n"
                    "Atenciosamente,\n3ª CAP"
                )

                send_email_with_multiple_attachments(
                    managers_list,
                    "Planilhas Gerais e Individuais de Processos",
                    body_managers,
                    attachments
                )

            # 2) Informantes (se selecionado)
            if modo_envio == "Produção - Gestores e Informantes":
                for inf in set(pre_individual_files.keys()) | set(res_individual_files.keys()):
                    email_dest = informantes_emails.get(inf.upper(), "")
                    if not email_dest:
                        continue

                    att_pre = pre_individual
