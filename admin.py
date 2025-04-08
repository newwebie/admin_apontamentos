import streamlit as st
import pandas as pd
from datetime import datetime
import io
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential

# === Configurações do SharePoint ===
username = st.secrets["sharepoint"]["username"]
password = st.secrets["sharepoint"]["password"]
site_url = st.secrets["sharepoint"]["site_url"]
file_name = st.secrets["sharepoint"]["file_name"]  # Arquivo Excel com duas abas:
                                                # "Staff Operações Clínica" e "Colaboradores"

# -------------------------------------------------------------------
# Função para ler as duas abas do Excel armazenado no SharePoint.
# Usamos o @st.cache_data sem ttl para manter o cache até que seja limpo.
# -------------------------------------------------------------------
@st.cache_data
def read_excel_sheets_from_sharepoint():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, file_name)
        xls = pd.ExcelFile(io.BytesIO(response.content))
        # Aba com as informações de limite de colaboradores
        staff_df = pd.read_excel(xls, sheet_name="Staff Operações Clínica")
        # Aba com os colaboradores já cadastrados
        colaboradores_df = pd.read_excel(xls, sheet_name="Colaboradores")
        return staff_df, colaboradores_df
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo ou ler as planilhas no SharePoint: {e}")
        return pd.DataFrame(), pd.DataFrame()

# -------------------------------------------------------------------
# Função para atualizar a aba "Colaboradores" dentro do mesmo Excel do SharePoint.
# Essa função preserva a aba "Staff Operações Clínica".
# -------------------------------------------------------------------
def update_colaboradores_sheet(colaboradores_df):
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        # Recarrega o arquivo para preservar a aba de Staff
        response = File.open_binary(ctx, file_name)
        xls = pd.ExcelFile(io.BytesIO(response.content))
        staff_df = pd.read_excel(xls, sheet_name="Staff Operações Clínica")
        
        # Cria um novo arquivo Excel em memória com as duas abas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            staff_df.to_excel(writer, sheet_name="Staff Operações Clínica", index=False)
            colaboradores_df.to_excel(writer, sheet_name="Colaboradores", index=False)
        output.seek(0)
        file_content = output.read()
        
        folder_path = "/".join(file_name.split("/")[:-1])
        file_name_only = file_name.split("/")[-1]
        target_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        target_folder.upload_file(file_name_only, file_content).execute_query()
        st.success("Arquivo atualizado com sucesso!")
    except Exception as e:
        st.error(f"Erro ao atualizar a planilha de colaboradores no SharePoint: {e}")

# -------------------------------------------------------------------
# Função principal com o formulário de cadastro
# -------------------------------------------------------------------
def main():
    st.title("")
    tabs = st.tabs(["Cadastro", "Base Colaboradores"])

    with tabs[0]:
        st.title("Cadastrar novo Colaborador")
        # Lê os dados do Excel com cache
        staff_df, colaboradores_df = read_excel_sheets_from_sharepoint()
        
        if staff_df.empty:
            st.error("Não foi possível carregar a planilha 'Staff Operações Clínica'.")
            return
        
        # Campos do formulário
        nome = st.text_input("Nome Completo do colaborador")
        cpf = st.text_input("CPF ou CNPJ", placeholder="Apenas números")
        
        # Para os selects, usamos os valores únicos da planilha de Staff
        cargos_unicos = sorted(staff_df["Cargo"].unique())
        cargo = st.selectbox("Cargo", cargos_unicos)
        
        escalas_unicas = sorted(staff_df["Escala"].unique())
        escala = st.selectbox("Escala", escalas_unicas)
        
        horarios_unicos = sorted(staff_df["Horário"].unique())
        horario = st.selectbox("Horário", horarios_unicos)
        
        turmas_unicas = sorted(staff_df["Turma"].unique())
        turma = st.selectbox("Turma", turmas_unicas)
        
        registro = st.selectbox("Tipo de Registro", ["Entrada", "Saída", "Atualização"])
        
        # Configura as datas conforme o tipo de registro
        if registro == "Entrada":
            entrada = st.date_input("Data da Entrada", value=datetime.today(), format='DD/MM/YYYY')
            saida = None
            att = None
        elif registro == "Saída":
            saida = st.date_input("Data da Saída", value=datetime.today(), format='DD/MM/YYYY')
            entrada = None
            att = None
        else:
            att = st.date_input("Data da Atualização", value=datetime.today(), format='DD/MM/YYYY')
            entrada = None
            saida = None
        
        contrato = st.selectbox("Tipo de Contrato", ["CLT", "Autonomo", "Horista"])
        
        supervisao = st.text_input("Supervisão Direta")
        status_prof = st.selectbox("Status do Profissional", 
                                ["Em Treinamento", "Apto", "Afastado", "Desistiu antes do onboarding"])
        responsavel = st.text_input("Responsável pela Inclusão dos dados")
        
        if st.button("Enviar"):
            # Validação dos campos obrigatórios
            if not nome.strip() or not cpf.strip() or not supervisao.strip() or not responsavel.strip():
                st.error("Preencha os campos obrigatórios: Nome, CPF/CNPJ, Supervisão Direta e Responsável.")
                return
            
            # Verifica se a combinação (Escala, Horário, Turma, Cargo) existe na aba de Staff
            filtro_staff = staff_df[
                (staff_df["Escala"] == escala) &
                (staff_df["Horário"] == horario) &
                (staff_df["Turma"] == turma) &
                (staff_df["Cargo"] == cargo)
            ]
            if filtro_staff.empty:
                st.error("Essa combinação de Escala / Horário / Turma / Cargo não existe na planilha base.")
                return
            
            # Pega a quantidade máxima permitida para essa combinação
            max_colabs = int(filtro_staff["Quantidade Staff"].iloc[0])
            
            # Conta quantos colaboradores já foram cadastrados para essa combinação
            filtro_colab = colaboradores_df[
                (colaboradores_df["Escala"] == escala) &
                (colaboradores_df["Horário"] == horario) &
                (colaboradores_df["Turma"] == turma) &
                (colaboradores_df["Cargo"] == cargo)
            ]
            count_atual = filtro_colab.shape[0]
            if count_atual >= max_colabs:
                st.error(f"Limite de colaboradores atingido para essa combinação: {max_colabs}")
                return


            duplicado = colaboradores_df[colaboradores_df["CPF ou CNPJ"] == cpf]
            if not duplicado.empty:
                st.warning("Este CPF já foi cadastrado.")
            
            # Formata a data de cadastro
            data_formatada = datetime.today().strftime("%d/%m/%Y")
            

            # Cria o novo registro do colaborador
            novo_colaborador = {
                "Nome Completo do Profissional": nome,
                "CPF ou CNPJ": cpf,
                "Cargo": cargo,
                "Escala": escala,
                "Horário": horario,
                "Turma": turma,
                "Tipo de Registro": registro,
                "Entrada": entrada,
                "Saída": saida,
                "Atualização": att,
                "Tipo de Contrato": contrato,
                "Supervisão Direta": supervisao,
                "Status do Profissional": status_prof,
                "Responsável pela Inclusão dos dados": responsavel,
                "CreatedAt": data_formatada
            }
            
            # Concatena o novo colaborador com os já cadastrados
            novo_df = pd.DataFrame([novo_colaborador])
            colaboradores_df = pd.concat([colaboradores_df, novo_df], ignore_index=True)
            
            # Atualiza a aba "Colaboradores" no Excel do SharePoint
            update_colaboradores_sheet(colaboradores_df)
            
            # Limpa o cache para que a próxima leitura traga os dados atualizados e
            # força uma recarga da aplicação (F5 automático)
            st.cache_data.clear()
    
    with tabs[1]:
        st.title('Base Colaboradores')
        staff_df, df = read_excel_sheets_from_sharepoint()
        if df.empty:
            st.info("Não há colaboradores na base")
        else:
            date_cols = ["Data de Entrada", "Data Saída", "Data Atualização"]
            for col in date_cols:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
            df.index = range(1, len(df) + 1)
            st.dataframe(df)

if __name__ == "__main__":
    main()
