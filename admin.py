import streamlit as st
import pandas as pd
from datetime import datetime, date
import io
import re
import csv
import time
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential


# === Configurações do SharePoint ===
username = st.secrets["sharepoint"]["username"]
password = st.secrets["sharepoint"]["password"]
site_url = st.secrets["sharepoint"]["site_url"]
file_name = st.secrets["sharepoint"]["file_name"]
apontamentos_file = st.secrets["sharepoint"]["apontamentos_file"]
bio_file = st.secrets["sharepoint"]["bio_file"]



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

def update_staff_sheet(staff_df):
    while True:
        try:
            ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

            # Lê o workbook para preservar aba "Colaboradores"
            response = File.open_binary(ctx, file_name)
            xls = pd.ExcelFile(io.BytesIO(response.content))
            colaboradores_df = pd.read_excel(xls, sheet_name="Colaboradores")

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                staff_df.to_excel(w, sheet_name="Staff Operações Clínica", index=False)
                colaboradores_df.to_excel(w, sheet_name="Colaboradores", index=False)
            out.seek(0)

            folder = "/".join(file_name.split("/")[:-1])
            name   = file_name.split("/")[-1]
            ctx.web.get_folder_by_server_relative_url(folder).upload_file(name, out.read()).execute_query()

            # sucesso
            st.cache_data.clear()
            st.success("Alterações submetidas com sucesso!")
            break

        except Exception as e:
            locked = (
                getattr(e, "response_status", None) == 423
                or "-2147018894" in str(e)
                or "lock" in str(e).lower()
            )
            if locked:
                st.warning("Arquivo em uso. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue
            else:
                st.error(f"Erro ao atualizar a planilha de Staff no SharePoint: {e}")
                break


def update_colaboradores_sheet(colaboradores_df):
    while True:
        try:
            ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

            # Recarrega o arquivo para preservar a aba de Staff
            response = File.open_binary(ctx, file_name)
            xls = pd.ExcelFile(io.BytesIO(response.content))
            staff_df = pd.read_excel(xls, sheet_name="Staff Operações Clínica")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                staff_df.to_excel(writer, sheet_name="Staff Operações Clínica", index=False)
                colaboradores_df.to_excel(writer, sheet_name="Colaboradores", index=False)
            output.seek(0)

            folder_path     = "/".join(file_name.split("/")[:-1])
            file_name_only  = file_name.split("/")[-1]
            target_folder   = ctx.web.get_folder_by_server_relative_url(folder_path)
            target_folder.upload_file(file_name_only, output.read()).execute_query()

            # sucesso
            st.success("Alterações submetidas com sucesso!")
            st.cache_data.clear()
            break

        except Exception as e:
            locked = (
                getattr(e, "response_status", None) == 423
                or "-2147018894" in str(e)
                or "lock" in str(e).lower()
            )
            if locked:
                st.warning("Arquivo em uso. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue
            else:
                st.error(f"Erro ao atualizar a planilha de Colaboradores no SharePoint: {e}")
                break


@st.cache_data
def get_sharepoint_file():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, apontamentos_file)
        df = pd.read_excel(io.BytesIO(response.content))
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo de apontamentos: {e}")
        return pd.DataFrame()



def update_sharepoint_file(df_editado):
    while True:
        try:
            ctx = ClientContext(site_url).with_credentials(
                UserCredential(username, password)
            )

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_editado.to_excel(writer, index=False)
            output.seek(0)

            folder_path    = "/".join(apontamentos_file.split("/")[:-1])
            file_name_only = apontamentos_file.split("/")[-1]
            target_folder  = ctx.web.get_folder_by_server_relative_url(folder_path)
            target_folder.upload_file(file_name_only, output.read()).execute_query()

            # sucesso
            st.cache_data.clear()
            st.session_state["df_apontamentos"] = df_editado
            st.success("Apontamentos atualizados com sucesso!")
            break

        except Exception as e:
            locked = (
                getattr(e, "response_status", None) == 423
                or "-2147018894" in str(e)
                or "lock" in str(e).lower()
            )
            if locked:
                st.warning("Arquivo em uso. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue
            else:
                st.error(f"Erro ao salvar o arquivo de apontamentos: {e}")
                break



def get_deslig_state(colab_key: str, default_date: date | None, default_reason: str):
    """
    Garante que st.session_state tenha chaves exclusivas por colaborador:
      • ds_data_<colab_key>   → date (data de desligamento)
      • ds_reason_<colab_key> → str  (motivo desligamento/distrato)

    Retorna (key_date, key_reason) para usar nos widgets.
    """
    k_date   = f"ds_data_{colab_key}"
    k_reason = f"ds_reason_{colab_key}"

    if k_date not in st.session_state:
        st.session_state[k_date] = default_date or date.today()
    if k_reason not in st.session_state:
        st.session_state[k_reason] = default_reason

    return k_date, k_reason

#função pra ler o CPF corretamente
def so_digitos(v):
    return re.sub(r"\D", "", str(v))


def clear_cache_and_reload():
    st.cache_data.clear()       # limpa st.cache_data




# -------------------------------------------------------------------
# Formulário de cadastro
# -------------------------------------------------------------------
st.set_page_config(layout="wide")  

def main():
    st.title("📋 Painel ADM")
    tabs = st.tabs(["Apontamentos", "Posições", "Atualizar Colaborador", "Novo Colaborador"])

    with tabs[3]:
        spacer_left, main, spacer_right = st.columns([2, 4, 2])
        with main:
            st.title("Cadastrar Colaborador")

             # Lê os dados do Excel com cache
            staff_df, colaboradores_df = read_excel_sheets_from_sharepoint()

            # Se a planilha Staff está vazia, aborta
            if staff_df.empty:
                st.error("Não foi possível carregar a planilha 'Staff Operações Clínica'.")
                return

            # ===========================
            # Formulário de cadastro
            # ===========================

            

            # 👉 Novo: Selecionar somente o ID da vaga
            id_vagas = sorted(staff_df["ID Vaga"].dropna().unique())
            id_vaga = st.selectbox("ID Vaga", id_vagas)

            

            # Dados da vaga selecionada
            vaga_info = staff_df.loc[staff_df["ID Vaga"] == id_vaga].iloc[0]
            # Calculo vagas disponiveis  
            disponiveis = vaga_info["Quantidade Staff"] - vaga_info["Ativos"]
            st.text_input("Vagas Diponiveis", disponiveis, disabled=True)
            st.markdown("---")

            

            nome = st.text_input("Nome Completo do colaborador")
            cpf = str(st.text_input("CPF ou CNPJ", placeholder="Apenas números"))

            # Campos somente exibição
            
            st.text_input("Cargo", vaga_info["Cargo"], disabled=True)
            st.text_input("Turma", vaga_info["Departamento"], disabled=True)
            st.text_input("Escala", vaga_info["Escala"], disabled=True)
            st.text_input("Horário", vaga_info["Horário"], disabled=True)
            st.text_input("Turma", vaga_info["Turma"], disabled=True)
            st.text_input("Plantão", vaga_info["Plantão"], disabled=True)
            st.text_input("Supervisão", vaga_info["Supervisora"], disabled=True)
            

            entrada = st.date_input(
                "Data da Entrada",  
                format="DD/MM/YYYY",
            )

            contrato = st.selectbox(
                "Tipo de Contrato",
                ["CLT", "Autonomo", "Horista"],
            )

            responsavel = st.text_input("Responsável pela Inclusão dos dados")

            if st.button("Enviar"):
                # -----------------------
                # Validações
                # -----------------------
                if not nome.strip() or not responsavel.strip() or not cpf.strip():
                    st.error(
                        "Preencha os campos obrigatórios: Nome, Supervisão Direta e Responsável."
                    )
                    return

                colab_cpfs = colaboradores_df["CPF ou CNPJ"].apply(so_digitos)
                if cpf in colab_cpfs.values:
                    st.error(
                        "Já existe um colaborador cadastrado com este CPF/CNPJ."
                    )
                    return

                # Limite de colaboradores por vaga
                max_colabs = int(vaga_info["Quantidade Staff"])
                status_col = "Ativos"

                filtro_colab = colaboradores_df[
                    (colaboradores_df["ID Vaga"] == id_vaga)
                    & (colaboradores_df[status_col] == "Sim")
                ]

                if filtro_colab.shape[0] >= max_colabs:
                    st.error(
                        f"Limite de colaboradores atingido para essa vaga: {max_colabs}"
                    )
                    return

                # -----------------------
                # Cria novo colaborador
                # -----------------------
                novo_colaborador = {
                    "ID Vaga": id_vaga,
                    "Nome Completo do Profissional": nome,
                    "CPF ou CNPJ": cpf,
                    "Cargo": vaga_info["Cargo"],
                    "Departamento": vaga_info["Departamento"],
                    "Escala": vaga_info["Escala"],
                    "Horário": vaga_info["Horário"],
                    "Turma": vaga_info["Turma"],
                    "Plantão": vaga_info["Plantão"],
                    "Supervisão Direta": vaga_info["Supervisora"],
                    "Data Entrada": entrada,
                    "Tipo de Contrato": contrato,
                    "Responsável pela Inclusão dos dados": responsavel,
                    status_col: "Sim",
                    "Status do Profissional": "Menos de 3 meses",
                }

                colaboradores_df = pd.concat(
                    [colaboradores_df, pd.DataFrame([novo_colaborador])],
                    ignore_index=True,
                )

                # 5️⃣  Salva de volta no SharePoint, limpa cache e avisa
                update_colaboradores_sheet(colaboradores_df)
                st.cache_data.clear()
# --------------------------------------------------------------------
# Edição de colaboradores
# --------------------------------------------------------------------

        with tabs[2]:
            spacer_left, main_col, spacer_right = st.columns([2, 4, 2])
            with main_col:
                if colaboradores_df.empty:
                    st.info("Não há colaboradores na base")
                    st.stop()

                st.title("Atualizar Colaborador")

                # 1) Seleciona colaborador
                nomes = colaboradores_df["Nome Completo do Profissional"].dropna().sort_values().unique()
                selec_nome = st.selectbox("Selecione o colaborador", nomes, key="sel_colab")

                linha = colaboradores_df.loc[colaboradores_df["Nome Completo do Profissional"] == selec_nome].iloc[0]

                old_id_vaga = linha.get("ID Vaga", "")
                old_ativo = linha.get("Ativos", "Não")

                # 2) Seleciona a vaga
                id_vagas = sorted(staff_df["ID Vaga"].dropna().unique())
                idx_vaga = id_vagas.index(old_id_vaga) if old_id_vaga in id_vagas else 0
                id_vaga = st.selectbox("ID Vaga", id_vagas, index=idx_vaga, key=f"idvaga_{selec_nome}")

                vaga_info = staff_df.loc[staff_df["ID Vaga"] == id_vaga].iloc[0]

                # 3) Campos editáveis
                nome = st.text_input("Nome Completo do Profissional", linha["Nome Completo do Profissional"], key=f"nome_{selec_nome}")
                cpf = st.text_input("CPF ou CNPJ", linha["CPF ou CNPJ"], key=f"cpf_{selec_nome}")

                lista_status = ["Em Treinamento", "Apto", "Afastado", "Desistiu antes do onboarding", "Desligado"]
                status_prof = st.selectbox(
                    "Status do Profissional",
                    lista_status,
                    index=lista_status.index(linha["Status do Profissional"]) if linha["Status do Profissional"] in lista_status else 0,
                    key=f"status_{selec_nome}",
                )

                tipo_contrato = st.selectbox(
                    "Tipo de Contrato",
                    ["CLT", "Autonomo", "Horista"],
                    index=["CLT", "Autonomo", "Horista"].index(linha["Tipo de Contrato"]) if linha["Tipo de Contrato"] in ["CLT", "Autonomo", "Horista"] else 0,
                    key=f"contrato_{selec_nome}",
                )

                responsavel_att = st.text_input("Responsável pela Atualização dos dados", key=f"resp_{selec_nome}")

                # 4) Campos SOMENTE‑LEITURA provenientes da vaga
                st.text_input("Departamento", vaga_info["Departamento"], disabled=True, key=f"dep_{selec_nome}")
                st.text_input("Cargo", vaga_info["Cargo"], disabled=True, key=f"cargo_{selec_nome}")
                st.text_input("Escala", vaga_info["Escala"], disabled=True, key=f"escala_{selec_nome}")
                st.text_input("Horário", vaga_info["Horário"], disabled=True, key=f"hora_{selec_nome}")
                st.text_input("Turma", vaga_info["Turma"], disabled=True, key=f"turma_{selec_nome}")
                st.text_input("Supervisão Direta", vaga_info["Supervisora"], disabled=True, key=f"sup_{selec_nome}")
                st.text_input("Plantão", vaga_info["Plantão"], disabled=True, key=f"plantao_{selec_nome}")

                # 5) Disponibilidade da vaga
                max_colabs = int(vaga_info["Quantidade Staff"])
                ativos_na_vaga = colaboradores_df[
                    (colaboradores_df["ID Vaga"] == id_vaga)
                    & (colaboradores_df["Ativos"] == "Sim")
                ].shape[0]

                disponiveis = max_colabs - ativos_na_vaga
                st.info(f"Disponíveis: {disponiveis} / {max_colabs}")
                if disponiveis <= 0 and status_prof != "Desligado":
                    st.warning("Esta vaga está lotada. Só Será possível se marcar o colaborador como 'Desligado'.", icon="⚠️")

                # 6) Regras para desligamento
                data_deslig = linha.get("Data Desligamento", None)
                motivo_clt = linha.get("Desligamento CLT", "")
                motivo_auto = linha.get("Saída Autonomo", "")

                if status_prof == "Desligado":
                    key_date, key_reason = get_deslig_state(
                        selec_nome,
                        linha.get("Atualização", datetime.now()).date(),
                        motivo_clt or motivo_auto,
                    )

                    data_deslig = st.date_input(
                        "Data do desligamento",
                        format="DD/MM/YYYY",
                    )

                    if tipo_contrato.lower() == "clt":
                        lista_clt = ["Solicitação de Desligamento", "Desligamento pela Gestão"]
                        if st.session_state.get(key_reason) not in lista_clt:
                            st.session_state[key_reason] = lista_clt[0]
                        motivo_clt = st.selectbox("Motivo do desligamento (CLT)", lista_clt, key=key_reason)
                        motivo_auto = ""
                    elif tipo_contrato.lower() == "autonomo":
                        lista_auto = ["Distrato", "Solicitação de Distrato", "Distrato pela Gestão"]
                        if st.session_state.get(key_reason) not in lista_auto:
                            st.session_state[key_reason] = lista_auto[0]
                        motivo_auto = st.selectbox("Motivo do distrato (Autônomo)", lista_auto, key=key_reason)
                        motivo_clt = ""
                    else:
                        motivo_clt = motivo_auto = ""

                # 7) Botão SALVAR
                if st.button("Salvar alterações", key=f"save_{selec_nome}"):
                    if not responsavel_att.strip():
                        st.error("Preencha o campo Responsável pela Atualização dos dados.")
                        st.stop()

                    novo_ativo = "Não" if status_prof == "Desligado" else "Sim"

                    # ➤ Regras de ajuste de Ativos no staff_df
                    if status_prof == "Desligado":
                        staff_df.loc[staff_df["ID Vaga"] == old_id_vaga, "Ativos"] = (
                            staff_df.loc[staff_df["ID Vaga"] == old_id_vaga, "Ativos"].astype(int) - 1
                        )
                    elif old_id_vaga != id_vaga:
                        staff_df.loc[staff_df["ID Vaga"] == old_id_vaga, "Ativos"] = (
                            staff_df.loc[staff_df["ID Vaga"] == old_id_vaga, "Ativos"].astype(int) - 1
                        )
                        staff_df.loc[staff_df["ID Vaga"] == id_vaga, "Ativos"] = (
                            staff_df.loc[staff_df["ID Vaga"] == id_vaga, "Ativos"].astype(int) + 1
                        )

                    # Atualiza linha no colaboradores_df
                    colaboradores_df.loc[linha.name, [
                        "ID Vaga", "Nome Completo do Profissional", "CPF ou CNPJ", "Cargo", "Departamento",
                        "Escala", "Horário", "Turma", "Tipo de Contrato", "Supervisão Direta", "Plantão",
                        "Status do Profissional", "Ativos", "Responsável Atualização", "Atualização"
                    ]] = [
                        id_vaga, nome, cpf, vaga_info["Cargo"], vaga_info["Departamento"],
                        vaga_info["Escala"], vaga_info["Horário"], vaga_info["Turma"],
                        tipo_contrato, vaga_info["Supervisora"], vaga_info["Plantão"],
                        status_prof, novo_ativo, responsavel_att, datetime.now()
                    ]

                    # Campos de desligamento
                    if status_prof == "Desligado":
                        colaboradores_df.loc[linha.name, "Ativos"] = "Não"
                        colaboradores_df.loc[linha.name, "Data Desligamento"] = data_deslig
                        if tipo_contrato.lower() == "clt":
                            colaboradores_df.loc[linha.name, "Desligamento CLT"] = motivo_clt
                            colaboradores_df.loc[linha.name, "Saída Autonomo"] = ""
                        elif tipo_contrato.lower() == "autonomo":
                            colaboradores_df.loc[linha.name, "Saída Autonomo"] = motivo_auto
                            colaboradores_df.loc[linha.name, "Desligamento CLT"] = ""
                        else:
                            colaboradores_df.loc[linha.name, ["Desligamento CLT", "Saída Autonomo"]] = ""

                    # Salva nos sheets
                    update_colaboradores_sheet(colaboradores_df)
                    update_staff_sheet(staff_df)


# --------------------------------------------------------------------
# Edição apontamentos        
# --------------------------------------------------------------------

    

        with tabs[0]:
            st.title("Lista de Apontamentos")

            # 1️⃣  Carrega arquivo e garante 'orig_idx' --------------------------------
            df = get_sharepoint_file()                # <- sua função de leitura

            if "orig_idx" not in df.columns:          # primeira execução
                df["orig_idx"] = range(len(df))       # cria IDs

            df.set_index("orig_idx", inplace=True)    # índice permanente = ID

            if df.empty:
                st.info("Nenhum apontamento encontrado!")
            else:
                # 2️⃣  Conversão de colunas de data -----------------------------------
                colunas_data = [
                    "Data do Apontamento", "Prazo Para Resolução", "Data de Verificação",
                    "Data Resolução", "Data Atualização", "Disponibilizado para Verificação"
                ]

                for col in colunas_data:
                    if col in df.columns:
                        df[col] = (
                            pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")
                            .dt.date
                        )

                # 3️⃣  Cópia para filtros --------------------------------------------
                df_filtrado = df.copy()

                # (demais filtros / botões PENDENTE & VERIFICANDO — **inalterados**)
                # ---------------------------------------------------------------------
                def toggle_pending():
                    st.session_state.show_pending = not st.session_state.get("show_pending", False)
                    st.session_state.show_verificando = False

                def toggle_verificando():
                    st.session_state.show_verificando = not st.session_state.get("show_verificando", False)
                    st.session_state.show_pending = False

                st.session_state.setdefault("show_pending", False)
                st.session_state.setdefault("show_verificando", False)

                col_btn1, col_btn2, col_btn3,*_ = st.columns(6)

                with col_btn1:
                    label_pend = (
                        "🔍  Filtrar Pendentes"
                        if not st.session_state.show_pending
                        else "📄  Mostrar todos"
                    )
                    st.button(label_pend, key="btn_toggle_pendentes", on_click=toggle_pending)

                with col_btn2:
                    label_verif = (
                        "🔎  Filtrar Verificando"
                        if not st.session_state.show_verificando
                        else "📄  Mostrar todos"
                    )
                    st.button(label_verif, key="btn_toggle_verificando", on_click=toggle_verificando)
                
                with col_btn3:
                    st.button(
                        "🔄  Atualizar",
                        key="btn_clear_cache",
                        on_click=clear_cache_and_reload
                    )

                # 4️⃣  Colunas a exibir ------------------------------------------------
                columns_to_display = [
                    "Status", "Código do Estudo", "Responsável Pela Correção", "Plantão",
                    "Participante", "Período", "Grau De Criticidade Do Apontamento",
                    "Documentos", "Apontamento", "Data do Apontamento",
                    "Disponibilizado para Verificação", "Prazo Para Resolução",
                    "Data Resolução", "Justificativa", "Responsável Pelo Apontamento",
                    "Origem Do Apontamento", "Data Atualização", "Responsável Atualização"
                ]
                df_filtrado = df_filtrado[columns_to_display]

                # 5️⃣  Aplica filtros (pendente / verificando) -------------------------
                if st.session_state.show_pending:
                    df_view = df_filtrado[df_filtrado["Status"] == "PENDENTE"].copy()
                elif st.session_state.show_verificando:
                    df_view = df_filtrado[df_filtrado["Status"] == "VERIFICANDO"].copy()
                else:
                    df_view = df_filtrado.copy()

                # 6️⃣  Filtro por Código do Estudo ------------------------------------
                if "Código do Estudo" in df.columns:
                    opcoes_estudos = ["Todos"] + sorted(df["Código do Estudo"].dropna().unique())
                    estudo_sel = st.selectbox("Selecione o Estudo", opcoes_estudos, key="estudo_selecionado")

                    if estudo_sel != "Todos":
                        df_view = df_view[df_view["Código do Estudo"] == estudo_sel]

                # 7️⃣  Garante coluna 'orig_idx' VISÍVEL -------------------------------
                #    (reset_index move o índice → coluna; índice ficará oculto)
                df_view = df_view.reset_index()               # cria coluna orig_idx
                first = df_view.pop("orig_idx")               # move p/ primeira posição
                df_view.insert(0, "orig_idx", first)

                # -------------------------------------------------
                # 3️⃣  Configura colunas (idêntico, mas usa df_view)
                # -------------------------------------------------
                selectbox_columns_opcoes = {
                "Status": [
                    "REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "VERIFICANDO", "PENDENTE", "NÃO APLICÁVEL"
                ],

                "Origem Do Apontamento": [
                    "Documentação Clínica", "Excelência Operacional", "Operações Clínicas", "Patrocinador / Monitor", "Garantia Da Qualidade"
                ],

                "Documentos": [
                    "Acompanhamento da Administração da Medicação", "Ajuste dos Relógios", "Anotação de enfermagem",
                    "Aplicação do TCLE", "Ausência de Período", "Avaliação Clínica Pré Internação", "Avaliação de Alta Clínica",
                    "Controle de Eliminações fisiológicas", "Controle de Glicemia", "Controle de Ausente de Período",
                    "Controle de DropOut", "Critérios de Inclusão e Exclusão", "Desvio de ambulação", "Dieta",
                    "Diretrizes do Protocolo", "Tabela de Controle de Preparo de Heparina", "TIME", "TCLE", "ECG",
                    "Escala de Enfermagem", "Evento Adverso", "Ficha de internação", "Formulário de conferência das amostras",
                    "Teste de HCG", "Teste de Drogas", "Teste de Álcool", "Término Prematuro",
                    "Medicação para tratamento dos Eventos Adversos", "Orientação por escrito", "Prescrição Médica",
                    "Registro de Temperatura da Enfermaria", "Relação dos Profissionais", "Sinais Vitais Pós Estudo",
                    "SAE", "SINEB", "FOR 104", "FOR 123", "FOR 166", "FOR 217", "FOR 233", "FOR 234", "FOR 235",
                    "FOR 236", "FOR 240", "FOR 241", "FOR 367", "Outros"
                ],

                "Participante": [
                    'N/A','PP01', 'PP02', 'PP03', 'PP04', 'PP05', 'PP06', 'PP07', 'PP08', 'PP09', 'PP10', 'PP11', 'PP12', 'PP13', 'PP14', 'PP15', 'PP16', 'PP17', 'PP18', 'PP19', 'PP20', 'PP21', 'PP22', 'PP23', 'PP24', 'PP25', 'PP26', 'PP27', 'PP28', 'PP29', 'PP30', 'PP31', 'PP32', 'PP33', 'PP34', 'PP35', 'PP36', 'PP37', 'PP38', 'PP39', 'PP40', 'PP41', 'PP42', 'PP43', 'PP44', 'PP45', 'PP46', 'PP47', 'PP48', 'PP49', 'PP50', 'PP51', 'PP52', 'PP53', 'PP54', 'PP55', 'PP56', 'PP57', 'PP58', 'PP59', 'PP60', 'PP61', 'PP62', 'PP63', 'PP64', 'PP65', 'PP66', 'PP67', 'PP68', 'PP69', 'PP70', 'PP71', 'PP72', 'PP73', 'PP74', 'PP75', 'PP76', 'PP77', 'PP78', 'PP79', 'PP80', 'PP81', 'PP82', 'PP83', 'PP84', 'PP85', 'PP86', 'PP87', 'PP88', 'PP89', 'PP90', 'PP91', 'PP92', 'PP93', 'PP94', 'PP95', 'PP96', 'PP97', 'PP98', 'PP99', 'PP100', 'PP101', 'PP102', 'PP103', 'PP104', 'PP105', 'PP106', 'PP107', 'PP108', 'PP109', 'PP110', 'PP111', 'PP112', 'PP113', 'PP114', 'PP115', 'PP116', 'PP117', 'PP118', 'PP119', 'PP120', 'PP121', 'PP122', 'PP123', 'PP124', 'PP125', 'PP126', 'PP127', 'PP128', 'PP129', 'PP130', 'PP131', 'PP132', 'PP133', 'PP134', 'PP135', 'PP136', 'PP137', 'PP138', 'PP139', 'PP140', 'PP141', 'PP142', 'PP143', 'PP144', 'PP145', 'PP146', 'PP147', 'PP148', 'PP149', 'PP150', 'PP151', 'PP152', 'PP153', 'PP154', 'PP155', 'PP156', 'PP157', 'PP158', 'PP159', 'PP160', 'PP161', 'PP162', 'PP163', 'PP164', 'PP165', 'PP166', 'PP167', 'PP168', 'PP169', 'PP170', 'PP171', 'PP172', 'PP173', 'PP174', 'PP175', 'PP176', 'PP177', 'PP178', 'PP179', 'PP180', 'PP181', 'PP182', 'PP183', 'PP184', 'PP185', 'PP186', 'PP187', 'PP188', 'PP189', 'PP190', 'PP191', 'PP192', 'PP193', 'PP194', 'PP195', 'PP196', 'PP197', 'PP198', 'PP199', 'PP200', 'PP201', 'PP202', 'PP203', 'PP204', 'PP205', 'PP206', 'PP207', 'PP208', 'PP209', 'PP210', 'PP211', 'PP212', 'PP213', 'PP214', 'PP215', 'PP216', 'PP217', 'PP218', 'PP219', 'PP220', 'PP221', 'PP222', 'PP223', 'PP224', 'PP225', 'PP226', 'PP227', 'PP228', 'PP229', 'PP230', 'PP231', 'PP232', 'PP233', 'PP234', 'PP235', 'PP236', 'PP237', 'PP238', 'PP239', 'PP240', 'PP241', 'PP242', 'PP243', 'PP244', 'PP245', 'PP246', 'PP247', 'PP248', 'PP249', 'PP250', 'PP251', 'PP252', 'PP253', 'PP254', 'PP255', 'PP256', 'PP257', 'PP258', 'PP259', 'PP260', 'PP261', 'PP262', 'PP263', 'PP264', 'PP265', 'PP266', 'PP267', 'PP268', 'PP269', 'PP270', 'PP271', 'PP272', 'PP273', 'PP274', 'PP275', 'PP276', 'PP277', 'PP278', 'PP279', 'PP280', 'PP281', 'PP282', 'PP283', 'PP284', 'PP285', 'PP286', 'PP287', 'PP288', 'PP289', 'PP290', 'PP291', 'PP292', 'PP293', 'PP294', 'PP295', 'PP296', 'PP297', 'PP298', 'PP299', 'PP300', 'PP301', 'PP302', 'PP303', 'PP304', 'PP305', 'PP306', 'PP307', 'PP308', 'PP309', 'PP310', 'PP311', 'PP312', 'PP313', 'PP314', 'PP315', 'PP316', 'PP317', 'PP318', 'PP319', 'PP320', 'PP321', 'PP322', 'PP323', 'PP324', 'PP325', 'PP326', 'PP327', 'PP328', 'PP329', 'PP330', 'PP331', 'PP332', 'PP333', 'PP334', 'PP335', 'PP336', 'PP337', 'PP338', 'PP339', 'PP340', 'PP341', 'PP342', 'PP343', 'PP344', 'PP345', 'PP346', 'PP347', 'PP348', 'PP349', 'PP350', 'PP351', 'PP352', 'PP353', 'PP354', 'PP355', 'PP356', 'PP357', 'PP358', 'PP359', 'PP360', 'PP361', 'PP362', 'PP363', 'PP364', 'PP365', 'PP366', 'PP367', 'PP368', 'PP369', 'PP370', 'PP371', 'PP372', 'PP373', 'PP374', 'PP375', 'PP376', 'PP377', 'PP378', 'PP379', 'PP380', 'PP381', 'PP382', 'PP383', 'PP384', 'PP385', 'PP386', 'PP387', 'PP388', 'PP389', 'PP390', 'PP391', 'PP392', 'PP393', 'PP394', 'PP395', 'PP396', 'PP397', 'PP398', 'PP399', 'PP400', 'PP401', 'PP402', 'PP403', 'PP404', 'PP405', 'PP406', 'PP407', 'PP408', 'PP409', 'PP410', 'PP411', 'PP412', 'PP413', 'PP414', 'PP415', 'PP416', 'PP417', 'PP418', 'PP419', 'PP420', 'PP421', 'PP422', 'PP423', 'PP424', 'PP425', 'PP426', 'PP427', 'PP428', 'PP429', 'PP430', 'PP431', 'PP432', 'PP433', 'PP434', 'PP435', 'PP436', 'PP437', 'PP438', 'PP439', 'PP440', 'PP441', 'PP442', 'PP443', 'PP444', 'PP445', 'PP446', 'PP447', 'PP448', 'PP449', 'PP450', 'PP451', 'PP452', 'PP453', 'PP454', 'PP455', 'PP456', 'PP457', 'PP458', 'PP459', 'PP460', 'PP461', 'PP462', 'PP463', 'PP464', 'PP465', 'PP466', 'PP467', 'PP468', 'PP469', 'PP470', 'PP471', 'PP472', 'PP473', 'PP474', 'PP475', 'PP476', 'PP477', 'PP478', 'PP479', 'PP480', 'PP481', 'PP482', 'PP483', 'PP484', 'PP485', 'PP486', 'PP487', 'PP488', 'PP489', 'PP490', 'PP491', 'PP492', 'PP493', 'PP494', 'PP495', 'PP496', 'PP497', 'PP498', 'PP499', 'PP500', 'PP501', 'PP502', 'PP503', 'PP504', 'PP505', 'PP506', 'PP507', 'PP508', 'PP509', 'PP510', 'PP511', 'PP512', 'PP513', 'PP514', 'PP515', 'PP516', 'PP517', 'PP518', 'PP519', 'PP520', 'PP521', 'PP522', 'PP523', 'PP524', 'PP525', 'PP526', 'PP527', 'PP528', 'PP529', 'PP530', 'PP531', 'PP532', 'PP533', 'PP534', 'PP535', 'PP536', 'PP537', 'PP538', 'PP539', 'PP540', 'PP541', 'PP542', 'PP543', 'PP544', 'PP545', 'PP546', 'PP547', 'PP548', 'PP549', 'PP550', 'PP551', 'PP552', 'PP553', 'PP554', 'PP555', 'PP556', 'PP557', 'PP558', 'PP559', 'PP560', 'PP561', 'PP562', 'PP563', 'PP564', 'PP565', 'PP566', 'PP567', 'PP568', 'PP569', 'PP570', 'PP571', 'PP572', 'PP573', 'PP574', 'PP575', 'PP576', 'PP577', 'PP578', 'PP579', 'PP580', 'PP581', 'PP582', 'PP583', 'PP584', 'PP585', 'PP586', 'PP587', 'PP588', 'PP589', 'PP590', 'PP591', 'PP592', 'PP593', 'PP594', 'PP595', 'PP596', 'PP597', 'PP598', 'PP599', 'PP600', 'PP601', 'PP602', 'PP603', 'PP604', 'PP605', 'PP606', 'PP607', 'PP608', 'PP609', 'PP610', 'PP611', 'PP612', 'PP613', 'PP614', 'PP615', 'PP616', 'PP617', 'PP618', 'PP619', 'PP620', 'PP621', 'PP622', 'PP623', 'PP624', 'PP625', 'PP626', 'PP627', 'PP628', 'PP629', 'PP630', 'PP631', 'PP632', 'PP633', 'PP634', 'PP635', 'PP636', 'PP637', 'PP638', 'PP639', 'PP640', 'PP641', 'PP642', 'PP643', 'PP644', 'PP645', 'PP646', 'PP647', 'PP648', 'PP649', 'PP650', 'PP651', 'PP652', 'PP653', 'PP654', 'PP655', 'PP656', 'PP657', 'PP658', 'PP659', 'PP660', 'PP661', 'PP662', 'PP663', 'PP664', 'PP665', 'PP666', 'PP667', 'PP668', 'PP669', 'PP670', 'PP671', 'PP672', 'PP673', 'PP674', 'PP675', 'PP676', 'PP677', 'PP678', 'PP679', 'PP680', 'PP681', 'PP682', 'PP683', 'PP684', 'PP685', 'PP686', 'PP687', 'PP688', 'PP689', 'PP690', 'PP691', 'PP692', 'PP693', 'PP694', 'PP695', 'PP696', 'PP697', 'PP698', 'PP699', 'PP700', 'PP701', 'PP702', 'PP703', 'PP704', 'PP705', 'PP706', 'PP707', 'PP708', 'PP709', 'PP710', 'PP711', 'PP712', 'PP713', 'PP714', 'PP715', 'PP716', 'PP717', 'PP718', 'PP719', 'PP720', 'PP721', 'PP722', 'PP723', 'PP724', 'PP725', 'PP726', 'PP727', 'PP728', 'PP729', 'PP730', 'PP731', 'PP732', 'PP733', 'PP734', 'PP735', 'PP736', 'PP737', 'PP738', 'PP739', 'PP740', 'PP741', 'PP742', 'PP743', 'PP744', 'PP745', 'PP746', 'PP747', 'PP748', 'PP749', 'PP750', 'PP751', 'PP752', 'PP753', 'PP754', 'PP755', 'PP756', 'PP757', 'PP758', 'PP759', 'PP760', 'PP761', 'PP762', 'PP763', 'PP764', 'PP765', 'PP766', 'PP767', 'PP768', 'PP769', 'PP770', 'PP771', 'PP772', 'PP773', 'PP774', 'PP775', 'PP776', 'PP777', 'PP778', 'PP779', 'PP780', 'PP781', 'PP782', 'PP783', 'PP784', 'PP785', 'PP786', 'PP787', 'PP788', 'PP789', 'PP790', 'PP791', 'PP792', 'PP793', 'PP794', 'PP795', 'PP796', 'PP797', 'PP798', 'PP799', 'PP800', 'PP801', 'PP802', 'PP803', 'PP804', 'PP805', 'PP806', 'PP807', 'PP808', 'PP809', 'PP810', 'PP811', 'PP812', 'PP813', 'PP814', 'PP815', 'PP816', 'PP817', 'PP818', 'PP819', 'PP820', 'PP821', 'PP822', 'PP823', 'PP824', 'PP825', 'PP826', 'PP827', 'PP828', 'PP829', 'PP830', 'PP831', 'PP832', 'PP833', 'PP834', 'PP835', 'PP836', 'PP837', 'PP838', 'PP839', 'PP840', 'PP841', 'PP842', 'PP843', 'PP844', 'PP845', 'PP846', 'PP847', 'PP848', 'PP849', 'PP850', 'PP851', 'PP852', 'PP853', 'PP854', 'PP855', 'PP856', 'PP857', 'PP858', 'PP859', 'PP860', 'PP861', 'PP862', 'PP863', 'PP864', 'PP865', 'PP866', 'PP867', 'PP868', 'PP869', 'PP870', 'PP871', 'PP872', 'PP873', 'PP874', 'PP875', 'PP876', 'PP877', 'PP878', 'PP879', 'PP880', 'PP881', 'PP882', 'PP883', 'PP884', 'PP885', 'PP886', 'PP887', 'PP888', 'PP889', 'PP890', 'PP891', 'PP892', 'PP893', 'PP894', 'PP895', 'PP896', 'PP897', 'PP898', 'PP899', 'PP900', 'PP901', 'PP902', 'PP903', 'PP904', 'PP905', 'PP906', 'PP907', 'PP908', 'PP909', 'PP910', 'PP911', 'PP912', 'PP913', 'PP914', 'PP915', 'PP916', 'PP917', 'PP918', 'PP919', 'PP920', 'PP921', 'PP922', 'PP923', 'PP924', 'PP925', 'PP926', 'PP927', 'PP928', 'PP929', 'PP930', 'PP931', 'PP932', 'PP933', 'PP934', 'PP935', 'PP936', 'PP937', 'PP938', 'PP939', 'PP940', 'PP941', 'PP942', 'PP943', 'PP944', 'PP945', 'PP946', 'PP947', 'PP948', 'PP949', 'PP950', 'PP951', 'PP952', 'PP953', 'PP954', 'PP955', 'PP956', 'PP957', 'PP958', 'PP959', 'PP960', 'PP961', 'PP962', 'PP963', 'PP964', 'PP965', 'PP966', 'PP967', 'PP968', 'PP969', 'PP970', 'PP971', 'PP972', 'PP973', 'PP974', 'PP975', 'PP976', 'PP977', 'PP978', 'PP979', 'PP980', 'PP981', 'PP982', 'PP983', 'PP984', 'PP985', 'PP986', 'PP987', 'PP988', 'PP989', 'PP990', 'PP991', 'PP992', 'PP993', 'PP994', 'PP995', 'PP996', 'PP997', 'PP998', 'PP999'
                ],

                "Período": [
                    '1° Período', '2° Período', '3° Período', '4° Período', '5° Período', 
                    '6° Período', '7° Período', '8° Período', '9° Período', '10° Período'
                ],

                "Grau De Criticidade Do Apontamento": [
                    "Baixo", "Médio", "Alto"
                ],
                }


                columns_config = {}
                for col in df_view.columns:
                    if col in selectbox_columns_opcoes:
                        columns_config[col] = st.column_config.SelectboxColumn(
                            col, options=selectbox_columns_opcoes[col], disabled=False
                        )
                    elif col in colunas_data:
                        columns_config[col] = st.column_config.DateColumn(col, format="DD/MM/YYYY")
                    elif col == "orig_idx":
                        columns_config[col] = st.column_config.NumberColumn(
                            "ID", disabled=True           # coluna ID fixo, não editável
                        )
                    else:
                        df_view[col] = df_view[col].astype(str).replace("nan", "")
                        columns_config[col] = st.column_config.TextColumn(col)

                columns_config["Data Atualização"] = st.column_config.DateColumn(
                    "Data Atualização", format="DD/MM/YYYY", disabled=True
                )
                columns_config["Responsável Atualização"] = st.column_config.TextColumn(
                    "Responsável Atualização", disabled=True
                )

                # 9️⃣  FOTO IMUTÁVEL p/ comparação ------------------------------------
                snapshot = df_view.copy(deep=True)
                cols_cmp = [c for c in snapshot.columns if c not in ("orig_idx", "Data Atualização", "Responsável Atualização")]

                # 🔟  Formulário e Editor ---------------------------------------------
                with st.form("grade"):
                    responsavel_att = st.selectbox(
                        "Responsável pela Atualização dos dados",
                        ["", "Glaucia Araujo","Guilherme Gonçalves", "Sandra de Souza"],
                        key="resp_att"
                    )

                    df_editado = st.data_editor(
                        snapshot,
                        column_config=columns_config,
                        num_rows="dynamic",
                        key="apontamentos",
                        hide_index=True       # <- oculta índice → só 1 coluna de ID
                    )
                    submitted = st.form_submit_button("Submeter Edições")

                # 11️⃣  Salva apenas se algo mudou ------------------------------------
                if submitted:
                    if responsavel_att.strip() == "":
                        st.warning("Escolha quem é o responsável antes de submeter.")
                        st.stop()

                    # helper normalização
                    def _norm(df_like: pd.DataFrame) -> pd.DataFrame:
                        return (
                            df_like[cols_cmp]
                            .astype(str)
                            .apply(lambda s: s.str.strip().replace("nan", ""))
                        )

                    if _norm(snapshot).equals(_norm(df_editado)):
                        st.toast("Nenhuma alteração detectada. Nada foi salvo!")
                        st.stop()

                    data_atual = datetime.now()

                    snap_idx = snapshot.set_index("orig_idx")
                    edit_idx = df_editado.set_index("orig_idx")

                    # linha vazia → exclusão
                    linhas_para_excluir = edit_idx.index[
                        _norm(edit_idx)[cols_cmp]
                        .replace("", pd.NA)
                        .isna()
                        .all(axis=1)
                    ]
                    if len(linhas_para_excluir) > 0:
                        edit_idx = edit_idx.drop(linhas_para_excluir)

                    removidos = snap_idx.index.difference(edit_idx.index).union(linhas_para_excluir)

                    comuns = snap_idx.index.intersection(edit_idx.index)
                    snap_cmp = _norm(snap_idx.loc[comuns].reset_index())
                    edit_cmp = _norm(edit_idx.loc[comuns].reset_index())
                    diff_mask = snap_cmp.ne(edit_cmp).any(axis=1)
                    linhas_alteradas = edit_idx.loc[comuns].reset_index().loc[diff_mask]

                    idx_modificados = []
                    df[cols_cmp] = df[cols_cmp].astype(object)

                    for _, row in linhas_alteradas.iterrows():
                        orig_idx = int(row["orig_idx"])
                        if not _norm(df.loc[[orig_idx]]).equals(_norm(row.to_frame().T)):
                            status_antigo = str(df.loc[orig_idx, "Status"]).strip().upper()
                            df.loc[orig_idx, cols_cmp] = row[cols_cmp].values

                            novo_status = str(row.get("Status", "")).strip().upper()
                            if novo_status == "VERIFICANDO" and status_antigo != "VERIFICANDO":
                                df.loc[orig_idx, "Disponibilizado para Verificação"] = data_atual

                            idx_modificados.append(orig_idx)

                    mudou = False
                    if idx_modificados:
                        df.loc[idx_modificados, "Data Atualização"] = data_atual
                        df.loc[idx_modificados, "Responsável Atualização"] = responsavel_att.strip()
                        mudou = True

                    if len(removidos) > 0:
                        df.drop(index=list(removidos), inplace=True)
                        mudou = True

                    if mudou:
                        update_sharepoint_file(df.reset_index())  # salva com coluna 'orig_idx'
                        st.cache_data.clear()
                    else:
                        st.toast("Nenhuma alteração detectada. Nada foi salvo!")

#---------------------------------------------------------------------
# Edição de Staff
#---------------------------------------------------------------------

        with tabs[1]:
            st.title("Relação de Vagas")

            # 1) Garante que ID Vaga não está como índice
            for df in (staff_df, colaboradores_df):
                if df.index.name == "ID Vaga":
                    df.reset_index(inplace=True)

            # 2) Colunas-chave
            if "ID Vaga" not in staff_df.columns:
                st.error("'ID Vaga' não está em staff_df")
                st.stop()
            if "ID Vaga" not in colaboradores_df.columns:
                st.error("'ID Vaga' não está em colaboradores_df")
                st.stop()

            # 3) Recalcula staff_df["Ativos"]
            ativos = colaboradores_df[colaboradores_df["Ativos"] == "Sim"]
            contagem = ativos.groupby("ID Vaga").size()
            staff_df["Ativos"] = staff_df["ID Vaga"].map(contagem).fillna(0).astype(int)

            # 4) Define colunas não editáveis
            non_editable_cols = ["Ativos"]
            column_config = {
                col: st.column_config.Column(disabled=True) if col in non_editable_cols else st.column_config.Column()
                for col in staff_df.columns
            }

            # Guarda cópia original para detectar mudanças
            original_df = staff_df.copy()

            # 5) Exibe editor com opção de excluir linhas
            edited_df = st.data_editor(
                staff_df,
                column_config=column_config,
                hide_index=True,
                num_rows="dynamic",  # permite excluir linhas
                use_container_width=True,
                key="staff_editor",
            )

            # 6) Botão SALVAR – salva planilha inteira
            if st.button("Salvar alterações", key="save_staff"):
                # Garante que coluna Ativos permanece correta (não salva alterações manuais)
                edited_df["Ativos"] = edited_df["ID Vaga"].map(contagem).fillna(0).astype(int)

                if edited_df.equals(original_df):
                    st.warning("Nenhuma alteração detectada – nada foi salvo.")
                    st.stop()

                if "Quantidade Staff" in edited_df.columns:
                    edited_df["Quantidade Staff"] = (
                        pd.to_numeric(edited_df["Quantidade Staff"], errors="coerce")
                        .fillna(0)
                        .astype(int)
                    )

                # Grava planilha completa
                update_staff_sheet(edited_df)
                st.cache_data.clear()


if __name__ == "__main__":
    main()