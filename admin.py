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
username           = st.secrets["sharepoint"]["username"]
password           = st.secrets["sharepoint"]["password"]
site_url           = st.secrets["sharepoint"]["site_url"]
file_name          = st.secrets["sharepoint"]["file_name"]
apontamentos_file  = st.secrets["sharepoint"]["apontamentos_file"]
bio_file           = st.secrets["sharepoint"]["bio_file"]


# --------------------------------------------------------------------
# Utilidades gerais
# --------------------------------------------------------------------
@st.cache_data
def read_excel_sheets_from_sharepoint():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, file_name)
        xls = pd.ExcelFile(io.BytesIO(response.content))
        staff_df        = pd.read_excel(xls, sheet_name="Staff Operações Clínica")
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
            ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_editado.to_excel(writer, index=False)
            output.seek(0)

            folder_path    = "/".join(apontamentos_file.split("/")[:-1])
            file_name_only = apontamentos_file.split("/")[-1]
            target_folder  = ctx.web.get_folder_by_server_relative_url(folder_path)
            target_folder.upload_file(file_name_only, output.read()).execute_query()

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
    k_date   = f"ds_data_{colab_key}"
    k_reason = f"ds_reason_{colab_key}"

    if k_date not in st.session_state:
        st.session_state[k_date] = default_date or date.today()
    if k_reason not in st.session_state:
        st.session_state[k_reason] = default_reason

    return k_date, k_reason


def so_digitos(v):
    return re.sub(r"\D", "", str(v))


def clear_cache_and_reload():
    st.cache_data.clear()


# --------------------------------------------------------------------
# Interface principal
# --------------------------------------------------------------------
st.set_page_config(layout="wide")

def main():
    st.title("📋 Painel ADM")
    tabs = st.tabs(["Apontamentos", "Posições", "Atualizar Colaborador", "Novo Colaborador"])

    # -----------------------------------------------------------------
    # TAB ‑ NOVO COLABORADOR
    # -----------------------------------------------------------------
    with tabs[3]:
        spacer_left, main, spacer_right = st.columns([2, 4, 2])
        with main:
            st.title("Cadastrar Colaborador")

            staff_df, colaboradores_df = read_excel_sheets_from_sharepoint()

            if staff_df.empty:
                st.error("Não foi possível carregar a planilha 'Staff Operações Clínica'.")
                return

            id_vagas = sorted(staff_df["ID Vaga"].dropna().unique())
            id_vaga  = st.selectbox("ID Vaga", id_vagas)

            vaga_info   = staff_df.loc[staff_df["ID Vaga"] == id_vaga].iloc[0]
            disponiveis = vaga_info["Quantidade Staff"] - vaga_info["Ativos"]
            st.text_input("Vagas Disponíveis", disponiveis, disabled=True)
            st.markdown("---")

            nome = st.text_input("Nome Completo do colaborador")
            cpf  = str(st.text_input("CPF ou CNPJ", placeholder="Apenas números"))

            st.text_input("Cargo", vaga_info["Cargo"], disabled=True)
            st.text_input("Turma", vaga_info["Departamento"], disabled=True)
            st.text_input("Escala", vaga_info["Escala"], disabled=True)
            st.text_input("Horário", vaga_info["Horário"], disabled=True)
            st.text_input("Turma", vaga_info["Turma"], disabled=True)
            st.text_input("Plantão", vaga_info["Plantão"], disabled=True)
            st.text_input("Supervisão", vaga_info["Supervisora"], disabled=True)

            entrada = st.date_input("Data da Entrada", format="DD/MM/YYYY")

            contrato = st.selectbox("Tipo de Contrato", ["CLT", "Autonomo", "Horista"])

            responsavel = st.text_input("Responsável pela Inclusão dos dados")

            if st.button("Enviar"):
                if not nome.strip() or not responsavel.strip() or not cpf.strip():
                    st.error("Preencha os campos obrigatórios: Nome, Supervisão Direta e Responsável.")
                    return

                colab_cpfs = colaboradores_df["CPF ou CNPJ"].apply(so_digitos)
                if cpf in colab_cpfs.values:
                    st.error("Já existe um colaborador cadastrado com este CPF/CNPJ.")
                    return

                max_colabs = int(vaga_info["Quantidade Staff"])
                status_col = "Ativos"

                filtro_colab = colaboradores_df[
                    (colaboradores_df["ID Vaga"] == id_vaga) &
                    (colaboradores_df[status_col] == "Sim")
                ]

                if filtro_colab.shape[0] >= max_colabs:
                    st.error(f"Limite de colaboradores atingido para essa vaga: {max_colabs}")
                    return

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

                update_colaboradores_sheet(colaboradores_df)
                st.cache_data.clear()

    # -----------------------------------------------------------------
    # TAB ‑ ATUALIZAR COLABORADOR
    # -----------------------------------------------------------------
    with tabs[2]:
        spacer_left, main_col, spacer_right = st.columns([2, 4, 2])
        with main_col:
            staff_df, colaboradores_df = read_excel_sheets_from_sharepoint()

            if colaboradores_df.empty:
                st.info("Não há colaboradores na base")
                st.stop()

            st.title("Atualizar Colaborador")

            nomes = colaboradores_df["Nome Completo do Profissional"].dropna().sort_values().unique()
            selec_nome = st.selectbox("Selecione o colaborador", nomes, key="sel_colab")

            linha     = colaboradores_df.loc[colaboradores_df["Nome Completo do Profissional"] == selec_nome].iloc[0]
            old_id_vaga = linha.get("ID Vaga", "")
            old_ativo    = linha.get("Ativos", "Não")

            id_vagas  = sorted(staff_df["ID Vaga"].dropna().unique())
            idx_vaga  = id_vagas.index(old_id_vaga) if old_id_vaga in id_vagas else 0
            id_vaga   = st.selectbox("ID Vaga", id_vagas, index=idx_vaga, key=f"idvaga_{selec_nome}")

            vaga_info = staff_df.loc[staff_df["ID Vaga"] == id_vaga].iloc[0]

            nome   = st.text_input("Nome Completo do Profissional", linha["Nome Completo do Profissional"], key=f"nome_{selec_nome}")
            cpf    = st.text_input("CPF ou CNPJ", linha["CPF ou CNPJ"], key=f"cpf_{selec_nome}")

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

            st.text_input("Departamento", vaga_info["Departamento"], disabled=True, key=f"dep_{selec_nome}")
            st.text_input("Cargo", vaga_info["Cargo"], disabled=True, key=f"cargo_{selec_nome}")
            st.text_input("Escala", vaga_info["Escala"], disabled=True, key=f"escala_{selec_nome}")
            st.text_input("Horário", vaga_info["Horário"], disabled=True, key=f"hora_{selec_nome}")
            st.text_input("Turma", vaga_info["Turma"], disabled=True, key=f"turma_{selec_nome}")
            st.text_input("Supervisão Direta", vaga_info["Supervisora"], disabled=True, key=f"sup_{selec_nome}")
            st.text_input("Plantão", vaga_info["Plantão"], disabled=True, key=f"plantao_{selec_nome}")

            max_colabs      = int(vaga_info["Quantidade Staff"])
            ativos_na_vaga  = colaboradores_df[
                (colaboradores_df["ID Vaga"] == id_vaga) &
                (colaboradores_df["Ativos"] == "Sim")
            ].shape[0]

            disponiveis = max_colabs - ativos_na_vaga
            st.info(f"Disponíveis: {disponiveis} / {max_colabs}")
            if disponiveis <= 0 and status_prof != "Desligado":
                st.warning("Esta vaga está lotada. Só será possível se marcar o colaborador como 'Desligado'.", icon="⚠️")

            data_deslig  = linha.get("Data Desligamento", None)
            motivo_clt   = linha.get("Desligamento CLT", "")
            motivo_auto  = linha.get("Saída Autonomo", "")

            if status_prof == "Desligado":
                key_date, key_reason = get_deslig_state(
                    selec_nome,
                    linha.get("Atualização", datetime.now()).date(),
                    motivo_clt or motivo_auto,
                )

                data_deslig = st.date_input("Data do desligamento", format="DD/MM/YYYY")

                if tipo_contrato.lower() == "clt":
                    lista_clt = ["Solicitação de Desligamento", "Desligamento pela Gestão"]
                    if st.session_state.get(key_reason) not in lista_clt:
                        st.session_state[key_reason] = lista_clt[0]
                    motivo_clt  = st.selectbox("Motivo do desligamento (CLT)", lista_clt, key=key_reason)
                    motivo_auto = ""
                elif tipo_contrato.lower() == "autonomo":
                    lista_auto = ["Distrato", "Solicitação de Distrato", "Distrato pela Gestão"]
                    if st.session_state.get(key_reason) not in lista_auto:
                        st.session_state[key_reason] = lista_auto[0]
                    motivo_auto = st.selectbox("Motivo do distrato (Autônomo)", lista_auto, key=key_reason)
                    motivo_clt  = ""
                else:
                    motivo_clt = motivo_auto = ""

            if st.button("Salvar alterações", key=f"save_{selec_nome}"):
                if not responsavel_att.strip():
                    st.error("Preencha o campo Responsável pela Atualização dos dados.")
                    st.stop()

                novo_ativo = "Não" if status_prof == "Desligado" else "Sim"

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

                if status_prof == "Desligado":
                    colaboradores_df.loc[linha.name, "Ativos"] = "Não"
                    colaboradores_df.loc[linha.name, "Data Desligamento"] = data_deslig
                    if tipo_contrato.lower() == "clt":
                        colaboradores_df.loc[linha.name, "Desligamento CLT"] = motivo_clt
                        colaboradores_df.loc[linha.name, "Saída Autonomo"]   = ""
                    elif tipo_contrato.lower() == "autonomo":
                        colaboradores_df.loc[linha.name, "Saída Autonomo"]   = motivo_auto
                        colaboradores_df.loc[linha.name, "Desligamento CLT"] = ""
                    else:
                        colaboradores_df.loc[linha.name, ["Desligamento CLT", "Saída Autonomo"]] = ""

                update_colaboradores_sheet(colaboradores_df)
                update_staff_sheet(staff_df)

    # -----------------------------------------------------------------
    # TAB ‑ APONTAMENTOS (REFATORADO COM COLUNA 'ID')
    # -----------------------------------------------------------------
    
    with tabs[0]:
        st.title("Lista de Apontamentos")

        df = get_sharepoint_file()

        # 1) Garante coluna ID fixa --------------------------------------------------
        if "ID" not in df.columns:
            df.insert(0, "ID", range(1, len(df) + 1))  # primeira criação

        # ID deve ser numérico (Int64 para permitir NA em novas linhas)
        df["ID"] = pd.to_numeric(df["ID"], errors="coerce").astype("Int64")

        # 2) Converte colunas de data ------------------------------------------------
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

        # 3) Cópia para filtros ------------------------------------------------------
        df_filtrado = df.copy()

        # --- Botões rápidos (pendente / verificando) -------------------------------
        def toggle_pending():
            st.session_state.show_pending = not st.session_state.get("show_pending", False)
            st.session_state.show_verificando = False

        def toggle_verificando():
            st.session_state.show_verificando = not st.session_state.get("show_verificando", False)
            st.session_state.show_pending = False

        st.session_state.setdefault("show_pending", False)
        st.session_state.setdefault("show_verificando", False)

        col_btn1, col_btn2, col_btn3, *_ = st.columns(6)

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
            st.button("🔄  Atualizar", key="btn_clear_cache", on_click=clear_cache_and_reload)

        # 4) Filtro por Código do Estudo --------------------------------------------
        columns_to_display = [
            "Status", "Código do Estudo", "Responsável Pela Correção", "Plantão",
            "Participante", "Período", "Grau De Criticidade Do Apontamento",
            "Documentos", "Apontamento", "Data do Apontamento",
            "Disponibilizado para Verificação", "Prazo Para Resolução",
            "Data Resolução", "Justificativa", "Responsável Pelo Apontamento",
            "Origem Do Apontamento", "Data Atualização", "Responsável Atualização"
        ]
        df_filtrado = df_filtrado[["ID"] + columns_to_display]

        if st.session_state.show_pending:
            df_view = df_filtrado[df_filtrado["Status"] == "PENDENTE"].copy()
        elif st.session_state.show_verificando:
            df_view = df_filtrado[df_filtrado["Status"] == "VERIFICANDO"].copy()
        else:
            df_view = df_filtrado.copy()

        if "Código do Estudo" in df.columns:
            opcoes_estudos = ["Todos"] + sorted(df["Código do Estudo"].dropna().unique())
            estudo_sel = st.selectbox("Selecione o Estudo", opcoes_estudos, key="estudo_selecionado")

            if estudo_sel != "Todos":
                df_view = df_view[df_view["Código do Estudo"] == estudo_sel]

        # 5) Move coluna ID para primeira posição e cria snapshot --------------------
        first = df_view.pop("ID")
        df_view.insert(0, "ID", first)

        selectbox_columns_opcoes = {
            "Status": [
                "REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "VERIFICANDO", "PENDENTE", "NÃO APLICÁVEL"
            ],
            "Origem Do Apontamento": [
                "Documentação Clínica", "Excelência Operacional", "Operações Clínicas",
                "Patrocinador / Monitor", "Garantia Da Qualidade"
            ],
            "Participante": [
                'N/A','PP01', 'PP02', 'PP03', 'PP04', 'PP05', 'PP06', 'PP07', 'PP08', 'PP09', 'PP10',
                'PP11', 'PP12', 'PP13', 'PP14', 'PP15', 'PP16', 'PP17', 'PP18', 'PP19', 'PP20',
                'PP990', 'PP991', 'PP992', 'PP993', 'PP994', 'PP995', 'PP996', 'PP997', 'PP998', 'PP999'
            ],
            "Período": [
                '1° Período', '2° Período', '3° Período', '4° Período', '5° Período',
                '6° Período', '7° Período', '8° Período', '9° Período', '10° Período'
            ],
            "Grau De Criticidade Do Apontamento": ["Baixo", "Médio", "Alto"],
        }

        columns_config = {}
        for col in df_view.columns:
            if col in selectbox_columns_opcoes:
                columns_config[col] = st.column_config.SelectboxColumn(
                    col, options=selectbox_columns_opcoes[col], disabled=False
                )
            elif col in colunas_data:
                columns_config[col] = st.column_config.DateColumn(col, format="DD/MM/YYYY")
            elif col == "ID":
                columns_config[col] = st.column_config.NumberColumn("ID", disabled=True)
            else:
                df_view[col] = df_view[col].astype(str).replace("nan", "")
                columns_config[col] = st.column_config.TextColumn(col)

        columns_config["Data Atualização"] = st.column_config.DateColumn(
            "Data Atualização", format="DD/MM/YYYY", disabled=True
        )
        columns_config["Responsável Atualização"] = st.column_config.TextColumn(
            "Responsável Atualização", disabled=True
        )

        snapshot = df_view.copy(deep=True)
        cols_cmp = [c for c in snapshot.columns if c not in ("ID", "Data Atualização", "Responsável Atualização")]

        with st.form("grade"):
            responsavel_att = st.selectbox(
                "Responsável pela Atualização dos dados",
                ["", "Glaucia Araujo", "Guilherme Gonçalves", "Sandra de Souza"],
                key="resp_att"
            )

            df_editado = st.data_editor(
                snapshot,
                column_config=columns_config,
                num_rows="dynamic",
                key="apontamentos",
                hide_index=True
            )
            submitted = st.form_submit_button("Submeter Edições")

        # -------------------------------------------------------------------------
        # PROCESSAMENTO DAS ALTERAÇÕES
        # -------------------------------------------------------------------------
        if submitted:
            if responsavel_att.strip() == "":
                st.warning("Escolha quem é o responsável antes de submeter.")
                st.stop()

            # 1) Converte IDs novos vazios para NA  ---------------------------------
            df_editado["ID"] = pd.to_numeric(df_editado["ID"], errors="coerce").astype("Int64")

            # 2) Gera IDs para linhas recém‑criadas --------------------------------
            if df["ID"].notna().any():
                proximo_id = int(df["ID"].max()) + 1
            else:
                proximo_id = 1

            linhas_sem_id = df_editado["ID"].isna()
            qtd_novas = linhas_sem_id.sum()
            if qtd_novas:
                df_editado.loc[linhas_sem_id, "ID"] = range(proximo_id, proximo_id + qtd_novas)

            df_editado["ID"] = df_editado["ID"].astype(int)

            # 3) Função de normalização --------------------------------------------
            def _norm(df_like: pd.DataFrame) -> pd.DataFrame:
                return (
                    df_like[cols_cmp]
                    .astype(str)
                    .apply(lambda s: s.str.strip().replace("nan", ""))
                )

            # 4) Detecção de alterações ---------------------------------------------
            # Verifica se há novas linhas pelo número de linhas
            if len(df_editado) <= len(snapshot) and _norm(snapshot).equals(_norm(df_editado)):
                st.toast("Nenhuma alteração detectada. Nada foi salvo!")
                st.stop()

            data_atual = datetime.now()

            # Reindexação para comparação
            snap_idx = snapshot.set_index("ID")
            edit_idx = df_editado.set_index("ID")

            # Linhas marcadas como vazias → será considerado exclusão ----------------
            linhas_vazias = edit_idx.index[
                _norm(edit_idx)[cols_cmp].replace("", pd.NA).isna().all(axis=1)
            ]
            if len(linhas_vazias) > 0:
                edit_idx = edit_idx.drop(linhas_vazias)

            removidos = snap_idx.index.difference(edit_idx.index).union(linhas_vazias)

            # Linhas em comum alteradas ---------------------------------------------
            comuns = snap_idx.index.intersection(edit_idx.index)
            snap_cmp = _norm(snap_idx.loc[comuns].reset_index())
            edit_cmp = _norm(edit_idx.loc[comuns].reset_index())
            diff_mask = snap_cmp.ne(edit_cmp).any(axis=1)
            linhas_alt = edit_idx.loc[comuns].reset_index().loc[diff_mask]

            idx_modificados = []
            df[cols_cmp] = df[cols_cmp].astype(object)

            # Processa novas linhas
            novas_linhas = edit_idx.index.difference(snap_idx.index)
            if len(novas_linhas) > 0:
                novas = edit_idx.loc[novas_linhas].reset_index()
                for _, row in novas.iterrows():
                    df = pd.concat([df, row.to_frame().T], ignore_index=True)
                    idx_modificados.append(int(row["ID"]))

            # Processa linhas alteradas
            for _, row in linhas_alt.iterrows():
                rid = int(row["ID"])
                if not _norm(df.loc[df["ID"] == rid]).equals(_norm(row.to_frame().T)):
                    status_ant = str(df.loc[df["ID"] == rid, "Status"].iloc[0]).strip().upper()
                    df.loc[df["ID"] == rid, cols_cmp] = row[cols_cmp].values

                    novo_status = str(row.get("Status", "")).strip().upper()
                    if novo_status == "VERIFICANDO" and status_ant != "VERIFICANDO":
                        df.loc[df["ID"] == rid, "Disponibilizado para Verificação"] = data_atual

                    idx_modificados.append(rid)

            mudou = False
            if idx_modificados:
                df.loc[df["ID"].isin(idx_modificados), "Data Atualização"] = data_atual
                df.loc[df["ID"].isin(idx_modificados), "Responsável Atualização"] = responsavel_att.strip()
                mudou = True

            if len(removidos) > 0:
                df = df[~df["ID"].isin(removidos)]
                mudou = True

            if mudou:
                update_sharepoint_file(df.reset_index(drop=True))
                st.cache_data.clear()
            else:
                st.toast("Nenhuma alteração detectada. Nada foi salvo!")

    # -----------------------------------------------------------------
    # TAB ‑ POSIÇÕES (STAFF)
    # -----------------------------------------------------------------
    with tabs[1]:
        st.title("Relação de Vagas")

        staff_df, colaboradores_df = read_excel_sheets_from_sharepoint()

        for df_temp in (staff_df, colaboradores_df):
            if df_temp.index.name == "ID Vaga":
                df_temp.reset_index(inplace=True)

        if "ID Vaga" not in staff_df.columns:
            st.error("'ID Vaga' não está em staff_df")
            st.stop()
        if "ID Vaga" not in colaboradores_df.columns:
            st.error("'ID Vaga' não está em colaboradores_df")
            st.stop()

        ativos      = colaboradores_df[colaboradores_df["Ativos"] == "Sim"]
        contagem    = ativos.groupby("ID Vaga").size()
        staff_df["Ativos"] = staff_df["ID Vaga"].map(contagem).fillna(0).astype(int)

        non_editable_cols = ["Ativos"]
        column_config = {
            col: st.column_config.Column(disabled=True) if col in non_editable_cols
            else st.column_config.Column()
            for col in staff_df.columns
        }

        original_df = staff_df.copy()

        edited_df = st.data_editor(
            staff_df,
            column_config=column_config,
            hide_index=True,
            num_rows="dynamic",
            use_container_width=True,
            key="staff_editor",
        )

        if st.button("Salvar alterações", key="save_staff"):
            edited_df["Quantidade Staff"] = (
                pd.to_numeric(edited_df["Quantidade Staff"], errors="coerce")
                .fillna(0)
                .astype(int)
            )
            edited_df["Ativos"] = edited_df["ID Vaga"].map(contagem).fillna(0).astype(int)

            if edited_df.equals(original_df):
                st.warning("Nenhuma alteração detectada – nada foi salvo.")
                st.stop()

            update_staff_sheet(edited_df)
            st.cache_data.clear()


if __name__ == "__main__":
    main()
