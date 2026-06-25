import streamlit as st
import pandas as pd
import requests
import time
import re
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CNPJ Enterprise Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicialização de Variáveis de Estado
if "rodando" not in st.session_state: st.session_state.rodando = False
if "logs" not in st.session_state: st.session_state.logs = []

ARQUIVO_SAIDA = "RESULTADOS_CNPJ.xlsx"

# --- FUNÇÕES AUXILIARES ---
def limpar_cnpj(cnpj):
    so_numeros = re.sub(r'\D', '', str(cnpj))
    return so_numeros.zfill(14) if len(so_numeros) > 0 else ""

def _get(dic, *chaves):
    for chave in chaves:
        if isinstance(dic, dict): dic = dic.get(chave, "")
        else: return ""
    return dic if dic != {} else ""

def formatar_eta(segundos):
    if segundos <= 0: return "00:00:00"
    dias = int(segundos // 86400)
    resto = segundos % 86400
    horas = int(resto // 3600)
    resto %= 3600
    minutos = int(resto // 60)
    segs = int(resto % 60)
    if dias > 0: return f"{dias}d {horas:02d}h {minutos:02d}m {segs:02d}s"
    return f"{horas:02d}h {minutos:02d}m {segs:02d}s"

def extrair_dados_json(dados):
    estab = dados.get('estabelecimento', {})
    socios = " \n; ".join([f"Nome: {s.get('nome')} | Doc: {s.get('cpf_cnpj_socio')} | Qualificação: {_get(s, 'qualificacao_socio', 'descricao')}" for s in dados.get('socios', [])])
    secundarias = " \n; ".join([f"{act.get('subclasse')} - {act.get('descricao')}" for act in estab.get('atividades_secundarias', [])])
    
    linha = {
        "CNPJ Raiz": _get(dados, 'cnpj_raiz'), "Razão Social": _get(dados, 'razao_social'),
        "Capital Social": _get(dados, 'capital_social'), "Responsável Federativo": _get(dados, 'responsavel_federativo'),
        "Atualizado Em (Geral)": _get(dados, 'atualizado_em'), "Porte Descrição": _get(dados, 'porte', 'descricao'), 
        "Natureza Jurídica Descrição": _get(dados, 'natureza_juridica', 'descricao'), "Simples Nacional / MEI": _get(dados, 'simples'),
        "CNPJ Completo": _get(estab, 'cnpj'), "Situação Cadastral": _get(estab, 'situacao_cadastral'),
        "Data Situação Cadastral": _get(estab, 'data_situacao_cadastral'), "Data Início Atividade": _get(estab, 'data_inicio_atividade'),
        "Logradouro": _get(estab, 'logradouro'), "Número": _get(estab, 'numero'), "Bairro": _get(estab, 'bairro'),
        "CEP": _get(estab, 'cep'), "Telefone": _get(estab, 'telefone1'), "E-mail": _get(estab, 'email'), 
        "Atividade Principal": _get(estab, 'atividade_principal', 'descricao'), "Estado (UF)": _get(estab, 'estado', 'sigla'), 
        "Cidade": _get(estab, 'cidade', 'nome'), "Atividades Secundárias": secundarias, "Quadro de Sócios": socios
    }

    uf_sede = _get(estab, 'estado', 'sigla')
    lista_ies_ativas = [ie for ie in estab.get('inscricoes_estaduais', []) if ie.get('ativo') is True]

    if not lista_ies_ativas:
        linha[f"IE {uf_sede}" if uf_sede else "IE PRINCIPAL"] = "ISENTO"
    else:
        for ie in lista_ies_ativas:
            sigla_estado = _get(ie, 'estado', 'sigla')
            if not sigla_estado: continue
            num_ie = re.sub(r'\D', '', str(ie.get('inscricao_estadual', '')))
            nome_coluna = f"IE {sigla_estado}"
            if num_ie:
                if nome_coluna in linha: 
                    if num_ie not in linha[nome_coluna].split(" ; "):
                        linha[nome_coluna] += f" ; {num_ie}"
                else: 
                    linha[nome_coluna] = num_ie
    return linha

# --- MENU LATERAL DE NAVEGAÇÃO ---
st.sidebar.title("Navegação")
modo_app = st.sidebar.radio("Escolha o Módulo:", ["🗂️ Consulta em Lote (Excel)", "📇 Consulta Única (Cartão)"])
st.sidebar.markdown("---")


# =========================================================================
# MÓDULO 1: CONSULTA ÚNICA (CARTÃO DE CNPJ)
# =========================================================================
if modo_app == "📇 Consulta Única (Cartão)":
    st.title("📇 Cartão de Consulta CNPJ")
    st.markdown("Consulte uma única empresa rapidamente e exporte a ficha cadastral.")

    cnpj_input = st.text_input("Digite o CNPJ (com ou sem pontuação):", max_chars=18)
    
    if st.button("🔍 Pesquisar Empresa", type="primary"):
        cnpj_limpo = limpar_cnpj(cnpj_input)
        if len(cnpj_limpo) != 14:
            st.warning("⚠️ O CNPJ deve conter 14 dígitos válidos.")
        else:
            with st.spinner("Buscando dados na Receita..."):
                try:
                    res = requests.get(f"https://publica.cnpj.ws/cnpj/{cnpj_limpo}", timeout=15)
                    if res.status_code == 200:
                        dados_empresa = extrair_dados_json(res.json())
                        
                        st.success("✅ Empresa Localizada!")
                        
                        # Renderização do "Cartão" Visual
                        with st.container(border=True):
                            st.subheader(f"🏢 {dados_empresa['Razão Social']}")
                            st.write(f"**CNPJ:** {dados_empresa['CNPJ Completo']} | **Situação:** {dados_empresa['Situação Cadastral']} ({dados_empresa['Data Situação Cadastral']})")
                            st.markdown("---")
                            
                            col1, col2, col3 = st.columns(3)
                            col1.metric("Capital Social", f"R$ {dados_empresa['Capital Social']}")
                            col2.metric("Porte", dados_empresa['Porte Descrição'])
                            col3.metric("Natureza", dados_empresa['Natureza Jurídica Descrição'])
                            
                            st.markdown("---")
                            st.markdown(f"📍 **Endereço:** {dados_empresa['Logradouro']}, {dados_empresa['Número']} - {dados_empresa['Bairro']}, {dados_empresa['Cidade']}/{dados_empresa['Estado (UF)']} - CEP: {dados_empresa['CEP']}")
                            st.markdown(f"📞 **Contato:** {dados_empresa['Telefone']} | ✉️ {dados_empresa['E-mail']}")
                            st.markdown(f"🛠️ **Atividade Principal:** {dados_empresa['Atividade Principal']}")
                            
                            # Botão para baixar apenas essa empresa
                            df_unica = pd.DataFrame([dados_empresa])
                            # Cria o arquivo na memória para não misturar com o lote
                            from io import BytesIO
                            buffer = BytesIO()
                            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                df_unica.to_excel(writer, index=False)
                            
                            st.download_button(
                                label="📥 Baixar Cartão em Excel",
                                data=buffer.getvalue(),
                                file_name=f"CNPJ_{cnpj_limpo}.xlsx",
                                mime="application/vnd.ms-excel"
                            )
                    elif res.status_code == 429:
                        st.error("⚠️ Limite de requisições excedido. Aguarde 1 minuto e tente novamente.")
                    else:
                        st.error("❌ CNPJ Inválido ou não encontrado na Receita Federal.")
                except Exception as e:
                    st.error(f"Erro de conexão: {e}")

# =========================================================================
# MÓDULO 2: CONSULTA EM LOTE (EXCEL)
# =========================================================================
elif modo_app == "🗂️ Consulta em Lote (Excel)":
    st.title("📊 Consulta Automática em Lote")
    
    st.sidebar.title("⚙️ Origem dos Dados")
    arquivo_carregado = st.sidebar.file_uploader("Suba a planilha Excel (Entrada)", type=["xlsx", "xls"])
    df_entrada = None
    coluna_selecionada = None

    if arquivo_carregado:
        try:
            excel_file = pd.ExcelFile(arquivo_carregado)
            aba = st.sidebar.selectbox("Escolha a ABA:", excel_file.sheet_names)
            df_entrada = pd.read_excel(arquivo_carregado, sheet_name=aba)
            coluna_selecionada = st.sidebar.selectbox("Qual COLUNA possui os CNPJs?", df_entrada.columns)
        except Exception as e:
            st.error(f"Erro ao carregar arquivo: {e}")

    # Lógica do Botão de Download Constante
    st.sidebar.markdown("---")
    st.sidebar.subheader("💾 Download de Backup")
    if os.path.exists(ARQUIVO_SAIDA):
        with open(ARQUIVO_SAIDA, "rb") as f:
            st.sidebar.download_button("📥 Baixar Planilha Salva no Servidor", f, file_name="LOTE_RESULTADOS_CNPJ.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

    # --- O truque do Botão que "Some" ---
    area_botoes = st.empty() # Cria um espaço vazio reservado
    
    if not st.session_state.rodando:
        # Se NÃO está rodando, o botão Iniciar aparece dentro do espaço reservado
        if area_botoes.button("▶️ Iniciar Processamento", disabled=(df_entrada is None), use_container_width=True):
            st.session_state.rodando = True
            st.session_state.logs = ["🚀 Processamento iniciado..."]
            st.rerun()
    else:
        # Se ESTÁ rodando, substituímos o botão pelo aviso vermelho
        area_botoes.error("⏳ Processamento em andamento... (Para interromper, clique em 'Stop' no canto superior direito do site)")

    # --- LÓGICA DE PROCESSAMENTO EM LOTE ---
    if st.session_state.rodando:
        lista_bruta = df_entrada[coluna_selecionada].dropna().tolist()
        resultados_atuais = []
        cnpjs_processados = set()

        # Anti-queda: Tenta carregar progresso anterior salvo no servidor
        if os.path.exists(ARQUIVO_SAIDA):
            try:
                df_existente = pd.read_excel(ARQUIVO_SAIDA)
                resultados_atuais = df_existente.to_dict('records')
                if 'CNPJ Completo' in df_existente.columns:
                    cnpjs_processados.update([limpar_cnpj(c) for c in df_existente['CNPJ Completo'].dropna()])
                st.session_state.logs.append(f"🔄 Sistema Anti-Queda: {len(cnpjs_processados)} CNPJs carregados da memória. Retomando de onde parou.")
            except: pass

        pendentes = list(dict.fromkeys([limpar_cnpj(x) for x in lista_bruta if len(limpar_cnpj(x)) == 14 and limpar_cnpj(x) not in cnpjs_processados]))
        total = len(pendentes)

        if total == 0:
            st.success("✅ Todos os CNPJs desta planilha já foram processados!")
            st.session_state.rodando = False
            st.rerun()
        else:
            # Layout de Métricas Visuais
            m1, m2, m3, m4 = st.columns(4)
            metric_sucesso = m1.empty()
            metric_falha = m2.empty()
            metric_req_min = m3.empty()
            metric_t_medio = m4.empty()
            
            progress_bar = st.progress(0)
            eta_placeholder = st.empty()
            log_placeholder = st.empty()

            session = requests.Session()
            session.headers.update({"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
            
            sucessos = 0
            falhas = 0
            t_inicio_global = time.time()
            tempo_ciclo_fixo = 20.5

            for index, cnpj in enumerate(pendentes):
                if not st.session_state.rodando: break
                
                t_req_inicio = time.time()
                sucesso_req = False
                status_erro = ""

                for tentativa in range(3):
                    try:
                        res = session.get(f"https://publica.cnpj.ws/cnpj/{cnpj}", timeout=15)
                        if res.status_code == 200:
                            dados = extrair_dados_json(res.json())
                            resultados_atuais.append(dados)
                            sucesso_req = True
                            sucessos += 1
                            st.session_state.logs.append(f"✅ {dados.get('Razão Social', cnpj)}")
                            break
                        elif res.status_code == 400:
                            status_erro = "Inexistente/Inválido"
                            break
                        elif res.status_code == 429:
                            r_after = int(res.headers.get("Retry-After", 25))
                            st.session_state.logs.append(f"⚠️ Limite API. Pausa de {r_after}s...")
                            time.sleep(r_after)
                        else:
                            status_erro = f"Erro {res.status_code}"
                            time.sleep(2)
                    except:
                        status_erro = "Falha de Rede"
                        time.sleep(2)

                if not sucesso_req:
                    falhas += 1
                    st.session_state.logs.append(f"❌ CNPJ {cnpj}: {status_erro}")
                    if status_erro != "Inexistente/Inválido":
                        resultados_atuais.append({"CNPJ Completo": cnpj, "Razão Social": f"FALHA ({status_erro})", "Situação Cadastral": "ERRO"})

                # --- Atualiza o visual em tempo real ---
                processados_agora = index + 1
                t_decorrido = max(time.time() - t_inicio_global, 1)
                req_min = processados_agora / (t_decorrido / 60)
                t_medio = t_decorrido / processados_agora
                s_restantes = (total - processados_agora) * max(t_medio, tempo_ciclo_fixo)

                metric_sucesso.metric("Sucessos ✅", str(sucessos))
                metric_falha.metric("Falhas ❌", str(falhas))
                metric_req_min.metric("Velocidade ⚡", f"{req_min:.1f} req/m")
                metric_t_medio.metric("Tempo Médio ⏱️", f"{t_medio:.1f}s")

                progress_bar.progress(processados_agora / total)
                eta_placeholder.write(f"**Progresso:** {processados_agora}/{total} | **Faltam:** {formatar_eta(s_restantes)}")
                
                # Exibe apenas os últimos 5 logs para a tela não ficar gigante
                log_placeholder.text_area("Logs em Tempo Real", value="\n".join(st.session_state.logs[-5:]), height=150)

                # Salva o progresso físico no disco do servidor a cada 5 CNPJs (Garante que não perca dados se a net cair)
                if sucesso_req and (index % 5 == 0 or index == total - 1):
                    try: pd.DataFrame(resultados_atuais).to_excel(ARQUIVO_SAIDA, index=False)
                    except: pass

                # Trava ética e matemática para não estourar as 3 requisições por minuto da Receita
                if index < total - 1:
                    time.sleep(tempo_ciclo_fixo)

            pd.DataFrame(resultados_atuais).to_excel(ARQUIVO_SAIDA, index=False)
            st.success("🎉 Processamento 100% Concluído! O arquivo final já está pronto para download na barra lateral esquerda.")
            st.session_state.rodando = False
            st.rerun()
