import streamlit as st
import pandas as pd
import requests
import time
import re
import os
import datetime


# 1. Configuração da Página
st.set_page_config(
    page_title="CNPJ Enterprise Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicialização segura do Session State
if "rodando" not in st.session_state:
    st.session_state.rodando = False
if "resultados" not in st.session_state:
    st.session_state.resultados = []
if "logs" not in st.session_state:
    st.session_state.logs = []
if "pasta_destino" not in st.session_state:
    st.session_state.pasta_destino = "" # Começa vazio até o usuário escolher

# Nome do arquivo de backup interno
ARQUIVO_BACKUP = "backup_interno_cnpj.xlsx"

# --- HELPERS DE TRATAMENTO DE DADOS ---
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

# --- INTERFACE DE CONFIGURAÇÃO (SIDEBAR) ---
st.sidebar.title("⚙️ Origem dos Dados")
arquivo_carregado = st.sidebar.file_uploader("Suba a planilha Excel (Entrada)", type=["xlsx", "xls"])

aba_selecionada = None
coluna_selecionada = None
df_entrada = None

if arquivo_carregado is not None:
    try:
        excel_file = pd.ExcelFile(arquivo_carregado)
        aba_selecionada = st.sidebar.selectbox("Em qual ABA estão os dados?", excel_file.sheet_names)
        
        if aba_selecionada is not None:
            df_entrada = pd.read_excel(arquivo_carregado, sheet_name=aba_selecionada)
            coluna_selecionada = st.sidebar.selectbox("Qual COLUNA possui os CNPJs?", df_entrada.columns)
    except Exception as e:
        st.sidebar.error(f"Erro ao ler o arquivo: {e}")

# --- SUBSTITUA A SUA SIDEBAR POR ESTE BLOCO LIMPO ---
st.sidebar.title("💾 Destino dos Resultados")
st.sidebar.info("Ao finalizar o processamento, o botão de download aparecerá aqui automaticamente.")

# Botão de Download Dinâmico (Lê o arquivo gerado e entrega ao usuário)
# O nome do arquivo gerado agora é constante para facilitar
NOME_ARQUIVO_FINAL = "RESULTADOS_CNPJ.xlsx"

if os.path.exists(NOME_ARQUIVO_FINAL):
    with open(NOME_ARQUIVO_FINAL, "rb") as f:
        bytes_excel = f.read()
    st.sidebar.download_button(
        label="📥 Baixar Planilha Consolidada",
        data=bytes_excel,
        file_name=NOME_ARQUIVO_FINAL,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

nome_arquivo_saida = st.sidebar.text_input("Nome do arquivo de saída", value="RESULTADOS_CNPJ.xlsx")
caminho_final_salvamento = NOME_ARQUIVO_FINAL  # sempre "RESULTADOS_CNPJ.xlsx"


st.sidebar.markdown("---")
# Botão de Download Alternativo
if os.path.exists(caminho_final_salvamento):
    with open(caminho_final_salvamento, "rb") as f:
        bytes_excel = f.read()
    st.sidebar.download_button(
        label="📥 Baixar Planilha",
        data=bytes_excel,
        file_name=nome_arquivo_saida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

# --- CORPO PRINCIPAL DO DASHBOARD ---
st.title("📊 Painel Avançado de Consulta - API CNPJ")
st.markdown("Automação corporativa resiliente para enriquecimento cadastral em massa.")

# Painel de Métricas
m1, m2, m3, m4 = st.columns(4)
metric_sucesso = m1.empty()
metric_falha = m2.empty()
metric_req_min = m3.empty()
metric_t_medio = m4.empty()

metric_sucesso.metric("Sucessos ✅", "0")
metric_falha.metric("Falhas ❌", "0")
metric_req_min.metric("Velocidade ⚡", "0.0 req/min")
metric_t_medio.metric("Tempo Médio ⏱️", "0.0s")

progress_placeholder = st.empty()
eta_placeholder = st.empty()

tab_dados, tab_logs = st.tabs(["👁️ Visualização de Dados (Preview)", "📜 Logs de Execução"])
preview_placeholder = tab_dados.empty()
log_placeholder = tab_logs.empty()

# Botões de Ação
col1, col2 = st.columns([1, 4])
# Só libera o botão INICIAR se o usuário selecionou o arquivo, a aba, a coluna e A PASTA DESTINO.
disponivel_para_rodar = (
    arquivo_carregado is not None
    and df_entrada is not None
)

with col1:
    btn_disparar = st.button("▶️ Iniciar Processamento", disabled=not disponivel_para_rodar or st.session_state.rodando, use_container_width=True)
with col2:
    if st.session_state.rodando:
        if st.button("⏸ Pausar / Parar", type="secondary"):
            st.session_state.rodando = False
            st.rerun()

if btn_disparar:
    st.session_state.rodando = True
    st.session_state.logs = ["🚀 Processamento iniciado..."]

# --- LÓGICA DE PROCESSAMENTO EM BACKGROUND ---
if st.session_state.rodando:
    lista_bruta = df_entrada[coluna_selecionada].dropna().tolist()
    cnpjs_processados = set()
    resultados_atuais = []

    # Recupera o progresso do arquivo na pasta de destino (Anti-queda)
    if os.path.exists(caminho_final_salvamento):
        try:
            df_existente = pd.read_excel(caminho_final_salvamento)
            resultados_atuais = df_existente.to_dict('records')
            if 'CNPJ Completo' in df_existente.columns:
                cnpjs_processados.update([limpar_cnpj(c) for c in df_existente['CNPJ Completo'].dropna()])
            st.session_state.logs.append(f"🔄 Arquivo existente detectado na pasta. {len(cnpjs_processados)} CNPJs já estão prontos.")
        except:
            pass

    cnpjs_pendentes = list(dict.fromkeys([limpar_cnpj(x) for x in lista_bruta if len(limpar_cnpj(x)) == 14 and limpar_cnpj(x) not in cnpjs_processados]))
    total = len(cnpjs_pendentes)

    if total == 0:
        st.success("✅ Todos os CNPJs já foram processados!")
        st.session_state.rodando = False
        st.rerun()
    else:
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
        
        sucessos = 0
        falhas = 0
        t_inicio_global = time.time()
        tempo_ciclo_fixo = 20.5

        for index, cnpj in enumerate(cnpjs_pendentes):
            t_inicio_req = time.time()
            url = f"https://publica.cnpj.ws/cnpj/{cnpj}"
            sucesso_req = False
            status_erro = "Desconhecido"

            for tentativa in range(3):
                try:
                    res = session.get(url, timeout=15)
                    if res.status_code == 200:
                        linha = extrair_dados_json(res.json())
                        resultados_atuais.append(linha)
                        sucesso_req = True
                        sucessos += 1
                        st.session_state.logs.append(f"[{index+1}/{total}] ✅ Sucesso: {linha['Razão Social']}")
                        break
                    elif res.status_code == 400:
                        status_erro = "Erro 400 (Inexistente)"
                        break
                    elif res.status_code == 429:
                        r_after = int(res.headers.get("Retry-After", 25))
                        st.session_state.logs.append(f"⚠️ Limite (429). Aguardando {r_after}s...")
                        time.sleep(r_after)
                    else:
                        status_erro = f"HTTP {res.status_code}"
                        time.sleep(2)
                except Exception:
                    status_erro = "Falha de Rede"
                    time.sleep(2)

            if not sucesso_req:
                falhas += 1
                st.session_state.logs.append(f"[{index+1}/{total}] ❌ Falha: {status_erro}")
                if "Erro 400" not in status_erro:
                    linha_erro = {"CNPJ Completo": cnpj, "Razão Social": f"FALHA ({status_erro})", "Situação Cadastral": "ERRO"}
                    resultados_atuais.append(linha_erro)

            # Atualiza Tela
            processados_agora = index + 1
            t_decorrido = max(time.time() - t_inicio_global, 1)
            req_min = processados_agora / (t_decorrido / 60)
            t_medio = t_decorrido / processados_agora
            s_restantes = (total - processados_agora) * max(t_medio, tempo_ciclo_fixo)

            metric_sucesso.metric("Sucessos ✅", str(sucessos))
            metric_falha.metric("Falhas ❌", str(falhas))
            metric_req_min.metric("Velocidade ⚡", f"{req_min:.1f} req/min")
            metric_t_medio.metric("Tempo Médio ⏱️", f"{t_medio:.1f}s")

            progress_placeholder.progress(processados_agora / total)
            eta_placeholder.write(f"**Progresso:** {processados_agora}/{total} | **Tempo Restante:** {formatar_eta(s_restantes)}")

            log_placeholder.text_area("Logs em Tempo Real", value="\n".join(st.session_state.logs[-12:]), height=250)
            
            df_preview = pd.DataFrame(resultados_atuais[-15:])
            if not df_preview.empty and "Razão Social" in df_preview.columns:
                preview_placeholder.dataframe(df_preview, use_container_width=True)

            # SALVA DIRETAMENTE NA PASTA ESCOLHIDA PELO EXPLORADOR DO WINDOWS
            if sucesso_req and (index % 5 == 0 or index == total - 1):
                try:
                    with pd.ExcelWriter(
                    caminho_final_salvamento,
                    engine="openpyxl"
                ) as writer:
                    pd.DataFrame(resultados_atuais).to_excel(
                        writer,
                        index=False
                    )
                except:
                    pass

            # Trava para não tomar erro 429
            if index < total - 1:
                time.sleep(tempo_ciclo_fixo)

        st.success(f"🎉 Processamento concluído! O arquivo foi salvo em: {caminho_final_salvamento}")
        st.session_state.rodando = False
        st.rerun()
