import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import requests
import time
import threading
import re
import os
import datetime

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class AppCNPJ(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Consulta de CNPJ - Enterprise Edition v3")
        self.geometry("950x750")
        self.resizable(False, False)

        # Estados do Sistema
        self.arquivo_entrada = None
        self.arquivo_saida = None
        self.df_entrada = None
        self.cancelar_processo = False
        
        # Controle de Pausa Seguro (Thread Event)
        self.is_paused = False
        self.pause_event = threading.Event()
        self.pause_event.set() # Inicialmente liberado (não pausado)

        # Variáveis de Telemetria
        self.stats = {"sucesso": 0, "falha": 0, "retry": 0, "inicio": 0, "processados": 0}

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "application/json"
        })

        self.setup_ui()

    # --- UI SETUP AVANÇADO ---
    def setup_ui(self):
        # Topbar com Tema e Exportação
        self.frame_top = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_top.pack(fill="x", padx=20, pady=(10, 0))
        
        self.lbl_titulo = ctk.CTkLabel(self.frame_top, text="Consulta Automática de CNPJ", font=("Roboto", 24, "bold"))
        self.lbl_titulo.pack(side="left")

        self.switch_tema = ctk.CTkSwitch(self.frame_top, text="Modo Claro", command=self.mudar_tema)
        self.switch_tema.pack(side="right")

        self.btn_exportar_csv = ctk.CTkButton(self.frame_top, text="📊 Exportar CSV", fg_color="#17a2b8", hover_color="#138496", command=self.exportar_csv, width=120)
        self.btn_exportar_csv.pack(side="right", padx=15)

        # Painel Duplo (Esquerda: Arquivos | Direita: Métricas)
        self.frame_mid = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_mid.pack(fill="x", padx=20, pady=10)

        # ESQUERDA: Arquivos
        self.frame_arquivos = ctk.CTkFrame(self.frame_mid)
        self.frame_arquivos.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.btn_selecionar_entrada = ctk.CTkButton(self.frame_arquivos, text="📁 Arquivo de Entrada", command=self.selecionar_entrada)
        self.btn_selecionar_entrada.pack(padx=10, pady=(15, 5), fill="x")
        self.lbl_caminho_entrada = ctk.CTkLabel(self.frame_arquivos, text="Nenhum arquivo...", text_color="gray")
        self.lbl_caminho_entrada.pack(padx=10, pady=(0, 10))

        self.btn_selecionar_saida = ctk.CTkButton(self.frame_arquivos, text="💾 Arquivo de Saída", command=self.selecionar_saida)
        self.btn_selecionar_saida.pack(padx=10, pady=(5, 5), fill="x")
        self.lbl_caminho_saida = ctk.CTkLabel(self.frame_arquivos, text="Nenhum arquivo...", text_color="gray")
        self.lbl_caminho_saida.pack(padx=10, pady=(0, 15))

        # DIREITA: Dashboard de Status Real
        self.frame_stats = ctk.CTkFrame(self.frame_mid)
        self.frame_stats.pack(side="right", fill="both", expand=True)

        self.lbl_sucessos = ctk.CTkLabel(self.frame_stats, text="✅ Sucessos: 0", font=("Roboto", 12, "bold"), text_color="#28a745")
        self.lbl_sucessos.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="w")
        
        self.lbl_falhas = ctk.CTkLabel(self.frame_stats, text="❌ Falhas: 0", font=("Roboto", 12, "bold"), text_color="#dc3545")
        self.lbl_falhas.grid(row=0, column=1, padx=15, pady=(15, 5), sticky="w")

        self.lbl_retries = ctk.CTkLabel(self.frame_stats, text="🔄 Retries: 0", font=("Roboto", 12))
        self.lbl_retries.grid(row=1, column=0, padx=15, pady=5, sticky="w")

        self.lbl_req_min = ctk.CTkLabel(self.frame_stats, text="⚡ Req/Min: 0.0", font=("Roboto", 12))
        self.lbl_req_min.grid(row=1, column=1, padx=15, pady=5, sticky="w")

        self.lbl_tempo_medio = ctk.CTkLabel(self.frame_stats, text="⏱️ Tempo Médio: 0.0s", font=("Roboto", 12))
        self.lbl_tempo_medio.grid(row=2, column=0, columnspan=2, padx=15, pady=(5, 15), sticky="w")

        # Botões de Ação
        self.frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_botoes.pack(pady=5)

        self.btn_iniciar = ctk.CTkButton(self.frame_botoes, text="▶ INICIAR", fg_color="#28a745", hover_color="#218838", command=self.iniciar_thread, state="disabled")
        self.btn_iniciar.pack(side="left", padx=5)

        self.btn_pausar = ctk.CTkButton(self.frame_botoes, text="⏸ PAUSAR", fg_color="#ffc107", text_color="black", hover_color="#e0a800", command=self.alternar_pausa, state="disabled")
        self.btn_pausar.pack(side="left", padx=5)

        self.btn_cancelar = ctk.CTkButton(self.frame_botoes, text="⏹ CANCELAR", fg_color="#dc3545", hover_color="#c82333", command=self.acionar_cancelamento, state="disabled")
        self.btn_cancelar.pack(side="left", padx=5)

        # Progresso Dinâmico
        self.frame_progresso = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_progresso.pack(fill="x", padx=20, pady=5)
        
        self.lbl_contagem = ctk.CTkLabel(self.frame_progresso, text="Progresso: 0 / 0 (0%)", font=("Roboto", 12, "bold"))
        self.lbl_contagem.pack(side="left")
        self.lbl_eta = ctk.CTkLabel(self.frame_progresso, text="ETA: --:--:--", text_color="#17a2b8")
        self.lbl_eta.pack(side="right")
        
        self.progress = ctk.CTkProgressBar(self, width=900)
        self.progress.pack(pady=(25, 0))
        self.progress.set(0)

        # Abas (Tabview) para Logs e Tabela Preview
        self.tabview = ctk.CTkTabview(self, height=220)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(5, 15))
        self.tabview.add("Visualização de Dados")
        self.tabview.add("Logs do Sistema")

        # Configuração da Caixa de Log
        self.caixa_log = ctk.CTkTextbox(self.tabview.tab("Logs do Sistema"), state="disabled")
        self.caixa_log.pack(fill="both", expand=True)

        # Configuração do Preview (Treeview native)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", borderwidth=0)
        style.map("Treeview", background=[("selected", "#1f538d")])
        style.configure("Treeview.Heading", background="#1f538d", foreground="white", font=("Roboto", 10, "bold"))

        colunas = ("CNPJ", "Razão Social", "Situação", "Status Req.")
        self.tree = ttk.Treeview(self.tabview.tab("Visualização de Dados"), columns=colunas, show="headings", height=8)
        for col in colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="w", width=200 if col == "Razão Social" else 100)
        self.tree.pack(fill="both", expand=True)

    # --- FUNÇÕES DE UI E TEMA ---
    def mudar_tema(self):
        novo_tema = "Light" if self.switch_tema.get() else "Dark"
        ctk.set_appearance_mode(novo_tema)
        bg_tree = "white" if novo_tema == "Light" else "#2b2b2b"
        fg_tree = "black" if novo_tema == "Light" else "white"
        style = ttk.Style()
        style.configure("Treeview", background=bg_tree, foreground=fg_tree, fieldbackground=bg_tree)

    def exportar_csv(self):
        if not self.arquivo_saida or not os.path.exists(self.arquivo_saida):
            messagebox.showwarning("Aviso", "Ainda não há dados salvos para exportar.")
            return
        try:
            df = pd.read_excel(self.arquivo_saida)
            csv_path = self.arquivo_saida.replace(".xlsx", ".csv")
            df.to_csv(csv_path, index=False, sep=";", encoding="utf-8-sig")
            messagebox.showinfo("Sucesso", f"CSV exportado com sucesso para:\n{csv_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {e}")

    # --- THREAD-SAFE UPDATES ---
    def safe_log(self, mensagem):
        self.after(0, lambda: self._update_log(mensagem))

    def _update_log(self, mensagem):
        self.caixa_log.configure(state="normal")
        self.caixa_log.insert("end", mensagem + "\n")
        self.caixa_log.see("end")
        self.caixa_log.configure(state="disabled")

    def safe_add_treeview(self, cnpj, razao, situacao, status_req):
        self.after(0, lambda: self.tree.insert("", "end", values=(cnpj, razao, situacao, status_req)))
        # Mantém apenas as últimas 50 linhas para não explodir memória
        self.after(0, lambda: self.limpar_treeview_antigo())

    def limpar_treeview_antigo(self):
        itens = self.tree.get_children()
        if len(itens) > 50:
            self.tree.delete(itens[0])
            self.tree.yview_moveto(1)

    def safe_update_telemetry(self):
        def _update():
            self.lbl_sucessos.configure(text=f"✅ Sucessos: {self.stats['sucesso']}")
            self.lbl_falhas.configure(text=f"❌ Falhas: {self.stats['falha']}")
            self.lbl_retries.configure(text=f"🔄 Retries: {self.stats['retry']}")
            
            tempo_decorrido = max(time.time() - self.stats['inicio'], 1)
            minutos = tempo_decorrido / 60
            req_min = self.stats['processados'] / minutos if minutos > 0 else 0
            t_medio = tempo_decorrido / self.stats['processados'] if self.stats['processados'] > 0 else 0
            
            self.lbl_req_min.configure(text=f"⚡ Req/Min: {req_min:.1f}")
            self.lbl_tempo_medio.configure(text=f"⏱️ Tempo Médio: {t_medio:.1f}s")
        self.after(0, _update)

    def safe_update_progress(self, index, total, eta_str, prog_val, cor="#1f538d"):
        def _update():
            pct = int((index / total) * 100) if total > 0 else 100
            self.lbl_contagem.configure(text=f"Progresso: {index} / {total} ({pct}%)")
            self.lbl_eta.configure(text=f"ETA: {eta_str}")
            self.progress.set(prog_val)
            self.progress.configure(progress_color=cor)
        self.after(0, _update)

    # --- CONTROLES GERAIS ---
    def verificar_prontidao(self):
        if self.arquivo_entrada and self.arquivo_saida:
            self.btn_iniciar.configure(state="normal")

    def selecionar_entrada(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if arquivo:
            try:
                df = pd.read_excel(arquivo)
                colunas = [str(c).upper().strip() for c in df.columns]
                if 'CNPJ' not in colunas: return messagebox.showerror("Erro", "Coluna 'CNPJ' não encontrada.")
                df.columns = colunas
                self.arquivo_entrada = arquivo
                self.df_entrada = df
                self.lbl_caminho_entrada.configure(text=os.path.basename(arquivo), text_color="white")
                self.verificar_prontidao()
            except Exception as e: messagebox.showerror("Erro", str(e))

    def selecionar_saida(self):
        arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="RESULTADOS_CNPJ.xlsx", filetypes=[("Excel", "*.xlsx")])
        if arquivo:
            self.arquivo_saida = arquivo
            self.lbl_caminho_saida.configure(text=os.path.basename(arquivo), text_color="white")
            self.verificar_prontidao()

    def alternar_pausa(self):
        if not self.is_paused:
            self.is_paused = True
            self.pause_event.clear() # Bloqueia a thread
            self.btn_pausar.configure(text="▶ RETOMAR", fg_color="#28a745")
            self.progress.configure(progress_color="#ffc107")
            self.safe_log("⏸ Sistema em Pausa...")
        else:
            self.is_paused = False
            self.pause_event.set() # Libera a thread
            self.btn_pausar.configure(text="⏸ PAUSAR", fg_color="#ffc107")
            self.progress.configure(progress_color="#1f538d")
            self.safe_log("▶ Sistema Retomado!")

    def acionar_cancelamento(self):
        self.cancelar_processo = True
        self.pause_event.set() # Garante que não cancele preso na pausa
        self.btn_cancelar.configure(state="disabled", text="CANCELANDO...")
        self.progress.configure(progress_color="#dc3545")

    # --- LÓGICA DE PROCESSAMENTO E EXTRAÇÃO ---
    def limpar_cnpj(self, cnpj):
        # Converte para string e quebra no ponto para remover decimais (.0) que o Pandas adiciona
        cnpj_str = str(cnpj).split('.')[0]
        # Limpa caracteres especiais e garante os 14 dígitos
        return re.sub(r'\D', '', cnpj_str).zfill(14)

    def _get(self, dic, *chaves):
        for chave in chaves:
            if isinstance(dic, dict): dic = dic.get(chave, "")
            else: return ""
        return dic if dic != {} else ""

    def extrair_dados_json(self, dados):
        estab = dados.get('estabelecimento', {})
        
        # Tratamento de listas estruturadas (Sócios e CNAEs)
        socios = " \n; ".join([f"Nome: {s.get('nome')} | Doc: {s.get('cpf_cnpj_socio')} | Qualificação: {self._get(s, 'qualificacao_socio', 'descricao')}" for s in dados.get('socios', [])])
        secundarias = " \n; ".join([f"{act.get('subclasse')} - {act.get('descricao')}" for act in estab.get('atividades_secundarias', [])])
        
        # Dicionário principal com todos os dados básicos mapeados
        linha_excel = {
            "CNPJ Raiz": self._get(dados, 'cnpj_raiz'), 
            "Razão Social": self._get(dados, 'razao_social'),
            "Capital Social": self._get(dados, 'capital_social'), 
            "Responsável Federativo": self._get(dados, 'responsavel_federativo'),
            "Atualizado Em (Geral)": self._get(dados, 'atualizado_em'), 
            "Porte ID": self._get(dados, 'porte', 'id'),
            "Porte Descrição": self._get(dados, 'porte', 'descricao'), 
            "Natureza Jurídica ID": self._get(dados, 'natureza_juridica', 'id'),
            "Natureza Jurídica Descrição": self._get(dados, 'natureza_juridica', 'descricao'), 
            "Simples Nacional / MEI": self._get(dados, 'simples'),
            "CNPJ Completo": self._get(estab, 'cnpj'), 
            "Situação Cadastral": self._get(estab, 'situacao_cadastral'),
            "Data Situação Cadastral": self._get(estab, 'data_situacao_cadastral'), 
            "Data Início Atividade": self._get(estab, 'data_inicio_atividade'),
            "Logradouro": self._get(estab, 'logradouro'), 
            "Número": self._get(estab, 'numero'),
            "Complemento": self._get(estab, 'complemento'), 
            "Bairro": self._get(estab, 'bairro'),
            "CEP": self._get(estab, 'cep'), 
            "Telefone": self._get(estab, 'telefone1'),
            "E-mail": self._get(estab, 'email'), 
            "Atividade Principal": self._get(estab, 'atividade_principal', 'descricao'),
            "Estado (UF)": self._get(estab, 'estado', 'sigla'), 
            "Cidade": self._get(estab, 'cidade', 'nome'),
            "Atividades Secundárias": secundarias, 
            "Quadro de Sócios": socios
        }

        # --- NOVO: SEPARAÇÃO DINÂMICA DAS IEs POR ESTADO ---
        uf_sede = self._get(estab, 'estado', 'sigla')
        lista_ies = estab.get('inscricoes_estaduais', [])

        if not lista_ies:
            # Se não possui nenhuma IE registrada, cria a coluna "IE [Estado Sede]" com valor ISENTO
            nome_coluna_isento = f"IE {uf_sede}" if uf_sede else "IE PRINCIPAL"
            linha_excel[nome_coluna_isento] = "ISENTO"
        else:
            # Se possui IEs, varre todas e joga nas colunas correspondentes de cada Estado
            for ie in lista_ies:
                sigla_estado = self._get(ie, 'estado', 'sigla')
                
                # Ignora se por erro da API a IE vier sem UF atrelada
                if not sigla_estado:
                    continue
                
                # Limpa a Inscrição Estadual deixando apenas números
                num_ie = re.sub(r'\D', '', str(ie.get('inscricao_estadual', '')))
                nome_coluna = f"IE {sigla_estado}"

                if num_ie:
                    # Verifica se já existe uma IE para esse estado preenchida (ex: uma ativa e uma baixada)
                    # Se sim, junta na mesma célula. Se não, preenche normalmente.
                    if nome_coluna in linha_excel:
                        linha_excel[nome_coluna] += f" ; {num_ie}"
                    else:
                        linha_excel[nome_coluna] = num_ie

        return linha_excel

    # --- CORE ---
    def iniciar_thread(self):
        self.cancelar_processo = False
        self.is_paused = False
        self.pause_event.set()
        
        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_cancelar.configure(state="normal", text="⏹ CANCELAR")
        self.stats = {"sucesso": 0, "falha": 0, "retry": 0, "inicio": time.time(), "processados": 0}
        
        for item in self.tree.get_children(): self.tree.delete(item)
        threading.Thread(target=self.processar_cnpjs, daemon=True).start()

    def processar_cnpjs(self):
        resultados_atuais = []
        cnpjs_processados = set()

        if os.path.exists(self.arquivo_saida):
            try:
                df_existente = pd.read_excel(self.arquivo_saida)
                resultados_atuais = df_existente.to_dict('records')
                if 'CNPJ Completo' in df_existente.columns:
                    cnpjs_processados.update([self.limpar_cnpj(c) for c in df_existente['CNPJ Completo'].dropna()])
            except: pass

        lista_bruta = self.df_entrada['CNPJ'].dropna().tolist()
        cnpjs_pendentes = [c for c in [self.limpar_cnpj(x) for x in lista_bruta] if len(c) == 14 and c not in cnpjs_processados]
        cnpjs_pendentes = list(dict.fromkeys(cnpjs_pendentes))
        total = len(cnpjs_pendentes)

        if total == 0:
            self.safe_log("✅ Tudo finalizado ou lista vazia!")
            self.after(0, lambda: self.btn_iniciar.configure(state="normal"))
            return

        for index, cnpj in enumerate(cnpjs_pendentes):
            self.pause_event.wait() # VERIFICA PAUSA AQUI
            if self.cancelar_processo: break

            self.safe_log(f"[{index+1}/{total}] Consultando: {cnpj}")
            url = f"https://publica.cnpj.ws/cnpj/{cnpj}"
            sucesso = False
            t_inicio_req = time.time()

            for tentativa in range(3):
                self.pause_event.wait()
                if self.cancelar_processo: break
                try:
                    res = self.session.get(url, timeout=15)
                    if res.status_code == 200:
                        linha = self.extrair_dados_json(res.json())
                        resultados_atuais.append(linha)
                        sucesso = True
                        self.stats["sucesso"] += 1
                        self.safe_add_treeview(cnpj, linha['Razão Social'], linha['Situação Cadastral'], "✅ OK")
                        break
                    elif res.status_code == 400:
                        self.stats["falha"] += 1
                        self.safe_add_treeview(cnpj, "N/A", "Inválido Receita", "❌ Erro 400")
                        break
                    elif res.status_code == 429:
                        r_after = int(res.headers.get("Retry-After", 25))
                        self.safe_log(f" ❌ Limite! Pausa de {r_after}s...")
                        time.sleep(r_after)
                    else:
                        self.stats["retry"] += 1
                        self.safe_update_telemetry()
                        time.sleep(2)
                except requests.exceptions.RequestException:
                    self.stats["retry"] += 1
                    time.sleep(2)

            if not sucesso and not self.cancelar_processo and res.status_code != 400:
                self.stats["falha"] += 1
                self.safe_add_treeview(cnpj, "Falha de Conexão", "-", "❌ Timeout")

            self.stats["processados"] += 1
            self.safe_update_telemetry()

            # Salva Excel periodicamente
            if sucesso and (index % 5 == 0 or index == total - 1):
                try: pd.DataFrame(resultados_atuais).to_excel(self.arquivo_saida, index=False)
                except: pass

            # Cálculos Dinâmicos de ETA e Delay (Limitador)
            t_gasto = time.time() - t_inicio_req
            t_ciclo = max(20.5, t_gasto) # Força um minimo de 20.5s para API
            s_restantes = (total - (index+1)) * (self.stats['processados'] / max(time.time() - self.stats['inicio'], 1) ** -1 if index > 5 else t_ciclo)
            
            cor_prog = "#1f538d" if not self.is_paused else "#ffc107"
            self.safe_update_progress(index + 1, total, str(datetime.timedelta(seconds=int(s_restantes))), (index+1)/total, cor_prog)

            if index < total - 1 and not self.cancelar_processo:
                t_espera = max(0, 20.5 - t_gasto)
                for _ in range(int(t_espera * 10)):
                    if self.cancelar_processo: break
                    time.sleep(0.1)

        self.btn_iniciar.configure(state="normal")
        self.btn_pausar.configure(state="disabled")
        self.btn_cancelar.configure(state="disabled", text="⏹ CANCELAR")
        self.progress.configure(progress_color="#28a745" if not self.cancelar_processo else "#dc3545")

if __name__ == "__main__":
    app = AppCNPJ()
    app.mainloop()