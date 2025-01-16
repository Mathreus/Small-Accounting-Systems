import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

# Variáveis globais
dados_estoque = None
usuario_logado = False
dados_atualizados = None  # Para armazenar dados com contagens atualizadas

CAMINHO_PLANILHA = r'G:\Drives compartilhados\AUDITORIA\18- Auditoria 2024\Inventarios 2024\Teste_Bigquery.xlsx'

# Funções
def realizar_login(usuario, senha):
    global usuario_logado
    if usuario == "mhmelo" and senha == "123":
        usuario_logado = True
        messagebox.showinfo("Login", "Login realizado com sucesso!")
        login_window.destroy()  # Fechar janela de login
        menu_principal()  # Abrir menu principal
    else:
        messagebox.showerror("Erro", "Usuário ou senha inválidos!")

# Função para tentar realizar o salvamento automático
def salvar_em_tempo_real():
    global dados_estoque
    try:
        # Salvar o DataFrame atualizado no arquivo Excel
        with pd.ExcelWriter(CAMINHO_PLANILHA, engine='openpyxl', mode='w') as writer:
            dados_estoque.to_excel(writer, index=False, sheet_name='Dados Atualizados')
        messagebox.showinfo("Sucesso", "Alterações salvas automaticamente!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar automaticamente: {str(e)}")

def carregar_planilha():
    global dados_estoque
    try:
        dados_estoque = pd.read_excel(CAMINHO_PLANILHA)
        dados_estoque.columns = dados_estoque.columns.str.strip()
        for col in ["Centro", "Material", "Texto_Breve_Material", "Deposito", "Grupo_Mercadorias", "Quantidade"]:
            if col in dados_estoque.columns:
                dados_estoque[col] = dados_estoque[col].astype(str).str.strip()

        if "Quantidade" in dados_estoque.columns:
            dados_estoque["Quantidade"] = pd.to_numeric(dados_estoque["Quantidade"], errors="coerce")

        # Filtrar apenas os itens com quantidade diferente de zero
        dados_estoque = dados_estoque[dados_estoque["Quantidade"] != 0]

        # Adicionar colunas necessárias
        dados_estoque["Grupo"] = dados_estoque["Grupo_Mercadorias"].apply(lambda x: x[0] if x else "")
        dados_estoque["Contagem"] = 0
        dados_estoque["Diferença"] = 0
        dados_estoque["Classificação"] = None
        dados_estoque["Endereco"] = ""
        dados_estoque["Observações"] = ""
        dados_estoque["Classificação_Justif"] = ""
        dados_estoque["Recontagem"] = ""
        dados_estoque["Diferença_Recontagem"] = ""

        messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo não encontrado: {CAMINHO_PLANILHA}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha: {str(e)}")

def atualizar_e_salvar(material, contagem, recontagem, endereco, observacao):
    global dados_estoque

    try:
        # Atualizar os valores no DataFrame
        dados_estoque.loc[dados_estoque["Material"] == material, "Endereco"] = endereco
        dados_estoque.loc[dados_estoque["Material"] == material, "Contagem"] = float(contagem)
        dados_estoque.loc[dados_estoque["Material"] == material, "Recontagem"] = float(recontagem)
        dados_estoque.loc[dados_estoque["Material"] == material, "Observações"] = observacao

        # Salvar automaticamente
        salvar_em_tempo_real()
    except ValueError:
        messagebox.showerror("Erro", "Erro ao atualizar os valores. Verifique os campos!")

def filtrar_dados():
    global dados_estoque

    if dados_estoque is None:
        messagebox.showerror("Erro", "Por favor, carregue a planilha antes!")
        return

    # Obter valores dos campos de entrada
    centro = entry_centro.get().strip()
    material = entry_material.get().strip()
    texto_breve = entry_texto_breve.get().strip()
    deposito = entry_deposito.get().strip()
    endereco = entry_endereco.get().strip()
    observacao = entry_observacao.get().strip()

    try:
        # Inicializar dados filtrados
        dados_filtrados = dados_estoque

        # Aplicar filtros
        if centro:
            dados_filtrados = dados_filtrados[dados_filtrados["Centro"] == centro]
        if material:
            dados_filtrados = dados_filtrados[dados_filtrados["Material"] == material]
        if texto_breve:
            dados_filtrados = dados_filtrados[dados_filtrados["Texto_Breve_Material"].str.contains(texto_breve, case=False, na=False)]
        if deposito:
            dados_filtrados = dados_filtrados[dados_filtrados["Deposito"] == deposito]
        if endereco:
            dados_filtrados = dados_filtrados[dados_filtrados["Endereco"].str.contains(endereco, case=False, na=False)]
        if observacao:
            dados_filtrados = dados_filtrados[dados_filtrados["Observações"].str.contains(observacao, case=False, na=False)]

        # Converter colunas para valores numéricos, tratando erros
        for col in ["Quantidade", "Contagem", "Recontagem"]:
            if col in dados_filtrados.columns:
                dados_filtrados[col] = pd.to_numeric(dados_filtrados[col], errors="coerce").fillna(0)

        # Atualizar diferença e classificação para Contagem - Quantidade
        if "Quantidade" in dados_filtrados.columns and "Contagem" in dados_filtrados.columns:
            dados_filtrados["Diferença"] = dados_filtrados["Contagem"] - dados_filtrados["Quantidade"]
            dados_filtrados["Classificação"] = dados_filtrados["Diferença"].apply(
                lambda x: "Sobra" if x > 0 else "Falta" if x < 0 else "OK"
            )

        # Atualizar diferença para Recontagem - Quantidade
        if "Quantidade" in dados_filtrados.columns and "Recontagem" in dados_filtrados.columns:
            dados_filtrados["Diferença_Recontagem"] = dados_filtrados["Recontagem"] - dados_filtrados["Quantidade"]
            dados_filtrados["Classificação_Justif"] = dados_filtrados["Diferença_Recontagem"].apply(
                lambda x: "Sobra" if x > 0 else "Falta" if x < 0 else "OK"
            )

        # Exibir resultados
        if dados_filtrados.empty:
            messagebox.showinfo("Resultado", "Nenhum dado encontrado com os filtros aplicados!")
        else:
            exibir_dados(dados_filtrados)

    except KeyError as e:
        messagebox.showerror("Erro", f"A coluna '{str(e)}' não foi encontrada na planilha!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao filtrar os dados: {str(e)}")

# Função para atualizar os valores da combobox dinamicamente
def atualizar_combobox(event):
    texto_digitado = combobox_material.get()
    # Filtra materiais que contêm o texto digitado
    material_filtrados = [material for material in material if texto_digitado.lower() in material.lower()]
    
    # Atualiza os valores da combobox
    combobox_material['values'] = material_filtrados
    
    # Mantém o texto já digitado na combobox
    combobox_material.set(texto_digitado)

# Função para exibir os dados e incluir a funcionalidade dinâmica na combobox
def exibir_dados(dados):
    global dados_atualizados, material, combobox_material

    # Inicialização da variável `dados_atualizados` com os dados originais
    dados_atualizados = dados.copy()

    # Reordenando as colunas conforme especificado
    ordem_colunas = [
        "Centro", "Deposito", "Grupo", "Grupo_Mercadorias", "Material", 
        "Texto_Breve_Material", "Endereco", "Quantidade", "Valor_Estoque_Total", 
        "Contagem", "Diferença", "Classificação", "Recontagem", "Diferença_Recontagem", 
        "Classificação_Justif", "Observações"
    ]
    
    # Validar se todas as colunas existem no DataFrame antes de reordenar
    colunas_existentes = [col for col in ordem_colunas if col in dados.columns]
    dados = dados[colunas_existentes]

    # Criar janela para exibir os dados
    janela_dados = tk.Toplevel(root)
    janela_dados.title("Dados Filtrados")

    frame = ttk.Frame(janela_dados)
    frame.pack(fill="both", expand=True)

    # Configuração do Treeview
    tree = ttk.Treeview(frame, columns=colunas_existentes, show="headings", height=15)
    for col in colunas_existentes:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for _, row in dados.iterrows():
        tree.insert("", "end", values=list(row))

    tree.pack(fill="both", expand=True)

    # Criar Frame para os campos de entrada de dados
    frame_inputs = ttk.Frame(janela_dados)
    frame_inputs.pack(fill="x", padx=10, pady=10)

    # Lista de materiais e combobox dinâmica
    material = dados["Material"].unique().tolist()
    descricao_material = dados["Texto_Breve_Material"].unique().tolist()

    # Campos organizados lado a lado usando grid
    tk.Label(frame_inputs, text="Material:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    combobox_material = ttk.Combobox(frame_inputs, values=material)
    combobox_material.grid(row=0, column=1, sticky="w", padx=5, pady=5)
    combobox_material.bind("<KeyRelease>", atualizar_combobox)  # Evento de digitação

    tk.Label(frame_inputs, text="Texto_Breve_Material:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    combobox_texto_breve_material = ttk.Combobox(frame_inputs, values=descricao_material)
    combobox_texto_breve_material.grid(row=1, column=1, sticky="w", padx=5, pady=5)
    combobox_texto_breve_material.bind("<KeyRelease>", atualizar_combobox)  # Evento de digitação

    tk.Label(frame_inputs, text="Contagem:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    entry_contagem_contagem = tk.Entry(frame_inputs)
    entry_contagem_contagem.grid(row=2, column=1, sticky="w", padx=5, pady=5)

    tk.Label(frame_inputs, text="Endereço:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    entry_endereco_contagem = tk.Entry(frame_inputs)
    entry_endereco_contagem.grid(row=3, column=1, sticky="w", padx=5, pady=5)

    tk.Label(frame_inputs, text="Observações:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
    entry_observacao_contagem = tk.Entry(frame_inputs)
    entry_observacao_contagem.grid(row=4, column=1, sticky="w", padx=5, pady=5)

    tk.Label(frame_inputs, text="Recontagem:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
    entry_recontagem_contagem = tk.Entry(frame_inputs)
    entry_recontagem_contagem.grid(row=5, column=1, sticky="w", padx=5, pady=5)

    # Adicionar campo para a descrição do produto "Sobra"
    tk.Label(frame_inputs, text="Descrição do Produto (Sobra):").grid(row=6, column=0, sticky="e", padx=5, pady=5)
    entry_descricao_sobra = tk.Entry(frame_inputs)
    entry_descricao_sobra.grid(row=6, column=1, sticky="w", padx=5, pady=5)

    # Função para adicionar a "Sobra"
    def adicionar_sobra():
        global dados_atualizados  # Garantir que usamos a variável global

        # Obter os valores dos campos
        item_selecionado = combobox_material.get().strip()
        texto_material = combobox_texto_breve_material.get().strip()
        endereco = entry_endereco_contagem.get().strip()
        contagem = entry_contagem_contagem.get().strip()
        recontagem = entry_recontagem_contagem.get().strip()
        observacoes = entry_observacao_contagem.get().strip()

        # Obter a descrição do produto (Sobra) do novo campo
        nova_descricao = entry_descricao_sobra.get().strip()  # Agora a descrição vem do campo "Descrição do Produto (Sobra)"

        # Criar um novo DataFrame com os dados da "Sobra"
        novo_dado = pd.DataFrame([{
            "Centro": None,  # Mantém vazio
            "Deposito": None,  # Mantém vazio
            "Grupo": None,  # Mantém vazio
            "Grupo_Mercadorias": None,  # Mantém vazio
            "Material": item_selecionado if item_selecionado else None,
            "Texto_Breve_Material": nova_descricao if nova_descricao else None,  # Adicionando a descrição inserida
            "Endereco": endereco if endereco else None,
            "Quantidade": None,  # Mantém vazio
            "Valor_Estoque_Total": None,  # Mantém vazio
            "Contagem": float(contagem) if contagem else None,
            "Diferença": None,  # Mantém vazio
            "Classificação": None,  # Mantém vazio
            "Recontagem": float(recontagem) if recontagem else None,
            "Diferença_Recontagem": None,  # Mantém vazio
            "Classificação_Justif": None,  # Mantém vazio
            "Observações": observacoes if observacoes else None
        }])

        # Concatenar os novos dados com os dados existentes
        dados_atualizados = pd.concat([dados_atualizados, novo_dado], ignore_index=True)

        # Atualizar a visualização com o novo dado
        tree.insert("", "end", values=list(novo_dado.iloc[0]))

    # Botão para adicionar a "Sobra"
    tk.Button(frame_inputs, text="Adicionar Sobra", command=adicionar_sobra).grid(row=7, column=0, columnspan=2, pady=10)

    def salvar_contagem():
        global dados_atualizados
        material = combobox_material.get().strip()
        texto_breve_material = combobox_texto_breve_material.get().strip()
        contagem = entry_contagem_contagem.get().strip()
        recontagem = entry_recontagem_contagem.get().strip()
        endereco = entry_endereco_contagem.get().strip()
        observacao = entry_observacao_contagem.get().strip()

        if not material or not contagem:
            messagebox.showerror("Erro", "Preencha os campos obrigatórios para salvar a contagem!")
            return

        try:
            contagem = float(contagem) if contagem else None
            recontagem = float(recontagem) if recontagem else None
        
            # Atualizar contagem no DataFrame
            dados_atualizados.loc[dados_atualizados["Material"] == material, "Endereco"] = endereco
            dados_atualizados.loc[dados_atualizados["Material"] == material, "Contagem"] = contagem
            dados_atualizados.loc[dados_atualizados["Material"] == material, "Recontagem"] = recontagem
            dados_atualizados.loc[dados_atualizados["Material"] == material, "Observações"] = observacao
        
            # Atualizar diferença e classificação no DataFrame
            if "Quantidade" in dados_atualizados.columns:
                dados_atualizados["Diferença"] = (
                    dados_atualizados["Contagem"] - dados_atualizados["Quantidade"]
                ).fillna(0)
                dados_atualizados["Classificação"] = dados_atualizados["Diferença"].apply(
                    lambda x: "Sobra" if x > 0 else "Falta" if x < 0 else "OK"
                )
        
            # Atualizar diferença e classificação para recontagem
            dados_atualizados["Diferença_Recontagem"] = (
                dados_atualizados["Recontagem"] - dados_atualizados["Quantidade"]
            ).fillna(0)
            dados_atualizados["Classificação_Justif"] = dados_atualizados["Diferença_Recontagem"].apply(
                lambda x: "Sobra" if x > 0 else "Falta" if x < 0 else "OK"
            )

            # Atualizar a visualização (treeview) sem alterar os botões
            atualizar_treeview()

            messagebox.showinfo("Sucesso", "Contagem atualizada!")

        except ValueError:
            messagebox.showerror("Erro", "Insira um valor numérico válido para a contagem e recontagem!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar a contagem: {str(e)}")

    def atualizar_treeview():
        # Limpar apenas os itens do treeview
        for item in tree.get_children():
            tree.delete(item)

        # Recarregar os dados do DataFrame no treeview
        for _, row in dados_atualizados.iterrows():
            tree.insert("", "end", values=list(row))

    # Botão para salvar a contagem
    btn_salvar = tk.Button(janela_dados, text="Salvar Contagem", command=salvar_contagem)
    btn_salvar.pack(pady=10)

    def exportar_dados():
        try:
            caminho_exportacao = r"G:\Drives compartilhados\AUDITORIA\Contagem_Atualizada.xlsx"
        
            # Ordem desejada das colunas
            ordem_colunas = [
                "Centro", "Deposito", "Grupo", "Grupo_Mercadorias", "Material", 
                "Texto_Breve_Material", "Endereco", "Quantidade", "Valor_Estoque_Total", 
                "Contagem", "Diferença", "Classificação", "Recontagem", 
                "Diferença_Recontagem", "Classificação_Justif", "Observações"
            ]
        
            # Reorganizar as colunas antes de exportar
            if set(ordem_colunas).issubset(dados_atualizados.columns):
                dados_exportar = dados_atualizados[ordem_colunas]
            else:
                missing_cols = list(set(ordem_colunas) - set(dados_atualizados.columns))
                raise KeyError(f"Faltam as seguintes colunas no DataFrame: {', '.join(missing_cols)}")
        
            # Exportar para Excel
            dados_exportar.to_excel(caminho_exportacao, index=False)
            messagebox.showinfo("Sucesso", f"Dados exportados para: {caminho_exportacao}")
        
        except KeyError as e:
            messagebox.showerror("Erro", f"Erro nas colunas: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar os dados: {str(e)}")
    
    tk.Button(janela_dados, text="Exportar Dados Atualizados", command=exportar_dados).pack(pady=10)

# Função para o menu principal
def menu_principal():
    global root
    # Criação da janela principal
    root = tk.Tk()
    root.title("Sistema de Contagem de Estoque")  # Definir o título da janela
    root.geometry("400x300")  # Definir o tamanho da janela (largura x altura)

    # Botões do menu principal
    # Botão para carregar a planilha
    tk.Button(root, text="Carregar Planilha", command=carregar_planilha, width=25, height=2).pack(pady=10)
    
    # Botão para exibir os filtros
    tk.Button(root, text="Exibir Filtros", command=exibir_filtros, width=25, height=2).pack(pady=10)

    # Iniciar a interface gráfica do Tkinter
    root.mainloop()

# Função para exibir a tela de filtros
def exibir_filtros():
    global entry_centro, entry_material, entry_texto_breve, entry_deposito, entry_endereco, entry_observacao
    janela_filtros = tk.Toplevel(root)
    janela_filtros.title("Aplicar Filtros")
    
    # Entradas de filtros
    tk.Label(janela_filtros, text="Centro:").pack(pady=5)
    entry_centro = tk.Entry(janela_filtros)
    entry_centro.pack(pady=5)

    tk.Label(janela_filtros, text="Material:").pack(pady=5)
    entry_material = tk.Entry(janela_filtros)
    entry_material.pack(pady=5)

    tk.Label(janela_filtros, text="Texto Breve Material:").pack(pady=5)
    entry_texto_breve = tk.Entry(janela_filtros)
    entry_texto_breve.pack(pady=5)

    tk.Label(janela_filtros, text="Depósito:").pack(pady=5)
    entry_deposito = tk.Entry(janela_filtros)
    entry_deposito.pack(pady=5)

    tk.Label(janela_filtros, text="Endereco:").pack(pady=5)
    entry_endereco = tk.Entry(janela_filtros)
    entry_endereco.pack(pady=5)

    tk.Label(janela_filtros, text="Observações:").pack(pady=5)
    entry_observacao = tk.Entry(janela_filtros)
    entry_observacao.pack(pady=5)

    tk.Button(janela_filtros, text="Filtrar", command=filtrar_dados, width=25, height=2).pack(pady=10)

# Tela de login
login_window = tk.Tk()
login_window.title("Tela de Login")
login_window.geometry("400x300")

tk.Label(login_window, text="Usuário:").pack(pady=10)
entry_usuario = tk.Entry(login_window)
entry_usuario.pack(pady=5)

tk.Label(login_window, text="Senha:").pack(pady=10)
entry_senha = tk.Entry(login_window, show="*")
entry_senha.pack(pady=5)

tk.Button(login_window, text="Login", command=lambda: realizar_login(entry_usuario.get(), entry_senha.get())).pack(pady=20)

login_window.mainloop()
