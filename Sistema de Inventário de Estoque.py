import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

# Variáveis globais
dados_estoque = None
usuario_logado = False
dados_atualizados = None  # Para armazenar dados com contagens atualizadas

CAMINHO_PLANILHA = r'G:\Drives compartilhados\AUDITORIA\18- Auditoria 2024\Inventarios 2024\Teste_Bigquery.xlsx'

def realizar_login(usuario, senha):
    global usuario_logado
    if usuario == "mhmelo" and senha == "1234":
        usuario_logado = True
        messagebox.showinfo("Login", "Login realizado com sucesso!")
        root.deiconify()
        login_window.destroy()
    else:
        messagebox.showerror("Erro", "Usuário ou senha inválidos!")

def carregar_planilha():
    global dados_estoque
    try:
        dados_estoque = pd.read_excel(CAMINHO_PLANILHA)
        dados_estoque.columns = dados_estoque.columns.str.strip()
        for col in ["Centro", "Material", "Texto_Breve_Material", "Deposito", "Quantidade"]:
            if col in dados_estoque.columns:
                dados_estoque[col] = dados_estoque[col].astype(str).str.strip()
        
        if "Quantidade" in dados_estoque.columns:
            dados_estoque["Quantidade"] = pd.to_numeric(dados_estoque["Quantidade"], errors="coerce")

        # Filtrar apenas os itens com quantidade diferente de zero
        dados_estoque = dados_estoque[dados_estoque["Quantidade"] != 0]

        dados_estoque["Contagem"] = None  # Adicionar coluna para contagem
        dados_estoque["Diferença"] = None  # Adicionar coluna para diferença
        messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo não encontrado: {CAMINHO_PLANILHA}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha: {str(e)}")

def filtrar_dados():
    global dados_estoque

    if dados_estoque is None:
        messagebox.showerror("Erro", "Por favor, carregue a planilha antes!")
        return

    centro = entry_centro.get().strip()
    material = entry_material.get().strip()
    texto_breve = entry_texto_breve.get().strip()
    deposito = entry_deposito.get().strip()

    try:
        dados_filtrados = dados_estoque
        if centro:
            dados_filtrados = dados_filtrados[dados_filtrados["Centro"] == centro]
        if material:
            dados_filtrados = dados_filtrados[dados_filtrados["Material"] == material]
        if texto_breve:
            dados_filtrados = dados_filtrados[dados_filtrados["Texto_Breve_Material"].str.contains(texto_breve, case=False, na=False)]
        if deposito:
            dados_filtrados = dados_filtrados[dados_filtrados["Deposito"] == deposito]

        # Filtrar apenas as linhas com diferença diferente de 0
        if "Quantidade" in dados_filtrados.columns and "Contagem" in dados_filtrados.columns:
            dados_filtrados["Diferença"] = dados_filtrados["Contagem"] - dados_filtrados["Quantidade"]
            dados_filtrados = dados_filtrados[dados_filtrados["Diferença"] != 0]

        if dados_filtrados.empty:
            messagebox.showinfo("Resultado", "Nenhum dado encontrado com os filtros aplicados!")
        else:
            exibir_dados(dados_filtrados)
    except KeyError as e:
        messagebox.showerror("Erro", f"A coluna {str(e)} não foi encontrada na planilha!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao filtrar os dados: {str(e)}")

def exibir_dados(dados):
    global dados_atualizados
    dados_atualizados = dados.copy()

    janela_dados = tk.Toplevel(root)
    janela_dados.title("Dados Filtrados")

    frame = ttk.Frame(janela_dados)
    frame.pack(fill="both", expand=True)

    tree = ttk.Treeview(frame, columns=list(dados.columns), show="headings", height=15)
    for col in dados.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for index, row in dados.iterrows():
        tree.insert("", "end", values=list(row))

    tree.pack(fill="both", expand=True)

    tk.Label(janela_dados, text="Insira o valor da contagem:").pack(pady=5)
    tk.Label(janela_dados, text="Material:").pack(pady=5)
    entry_material_contagem = tk.Entry(janela_dados)
    entry_material_contagem.pack(pady=5)
    
    tk.Label(janela_dados, text="Texto Breve Material:").pack(pady=5)
    entry_material_contagem = tk.Entry(janela_dados)
    entry_material_contagem.pack(pady=5)

    tk.Label(janela_dados, text="Valor da Contagem:").pack(pady=5)
    entry_contagem = tk.Entry(janela_dados)
    entry_contagem.pack(pady=5)

    def salvar_contagem():
        global dados_atualizados
        material = entry_material_contagem.get().strip()
        contagem = entry_contagem.get().strip()

        if not material or not contagem:
            messagebox.showerror("Erro", "Preencha todos os campos para salvar a contagem!")
            return

        try:
            contagem = float(contagem)
            # Atualizar contagem no DataFrame
            dados_atualizados.loc[dados_atualizados["Material"] == material, "Contagem"] = contagem

            # Atualizar diferença no DataFrame
            if "Quantidade" in dados_atualizados.columns:
                dados_atualizados["Diferença"] = (
                    dados_atualizados["Contagem"] - dados_atualizados["Quantidade"]
                )

            messagebox.showinfo("Sucesso", "Contagem e diferença atualizadas com sucesso!")
        except ValueError:
            messagebox.showerror("Erro", "Insira um valor numérico válido para a contagem!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar a contagem: {str(e)}")

    tk.Button(janela_dados, text="Salvar Contagem", command=salvar_contagem).pack(pady=10)

    def exportar_dados():
        try:
            caminho_exportacao = r"G:\Drives compartilhados\AUDITORIA\Contagem_Atualizada.xlsx"
            dados_atualizados.to_excel(caminho_exportacao, index=False)
            messagebox.showinfo("Sucesso", f"Dados exportados para: {caminho_exportacao}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar os dados: {str(e)}")

    tk.Button(janela_dados, text="Exportar Dados Atualizados", command=exportar_dados).pack(pady=10)

# Janela de login
login_window = tk.Tk()
login_window.title("Login")
login_window.geometry("300x200")

tk.Label(login_window, text="Usuário:").pack(pady=5)
entry_usuario = tk.Entry(login_window)
entry_usuario.pack(pady=5)

tk.Label(login_window, text="Senha:").pack(pady=5)
entry_senha = tk.Entry(login_window, show="*")
entry_senha.pack(pady=5)

tk.Button(login_window, text="Login", command=lambda: realizar_login(entry_usuario.get(), entry_senha.get())).pack(pady=10)

# Janela principal
root = tk.Tk()
root.title("Sistema de Contagem de Estoque")
root.geometry("600x500")
root.withdraw()

tk.Button(root, text="Carregar Planilha", command=carregar_planilha).pack(pady=10)

tk.Label(root, text="Digite o Centro:").pack(pady=5)
entry_centro = tk.Entry(root)
entry_centro.pack(pady=5)

tk.Label(root, text="Digite o Material:").pack(pady=5)
entry_material = tk.Entry(root)
entry_material.pack(pady=5)

tk.Label(root, text="Digite o Texto Breve do Material:").pack(pady=5)
entry_texto_breve = tk.Entry(root)
entry_texto_breve.pack(pady=5)

tk.Label(root, text="Digite o Depósito:").pack(pady=5)
entry_deposito = tk.Entry(root)
entry_deposito.pack(pady=5)

tk.Button(root, text="Filtrar Dados", command=filtrar_dados).pack(pady=10)

root.mainloop()