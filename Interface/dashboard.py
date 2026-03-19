"""
Dashboard para gerenciamento de planilhas Excel.

Este módulo contém a classe Dashboard que implementa a interface principal
do aplicativo de gerenciamento de planilhas. Inclui funcionalidades para:
- Criar novas planilhas Excel
- Ler e processar planilhas existentes
- Gerar pastas organizadas por cliente
- Visualizar dados em interface gráfica
- Criar relatórios consolidados
"""

import glob
import openpyxl
import os
import customtkinter
import sys
from Interface.Fonts.fonts import Fonts
from Interface.colors import DashboardColors, ViewColors, ExcelColors

class Dashboard(customtkinter.CTkFrame):
    """
    Classe principal do dashboard de gerenciamento de planilhas.
    
    Organiza a interface em dois painéis:
    - Esquerdo: Botões de ação
    - Direito: Lista de planilhas disponíveis e área de resultados
    """
    fonts = Fonts()
    
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        
        # Configurar o título principal
        self._setup_title()
        
        # Configurar o layout principal com frames esquerdo e direito
        self._setup_main_layout()
        
        # Configurar os botões no frame esquerdo
        self._setup_buttons()
        
        # Configurar a lista de planilhas no frame direito
        self._setup_excel_list()
    
    def _setup_title(self):
        """Configura o título principal com estilo Excel."""
        ctkTitle = customtkinter.CTkTextbox(
            self, 
            height=40, 
            font=("Inter", 25, "bold"), 
            fg_color=DashboardColors.TITLE_BACKGROUND,  # Azul Excel header
            text_color=DashboardColors.TITLE_TEXT  # Texto branco
        ) 
        ctkTitle.tag_config("center", justify='center')
        ctkTitle.insert("0.0", "Dashboard - Gerenciar Planilhas")
        ctkTitle.tag_add("center", "1.0", "end")
        ctkTitle.configure(state="disabled")
        ctkTitle.pack(fill="x", padx=20, pady=(20, 20))
    
    def _setup_main_layout(self):
        """Configura o layout principal com cores do Excel."""
        self.main_frame = customtkinter.CTkFrame(
            self,
            fg_color=DashboardColors.MAIN_FRAME_BACKGROUND,  # Fundo escuro como Excel
            border_width=1,
            border_color=DashboardColors.MAIN_FRAME_BORDER  # Borda azul Excel
        )
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Frame esquerdo para botões
        self.left_frame = customtkinter.CTkFrame(
            self.main_frame, 
            width=300,
            fg_color=DashboardColors.LEFT_FRAME_BACKGROUND,  # Fundo escuro
            border_width=1,
            border_color=DashboardColors.LEFT_FRAME_BORDER  # Borda azul Excel
        )
        self.left_frame.pack(side="left", fill="y", padx=10, pady=10)
        self.left_frame.pack_propagate(False)  # Manter largura fixa
        
        # Título do frame esquerdo
        self._setup_left_title()
        
        # Frame direito para lista de planilhas
        self.right_frame = customtkinter.CTkFrame(
            self.main_frame,
            fg_color=DashboardColors.RIGHT_FRAME_BACKGROUND,  # Fundo escuro
            border_width=1,
            border_color=DashboardColors.RIGHT_FRAME_BORDER  # Borda azul Excel
        )
        self.right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)
    
    def _setup_left_title(self):
        """Configura o título do painel esquerdo com estilo Excel."""
        ctkTitle = customtkinter.CTkTextbox(
            self.left_frame, 
            height=30, 
            font=("Inter", 18, "bold"), 
            text_color=DashboardColors.LEFT_TITLE_TEXT  # Texto branco
        ) 
        ctkTitle.tag_config("center", justify='center')
        ctkTitle.insert("0.0", "Ações")
        ctkTitle.tag_add("center", "1.0", "end")
        ctkTitle.configure(state="disabled")
        ctkTitle.pack(fill="x", padx=20, pady=(20, 10))
    
    def _setup_button(self, text:str, command=None):
        """Configura um botão com estilo Excel."""
        return customtkinter.CTkButton(
            self.left_frame, 
            fg_color=DashboardColors.BUTTON_BACKGROUND_SECONDARY,  # Azul Excel
            hover_color=DashboardColors.BUTTON_HOVER,  # Azul mais escuro no hover
            text_color=DashboardColors.BUTTON_TEXT,
            width=400
        )
        
    def _setup_render_button(self, text:str,fg_color, command=None):
        self.render = customtkinter.CTkButton(
            self.left_frame, 
            text=text, 
            fg_color=fg_color,  # Azul Excel
            hover_color=DashboardColors.BUTTON_HOVER,
            text_color=DashboardColors.BUTTON_TEXT,
            command=command
        )
        return self.render.pack(pady=10)
    
    def _setup_buttons(self):
        """Configura todos os botões no painel esquerdo com cores do Excel."""
        # Botão para relatório por cliente
        self._setup_render_button("Gerar relatórios por cliente", DashboardColors.BUTTON_BACKGROUND_SECONDARY, self.read_and_create_client_folders)
        
        # Botão para atualizar lista
        self._setup_render_button("Atualizar Lista", DashboardColors.BUTTON_BACKGROUND_SECONDARY, self.load_excel_files)
        
        # Botão para visualizar planilha selecionada
        self._setup_render_button("Visualizar Planilha", DashboardColors.BUTTON_BACKGROUND_SECONDARY, self.view_selected_excel)
        
        # Botão de logout/sair
        self._setup_render_button("Sair", DashboardColors.BUTTON_CRITICAL_BACKGROUND, self.logout)
        
    
    def _setup_excel_list(self):
        """Configura a lista de planilhas no painel direito com cores do Excel."""
        # Título do frame direito
        self.right_title = customtkinter.CTkLabel(
            self.right_frame, 
            text="Planilhas Disponíveis", 
            font=("Inter", 18, "bold"),
            text_color=DashboardColors.RIGHT_TITLE_TEXT  # Texto branco para contraste
        )
        self.right_title.pack(pady=10)
        
        # Scrollable frame para a lista
        self.scrollable_frame = customtkinter.CTkScrollableFrame(
            self.right_frame, 
            height=400,
            fg_color=DashboardColors.SCROLL_BACKGROUND,  # Fundo escuro para scroll
            border_width=1,
            border_color=DashboardColors.SCROLL_BORDER  # Borda azul Excel
        )
        self.scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Dicionário para armazenar checkboxes e seus caminhos
        self.checkboxes = {}
        self.checkbox_frames = {}  # Frames para background dos checkboxes
        
        # Carregar planilhas
        self.load_excel_files()
        
        # Área para resultados
        self.result_textbox = customtkinter.CTkTextbox(
            self.right_frame, 
            height=100,
            fg_color=DashboardColors.RESULT_TEXTBOX_BACKGROUND,  # Fundo escuro para textbox
            border_width=1,
            border_color=DashboardColors.RESULT_TEXTBOX_BORDER,  # Borda azul Excel
            text_color=DashboardColors.RESULT_TEXTBOX_TEXT  # Texto branco
        )
        self.result_textbox.pack(fill="x", padx=10, pady=10)
    
    def logout(self):
        """Retorna à tela de login."""
        self.pack_forget()
        self.master.meu_form.pack(pady=20, expand=True)  # type: ignore
    
    def _on_frame_hover(self, frame, is_hover):
        """
        Aplica efeito hover no frame usando cores do Excel no tema escuro.
        
        - Hover: Azul Excel claro (#6B9BD1) - feedback visual sutil
        - Normal: Mantém a cor atual baseada na seleção
        """
        if is_hover:
            # Salvar cor atual antes do hover
            current_color = frame.cget("fg_color")
            frame._original_color = current_color
            frame.configure(fg_color="#6B9BD1")  # Azul Excel hover (#6B9BD1)
        else:
            # Restaurar cor original
            if hasattr(frame, '_original_color'):
                frame.configure(fg_color=frame._original_color)
    
    def _update_checkbox_background(self, frame, checkbox):
        """
        Atualiza a cor de fundo do frame baseado no estado do checkbox.
        
        Usa paletas de cores do Excel no tema escuro:
        - Não selecionado: Fundo escuro (#2a2a2a)
        - Selecionado: Azul Excel (#4F81BD)
        - Hover: Azul mais claro (#6B9BD1)
        """
        if checkbox.get() == 1:  # Selecionado
            frame.configure(fg_color="#4F81BD")  # Azul Excel
        else:  # Não selecionado
            frame.configure(fg_color="#2a2a2a")  # Fundo escuro
    
    def load_excel_files(self):
        """
        Carrega e exibe a lista de arquivos Excel (.xlsx) encontrados.
        
        Procura recursivamente por arquivos .xlsx no diretório atual
        e subdiretórios, criando frames com checkboxes para seleção.
        Cada item tem background que muda quando selecionado.
        """
        # Limpar checkboxes anteriores
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.checkboxes = {}
        self.checkbox_frames = {}  # Novo dicionário para armazenar os frames
        
        # Procurar arquivos .xlsx no diretório atual e subdiretórios
        excel_files = glob.glob("**/*.xlsx", recursive=True)
        
        if not excel_files:
            no_files_label = customtkinter.CTkLabel(self.scrollable_frame, text="Nenhuma planilha encontrada.")
            no_files_label.pack(pady=20)
            return
        
        for file_path in excel_files:
            # Criar frame para cada item com background Excel
            item_frame = customtkinter.CTkFrame(
                self.scrollable_frame,
                fg_color="#2a2a2a",  # Fundo escuro Excel (#2a2a2a)
                corner_radius=5,
                border_width=1,
                border_color="#404040"  # Borda sutil
            )
            item_frame.pack(fill="x", padx=5, pady=2)
            
            # Adicionar eventos de mouse para hover effect
            item_frame.bind("<Enter>", lambda e, f=item_frame: self._on_frame_hover(f, True))
            item_frame.bind("<Leave>", lambda e, f=item_frame: self._on_frame_hover(f, False))
            
            # Checkbox dentro do frame
            checkbox = customtkinter.CTkCheckBox(
                item_frame, 
                text=os.path.basename(file_path)
            )
            checkbox.pack(anchor="w", padx=10, pady=5)
            
            # Configurar callback para mudança de background
            checkbox.configure(command=lambda f=item_frame, c=checkbox: self._update_checkbox_background(f, c))
            
            # Armazenar referências
            self.checkboxes[checkbox] = file_path
            self.checkbox_frames[checkbox] = item_frame
    
    def create_new_excel(self):
        """
        Cria uma nova planilha Excel formatada com cabeçalhos padrão.
        
        Abre um diálogo para salvar o arquivo na pasta Planilhas/,
        aplicando formatação visual (fonte, cores, alinhamento).
        """
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        import tkinter.filedialog as filedialog
        import os
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Clientes" # pyright: ignore[reportOptionalMemberAccess]
        
        # Headers
        headers = ["Nome do Cliente", "Nome do Produto", "Ativo (Sim/Não)", "Plano (Ativado/Vencimento)"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header) # type: ignore
            # Formatação bonita: fonte branca, fundo azul, centralizado
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        # Ajustar largura das colunas automaticamente
        for col in ws.columns: # pyright: ignore[reportOptionalMemberAccess]
            max_length = 0
            column = col[0].column_letter # pyright: ignore[reportAttributeAccessIssue]
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width # type: ignore
        
        # Garantir pasta Planilhas existente
        planilhas_dir = os.path.join(os.getcwd(), "Planilhas")
        clientes_dir = os.path.join(os.getcwd(), "Clientes")
        os.makedirs(planilhas_dir, exist_ok=True)
        os.makedirs(clientes_dir, exist_ok=True)

        # Salvar arquivo na pasta Planilhas (dados de fonte), pastas de cliente vão para Clientes
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=planilhas_dir
        )
        if file_path:
            wb.save(file_path)
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", f"Planilha criada e formatada: {file_path}")
            self.load_excel_files()  # Recarregar lista
        else:
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", "Criação cancelada.")
    
    def read_selected_excels(self):
        """
        Lê os arquivos Excel selecionados e conta os nomes de clientes.
        
        Processa múltiplas planilhas selecionadas via checkboxes,
        extraindo nomes da segunda coluna e exibindo contagem.
        """
        selected_files = [path for checkbox, path in self.checkboxes.items() if checkbox.get() == 1]
        
        if not selected_files:
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", "Nenhuma planilha selecionada.")
            return
        
        results = []
        for file_path in selected_files:
            try:
                wb = openpyxl.load_workbook(file_path)
                sheet = wb.active
                names = [row[0] for row in sheet.iter_rows(min_row=2, values_only=True) if row[0]] # type: ignore
                results.append(f"{os.path.basename(file_path)}: {len(names)} nomes lidos")
            except Exception as e:
                results.append(f"Erro ao ler {os.path.basename(file_path)}: {str(e)}")
        
        self.result_textbox.delete("0.0", "end")
        self.result_textbox.insert("0.0", "\n".join(results))
    
    def generate_folders(self):
        """
        Gera pastas organizadas por cliente a partir das planilhas selecionadas.
        
        Para cada planilha selecionada, cria uma estrutura de diretórios
        baseada nos nomes dos clientes encontrados.
        """
        selected_files = [path for checkbox, path in self.checkboxes.items() if checkbox.get() == 1]
        
        if not selected_files:
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", "Nenhuma planilha selecionada para gerar pastas.")
            return
        
        from Interface.Classes.folders import folder
        
        for file_path in selected_files:
            try:
                wb = openpyxl.load_workbook(file_path)
                sheet = wb.active
                names = [row[0] for row in sheet.iter_rows(min_row=2, values_only=True) if row[0]] # type: ignore
                if names:
                    folder_handler = folder(sheet.title) # pyright: ignore[reportOptionalMemberAccess]
                    folder_handler.generateFolders(sheet.title, names) # type: ignore
                    self.result_textbox.insert("end", f"Pastas geradas para {os.path.basename(file_path)}\n")
                else:
                    self.result_textbox.insert("end", f"Planilha {os.path.basename(file_path)} vazia.\n")
            except Exception as e:
                self.result_textbox.insert("end", f"Erro em {os.path.basename(file_path)}: {str(e)}\n")
    
    def read_and_create_client_folders(self):
        """
        Processa planilhas selecionadas e cria relatórios consolidados por cliente.
        
        Gera arquivos de resumo financeiro na pasta Clientes/,
        consolidando dados de mão de obra, domínio e hospedagem.
        """
        selected_files = [path for checkbox, path in self.checkboxes.items() if checkbox.get() == 1]
        if not selected_files:
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", "Nenhuma planilha selecionada para relatório.")
            return

        summary = []
        for file_path in selected_files:
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                sheet = wb.active

                clients = {}
                for row in sheet.iter_rows(min_row=2, values_only=True): # pyright: ignore[reportOptionalMemberAccess]
                    if not row or not row[0]:
                        continue
                    client = str(row[0]).strip()
                    mano_obra = self._parse_currency(row[2])  # type: ignore # Valor Mão de Obra
                    cobrado_dominio = self._parse_currency(row[4]) if len(row) > 4 else 0.0  # pyright: ignore[reportAttributeAccessIssue] # Cobrado (Domínio)
                    cobrado_hospedagem = self._parse_currency(row[6]) if len(row) > 6 else 0.0  # type: ignore # Cobrado (Hospedagem)

                    total = mano_obra + cobrado_dominio + cobrado_hospedagem
                    if client not in clients:
                        clients[client] = {
                            'mano_obra': 0.0,
                            'cobrado': 0.0,
                            'total': 0.0,
                            'linhas': 0
                        }
                    clients[client]['mano_obra'] += mano_obra
                    clients[client]['cobrado'] += (cobrado_dominio + cobrado_hospedagem)
                    clients[client]['total'] += total
                    clients[client]['linhas'] += 1

                for client, data in clients.items():
                    client_dir = os.path.join('Clientes', client)
                    os.makedirs(client_dir, exist_ok=True)

                    with open(os.path.join(client_dir, 'resumo.txt'), 'w', encoding='utf-8') as f:
                        f.write(f"Relatorio do(a) cliente: {client}\n")
                        f.write("="*30 + "\n")
                        f.write(f"Linhas: {data['linhas']}\n")
                        f.write(f"Total mão de obra: R$ {data['mano_obra']:.2f}\n")
                        f.write(f"Total cobrado (domínio+hospedagem): R$ {data['cobrado']:.2f}\n")
                        f.write(f"Total geral: R$ {data['total']:.2f}\n")

                summary.append(f"{os.path.basename(file_path)}: processado {len(clients)} clientes")

            except Exception as e:
                summary.append(f"Erro ao processar {os.path.basename(file_path)}: {str(e)}")

        self.result_textbox.delete("0.0", "end")
        self.result_textbox.insert("0.0", "\n".join(summary))

    def view_selected_excel(self):
        """
        Exibe uma visualização gráfica da primeira planilha selecionada.
        
        Abre uma nova janela com grid responsivo mostrando os dados,
        com cores dinâmicas baseadas no conteúdo (verde=positivo, laranja=atenção).
        """
        # 1. PEGAR OS ARQUIVOS SELECIONADOS (Correção do erro de variável)
        selected_files = [path for checkbox, path in self.checkboxes.items() if checkbox.get() == 1]
        
        if not selected_files:
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", "Nenhuma planilha selecionada para visualizar.")
            return
        
        file_path = selected_files[0]  # Pega a primeira selecionada
        
        try:
            # Carrega a planilha
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active
            
            # Criar janela de visualização com tema Excel
            view_window = customtkinter.CTkToplevel(self)
            view_window.title(f"DaniTechnologia - {os.path.basename(file_path)}")
            view_window.geometry("950x600")
            view_window.attributes("-topmost", True)
            view_window.configure(fg_color="#2b2b2b")  # Fundo escuro Excel

            # Frame com Scroll
            scroll_frame = customtkinter.CTkScrollableFrame(
                view_window,
                fg_color="#1e1e1e",  # Fundo escuro
                border_width=1,
                border_color="#4F81BD"  # Borda azul Excel
            )
            scroll_frame.pack(fill="both", expand=True, padx=20, pady=20)

            # 2. CONFIGURAR COLUNAS
            # Lemos a primeira linha para saber quantas colunas existem
            headers = [cell.value for cell in sheet[1]] # type: ignore
            for i in range(len(headers)):
                scroll_frame.grid_columnconfigure(i, weight=1, minsize=150)

            # 3. RENDERIZAR CABEÇALHO com cores Excel
            for i, header in enumerate(headers):
                h_label = customtkinter.CTkLabel(
                    scroll_frame, 
                    text=str(header or "").upper(), 
                    font=("Inter", 12, "bold"),
                    fg_color="#4F81BD",  # Azul Excel header
                    text_color="white",
                    corner_radius=5,
                    height=35
                )
                h_label.grid(row=0, column=i, padx=2, pady=10, sticky="nsew")

            # 4. RENDERIZAR DADOS
            # iter_rows começando da linha 2 (para pular o cabeçalho)
            for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1): # type: ignore
                for col_idx, cell_value in enumerate(row):
                    val_str = str(cell_value or "")
                    
                    # Cores baseadas no conteúdo (Sua estratégia de controle)
                    txt_color = "white"
                    if "Pago" in val_str or "Sim" in val_str:
                        txt_color = "#2ecc71" # Verde para sucesso
                    elif "Pendente" in val_str or "Não" in val_str:
                        txt_color = "#e67e22" # Laranja para atenção
                    
                    b_label = customtkinter.CTkLabel(
                        scroll_frame, 
                        text=val_str,
                        font=("Inter", 11),
                        text_color=txt_color,
                        fg_color="#3a3a3a",  # Fundo escuro para células
                        height=35
                    )
                    b_label.grid(row=row_idx, column=col_idx, padx=1, pady=1, sticky="nsew")

        except Exception as e:
            self.result_textbox.delete("0.0", "end")
            self.result_textbox.insert("0.0", f"Erro ao abrir {os.path.basename(file_path)}: {str(e)}")
    
    def _parse_currency(self, value):
        """
        Converte um valor para float, tratando diferentes formatos de moeda.
        
        Args:
            value: Valor a ser convertido (string ou numérico)
            
        Returns:
            float: Valor convertido ou 0.0 se inválido
        """
        if value is None:
            return 0.0
        try:
            # Remove símbolos de moeda e espaços
            str_value = str(value).replace('R$', '').replace('$', '').replace(' ', '').replace(',', '.')
            return float(str_value)
        except (ValueError, TypeError):
            return 0.0