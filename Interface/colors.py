"""
Paleta de Cores - Tema Excel para Leitor Automático

Este arquivo contém todas as cores da interface organizadas por categoria.
Para alterar o visual da aplicação, basta modificar os valores aqui.

Paleta baseada no Microsoft Excel:
- Azul Excel: #4F81BD (azul clássico do Excel)
- Fundo escuro: #2b2b2b (fundo moderno)
- Contraste alto: Branco em fundos escuros
"""

class ExcelColors:
    """Paleta de cores inspirada no Microsoft Excel"""

    # === CORES PRINCIPAIS DO EXCEL ===
    EXCEL_BLUE = "#4F81BD"          # Azul clássico do Excel
    EXCEL_BLUE_LIGHT = "#6B9BD1"    # Azul mais claro para hover
    EXCEL_BLUE_DARK = "#2E5C8A"     # Azul mais escuro para hover

    # === VERMELHO PARA AÇÕES CRÍTICAS ===
    EXCEL_RED = "#C00000"           # Vermelho Excel para sair/fechar
    EXCEL_RED_DARK = "#8B0000"      # Vermelho mais escuro

    # === VERDES E LARANJA PARA STATUS ===
    SUCCESS_GREEN = "#2ecc71"       # Verde para status positivo
    WARNING_ORANGE = "#e67e22"      # Laranja para atenção/pendente

    # === CINZAS E FUNDO ===
    GRAY_LIGHT = "#f8f9fa"          # Cinza muito claro (modo claro)
    GRAY_MEDIUM = "#e1e5e9"         # Cinza médio para bordas
    GRAY_DARK = "#404040"          # Cinza escuro para bordas

class ThemeColors:
    """Cores do tema escuro inspirado no Excel"""

    # === FUNDO DA APLICAÇÃO ===
    APP_BACKGROUND = "#2b2b2b"      # Fundo principal da aplicação

    # === FRAMES E PAINÉIS ===
    FRAME_BACKGROUND = "#1e1e1e"    # Fundo dos frames principais
    PANEL_BACKGROUND = "#2a2a2a"    # Fundo dos painéis internos
    INPUT_BACKGROUND = "#3a3a3a"    # Fundo dos campos de entrada

    # === SCROLLABLE FRAME ===
    SCROLL_BACKGROUND = "#2a2a2a"   # Fundo do scrollable frame

    # === TEXTO ===
    TEXT_PRIMARY = "#ffffff"        # Texto principal (branco)
    TEXT_SECONDARY = "#cccccc"      # Texto secundário (cinza claro)

    # === BORDAS ===
    BORDER_PRIMARY = ExcelColors.EXCEL_BLUE    # Borda principal
    BORDER_SECONDARY = "#404040"   # Borda secundária

    # === SELEÇÃO E HOVER ===
    SELECTION_BACKGROUND = ExcelColors.EXCEL_BLUE     # Fundo quando selecionado
    HOVER_BACKGROUND = ExcelColors.EXCEL_BLUE_LIGHT   # Fundo no hover

    # === BOTÕES ===
    BUTTON_PRIMARY = ExcelColors.EXCEL_BLUE           # Botão principal
    BUTTON_SECONDARY = ExcelColors.WARNING_ORANGE        # Botão secundário (azul mais claro)
    BUTTON_HOVER = ExcelColors.EXCEL_BLUE_DARK        # Hover do botão
    BUTTON_CRITICAL = ExcelColors.EXCEL_RED           # Botão crítico (sair)
    BUTTON_CRITICAL_HOVER = ExcelColors.EXCEL_RED_DARK # Hover crítico

class FormColors:
    """Cores específicas do formulário de login"""

    FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    TITLE_BACKGROUND = ExcelColors.EXCEL_BLUE
    TITLE_TEXT = ThemeColors.TEXT_PRIMARY

    LABEL_TEXT = ThemeColors.TEXT_PRIMARY

    ENTRY_BACKGROUND = ThemeColors.INPUT_BACKGROUND
    ENTRY_BORDER = ThemeColors.BORDER_PRIMARY
    ENTRY_TEXT = ThemeColors.TEXT_PRIMARY

    BUTTON_BACKGROUND = ThemeColors.BUTTON_PRIMARY
    BUTTON_HOVER = ThemeColors.BUTTON_HOVER
    BUTTON_TEXT = ThemeColors.TEXT_PRIMARY

class DashboardColors:
    """Cores específicas do dashboard"""

    TITLE_BACKGROUND = ExcelColors.EXCEL_BLUE
    TITLE_TEXT = ThemeColors.TEXT_PRIMARY

    MAIN_FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    MAIN_FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    LEFT_FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    LEFT_FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    RIGHT_FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    RIGHT_FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    LEFT_TITLE_BACKGROUND = ExcelColors.EXCEL_BLUE
    LEFT_TITLE_TEXT = ThemeColors.TEXT_PRIMARY

    BUTTON_BACKGROUND = ThemeColors.BUTTON_PRIMARY
    BUTTON_BACKGROUND_SECONDARY = ThemeColors.BUTTON_SECONDARY
    BUTTON_HOVER = ThemeColors.BUTTON_HOVER
    BUTTON_TEXT = ThemeColors.TEXT_PRIMARY

    BUTTON_CRITICAL_BACKGROUND = ThemeColors.BUTTON_CRITICAL
    BUTTON_CRITICAL_HOVER = ThemeColors.BUTTON_CRITICAL_HOVER
    BUTTON_CRITICAL_TEXT = ThemeColors.TEXT_PRIMARY

    RIGHT_TITLE_TEXT = ThemeColors.TEXT_PRIMARY

    SCROLL_BACKGROUND = ThemeColors.SCROLL_BACKGROUND
    SCROLL_BORDER = ThemeColors.BORDER_PRIMARY

    CHECKBOX_FRAME_BACKGROUND = ThemeColors.PANEL_BACKGROUND
    CHECKBOX_FRAME_SELECTED = ThemeColors.SELECTION_BACKGROUND
    CHECKBOX_FRAME_HOVER = ThemeColors.HOVER_BACKGROUND
    CHECKBOX_FRAME_BORDER = ThemeColors.BORDER_SECONDARY

    RESULT_TEXTBOX_BACKGROUND = ThemeColors.INPUT_BACKGROUND
    RESULT_TEXTBOX_BORDER = ThemeColors.BORDER_PRIMARY
    RESULT_TEXTBOX_TEXT = ThemeColors.TEXT_PRIMARY

class ViewColors:
    """Cores específicas da janela de visualização"""

    WINDOW_BACKGROUND = ThemeColors.APP_BACKGROUND

    SCROLL_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    SCROLL_BORDER = ThemeColors.BORDER_PRIMARY

    HEADER_BACKGROUND = ExcelColors.EXCEL_BLUE
    HEADER_TEXT = ThemeColors.TEXT_PRIMARY

    CELL_BACKGROUND = ThemeColors.INPUT_BACKGROUND
    CELL_TEXT = ThemeColors.TEXT_PRIMARY

    STATUS_SUCCESS = ExcelColors.SUCCESS_GREEN
    STATUS_WARNING = ExcelColors.WARNING_ORANGE

# === CONFIGURAÇÕES GLOBAIS ===
APPEARANCE_MODE = "dark"           # "light" ou "dark"
DEFAULT_COLOR_THEME = "blue"       # Tema base do CustomTkinter

# === FÁCIL ACESSO ===
# Para usar as cores, importe e use:
# from colors import DashboardColors, ExcelColors, etc.
#
# Exemplo:
# button.configure(fg_color=DashboardColors.BUTTON_BACKGROUND)
# frame.configure(fg_color=DashboardColors.FRAME_BACKGROUND)