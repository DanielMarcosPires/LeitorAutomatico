class InterfaceColors:
    """Paleta de cores da interface - Ajustada para Identidade ACELERA G&P"""

    # === CORES PRINCIPAIS (Substituído o 'Azul' pelo Amarelo da Logo) ===
    INTERFACE_BLUE = "#F9C319"          # Amarelo Principal da Logo
    INTERFACE_BLUE_LIGHT = "#FFD133"    # Amarelo mais claro para hover
    INTERFACE_BLUE_DARK = "#C79C14"     # Amarelo mais escuro para estados pressionados

    # === VERMELHO PARA AÇÕES CRÍTICAS ===
    # Ajustado para um coral/laranja vibrante que harmoniza com amarelo sem "brigar"
    INTERFACE_RED = "#FF6701"           # Vermelho moderno (Flat)
    INTERFACE_RED_DARK = "#C0392B"      # Vermelho escuro para hover

    # === VERDES E LARANJA PARA STATUS ===
    SUCCESS_GREEN = "#27AE60"           # Verde Esmeralda (Status Positivo)
    WARNING_ORANGE = "#F39C12"          # Laranja Flat (Atenção)

    # === CINZAS E FUNDO (Inspirado no fundo da logo) ===
    GRAY_LIGHT = "#2D2D2D"              # Cinza da logo (ideal para painéis laterais)
    GRAY_MEDIUM = "#3D3D3D"             # Cinza médio para separadores/bordas
    GRAY_DARK = "#1A1A1A"               # Cinza quase preto para o fundo principal

class ThemeColors:
    """Cores do tema escuro inspirado no Excel"""

    # === FUNDO DA APLICAÇÃO ===
    # Usei o tom de cinza carvão da logo para um visual mais "Pro"
    APP_BACKGROUND = "#1A1A1A"      # Fundo principal da aplicação

    # === FRAMES E PAINÉIS ===
    FRAME_BACKGROUND = "#242424"    # Fundo dos frames principais
    PANEL_BACKGROUND = "#2D2D2D"    # Fundo dos painéis internos
    INPUT_BACKGROUND = "#333333"    # Fundo dos campos de entrada

    # === SCROLLABLE FRAME ===
    SCROLL_BACKGROUND = "#242424"   # Fundo do scrollable frame

    # === TEXTO ===
    TEXT_PRIMARY = "#FFFFFF"        # Texto principal (branco)
    TEXT_SECONDARY = "#B0B0B0"      # Texto secundário (cinza claro)

    # === BORDAS ===
    BORDER_PRIMARY = InterfaceColors.INTERFACE_BLUE    # Borda principal (Amarelo)
    BORDER_SECONDARY = "#404040"   # Borda secundária

    # === SELEÇÃO E HOVER ===
    SELECTION_BACKGROUND = InterfaceColors.INTERFACE_BLUE     # Fundo quando selecionado
    HOVER_BACKGROUND = InterfaceColors.INTERFACE_BLUE_LIGHT   # Fundo no hover

    # === BOTÕES ===
    BUTTON_PRIMARY = InterfaceColors.INTERFACE_BLUE           # Botão principal
    BUTTON_SECONDARY = InterfaceColors.WARNING_ORANGE        # Botão secundário (Laranja)
    BUTTON_HOVER = InterfaceColors.INTERFACE_BLUE_DARK        # Hover do botão
    BUTTON_CRITICAL = InterfaceColors.INTERFACE_RED           # Botão crítico (sair)
    BUTTON_CRITICAL_HOVER = InterfaceColors.INTERFACE_RED_DARK # Hover crítico
class FormColors:
    """Cores específicas do formulário de login"""

    FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    TITLE_BACKGROUND = InterfaceColors.INTERFACE_BLUE
    TITLE_TEXT = ThemeColors.TEXT_PRIMARY

    LABEL_TEXT = ThemeColors.TEXT_PRIMARY

    ENTRY_BACKGROUND = ThemeColors.INPUT_BACKGROUND
    ENTRY_BORDER = ThemeColors.BORDER_PRIMARY
    ENTRY_TEXT = ThemeColors.TEXT_PRIMARY

    BUTTON_BACKGROUND = ThemeColors.BUTTON_HOVER
    BUTTON_HOVER = ThemeColors.BUTTON_PRIMARY
    BUTTON_TEXT = ThemeColors.TEXT_PRIMARY

class DashboardColors:
    """Cores específicas do dashboard"""

    TITLE_BACKGROUND = InterfaceColors.INTERFACE_BLUE
    TITLE_TEXT = ThemeColors.TEXT_PRIMARY

    MAIN_FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    MAIN_FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    LEFT_FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    LEFT_FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    RIGHT_FRAME_BACKGROUND = ThemeColors.FRAME_BACKGROUND
    RIGHT_FRAME_BORDER = ThemeColors.BORDER_PRIMARY

    LEFT_TITLE_BACKGROUND = InterfaceColors.INTERFACE_BLUE
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

    HEADER_BACKGROUND = InterfaceColors.INTERFACE_BLUE
    HEADER_TEXT = ThemeColors.TEXT_PRIMARY

    CELL_BACKGROUND = ThemeColors.INPUT_BACKGROUND
    CELL_TEXT = ThemeColors.TEXT_PRIMARY

    STATUS_SUCCESS = InterfaceColors.SUCCESS_GREEN
    STATUS_WARNING = InterfaceColors.WARNING_ORANGE

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