import sys

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from kpi_creator.ui.main_window import MainWindow


def _light_palette() -> QPalette:
    palette = QPalette()
    colors = {
        QPalette.ColorRole.Window: "#F4F6F8",
        QPalette.ColorRole.WindowText: "#172033",
        QPalette.ColorRole.Base: "#FFFFFF",
        QPalette.ColorRole.AlternateBase: "#F8FAFC",
        QPalette.ColorRole.ToolTipBase: "#FFFFFF",
        QPalette.ColorRole.ToolTipText: "#172033",
        QPalette.ColorRole.Text: "#172033",
        QPalette.ColorRole.Button: "#FFFFFF",
        QPalette.ColorRole.ButtonText: "#172033",
        QPalette.ColorRole.BrightText: "#B42318",
        QPalette.ColorRole.Link: "#0F766E",
        QPalette.ColorRole.Highlight: "#CCFBF1",
        QPalette.ColorRole.HighlightedText: "#0F766E",
        QPalette.ColorRole.PlaceholderText: "#98A2B3",
        QPalette.ColorRole.Mid: "#D0D5DD",
        QPalette.ColorRole.Dark: "#98A2B3",
        QPalette.ColorRole.Shadow: "#E4E7EC",
    }
    for role, color in colors.items():
        palette.setColor(role, QColor(color))
    for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor("#98A2B3"))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Base, QColor("#F8FAFC"))
    return palette


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("KPI Pivot Chart Creator")
    app.setOrganizationName("KPI Pivot Chart Creator")
    app.setStyle("Fusion")
    app.setPalette(_light_palette())
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
