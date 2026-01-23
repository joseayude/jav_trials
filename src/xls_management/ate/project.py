import os
from textual.app import App, ComposeResult
from textual.widgets import Select, Static
from textual.containers import Container
from xls_management.ate import PROJECTS

class ComboBoxApp(App):
    CSS = """
    Screen {
        align: center middle;
    }
    """

    def __init__(self):
        super().__init__()
        self.result: str = ""

    def compose(self) -> ComposeResult:
        # Create a combobox-like Select widget
        yield Container(
            Static("Choose a project:"),
            Select(
                options=[(p, p) for p in PROJECTS],
                prompt="Select project..."
            ),
            id="combo-container"
        )

    async def on_select_changed(self, event: Select.Changed) -> None:
        """Handle selection change event."""
        self.result = event.value
        await self.action_quit()
    
def project_combo_box() -> str:
    app = ComboBoxApp()
    app.run()
    return app.result

if __name__ == "__main__":
    ComboBoxApp().run()