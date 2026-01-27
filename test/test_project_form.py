from xls_management.tui.project_form import ProjectChoice


def test_project_form():
    name, evaluate = ProjectChoice().run()
    