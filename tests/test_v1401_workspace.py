from egomail.ui import MainWindow

def test_workspace_menu_target_exists():
    assert hasattr(MainWindow, "open_workspace_folder")
