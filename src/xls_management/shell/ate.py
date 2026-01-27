import cmd
from xls_management.ate import PROJECTS
from xls_management.ate.project import project_combo_box
from xls_management.utils.color import Color, ansi_color

class MyShell(cmd.Cmd):
    intro = (
        "Welcome to ate shell. Type help or ? to list commands.\n"
        f"Available projects: {', '.join(PROJECTS)}\n"
    )
    prompt = "(ate) "

    # Example command
    def do_project(self, arg):
        """project <name>"""
        if arg:
            print(arg)

    def do_choose(self, arg):
        """choose a project from a list using a combobox"""
        str_project = arg
        evalue_master_id: bool
        if not str_project:
            str_project, evalue_master_id = project_combo_box()
        print(f"project {ansi_color(str_project, Color.GREEN)} was chosen")
        msg = f"{ansi_color("evalue",Color.GREEN)} master id"
        if not evalue_master_id:
            msg = f"{ansi_color("do not evalue",Color.RED)} master id"
        print(msg)

    # Tab completion for 'project'
    def complete_project(self, text, line, begidx, endidx):
        """
        text: the current word being completed
        line: the full input line
        begidx, endidx: start and end indexes of the word
        """
        if not text:
            return PROJECTS  # Suggest all if nothing typed yet
        return [name for name in PROJECTS if name.lower().startswith(text.lower())]

    # Exit command
    def do_exit(self, arg):
        """Exit the shell"""
        print("Goodbye!")
        return True

def launch_shell():
    MyShell().cmdloop()

if __name__ == "__main__":
    MyShell().cmdloop()
