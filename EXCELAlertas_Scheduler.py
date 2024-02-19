# Script que deve ser executado quando o trigger é feito pelo agendador de tarefas do windows
from EXCELAlertas import *

command_executed = False
root = tk.Tk()
app = App(root)

app.execute_all(command_executed)