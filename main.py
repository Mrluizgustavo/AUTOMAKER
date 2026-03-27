import tkinter as tk
from services import interface_master

def main():
    root = tk.Tk()
    root.title("AutoPanel — Gerenciador de Automação")
    root.configure(bg="#0A0A0F")
    interface_master.tela_menu_principal(root)
    root.mainloop()

if __name__ == "__main__":
    main()