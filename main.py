import tkinter as tk
from services import main_interface
from services.despesas.services.dashboard_despesas import tela_dashboard_despesas
from services.ui_theme import BG, SURFACE, ACCENT, FONT_SMALL, TEXT, TEXT_SUB, FONT_BODY, divider

from services.main_interface import tela_menu_principal

def main():

    root = tk.Tk()
    root.title("AutoMaker — Gerenciador de Automação")
    root.configure(bg=SURFACE)
    root.geometry("960x660")
    #FRAMES PRINCIPAIS MAIN E SIDEBAR


    frame_sidebar = tk.Frame(root, bg=SURFACE, width=220)
    frame_sidebar.pack(side="left", fill="y")
    frame_sidebar.pack_propagate(False)

    frame_main = tk.Frame(root, bg=BG)
    frame_main.pack(side="left", fill="both" ,expand=True)


    nav_items = [
            ("🗂  Automações",    True, tela_menu_principal),
            ("📊  Dashboards",    False, None),
            ("📋  Outros",        False, None),
            ("⚙  Configurações",  False, None),
        ]

    
    def construir_sidebar(frame_sidebar, frame_principal_atual):

        tk.Label(frame_sidebar, text="⬡", font=("Segoe UI", 28), bg=SURFACE, fg=ACCENT).pack(pady=(32, 4))
        tk.Label(frame_sidebar, text="AutoMaker", font=("Segoe UI", 13, "bold"), bg=SURFACE, fg=TEXT).pack()
        tk.Label(frame_sidebar, text="Gerenciador de Automação", font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(pady=(2, 24))

        divider(frame_sidebar)

        for label, active, tela_fn in nav_items:

            f_color = ACCENT if active else SURFACE
            t_color = TEXT if active else TEXT_SUB
            item = tk.Frame(frame_sidebar, bg=f_color, cursor="hand2")
            item.pack(fill="x", padx=12, pady=2)
            tk.Label(item, text=label, bg=f_color, fg=t_color,
                    font=FONT_BODY, padx=14, pady=9, anchor="w").pack(fill="x")
            
            item.bind("<Button-1>", lambda e, destino = tela_fn: rotear_tela(frame_principal_atual, destino))
        


    def rotear_tela(frame_atual, destino):

        for widget in frame_atual.winfo_children():
            widget.destroy()

        if destino:
            destino(frame_atual)


    construir_sidebar(frame_sidebar, frame_main)
    rotear_tela(frame_main,lambda p: tela_menu_principal(p,rotear_tela))        
    
    root.mainloop()


if __name__ == "__main__":
    main()