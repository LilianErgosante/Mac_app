import tkinter as tk
import subprocess
import os

def run_script(script_name):
    parent_dir = os.path.dirname(os.path.abspath(os.getcwd()))
    subprocess.Popen(["python", f"{script_name}.py"], cwd=parent_dir)

def main():
    root = tk.Tk()
    root.title("Sélectionner une étude")

    # Configuration de la fenêtre principale
    root.geometry("400x300")
    root.configure(bg="#2C3E50")

    # Ajout d'un titre
    title = tk.Label(root, text="Sélectionnez une étude", font=("Helvetica", 16, "bold"), fg="#ECF0F1", bg="#2C3E50")
    title.pack(pady=20)

    # Configuration des boutons
    button_font = ("Helvetica", 12)
    button_bg = "#3498DB"
    button_fg = "#ECF0F1"

    tk.Button(root, text="Étude lombaire", command=lambda: run_script("lombaire"),
              font=button_font, bg=button_bg, fg=button_fg, relief="flat", padx=10, pady=5).pack(pady=10)
    tk.Button(root, text="Étude épaule", command=lambda: run_script("epaule"),
              font=button_font, bg=button_bg, fg=button_fg, relief="flat", padx=10, pady=5).pack(pady=10)
    tk.Button(root, text="Étude comparative HAPO", command=lambda: run_script("HAPO"),
              font=button_font, bg=button_bg, fg=button_fg, relief="flat", padx=10, pady=5).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()

