import tkinter as tk
from tkinter import messagebox, scrolledtext
from script_principal import lancer_traitement
import threading

# === Fonction pour afficher les logs dans l'interface ===
def log_vers_interface(message):
    log_area.insert(tk.END, message + "\n")
    log_area.see(tk.END)
    log_area.update_idletasks()

# === Fonctions des boutons (avec thread pour ne pas bloquer l'interface) ===
def lancer_test():
    if messagebox.askyesno("Confirmation", "Voulez-vous lancer le test de création des fichiers ?"):
        threading.Thread(target=lancer_traitement, kwargs={"envoi_actif": False, "callback_log": log_vers_interface}).start()

def lancer_envoi():
    if messagebox.askyesno("Confirmation", "Voulez-vous vraiment envoyer tous les emails ?"):
        threading.Thread(target=lancer_traitement, kwargs={"envoi_actif": True, "callback_log": log_vers_interface}).start()

# === Interface graphique ===
root = tk.Tk()
root.title("ENVOI DES RECUS FISCAUX")
try:
    # Windows : OK
    root.state('zoomed')
except Exception:
    # Linux/macOS : tenter -zoomed, sinon fallback fullscreen
    try:
        root.attributes('-zoomed', True)
    except Exception:
        root.attributes('-fullscreen', True)

# === Interface graphique ===
# === Titre ===
titre_label = tk.Label(
    root,
    text="CERCLE DES SENIORS NEUFLIZE OBC\n\nENVOI DES RECUS FISCAUX",
    font=("Arial", 28, "bold"),
    fg="#003366"
)
titre_label.pack(pady=20)

# === Zone de log ===
log_area = scrolledtext.ScrolledText(root, width=120, height=30, font=("Courier New", 11))
log_area.pack(padx=20, pady=20)

# === Boutons ===
frame_boutons = tk.Frame(root)
frame_boutons.pack()

btn_test = tk.Button(frame_boutons, text="Test création des fichiers", font=("Arial", 14), width=30, command=lancer_test)
btn_test.grid(row=0, column=0, padx=10, pady=10)

btn_envoi = tk.Button(frame_boutons, text="Envoi des emails", font=("Arial", 14), width=30, command=lancer_envoi)
btn_envoi.grid(row=1, column=0, padx=10, pady=10)

root.mainloop()
