import tkinter as tk
import time

start_time = time.time()
start_time_str = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime(start_time))

class UptimeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Temps Ordinateur")
        
        # Créer un cadre pour contenir les widgets
        self.frame = tk.Frame(root)
        self.frame.pack(expand=True, fill='both')

        # Label pour le texte "Temps d'Allumage"
        self.text_label = tk.Label(self.frame, text="Temps d'Allumage", font=("Helvetica", 20))
        self.text_label.pack(pady=(20, 5), anchor='center')

        # Ajouter un canevas pour dessiner une ligne
        self.canvas = tk.Canvas(self.frame, height=2, bg="black")
        self.canvas.pack(fill='x', pady=(5, 20))

        # Label pour afficher la date et heure du démarrage
        self.start_time_label = tk.Label(self.frame, text=f"Démarré le : {start_time_str}", font=("Helvetica", 10))
        self.start_time_label.pack(pady=(5, 5), anchor='center')

        # Label pour afficher le nombre de jours (initialement caché)
        self.days_label = tk.Label(self.frame, text="", font=("Helvetica", 36))
        self.days_label.pack(pady=(5, 5), anchor='center')
        
        # Label pour afficher le reste du temps (heures:minutes:secondes)
        self.time_label = tk.Label(self.frame, text="", font=("Helvetica", 48))
        self.time_label.pack(pady=(5, 20), anchor='center')

        self.update_uptime()

    def update_uptime(self):
        # Calculer le temps écoulé depuis le démarrage du programme
        current_time = time.time()
        uptime_duration = current_time - start_time
        
        # Convertir les secondes en jours, heures, minutes, secondes
        days = int(uptime_duration // (24 * 3600))
        uptime_duration %= (24 * 3600)
        hours = int(uptime_duration // 3600)
        uptime_duration %= 3600
        minutes = int(uptime_duration // 60)
        seconds = int(uptime_duration % 60)
        
        # Formater les jours et le reste du temps
        if days > 0:
            days_str = f"{days} jour{'s' if days > 1 else ''}"
            self.days_label.config(text=days_str)
            self.days_label.pack()  # Afficher le label des jours
        else:
            self.days_label.pack_forget()  # Cacher le label des jours

        time_str = f"{hours:02}:{minutes:02}:{seconds:02}"
        self.time_label.config(text=time_str)
        
        # Mettre à jour toutes les secondes
        self.root.after(1000, self.update_uptime)

if __name__ == "__main__":
    root = tk.Tk()
    app = UptimeApp(root)
    root.mainloop()
