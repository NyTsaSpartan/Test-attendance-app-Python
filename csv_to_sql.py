import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import sqlite3
import csv

class CSVtoDBConverter:
    def __init__(self, master):
        self.master = master
        self.master.title("CSV to DB Converter")

        self.label_source = tk.Label(self.master, text="Chemin du fichier CSV source:")
        self.label_source.grid(row=0, column=0, padx=10, pady=10)
        self.source_var = tk.StringVar()
        self.source_entry = tk.Entry(self.master, textvariable=self.source_var)
        self.source_entry.grid(row=0, column=1, padx=10, pady=10)
        self.button_browse_source = tk.Button(self.master, text="Parcourir", command=self.browse_source)
        self.button_browse_source.grid(row=0, column=2, padx=10, pady=10)

        self.label_destination = tk.Label(self.master, text="Chemin de sortie de la base de données:")
        self.label_destination.grid(row=1, column=0, padx=10, pady=10)
        self.destination_var = tk.StringVar()
        self.destination_entry = tk.Entry(self.master, textvariable=self.destination_var)
        #self.destination_entry.grid(row=1, column=1, padx=10, pady=10)
        self.button_browse_destination = tk.Button(self.master, text="Parcourir", command=self.browse_destination)
        #self.button_browse_destination.grid(row=1, column=2, padx=10, pady=10)

        self.button_convert = tk.Button(self.master, text="Convertir", command=self.convert_csv_to_db)
        self.button_convert.grid(row=2, column=0, columnspan=3, pady=10)

        self.button_display_db = tk.Button(self.master, text="Afficher le contenu de la base de données", command=self.display_db_content)
        self.button_display_db.grid(row=3, column=0, columnspan=3, pady=10)

        self.status_label = tk.Label(self.master, text="")
        self.status_label.grid(row=4, column=0, columnspan=3)

    def browse_source(self):
        file_path = filedialog.askopenfilename(filetypes=[("Fichiers CSV", "*.csv")])
        self.source_var.set(file_path)

    def browse_destination(self):
        folder_path = filedialog.askdirectory()
        self.destination_var.set(folder_path)

    def convert_csv_to_db(self):
        source_path = self.source_var.get()
        #destination_path = self.destination_var.get()
        destination_path = self.source_var.get()
        print(f"Source path : ", source_path)
        print(f"Destination path : ", destination_path)

        if not source_path or not destination_path:
            messagebox.showwarning("Attention", "Veuillez spécifier le chemin du fichier source et de destination.")
            return

        try:
            conn = sqlite3.connect(destination_path + "/car_pointeuse.db")
            cursor = conn.cursor()

            # Créer la table s'il n'existe pas
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS pointeuse (
                    id_personne INTEGER,
                    nom_personne TEXT,
                    service TEXT,
                    heure TEXT,
                    UNIQUE (id_personne, nom_personne, service, heure)
                )
            ''')

            with open(source_path, 'r', encoding='cp1252') as file:
                csv_reader = csv.reader(file)
                headers = next(csv_reader)

                for row in csv_reader:
                    # Assurez-vous d'ajuster les indices en fonction de votre fichier CSV
                    id_personne = row[0]
                    nom_personne = row[1]
                    service = row[2]
                    heure = row[3]

                    # Insérer les données dans la table (ignorer les doublons)
                    cursor.execute("INSERT OR IGNORE INTO pointeuse VALUES (?, ?, ?, ?)", (id_personne, nom_personne, service, heure))

            conn.commit()
            conn.close()
            self.status_label.config(text="Conversion terminée avec succès.", fg="green")

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")
            self.status_label.config(text="Erreur lors de la conversion.", fg="red")


    def display_db_content(self):
        destination_path = self.destination_var.get()

        if not destination_path:
            messagebox.showwarning("Attention", "Veuillez spécifier le chemin de la base de données.")
            return

        try:
            conn = sqlite3.connect(destination_path + "/car_pointeuse.db")
            cursor = conn.cursor()

            # Exécutez la requête SELECT
            cursor.execute("SELECT * FROM pointeuse")
            rows = cursor.fetchall()

            # Créez une fenêtre pour afficher les résultats
            display_window = tk.Toplevel(self.master)
            display_window.title("Contenu de la base de données")

            text_widget = scrolledtext.ScrolledText(display_window, wrap=tk.WORD)
            text_widget.pack(expand=True, fill='both')

            # Affichez les résultats dans le widget de texte
            for row in rows:
                text_widget.insert(tk.END, f"{row}\n")

            conn.close()

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")

# Le reste du code reste inchangé...

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVtoDBConverter(root)
    root.mainloop()
