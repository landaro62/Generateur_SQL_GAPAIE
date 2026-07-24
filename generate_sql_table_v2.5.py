import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
import sys
import datetime


def resource_path(relative_path):
    """Résout un chemin de ressource, que l'app tourne en script ou en .exe PyInstaller."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class App:
    def __init__(self, root):
        self.root = root
        root.title("Générateur SQL depuis CSV/Excel")
        root.resizable(True, True)
        try:
            root.iconbitmap(resource_path(os.path.join("medias", "icone.ico")))
        except tk.TclError:
            pass

        self.df = None
        self.table_name = None
        self.col_vars = {}
        self.key_vars = {}
        self.key_checkbuttons = {}
        self.query_mode = tk.StringVar(value="insert")

        padding = {'padx': 10, 'pady': 5}
        frame = ttk.Frame(root)
        frame.pack(fill="both", expand=True, **padding)

        # Bloc explicatif
        help_text = (
            "Ce petit utilitaire vous permet de transformer un fichier CSV ou Excel issu\n"
            "d'une table GAPAIE en un script SQL (INSERT ou UPDATE). Procédure :\n"
            "Formats acceptés : CSV (UTF8 ou ANSI, séparateur point-virgule), XLSX, XLS\n"
            "1) Cliquez sur “Charger fichier…” et sélectionnez votre fichier.\n"
            "2) Choisissez le type de requête (INSERT ou UPDATE).\n"
            "3) Cochez les colonnes que vous souhaitez inclure. En mode UPDATE,\n"
            "   cochez aussi “Clé (WHERE)” pour la ou les colonnes qui identifient\n"
            "   la ligne (utilisées dans la clause WHERE, exclues du SET).\n"
            "4) Cliquez sur “Générer SQL…” pour créer et sauvegarder\n"
            "   le fichier .sql contenant vos requêtes.\n\n"
            "Note : Toute valeur au format JJ/MM/AAAA sera automatiquement convertie\n"
            "      en AAAAMMJJ, et les nombres décimaux avec ',' seront convertis en '.'"
        )
        help_frame = ttk.LabelFrame(frame, text="Comment ça marche", padding=8)
        help_frame.grid(row=0, column=0, columnspan=2, sticky="ew", **padding)
        ttk.Label(help_frame, text=help_text, justify="left").pack(fill="x", expand=True)

        # Ligne de chargement du fichier source
        self.btn_load = ttk.Button(frame, text="Charger fichier…", command=self.load_csv)
        self.btn_load.grid(row=1, column=0, **padding)
        self.csv_label = ttk.Label(frame, text="Aucun fichier chargé", width=40)
        self.csv_label.grid(row=1, column=1, **padding)

        # Type de requête
        mode_frame = ttk.LabelFrame(frame, text="Type de requête", padding=8)
        mode_frame.grid(row=2, column=0, columnspan=2, sticky="ew", **padding)
        ttk.Radiobutton(
            mode_frame, text="INSERT", variable=self.query_mode, value="insert", command=self._on_mode_change
        ).pack(side="left", padx=(0, 15))
        ttk.Radiobutton(
            mode_frame, text="UPDATE", variable=self.query_mode, value="update", command=self._on_mode_change
        ).pack(side="left")

        # LabelFrame scrollable pour les colonnes
        self.cols_frame = ttk.LabelFrame(frame, text="Sélection des colonnes", padding=5)
        self.cols_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", **padding)

        # Barre Tout cocher / Tout décocher
        btns_bar = ttk.Frame(self.cols_frame)
        btns_bar.pack(side="top", fill="x", pady=(0, 5))
        self.btn_select_all = ttk.Button(btns_bar, text="Tout cocher", command=self.select_all, state="disabled")
        self.btn_select_all.pack(side="left", padx=(0, 5))
        self.btn_deselect_all = ttk.Button(btns_bar, text="Tout décocher", command=self.deselect_all, state="disabled")
        self.btn_deselect_all.pack(side="left")

        # Canvas + Scrollbar pour les checkboxes
        canvas_frame = ttk.Frame(self.cols_frame)
        canvas_frame.pack(side="top", fill="both", expand=True)
        self.canvas = tk.Canvas(canvas_frame, borderwidth=0, height=180)
        self.vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Frame interne dans le canvas
        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        # Molette de souris uniquement quand le curseur survole la zone de colonnes
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # Aperçu des données
        self.preview_frame = ttk.LabelFrame(frame, text="Aperçu des données (10 premières lignes)", padding=5)
        self.preview_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", **padding)

        tree_container = ttk.Frame(self.preview_frame, width=700, height=200)
        tree_container.pack(fill="both", expand=True)
        tree_container.grid_propagate(False)
        self.tree = ttk.Treeview(tree_container, show="headings", height=8)
        tree_vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        tree_hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_vsb.grid(row=0, column=1, sticky="ns")
        tree_hsb.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # Bouton Générer SQL (toujours visible, activé une fois un CSV chargé)
        self.btn_gen = ttk.Button(frame, text="Générer SQL…", command=self.generate_sql, state="disabled")
        self.btn_gen.grid(row=5, column=0, columnspan=2, pady=(10, 5))

        frame.grid_rowconfigure(3, weight=1)
        frame.grid_rowconfigure(4, weight=1)
        frame.grid_columnconfigure(1, weight=1)

    def load_csv(self):
        path = filedialog.askopenfilename(
            title="Sélectionnez le fichier source",
            filetypes=[
                ("Fichiers supportés", "*.csv *.xlsx *.xls"),
                ("CSV (point-virgule)", "*.csv"),
                ("Excel", "*.xlsx *.xls"),
                ("Tous fichiers", "*.*"),
            ]
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == '.csv':
                df = self._read_csv_auto(path)
            elif ext in ('.xlsx', '.xls'):
                df = self._read_excel_auto(path, ext)
            else:
                messagebox.showerror("Erreur", "Format de fichier non supporté : {}".format(ext))
                return
        except Exception as e:
            messagebox.showerror("Erreur", "Impossible de lire le fichier :\n{}".format(e))
            return

        self.df = df
        self.table_name = os.path.splitext(os.path.basename(path))[0]
        self.csv_label.config(text=os.path.basename(path))
        self._show_column_selection()
        self._show_preview()

        self.btn_select_all.config(state="normal")
        self.btn_deselect_all.config(state="normal")
        self.btn_gen.config(state="normal")

    def _read_csv_auto(self, path):
        """
        Essaie de lire le CSV en UTF-8, puis se rabat sur l'encodage ANSI
        (Windows-1252) si la lecture échoue à cause de l'encodage.
        """
        try:
            return pd.read_csv(path, sep=';', dtype=str, encoding='utf-8-sig').fillna('')
        except UnicodeDecodeError:
            return pd.read_csv(path, sep=';', dtype=str, encoding='cp1252').fillna('')

    def _read_excel_auto(self, path, ext):
        """
        Lit un fichier Excel (.xlsx via openpyxl, .xls via xlrd), première feuille,
        et convertit chaque cellule en texte (dates -> JJ/MM/AAAA, nombres entiers
        sans '.0' final) pour rester cohérent avec la lecture CSV.
        """
        engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
        df = pd.read_excel(path, dtype=object, engine=engine)
        return df.apply(lambda col: col.map(self._excel_cell_to_str))

    def _excel_cell_to_str(self, value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ''
        if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
            return value.strftime('%d/%m/%Y')
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def _show_column_selection(self):
        # Vider ancien contenu
        for w in self.inner.winfo_children():
            w.destroy()
        self.col_vars.clear()
        self.key_vars.clear()
        self.key_checkbuttons.clear()

        # Ajouter une checkbox "Inclure" et une checkbox "Clé (WHERE)" par colonne
        for i, col in enumerate(self.df.columns):
            var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(self.inner, text=col, variable=var, command=self._update_preview_columns)
            cb.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            self.col_vars[col] = var

            key_var = tk.BooleanVar(value=False)
            key_cb = ttk.Checkbutton(self.inner, text="Clé (WHERE)", variable=key_var)
            key_cb.grid(row=i, column=1, sticky="w", padx=5, pady=2)
            self.key_vars[col] = key_var
            self.key_checkbuttons[col] = key_cb

        self._update_key_columns_visibility()

        # Mettre à jour scrollregion
        self.inner.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def _on_mode_change(self):
        self._update_key_columns_visibility()

    def _update_key_columns_visibility(self):
        show_keys = self.query_mode.get() == "update"
        for cb in self.key_checkbuttons.values():
            if show_keys:
                cb.grid()
            else:
                cb.grid_remove()

    def select_all(self):
        for var in self.col_vars.values():
            var.set(True)
        self._update_preview_columns()

    def deselect_all(self):
        for var in self.col_vars.values():
            var.set(False)
        self._update_preview_columns()

    def _show_preview(self):
        cols = list(self.df.columns)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100, anchor="w", stretch=False)
        for _, row in self.df.head(10).iterrows():
            self.tree.insert("", "end", values=[row[c] for c in cols])
        self._update_preview_columns()

    def _update_preview_columns(self):
        selected = [c for c in self.df.columns if self.col_vars[c].get()]
        self.tree["displaycolumns"] = selected if selected else []

    def _on_mousewheel(self, event):
        # Windows / MacOS
        delta = int(-1 * (event.delta / 120))
        self.canvas.yview_scroll(delta, "units")

    def _format_value(self, value):
        """
        Si value correspond à un format JJ/MM/AAAA (avec séparateur '/'),
        retourne AAAAMMJJ. Si value est un nombre décimal avec ',' comme
        séparateur, retourne le nombre avec '.'. Sinon retourne value inchangé.
        """
        # Expression régulière pour détecter JJ/MM/AAAA
        m = re.match(r'^(\d{2})/(\d{2})/(\d{4})$', value)
        if m:
            jj, mm, aaaa = m.group(1), m.group(2), m.group(3)
            return aaaa + mm + jj
        # Expression régulière pour détecter un nombre décimal avec ','
        m = re.match(r'^-?\d+,\d+$', value)
        if m:
            return value.replace(',', '.')
        # Retourner tel quel si pas date ni nombre décimal
        return value

    def _quoted_value(self, row, col):
        raw = row[col] or ""
        formatted = self._format_value(raw)
        escaped = formatted.replace("'", "''")
        return "'" + escaped + "'"

    def generate_sql(self):
        if self.df is None:
            return

        mode = self.query_mode.get()
        selected = [c for c, v in self.col_vars.items() if v.get()]
        if not selected:
            messagebox.showerror("Erreur", "Veuillez sélectionner au moins une colonne.")
            return

        if mode == "update":
            keys = [c for c in selected if self.key_vars[c].get()]
            set_cols = [c for c in selected if c not in keys]
            if not keys:
                messagebox.showerror("Erreur", "Veuillez sélectionner au moins une colonne clé (WHERE).")
                return
            if not set_cols:
                messagebox.showerror("Erreur", "Veuillez sélectionner au moins une colonne à mettre à jour (hors clé).")
                return
            default_name = "update_{}.sql".format(self.table_name)
        else:
            default_name = "insert_{}.sql".format(self.table_name)

        sql_path = filedialog.asksaveasfilename(
            title="Enregistrer le script SQL sous…",
            defaultextension=".sql",
            initialfile=default_name,
            filetypes=[("Fichier SQL", "*.sql"), ("Tous fichiers", "*.*")]
        )
        if not sql_path:
            return

        queries = []
        if mode == "update":
            for _, row in self.df.iterrows():
                set_parts = ["{0} = {1}".format(c, self._quoted_value(row, c)) for c in set_cols]
                where_parts = ["{0} = {1}".format(c, self._quoted_value(row, c)) for c in keys]
                q = "UPDATE {0} SET {1} WHERE {2};".format(
                    self.table_name, ", ".join(set_parts), " AND ".join(where_parts)
                )
                queries.append(q)
        else:
            for _, row in self.df.iterrows():
                fields = ", ".join(selected)
                values = ", ".join(self._quoted_value(row, c) for c in selected)
                q = "INSERT INTO {0} ({1}) VALUES ({2});".format(self.table_name, fields, values)
                queries.append(q)

        try:
            with open(sql_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(queries))
            messagebox.showinfo("Succès", "Le script SQL a été généré :\n{}".format(sql_path))
        except Exception as e:
            messagebox.showerror("Erreur", "Impossible d’écrire le SQL :\n{}".format(e))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
