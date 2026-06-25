import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("xml2xlsx — Convertisseur XML vers Excel")
        self.resizable(False, False)
        self.configure(padx=20, pady=20)

        # ── Mode de sélection ──────────────────────────────────────────────
        mode_frame = ttk.LabelFrame(self, text="Source d'entrée", padding=10)
        mode_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        self.mode = tk.StringVar(value="dossier")
        ttk.Radiobutton(mode_frame, text="Dossier", variable=self.mode,
                        value="dossier", command=self._update_mode).grid(row=0, column=0, padx=8)
        ttk.Radiobutton(mode_frame, text="Fichier XML", variable=self.mode,
                        value="xml", command=self._update_mode).grid(row=0, column=1, padx=8)

        # ── Chemin XML ────────────────────────────────────────────────────
        path_frame = ttk.LabelFrame(self, text="Chemin", padding=10)
        path_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        path_frame.columnconfigure(0, weight=1)

        self.path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.path_var, width=55).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(path_frame, text="Parcourir…", command=self._browse_input).grid(row=0, column=1)

        # ── DTD (optionnel) ───────────────────────────────────────────────
        dtd_frame = ttk.LabelFrame(self, text="Schéma DTD (optionnel)", padding=10)
        dtd_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        dtd_frame.columnconfigure(0, weight=1)

        self.dtd_var = tk.StringVar()
        ttk.Entry(dtd_frame, textvariable=self.dtd_var, width=55).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(dtd_frame, text="Parcourir…", command=self._browse_dtd).grid(row=0, column=1)

        # ── Options ───────────────────────────────────────────────────────
        opt_frame = ttk.LabelFrame(self, text="Options", padding=10)
        opt_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        self.save_individual = tk.BooleanVar(value=True)
        self.cb_individual = ttk.Checkbutton(
            opt_frame,
            text="Générer un fichier XLSX individuel par fichier XML",
            variable=self.save_individual
        )
        self.cb_individual.grid(row=0, column=0, sticky="w")

        self.save_master = tk.BooleanVar(value=True)
        self.cb_master = ttk.Checkbutton(
            opt_frame,
            text="Générer un fichier XLSX maître (fusion de tous les fichiers)",
            variable=self.save_master
        )
        self.cb_master.grid(row=1, column=0, sticky="w")

        self.compute_position = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Calculer la colonne POSITION_INTRODD",
            variable=self.compute_position
        ).grid(row=2, column=0, sticky="w")

        # ── Lancer ────────────────────────────────────────────────────────
        self.run_btn = ttk.Button(self, text="▶  Lancer la conversion", command=self._run)
        self.run_btn.grid(row=4, column=0, columnspan=3, pady=(0, 10))

        # ── Barre de progression ──────────────────────────────────────────
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=400)
        self.progress.grid(row=5, column=0, columnspan=3, pady=(0, 8))

        # ── Journal ───────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self, text="Journal", padding=6)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="ew")

        self.log = tk.Text(log_frame, height=10, width=60, state="disabled",
                           font=("Courier", 9), bg="#1e1e1e", fg="#d4d4d4",
                           insertbackground="white")
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        self._update_mode()

    # ── Helpers ───────────────────────────────────────────────────────────

    def _update_mode(self):
        folder_only = self.mode.get() == "dossier"
        state = "normal" if folder_only else "disabled"
        self.cb_individual.configure(state=state)
        self.cb_master.configure(state=state)

    def _browse_input(self):
        if self.mode.get() == "dossier":
            path = filedialog.askdirectory(title="Sélectionner un dossier")
        else:
            path = filedialog.askopenfilename(
                title="Sélectionner un fichier XML",
                filetypes=[("Fichiers XML", "*.xml"), ("Tous les fichiers", "*.*")]
            )
        if path:
            self.path_var.set(path)

    def _browse_dtd(self):
        path = filedialog.askopenfilename(
            title="Sélectionner un fichier DTD",
            filetypes=[("Fichiers DTD", "*.dtd"), ("Tous les fichiers", "*.*")]
        )
        if path:
            self.dtd_var.set(path)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _run(self):
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("Chemin manquant", "Veuillez sélectionner un fichier ou un dossier.")
            return
        self.run_btn.configure(state="disabled")
        self.progress.start(10)
        self._log(f"▶ Démarrage… ({path})")
        threading.Thread(target=self._process, args=(path,), daemon=True).start()

    def _process(self, path):
        try:
            from xml2xlsx.xml2xlsx import (
                parse_dtd, extract_paragraphs, reorder_columns,
                PREFAB_CHILDREN, PREFAB_ATTRIBS, NO_LOWER
            )
            from xml2xlsx.format_excel import format_excel
            import pandas as pd
            import numpy as np

            dtd_path = self.dtd_var.get().strip()
            compute_position = self.compute_position.get()
            mode = self.mode.get()

            if dtd_path and os.path.isfile(dtd_path):
                children_map, attribs_map = parse_dtd(dtd_path)
                self._log(f"[schema] DTD chargé : {dtd_path}")
            else:
                children_map, attribs_map = PREFAB_CHILDREN, PREFAB_ATTRIBS
                self._log("[schema] Schéma PREFAB intégré")

            p_children = children_map.get('p', [])

            def process_rows(rows, out_path, xml_path = None,lowercased=True):
                df = pd.DataFrame(rows)
                df = df.replace('', np.nan)
                df.dropna(axis=1, how='all', inplace=True)
                if lowercased:
                    text_cols = [c for c in df.columns
                                 if c.endswith('_text') and c not in NO_LOWER]
                    df[text_cols] = df[text_cols].apply(
                        lambda col: col.map(lambda v: v.lower() if isinstance(v, str) else v))
                df = reorder_columns(df, p_children)
                from xml2xlsx.integrity_check import check_counts
                check_counts(df, xml_path, dtd_path)
                format_excel(df, out_path, p_children)
                self._log(f"    ✔ {os.path.basename(out_path)}")

            # ── Fichier XML unique ─────────────────────────────────────────
            if mode == "xml":
                rows = extract_paragraphs(path, children_map, attribs_map, compute_position)
                for row in rows:
                    row['source_file'] = os.path.basename(path)
                out = path.replace(".xml", ".xlsx")
                process_rows(rows, out, lowercased=False,xml_path=path)

            # ── Dossier ───────────────────────────────────────────────────
            elif mode == "dossier":
                xml_files = sorted(f for f in os.listdir(path) if f.endswith('.xml'))
                if not xml_files:
                    self._log("✘ Aucun fichier XML trouvé.")
                    return

                all_rows = []
                for f in xml_files:
                    file_path = os.path.join(path, f)
                    self._log(f"  → {f}")
                    rows = extract_paragraphs(file_path, children_map, attribs_map, compute_position)
                    if rows and self.save_individual.get():
                        ind_rows = [dict(r, source_file=f) for r in rows]
                        process_rows(ind_rows, file_path.replace(".xml", ".xlsx"),xml_path=file_path)
                    for row in rows:
                        row['source_file'] = f
                    all_rows.extend(rows)
           
                if self.save_master.get() and all_rows:
                    master_out = os.path.join(path, "master_output.xlsx")
                    process_rows(all_rows, master_out)
                    self._log(f"✔ Maître : {master_out}")

            self._log("✔ Conversion terminée.")

        except Exception as e:
            import traceback
            self._log(f"✘ Erreur : {e}")
            self._log(traceback.format_exc())

        finally:
            self.progress.stop()
            self.run_btn.configure(state="normal")


def main():
    App().mainloop()


if __name__ == '__main__':
    main()
