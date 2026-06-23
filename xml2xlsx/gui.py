import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys


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
        ttk.Radiobutton(mode_frame, text="Fichier XLSX (recalcul positions)", variable=self.mode,
                        value="xlsx", command=self._update_mode).grid(row=0, column=2, padx=8)

        # ── Chemin ────────────────────────────────────────────────────────
        path_frame = ttk.LabelFrame(self, text="Chemin", padding=10)
        path_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        path_frame.columnconfigure(0, weight=1)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var, width=55)
        self.path_entry.grid(row=0, column=0, padx=(0, 8))
        self.browse_btn = ttk.Button(path_frame, text="Parcourir…", command=self._browse)
        self.browse_btn.grid(row=0, column=1)

        # ── Options ───────────────────────────────────────────────────────
        opt_frame = ttk.LabelFrame(self, text="Options", padding=10)
        opt_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        self.ignore_multi_ppi = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Ignorer les paragraphes avec plusieurs PPI (ignore_multi_ppi)",
            variable=self.ignore_multi_ppi
        ).grid(row=0, column=0, sticky="w")

        self.save_individual = tk.BooleanVar(value=True)
        self.cb_individual = ttk.Checkbutton(
            opt_frame,
            text="Générer un fichier XLSX individuel par fichier XML",
            variable=self.save_individual
        )
        self.cb_individual.grid(row=1, column=0, sticky="w")

        self.save_master = tk.BooleanVar(value=True)
        self.cb_master = ttk.Checkbutton(
            opt_frame,
            text="Générer un fichier XLSX maître (fusion de tous les fichiers)",
            variable=self.save_master
        )
        self.cb_master.grid(row=2, column=0, sticky="w")

        # ── Lancer ────────────────────────────────────────────────────────
        self.run_btn = ttk.Button(self, text="▶  Lancer la conversion", command=self._run)
        self.run_btn.grid(row=3, column=0, columnspan=3, pady=(0, 10))

        # ── Barre de progression ──────────────────────────────────────────
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=400)
        self.progress.grid(row=4, column=0, columnspan=3, pady=(0, 8))

        # ── Journal ───────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self, text="Journal", padding=6)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="ew")

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
        mode = self.mode.get()
        # Activer/désactiver les checkboxes selon le mode
        folder_only = mode == "dossier"
        state = "normal" if folder_only else "disabled"
        self.cb_individual.configure(state=state)
        self.cb_master.configure(state=state)

    def _browse(self):
        mode = self.mode.get()
        if mode == "dossier":
            path = filedialog.askdirectory(title="Sélectionner un dossier")
        elif mode == "xml":
            path = filedialog.askopenfilename(
                title="Sélectionner un fichier XML",
                filetypes=[("Fichiers XML", "*.xml"), ("Tous les fichiers", "*.*")]
            )
        else:
            path = filedialog.askopenfilename(
                title="Sélectionner un fichier XLSX",
                filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
            )
        if path:
            self.path_var.set(path)

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

        thread = threading.Thread(target=self._process, args=(path,), daemon=True)
        thread.start()

    def _process(self, path):
        try:
            from xml2xlsx.xml2xlsx import (
                extract_paragraphs, get_introdd_position_dd,
                NO_LOWER
            )
            from xml2xlsx.format_excel import format_ppi_bold
            import pandas as pd
            import numpy as np

            ignore_multi = self.ignore_multi_ppi.get()
            mode = self.mode.get()

            self._log(f"[config] ignore_multi_ppi = {ignore_multi}")

            # ── Fichier XML unique ─────────────────────────────────────────
            if mode == "xml":
                rows = extract_paragraphs(path)
                for row in rows:
                    row['source_file'] = os.path.basename(path)
                df = pd.DataFrame(rows)
                df = df.replace('', np.nan)
                df.dropna(axis=1, how='all', inplace=True)
                if 'paragraph_text_dd' in df.columns:
                    pos = df.columns.get_loc('paragraph_text_dd') + 1
                    df.insert(pos, 'POSITION_INTRODD_DD',
                              df['paragraph_text_dd'].map(
                                  lambda t: get_introdd_position_dd(t, ignore_multi),
                                  na_action='ignore'))
                out = path.replace(".xml", ".xlsx")
                format_ppi_bold(df, out)
                self._log(f"✔ Fichier sauvegardé : {out}")

            # ── Fichier XLSX (recalcul) ───────────────────────────────────
            elif mode == "xlsx":
                df = pd.read_excel(path)
                if 'paragraph_text_dd' not in df.columns:
                    self._log("✘ Erreur : colonne 'paragraph_text_dd' introuvable.")
                    return
                if 'POSITION_INTRODD_DD' in df.columns:
                    df['POSITION_INTRODD_DD'] = df['paragraph_text_dd'].map(
                        lambda t: get_introdd_position_dd(t, ignore_multi), na_action='ignore')
                else:
                    pos = df.columns.get_loc('paragraph_text_dd') + 1
                    df.insert(pos, 'POSITION_INTRODD_DD', df['paragraph_text_dd'].map(
                        lambda t: get_introdd_position_dd(t, ignore_multi), na_action='ignore'))
                format_ppi_bold(df, path)
                self._log(f"✔ Fichier mis à jour : {path}")

            # ── Dossier ───────────────────────────────────────────────────
            elif mode == "dossier":
                xml_files = sorted([f for f in os.listdir(path) if f.endswith('.xml')])
                if not xml_files:
                    self._log("✘ Aucun fichier XML trouvé dans le dossier.")
                    return

                all_rows = []
                for f in xml_files:
                    file_path = os.path.join(path, f)
                    self._log(f"  → Traitement : {f}")
                    rows = extract_paragraphs(file_path)
                    if rows:
                        individual_df = pd.DataFrame(rows)
                        individual_df = individual_df.replace('', np.nan)
                        individual_df.dropna(axis=1, how='all', inplace=True)
                        text_cols = [c for c in individual_df.columns
                                     if c.endswith('_text') and c not in NO_LOWER]
                        individual_df[text_cols] = individual_df[text_cols].apply(
                            lambda col: col.map(lambda v: v.lower() if isinstance(v, str) else v)
                        )
                        if 'paragraph_text_dd' in individual_df.columns:
                            pos = individual_df.columns.get_loc('paragraph_text_dd') + 1
                            individual_df.insert(pos, 'POSITION_INTRODD_DD',
                                                 individual_df['paragraph_text_dd'].map(
                                                     lambda t: get_introdd_position_dd(t, ignore_multi),
                                                     na_action='ignore'))
                        if self.save_individual.get():
                            individual_out = file_path.replace(".xml", ".xlsx")
                            format_ppi_bold(individual_df, individual_out)
                            self._log(f"    ✔ Individuel : {os.path.basename(individual_out)}")

                    for row in rows:
                        row['source_file'] = f
                    all_rows.extend(rows)

                if self.save_master.get() and all_rows:
                    df = pd.DataFrame(all_rows)
                    cols = ['source_file'] + [c for c in df.columns if c != 'source_file']
                    df = df[cols]
                    text_cols = [c for c in df.columns if c.endswith('_text') and c not in NO_LOWER]
                    df[text_cols] = df[text_cols].apply(
                        lambda col: col.map(lambda v: v.lower() if isinstance(v, str) else v)
                    )
                    df = df.replace('', np.nan)
                    df.dropna(axis=1, how='all', inplace=True)
                    if 'paragraph_text_dd' in df.columns:
                        pos = df.columns.get_loc('paragraph_text_dd') + 1
                        df.insert(pos, 'POSITION_INTRODD_DD',
                                  df['paragraph_text_dd'].map(
                                      lambda t: get_introdd_position_dd(t, ignore_multi),
                                      na_action='ignore'))
                    master_out = os.path.join(path, "master_output.xlsx")
                    format_ppi_bold(df, master_out)
                    self._log(f"✔ Fichier maître : {master_out}")

            self._log("✔ Conversion terminée.")

        except Exception as e:
            self._log(f"✘ Erreur : {e}")

        finally:
            self.progress.stop()
            self.run_btn.configure(state="normal")


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
