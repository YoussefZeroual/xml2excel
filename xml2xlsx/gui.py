import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("xml2xlsx — Convertisseur XML vers Excel")
        self.resizable(True, True)
        self.configure(padx=10, pady=10)

        # Configure main grid: left (controls), right (log)
        self.columnconfigure(0, weight=0)  # Controls column - fixed
        self.columnconfigure(1, weight=1)  # Log column - expandable
        self.rowconfigure(0, weight=1)     # Make rows expandable for log

        # ════════════════════════════════════════════════════════════════════
        # LEFT SIDE: CONTROLS
        # ════════════════════════════════════════════════════════════════════
        left_frame = ttk.Frame(self)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # ── Mode de sélection ──────────────────────────────────────────────
        mode_frame = ttk.LabelFrame(left_frame, text="Source d'entrée", padding=10)
        mode_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        self.mode = tk.StringVar(value="dossier")
        ttk.Radiobutton(mode_frame, text="Dossier", variable=self.mode,
                        value="dossier", command=self._update_mode).grid(row=0, column=0, padx=4)
        ttk.Radiobutton(mode_frame, text="Fichier XML", variable=self.mode,
                        value="xml", command=self._update_mode).grid(row=0, column=1, padx=4)
        ttk.Radiobutton(mode_frame, text="Réécrire XML", variable=self.mode,
                        value="reverse", command=self._update_mode).grid(row=0, column=2, padx=4)

        # ── Chemin XML ────────────────────────────────────────────────────
        path_frame = ttk.LabelFrame(left_frame, text="Chemins", padding=10)
        path_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        path_frame.columnconfigure(0, weight=1)

        self.path_var = tk.StringVar()
        self.path_label = ttk.Label(path_frame, text="")
        self.path_label.grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Entry(path_frame, textvariable=self.path_var, width=35).grid(row=1, column=0, sticky="ew", pady=(0, 4))
        ttk.Button(path_frame, text="Parcourir…", command=self._browse_input).grid(row=1, column=1, padx=(4, 0))

        # ── XML original (pour mode reverse) ───────────────────────────────
        self.original_xml_var = tk.StringVar()
        self.original_xml_label = ttk.Label(path_frame, text="XML original (reverse)")
        self.original_xml_label.grid(row=2, column=0, sticky="w", pady=(8, 4))
        self.original_xml_entry = ttk.Entry(path_frame, textvariable=self.original_xml_var, width=35)
        self.original_xml_entry.grid(row=3, column=0, sticky="ew")
        self.original_xml_button = ttk.Button(path_frame, text="Parcourir…", command=self._browse_original_xml)
        self.original_xml_button.grid(row=3, column=1, padx=(4, 0))

        # ── DTD (optionnel) ───────────────────────────────────────────────
        dtd_frame = ttk.LabelFrame(left_frame, text="Schéma DTD (optionnel)", padding=10)
        dtd_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        dtd_frame.columnconfigure(0, weight=1)

        self.dtd_var = tk.StringVar()
        ttk.Entry(dtd_frame, textvariable=self.dtd_var, width=35).grid(row=0, column=0, sticky="ew")
        ttk.Button(dtd_frame, text="Parcourir…", command=self._browse_dtd).grid(row=0, column=1, padx=(4, 0))

        # ── Options ───────────────────────────────────────────────────────
        opt_frame = ttk.LabelFrame(left_frame, text="Options", padding=10)
        opt_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        self.save_individual = tk.BooleanVar(value=True)
        self.cb_individual = ttk.Checkbutton(
            opt_frame,
            text="XLSX individuel par fichier",
            variable=self.save_individual
        )
        self.cb_individual.grid(row=0, column=0, sticky="w")

        self.save_master = tk.BooleanVar(value=True)
        self.cb_master = ttk.Checkbutton(
            opt_frame,
            text="Fichier XLSX maître",
            variable=self.save_master
        )
        self.cb_master.grid(row=1, column=0, sticky="w")

        self.compute_position = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame,
            text="Calculer POSITION_INTRODD",
            variable=self.compute_position
        ).grid(row=2, column=0, sticky="w")

        # ── Lancer ────────────────────────────────────────────────────────
        self.run_btn = ttk.Button(left_frame, text="▶  Lancer", command=self._run, width=20)
        self.run_btn.grid(row=4, column=0, pady=(0, 10), sticky="ew")

        # ── Barre de progression ──────────────────────────────────────────
        self.progress = ttk.Progressbar(left_frame, mode="indeterminate", length=200)
        self.progress.grid(row=5, column=0, sticky="ew")

        # ════════════════════════════════════════════════════════════════════
        # RIGHT SIDE: LOG
        # ════════════════════════════════════════════════════════════════════
        log_frame = ttk.LabelFrame(self, text="Journal", padding=6)
        log_frame.grid(row=0, column=1, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log = tk.Text(log_frame, height=25, width=100, state="disabled",
                           font=("Courier", 10), bg="#1e1e1e", fg="#d4d4d4",
                           insertbackground="white", wrap=tk.WORD)
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        self._update_mode()

    # ── Helpers ───────────────────────────────────────────────────────────

    def _update_mode(self):
        mode = self.mode.get()
        if mode == "dossier":
            self.path_label.configure(text="Dossier XML")
            self.cb_individual.configure(state="normal")
            self.cb_master.configure(state="normal")
            self.original_xml_var.set("")
            self.original_xml_entry.configure(state="disabled")
            self.original_xml_button.configure(state="disabled")
            self.original_xml_label.configure(foreground="gray")
        elif mode == "xml":
            self.path_label.configure(text="Fichier XML")
            self.cb_individual.configure(state="disabled")
            self.cb_master.configure(state="disabled")
            self.original_xml_var.set("")
            self.original_xml_entry.configure(state="disabled")
            self.original_xml_button.configure(state="disabled")
            self.original_xml_label.configure(foreground="gray")
        elif mode == "reverse":
            self.path_label.configure(text="Fichier Excel modifié")
            self.cb_individual.configure(state="disabled")
            self.cb_master.configure(state="disabled")
            self.path_var.set("")
            self.original_xml_var.set("")
            self.dtd_var.set("")
            self.original_xml_entry.configure(state="normal")
            self.original_xml_button.configure(state="normal")
            self.original_xml_label.configure(foreground="black")

    def _browse_input(self):
        if self.mode.get() == "dossier":
            path = filedialog.askdirectory(title="Sélectionner un dossier")
        elif self.mode.get() == "reverse":
            path = filedialog.askopenfilename(
                title="Sélectionner un fichier Excel modifié",
                filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
            )
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

    def _browse_original_xml(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier XML original",
            filetypes=[("Fichiers XML", "*.xml"), ("Tous les fichiers", "*.*")]
        )
        if path:
            self.original_xml_var.set(path)

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
        
        if self.mode.get() == "reverse":
            xml_path = self.original_xml_var.get().strip()
            if not xml_path:
                messagebox.showwarning("XML original manquant", "Veuillez sélectionner le fichier XML original.")
                return
        
        self.run_btn.configure(state="disabled")
        self.progress.start(10)
        self._log(f"▶ Démarrage… ({path})")
        threading.Thread(target=self._process, args=(path,), daemon=True).start()

    def _process(self, path):
        try:
            from xml2xlsx.xml2xlsx import (
                parse_dtd, infer_schema_from_xml, extract_paragraphs,
                reorder_columns, NO_LOWER, xlsx2xml
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
                if mode == "reverse":
                    probe = self.original_xml_var.get().strip() or None
                    probe = [probe] if probe else []
                elif os.path.isfile(path) and path.endswith('.xml'):
                    probe = [path]
                else:
                    probe = [os.path.join(path, f)
                             for f in sorted(os.listdir(path))
                             if f.endswith('.xml')]
                children_map, attribs_map = infer_schema_from_xml(probe)
                self._log(f"[schema] Schéma inféré depuis {len(probe)} fichier(s)")

            p_children = children_map.get('p', [])

            def process_rows(rows, out_path, xml_path=None, lowercased=True):
                df = pd.DataFrame(rows)
                df = df.replace('', np.nan)
                df.dropna(axis=1, how='all', inplace=True)
                if lowercased:
                    text_cols = [c for c in df.columns
                                 if c.endswith('_text') and c not in NO_LOWER]
                    df[text_cols] = df[text_cols].apply(
                        lambda col: col.map(lambda v: v.lower() if isinstance(v, str) else v))
                df = reorder_columns(df, p_children)
                
                # Integrity check - capture messages for GUI log
                from xml2xlsx.integrity_check import check_counts
                messages = check_counts(df, xml_path, children_map)
                for msg in messages:
                    self._log(f"    {msg}")
                
                format_excel(df, out_path, p_children)
                self._log(f"    ✔ {os.path.basename(out_path)}")

            # ── Mode reverse ──────────────────────────────────────────────
            if mode == "reverse":
                excel_path = path
                xml_path = self.original_xml_var.get().strip()
                if not xml_path:
                    self._log("✘ Fichier XML original requis.")
                    return
                output_path = xml_path.replace('.xml', '_updated.xml')
                self._log(f"▶ Application des modifications Excel à XML…")
                xlsx2xml(excel_path, xml_path, output_path, children_map)
                self._log(f"✔ Fichier XML mis à jour : {output_path}")
                
                # Integrity check for reverse mode
                self._log("\n▶ Vérification de l'intégrité…")
                from xml2xlsx.integrity_check_reverse import check_reverse_integrity
                result = check_reverse_integrity(xml_path, output_path, children_map)
                for msg in result['messages']:
                    self._log(f"  {msg}")
                
                if not result['valid']:
                    self._log("\n⚠️  ATTENTION: Des problèmes ont été détectés!")

            # ── Fichier XML unique ─────────────────────────────────────────
            elif mode == "xml":
                rows = extract_paragraphs(path, children_map, attribs_map, compute_position)
                for row in rows:
                    row['source_file'] = os.path.basename(path)
                out = path.replace(".xml", ".xlsx")
                process_rows(rows, out, lowercased=False, xml_path=path)

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
                        process_rows(ind_rows, file_path.replace(".xml", ".xlsx"), xml_path=file_path,lowercased=False)
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
