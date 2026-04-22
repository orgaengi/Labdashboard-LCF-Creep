"""
Tool 1 — LCF & Creep Lab Dashboard
====================================
Tracks LCF (capacity=50/yr) and Creep (capacity=22/yr) labs.
Run: python tool1_lcf_creep.py
"""

import os, sys, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl

# Add current directory to path so lab_core is found
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lab_core as core

# ── Tool-specific config ──────────────────────────────────
TOOL_TITLE   = "LCF & Creep Lab Dashboard"
ALLOWED_TYPES = ["LCF", "Creep"]
OUTPUT_FILE   = "LCF_Creep_Dashboard.xlsx"

DEFAULT_CAPS = {
    "LCF":   50,   # 50 samples/year max
    "Creep": 22,   # 22 samples/year max
}

# ═══════════════════════════════════════════════════════════
#  DASHBOARD GENERATOR
# ═══════════════════════════════════════════════════════════

def generate(input_path, capacities, theme, progress_cb=None):
    def p(msg):
        if progress_cb: progress_cb(msg)

    p("📖 Reading file…")
    df, col_map, errors, warns = core.load_and_filter(input_path, ALLOWED_TYPES)
    if errors:
        raise ValueError("\n".join(errors))

    p("⚙️  Computing weekly demand…")
    weekly_df, util_df, types, years = core.build_weekly(df, col_map, capacities)

    p("📗 Building workbook…")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    p("📄 Sheet 1/6 — Weekly Planner…")
    core.write_weekly_planner(wb, weekly_df, types, TOOL_TITLE)

    p("🎨 Sheet 2/6 — Utilization…")
    core.write_utilization_sheet(wb, util_df, types, theme)

    p("📋 Sheet 3/6 — Summary…")
    core.write_summary_sheet(wb, weekly_df, util_df, types, capacities, years,
                              TOOL_TITLE,
                              cap_note=f"LCF max={capacities['LCF']}/yr | "
                                       f"Creep max={capacities['Creep']}/yr")

    p("📈 Sheet 4/6 — Utilization Chart…")
    core.write_utilization_chart(wb, util_df, types, years)

    p("📊 Sheet 5/6 — Capacity vs Demand…")
    core.write_capacity_chart(wb, weekly_df, types, capacities, years)

    p("📊 Sheet 6/6 — Year-on-Year…")
    core.write_yoy_chart(wb, weekly_df, types, years)

    # Gantt for current year (if data available)
    current_year = core.CURRENT_YEAR
    current_week = core.CURRENT_WEEK
    if current_year in years:
        p(f"🗓️  Gantt — {current_year}…")
        core.write_gantt_current_year(wb, weekly_df, types, capacities,
                                       current_year, current_week)
    else:
        # Use last year in data as "current"
        last_yr = max(years)
        p(f"🗓️  Gantt — {last_yr} (latest available)…")
        core.write_gantt_current_year(wb, weekly_df, types, capacities,
                                       last_yr, 52)

    p("💾 Saving…")
    out = core.save_workbook(wb, input_path, OUTPUT_FILE)
    return out, warns


# ═══════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"🧪  {TOOL_TITLE}")
        self.geometry("700x620")
        self.resizable(False, False)
        self.configure(bg="#F0F4F8")
        self.file_path    = tk.StringVar(value="No file selected")
        self.caps         = {k: tk.IntVar(value=v) for k, v in DEFAULT_CAPS.items()}
        self.theme_var    = tk.StringVar(value=list(core.COLOR_THEMES.keys())[0])
        self.progress_var = tk.StringVar(value="")
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg="#1A4E8A", height=70)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🧪  LCF & Creep Lab Dashboard",
                 bg="#1A4E8A", fg="white",
                 font=("Arial", 14, "bold")).pack(pady=10)
        tk.Label(hdr, text="LCF capacity: 50 samples/yr  |  Creep capacity: 22 samples/yr",
                 bg="#1A4E8A", fg="#BDD7EE",
                 font=("Arial", 9)).pack()

        main = tk.Frame(self, bg="#F0F4F8", padx=20, pady=10)
        main.pack(fill="both", expand=True)

        # File picker
        self._sec(main, "📂 Step 1 — Select Input Excel File")
        ff = tk.Frame(main, bg="#F0F4F8"); ff.pack(fill="x", pady=(0,8))
        self._btn(ff, "📁 Browse…", self._pick).pack(side="left")
        tk.Label(ff, textvariable=self.file_path, bg="#F0F4F8", fg="#444",
                 font=("Arial", 9), wraplength=500).pack(side="left", padx=10)

        ttk.Separator(main).pack(fill="x", pady=6)

        # Capacities
        self._sec(main, "⚙️ Step 2 — Lab Capacities (samples/year)")
        cf = tk.Frame(main, bg="#F0F4F8"); cf.pack(fill="x", pady=(0,8))
        for i, (name, var) in enumerate(self.caps.items()):
            card = tk.Frame(cf, bg="#FFFFFF", bd=1, relief="groove", padx=14, pady=10)
            card.grid(row=0, column=i, padx=10, sticky="ew")
            cf.columnconfigure(i, weight=1)
            tk.Label(card, text=name, bg="#FFFFFF",
                     font=("Arial", 10, "bold"), fg="#1A4E8A").pack(anchor="w")
            tk.Label(card, text="Max samples / year", bg="#FFFFFF",
                     font=("Arial", 8), fg="#666").pack(anchor="w")
            tk.Spinbox(card, from_=1, to=9999, textvariable=var,
                       width=10, font=("Arial", 11),
                       relief="flat", bd=1).pack(anchor="w", pady=(4,0))

        ttk.Separator(main).pack(fill="x", pady=6)

        # Theme
        self._sec(main, "🎨 Step 3 — Color Theme")
        tf = tk.Frame(main, bg="#F0F4F8"); tf.pack(fill="x", pady=(0,8))
        for i, name in enumerate(core.COLOR_THEMES):
            tk.Radiobutton(tf, text=name, variable=self.theme_var,
                           value=name, bg="#F0F4F8",
                           font=("Arial", 9)).grid(row=0, column=i, padx=8, sticky="w")

        ttk.Separator(main).pack(fill="x", pady=6)

        # Progress
        self.prog_label = tk.Label(main, textvariable=self.progress_var,
                                   bg="#F0F4F8", fg="#1A4E8A",
                                   font=("Arial", 9, "italic"))
        self.prog_label.pack(anchor="w")
        self.prog_bar = ttk.Progressbar(main, mode="indeterminate", length=640)
        self.prog_bar.pack(fill="x", pady=4)

        self.gen_btn = self._btn(main, "⚡ Generate LCF & Creep Dashboard",
                                 self._run, color="#1A4E8A", size=11, pady=13)
        self.gen_btn.pack(pady=10, fill="x")

        # Footer
        ftr = tk.Frame(self, bg="#DDE5EF", height=26)
        ftr.pack(fill="x", side="bottom")
        tk.Label(ftr,
                 text=f"Output → Lab_Dashboard_Output/{OUTPUT_FILE}",
                 bg="#DDE5EF", fg="#666", font=("Arial", 8)).pack(pady=5)

    def _sec(self, p, title):
        tk.Label(p, text=title, bg="#F0F4F8",
                 font=("Arial", 10, "bold"), fg="#1A4E8A").pack(anchor="w", pady=(6,2))

    def _btn(self, p, text, cmd, color="#1A4E8A", size=9, pady=8):
        return tk.Button(p, text=text, command=cmd,
                         bg=color, fg="white", activebackground=color,
                         font=("Arial", size, "bold"), relief="flat",
                         cursor="hand2", padx=16, pady=pady)

    def _pick(self):
        path = filedialog.askopenfilename(
            title="Select Lab Data Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.file_path.set(path)
            self.progress_var.set("")

    def _run(self):
        path = self.file_path.get()
        if path == "No file selected" or not os.path.isfile(path):
            messagebox.showerror("No File", "Please select a valid Excel file.")
            return
        caps  = {k: v.get() for k, v in self.caps.items()}
        theme = core.COLOR_THEMES[self.theme_var.get()]
        self.gen_btn.config(state="disabled", text="⏳ Generating…")
        self.prog_bar.start(10)

        def task():
            try:
                out, warns = generate(path, caps, theme,
                    progress_cb=lambda m: self.after(0, lambda msg=m: self.progress_var.set(msg)))
                self.after(0, lambda: self._ok(out, warns))
            except Exception as e:
                self.after(0, lambda err=str(e): self._err(err))

        threading.Thread(target=task, daemon=True).start()

    def _ok(self, out, warns):
        self.prog_bar.stop()
        self.gen_btn.config(state="normal", text="⚡ Generate LCF & Creep Dashboard")
        self.progress_var.set("✅ Done!")
        msg = f"✅ Dashboard created!\n\n📁 {out}\n\n"
        msg += "Sheets: Weekly Planner, Utilization, Summary,\n"
        msg += "Utilization Chart, Capacity Chart, YoY Chart, Gantt"
        if warns: msg += "\n\n⚠️ " + "\n".join(warns)
        messagebox.showinfo("Success", msg)

    def _err(self, e):
        self.prog_bar.stop()
        self.gen_btn.config(state="normal", text="⚡ Generate LCF & Creep Dashboard")
        self.progress_var.set("❌ Error")
        messagebox.showerror("Error", e)


if __name__ == "__main__":
    App().mainloop()
