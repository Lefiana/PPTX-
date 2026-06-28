import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os

import phase0_inspector
import phase1_ingestor
import phase2_crop_pilot
import phase3_assembler

class GraduationCommandCenter:
    def __init__(self, root):
        self.root = root
        self.root.title("Graduation Automation Suite")
        self.root.geometry("1200x800")
        
        self.sidebar = ttk.Frame(self.root, width=200, padding=10, relief="ridge")
        self.sidebar.pack(side="left", fill="y")
        
        self.main_content = ttk.Frame(self.root, padding=20)
        self.main_content.pack(side="right", fill="both", expand=True)
        
        self.frames = {}
        self.build_sidebar()
        
        self.build_phase0_view()
        self.build_phase1_view()
        self.build_phase2_view()
        self.build_phase3_view()
        
        self.show_frame("Phase 0")

    def build_sidebar(self):
        ttk.Label(self.sidebar, text="Command Center", font=("Arial", 14, "bold")).pack(pady=(0, 20))
        
        ttk.Button(self.sidebar, text="Phase 0: Inspector", command=lambda: self.show_frame("Phase 0")).pack(fill="x", pady=5)
        ttk.Button(self.sidebar, text="Phase 1: Ingest & Sort", command=lambda: self.show_frame("Phase 1")).pack(fill="x", pady=5)
        ttk.Button(self.sidebar, text="Phase 2: Crop Pilot", command=lambda: self.show_frame("Phase 2")).pack(fill="x", pady=5)
        ttk.Button(self.sidebar, text="Phase 3: Assemble", command=lambda: self.show_frame("Phase 3")).pack(fill="x", pady=5)

    def show_frame(self, frame_name):
        for frame in self.frames.values():
            frame.pack_forget()
        self.frames[frame_name].pack(fill="both", expand=True)

    def create_log_box(self, parent):
        log_text = tk.Text(parent, height=15, state="disabled", bg="black", fg="lime")
        log_text.pack(fill="both", expand=True, pady=10)
        return log_text

    def gui_log(self, log_widget, msg):
        log_widget.config(state="normal")
        log_widget.insert(tk.END, str(msg) + "\n")
        log_widget.see(tk.END)
        log_widget.config(state="disabled")
        self.root.update_idletasks()

    # ================= PHASE 0 =================
    def build_phase0_view(self):
        frame = ttk.Frame(self.main_content)
        self.frames["Phase 0"] = frame
        ttk.Label(frame, text="Phase 0: Layout Inspector", font=("Arial", 18, "bold")).pack(anchor="w", pady=(0, 20))
        
        self.p0_pptx_var = tk.StringVar()
        self.p0_excel_var = tk.StringVar()
        
        self.make_file_row(frame, "Select PPTX Template", self.p0_pptx_var, filetypes=[("PowerPoint", "*.pptx")])
        self.make_file_row(frame, "Select Excel Data", self.p0_excel_var, filetypes=[("Excel", "*.xlsx")])
        
        self.p0_log = self.create_log_box(frame)
        ttk.Button(frame, text="Run Inspector", command=self.run_p0).pack(anchor="w")

    def run_p0(self):
        self.gui_log(self.p0_log, "Starting Inspection...")
        phase0_inspector.inspect_environment(self.p0_pptx_var.get(), self.p0_excel_var.get())
        self.gui_log(self.p0_log, "Inspection Complete. config.json updated.")

    # ================= PHASE 1 =================
    def build_phase1_view(self):
        frame = ttk.Frame(self.main_content)
        self.frames["Phase 1"] = frame
        ttk.Label(frame, text="Phase 1: Folder Sorter", font=("Arial", 18, "bold")).pack(anchor="w", pady=(0, 20))
        
        self.p1_src_var = tk.StringVar()
        self.p1_dest_var = tk.StringVar()
        self.p1_preset_var = tk.StringVar(value="COLLEGE")
        
        self.make_dir_row(frame, "Select Raw Images Folder", self.p1_src_var)
        self.make_dir_row(frame, "Select Filtered Destination", self.p1_dest_var)
        
        pf = ttk.Frame(frame)
        pf.pack(fill="x", pady=5)
        ttk.Label(pf, text="Preset:").pack(side="left")
        ttk.Combobox(pf, textvariable=self.p1_preset_var, values=["COLLEGE", "SHS"], state="readonly").pack(side="left", padx=10)
        
        self.p1_log = self.create_log_box(frame)
        ttk.Button(frame, text="Run Ingestion", command=self.run_p1).pack(anchor="w")

    def run_p1(self):
        phase1_ingestor.run_phase1_ingestor(
            self.p1_src_var.get(), self.p1_dest_var.get(), self.p1_preset_var.get(), 
            log_callback=lambda msg: self.gui_log(self.p1_log, msg)
        )

    # ================= PHASE 2 =================
    def build_phase2_view(self):
        frame = ttk.Frame(self.main_content)
        self.frames["Phase 2"] = frame
        self.crop_pilot = phase2_crop_pilot.CropPilotApp(frame)

    # ================= PHASE 3 =================
    def build_phase3_view(self):
        frame = ttk.Frame(self.main_content)
        self.frames["Phase 3"] = frame
        ttk.Label(frame, text="Phase 3: Assembly Engine", font=("Arial", 18, "bold")).pack(anchor="w", pady=(0, 20))
        
        self.p3_dir_var = tk.StringVar()
        self.p3_pptx_var = tk.StringVar()
        self.p3_excel_var = tk.StringVar()
        
        self.make_dir_row(frame, "Select Filtered Folder (Phase 1 Output)", self.p3_dir_var)
        self.make_file_row(frame, "Select PPTX Template", self.p3_pptx_var, filetypes=[("PowerPoint", "*.pptx")])
        self.make_file_row(frame, "Select Excel Data", self.p3_excel_var, filetypes=[("Excel", "*.xlsx")])
        
        self.p3_log = self.create_log_box(frame)
        ttk.Button(frame, text="Assemble Final PPTX", command=self.run_p3).pack(anchor="w")

    def run_p3(self):
        out_path = os.path.join(os.path.dirname(self.p3_pptx_var.get()), "FINAL_GRADUATION_OUTPUT.pptx")
        phase3_assembler.run_phase3_assembler(
            self.p3_dir_var.get(), self.p3_pptx_var.get(), self.p3_excel_var.get(), out_path,
            log_callback=lambda msg: self.gui_log(self.p3_log, msg)
        )

    # --- UI Helpers ---
    def make_file_row(self, parent, label, var, filetypes):
        r = ttk.Frame(parent)
        r.pack(fill="x", pady=5)
        ttk.Button(r, text=label, command=lambda: var.set(filedialog.askopenfilename(filetypes=filetypes))).pack(side="left")
        ttk.Label(r, textvariable=var, foreground="gray").pack(side="left", padx=10)
        
    def make_dir_row(self, parent, label, var):
        r = ttk.Frame(parent)
        r.pack(fill="x", pady=5)
        ttk.Button(r, text=label, command=lambda: var.set(filedialog.askdirectory())).pack(side="left")
        ttk.Label(r, textvariable=var, foreground="gray").pack(side="left", padx=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = GraduationCommandCenter(root)
    root.mainloop()