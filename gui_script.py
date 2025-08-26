#!/usr/bin/env python3
# electrochemistry_automation_gui.py

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import json
import os
from datetime import datetime
from pathlib import Path
import subprocess
import threading
import time
import sys
import serial.tools.list_ports

# Import the local tecancavro library
try:
    from tecancavro import XCaliburD, TecanAPISerial
    PUMP_AVAILABLE = True
except ImportError:
    PUMP_AVAILABLE = False
    print("Warning: tecancavro library not found. Pump features disabled.")

class ElectrochemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Electrochemistry Automation System")
        self.root.geometry("1400x800")
        
        # Set working directory to script location
        self.script_dir = Path(__file__).parent.absolute()
        os.chdir(self.script_dir)
        print(f"Working directory: {os.getcwd()}")
        
        # Setup file structure
        self.base_path = Path("methods")
        self.base_path.mkdir(exist_ok=True)
        
        # Queue management
        self.measurement_queue = []
        self.is_running = False
        self.current_script = ""
        
        # Pump controller
        self.pump = None
        self.pump_com = None
        
        # Create GUI elements
        self.setup_gui()
        
    def setup_gui(self):
        # Main container with tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Tab 1: Method Creation
        self.method_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.method_frame, text="Method Creation")
        self.setup_method_tab()
        
        # Tab 2: Queue Management
        self.queue_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.queue_frame, text="Queue")
        self.setup_queue_tab()
        
        # Tab 3: Pump Control (if available)
        if PUMP_AVAILABLE:
            self.pump_frame = ttk.Frame(self.notebook)
            self.notebook.add(self.pump_frame, text="Pump Control")
            self.setup_pump_tab()
        
        # Tab 4: Script Preview
        self.script_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.script_frame, text="Script Preview")
        self.setup_script_tab()
        
    def setup_method_tab(self):
        # Left side - Technique selection
        left_frame = ttk.Frame(self.method_frame)
        left_frame.pack(side='left', fill='both', expand=True, padx=5)
        
        ttk.Label(left_frame, text="Select Technique:", 
                 font=('Arial', 12, 'bold')).pack(pady=5)
        
        technique_frame = ttk.Frame(left_frame)
        technique_frame.pack(pady=10)
        
        ttk.Button(technique_frame, text="Cyclic Voltammetry (CV)", 
                  command=self.show_cv_params, width=25).pack(pady=5)
        ttk.Button(technique_frame, text="Square Wave Voltammetry (SWV)", 
                  command=self.show_swv_params, width=25).pack(pady=5)
        
        # Device status
        self.device_status = ttk.Label(left_frame, text="", foreground="blue")
        self.device_status.pack(pady=10)
        
        ttk.Button(left_frame, text="Check Device Connection", 
                  command=self.check_device).pack(pady=5)
        
        # Right side - Parameters
        self.params_frame = ttk.LabelFrame(self.method_frame, 
                                          text="Parameters", padding=10)
        self.params_frame.pack(side='right', fill='both', expand=True, padx=5)
        
        # Initialize with CV parameters
        self.show_cv_params()
        
    def check_device(self):
        """Check for connected serial devices"""
        ports = list(serial.tools.list_ports.comports())
        if ports:
            device_list = "\n".join([f"{p.device}: {p.description}" for p in ports])
            self.device_status.config(text="Devices found (check console)", foreground="green")
            print("Available serial devices:")
            print(device_list)
        else:
            self.device_status.config(text="No devices found", foreground="red")
            
    def show_cv_params(self):
        self.clear_params_frame()
        self.current_technique = "CV"
        
        # CV specific parameters
        self.cv_params = {}
        
        params = [
            ("Begin Potential (V):", "begin_potential", "0"),
            ("Vertex 1 (V):", "vertex1", "-0.5"),
            ("Vertex 2 (V):", "vertex2", "0.5"),
            ("Step Potential (V):", "step_potential", "0.01"),
            ("Scan Rate (V/s):", "scan_rate", "0.1"),
            ("Number of Scans:", "n_scans", "1"),
            ("Conditioning Potential (V):", "cond_potential", "0"),
            ("Conditioning Time (s):", "cond_time", "0"),
        ]
        
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self.params_frame, text=label).grid(row=i, column=0, 
                                                          sticky='w', pady=2)
            entry = ttk.Entry(self.params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.cv_params[key] = entry
        
        # Buttons
        button_frame = ttk.Frame(self.params_frame)
        button_frame.grid(row=len(params), column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Generate Script", 
                  command=self.generate_cv_script).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Add to Queue", 
                  command=self.add_cv_to_queue).pack(side='left', padx=5)
        
    def show_swv_params(self):
        self.clear_params_frame()
        self.current_technique = "SWV"
        
        # SWV specific parameters
        self.swv_params = {}
        
        params = [
            ("Begin Potential (V):", "begin_potential", "-0.5"),
            ("End Potential (V):", "end_potential", "0.5"),
            ("Step Potential (V):", "step_potential", "0.01"),
            ("Amplitude (V):", "amplitude", "0.02"),
            ("Frequency (Hz):", "frequency", "15"),
            ("Conditioning Potential (V):", "cond_potential", "0"),
            ("Conditioning Time (s):", "cond_time", "0"),
        ]
        
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self.params_frame, text=label).grid(row=i, column=0, 
                                                          sticky='w', pady=2)
            entry = ttk.Entry(self.params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.swv_params[key] = entry
        
        # Buttons
        button_frame = ttk.Frame(self.params_frame)
        button_frame.grid(row=len(params), column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Generate Script", 
                  command=self.generate_swv_script).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Add to Queue", 
                  command=self.add_swv_to_queue).pack(side='left', padx=5)
        
    def setup_pump_tab(self):
        # Pump connection frame
        conn_frame = ttk.LabelFrame(self.pump_frame, text="Pump Connection", padding=10)
        conn_frame.pack(fill='x', padx=10, pady=10)
        
        # Serial port selection
        port_frame = ttk.Frame(conn_frame)
        port_frame.pack(fill='x', pady=5)
        
        ttk.Label(port_frame, text="Serial Port:").pack(side='left', padx=5)
        self.pump_port = ttk.Combobox(port_frame, width=25)
        self.pump_port.pack(side='left', padx=5)
        
        ttk.Button(port_frame, text="Refresh Ports", 
                  command=self.refresh_pump_ports).pack(side='left', padx=5)
        ttk.Button(port_frame, text="Connect", 
                  command=self.connect_pump).pack(side='left', padx=5)
        ttk.Button(port_frame, text="Initialize", 
                  command=self.init_pump).pack(side='left', padx=5)
        
        # Pump settings
        settings_frame = ttk.Frame(conn_frame)
        settings_frame.pack(fill='x', pady=5)
        
        ttk.Label(settings_frame, text="Syringe Size (µL):").pack(side='left', padx=5)
        self.syringe_size = ttk.Entry(settings_frame, width=10)
        self.syringe_size.insert(0, "1000")
        self.syringe_size.pack(side='left', padx=5)
        
        ttk.Label(settings_frame, text="Num Ports:").pack(side='left', padx=5)
        self.num_ports = ttk.Entry(settings_frame, width=10)
        self.num_ports.insert(0, "9")
        self.num_ports.pack(side='left', padx=5)
        
        self.pump_status = ttk.Label(conn_frame, text="Not connected", foreground="red")
        self.pump_status.pack(pady=5)
        
        # Manual pump controls
        control_frame = ttk.LabelFrame(self.pump_frame, text="Manual Control", padding=10)
        control_frame.pack(fill='x', padx=10, pady=10)
        
        # Port and volume controls
        params_frame = ttk.Frame(control_frame)
        params_frame.pack(fill='x', pady=5)
        
        ttk.Label(params_frame, text="Port:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.pump_port_select = ttk.Spinbox(params_frame, from_=1, to=9, width=10)
        self.pump_port_select.set(1)
        self.pump_port_select.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(params_frame, text="Volume (µL):").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.pump_volume = ttk.Entry(params_frame, width=10)
        self.pump_volume.insert(0, "100")
        self.pump_volume.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(params_frame, text="Speed Code (0-40):").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.pump_speed = ttk.Spinbox(params_frame, from_=0, to=40, width=10)
        self.pump_speed.set(10)
        self.pump_speed.grid(row=2, column=1, padx=5, pady=5)
        
        # Manual control buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Extract", 
                  command=self.pump_extract_manual).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Dispense", 
                  command=self.pump_dispense_manual).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Empty to Waste", 
                  command=self.pump_empty).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Change Port", 
                  command=self.pump_change_port).pack(side='left', padx=5)
        
        # Add pump action to queue
        queue_frame = ttk.LabelFrame(self.pump_frame, text="Add to Queue", padding=10)
        queue_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(queue_frame, text="Action:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.pump_action = ttk.Combobox(queue_frame, 
                                       values=["Extract", "Dispense", "Change Port", "Wait", "Empty to Waste"],
                                       width=15)
        self.pump_action.set("Dispense")
        self.pump_action.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(queue_frame, text="Port:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.queue_port = ttk.Spinbox(queue_frame, from_=1, to=9, width=10)
        self.queue_port.set(1)
        self.queue_port.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(queue_frame, text="Volume/Time:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.queue_param = ttk.Entry(queue_frame, width=10)
        self.queue_param.insert(0, "100")
        self.queue_param.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(queue_frame, text="(µL for Extract/Dispense, seconds for Wait)").grid(
            row=2, column=2, padx=5, pady=5)
        
        ttk.Button(queue_frame, text="Add to Queue", 
                  command=self.add_pump_to_queue).grid(row=3, column=0, columnspan=2, pady=10)
        
        # Refresh ports on startup
        self.refresh_pump_ports()
    
    def setup_queue_tab(self):
        # Queue control buttons
        control_frame = ttk.Frame(self.queue_frame)
        control_frame.pack(pady=10)
        
        ttk.Button(control_frame, text="Run Queue", 
                  command=self.run_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Stop", 
                  command=self.stop_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Clear Queue", 
                  command=self.clear_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Save Queue", 
                  command=self.save_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Load Queue", 
                  command=self.load_queue).pack(side='left', padx=5)
        
        # Queue display
        self.queue_tree = ttk.Treeview(self.queue_frame, 
                                       columns=('Type', 'Status', 'Details'), 
                                       show='tree headings', height=15)
        self.queue_tree.heading('#0', text='#')
        self.queue_tree.heading('Type', text='Type')
        self.queue_tree.heading('Status', text='Status')
        self.queue_tree.heading('Details', text='Details')
        
        self.queue_tree.column('#0', width=50)
        self.queue_tree.column('Type', width=150)
        self.queue_tree.column('Status', width=100)
        self.queue_tree.column('Details', width=400)
        
        self.queue_tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Status bar
        self.status_label = ttk.Label(self.queue_frame, 
                                     text="Status: Ready", 
                                     relief='sunken')
        self.status_label.pack(fill='x', padx=10, pady=5)
    
    def setup_script_tab(self):
        # Script preview text area
        text_frame = ttk.Frame(self.script_frame)
        text_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.script_text = tk.Text(text_frame, wrap='none', font=('Courier', 11))
        self.script_text.pack(fill='both', expand=True)
        
        # Buttons
        button_frame = ttk.Frame(self.script_frame)
        button_frame.pack(pady=5)
        
        ttk.Button(button_frame, text="Save Script", 
                  command=self.save_script).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Copy to Clipboard", 
                  command=self.copy_script).pack(side='left', padx=5)
    
    def clear_params_frame(self):
        for widget in self.params_frame.winfo_children():
            widget.destroy()
    
    def generate_cv_script(self):
        """Generate MethodSCRIPT for CV measurement"""
        try:
            script = self.create_cv_methodscript()
            self.current_script = script
            self.update_script_preview(script)
            self.notebook.select(self.notebook.index(self.script_frame))
            return script
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate script: {str(e)}")
            return None
    
    def create_cv_methodscript(self):
        """Create MethodSCRIPT for CV"""
        # Get parameters
        begin = self.cv_params['begin_potential'].get()
        v1 = self.cv_params['vertex1'].get()
        v2 = self.cv_params['vertex2'].get()
        step = self.cv_params['step_potential'].get()
        scan_rate = self.cv_params['scan_rate'].get()
        n_scans = self.cv_params['n_scans'].get()
        cond_pot = self.cv_params['cond_potential'].get()
        cond_time = self.cv_params['cond_time'].get()
        
        # Build script
        script = "e\n"
        script += "var c\n"
        script += "var p\n"
        script += "set_pgstat_mode 2\n"
        script += "set_max_bandwidth 40\n"
        script += "set_range ba 100u\n"
        script += "set_autoranging ba 1n 100u\n"
        
        if float(cond_time) > 0:
            script += f"set_e {cond_pot}\n"
            script += "cell_on\n"
            script += f"#autorange for {cond_time}s prior to CV\n"
            script += f"meas_loop_ca p c {cond_pot} 100m {cond_time}\n"
            script += "endloop\n"
        else:
            script += "set_e 0m\n"
            script += "cell_on\n"
        
        script += f"# E_res, I_res, E_begin, E_vtx1, E_vtx2, E_step, scan_rate\n"
        script += f"meas_loop_cv p c {begin} {v1} {v2} {step} {scan_rate}"
        
        if int(n_scans) > 1:
            script += f" nscans({n_scans})"
        
        script += "\n"
        script += "\tpck_start\n"
        script += "\tpck_add p\n"
        script += "\tpck_add c\n"
        script += "\tpck_end\n"
        script += "endloop\n"
        script += "on_finished:\n"
        script += "cell_off\n\n"
        
        return script

    def generate_swv_script(self):
        """Generate MethodSCRIPT for SWV measurement"""
        try:
            script = self.create_swv_methodscript()
            self.current_script = script
            self.update_script_preview(script)
            self.notebook.select(self.notebook.index(self.script_frame))
            return script
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate script: {str(e)}")
            return None
    
    def create_swv_methodscript(self):
        """Create MethodSCRIPT for SWV"""
        # Get parameters
        begin = self.swv_params['begin_potential'].get()
        end = self.swv_params['end_potential'].get()
        step = self.swv_params['step_potential'].get()
        amplitude = self.swv_params['amplitude'].get()
        frequency = self.swv_params['frequency'].get()
        cond_pot = self.swv_params['cond_potential'].get()
        cond_time = self.swv_params['cond_time'].get()
        
        # Build script
        script = "e\n"
        script += "var c\n"
        script += "var p\n"
        script += "var f\n"
        script += "var r\n"
        script += "var i\n"
        script += "store_var i 0i ja\n"
        script += "set_pgstat_chan 0\n"
        script += "set_pgstat_mode 2\n"
        script += "set_max_bandwidth 40\n"
        
        # Calculate potential range
        min_pot = min(float(begin), float(end)) - float(amplitude)
        max_pot = max(float(begin), float(end)) + float(amplitude)
        min_pot_mv = int(min_pot * 1000)
        max_pot_mv = int(max_pot * 1000)
        
        script += f"set_range_minmax da {min_pot_mv}m {max_pot_mv}m\n"
        script += "set_range ba 100u\n"
        script += "set_autoranging ba 1n 100u\n"
        script += "cell_on\n"
        
        if float(cond_time) > 0:
            script += f"#Equilibrate at {begin} and autorange for {cond_time}s prior to SWV\n"
            script += f"meas_loop_ca p c {begin} 100m {cond_time}\n"
            script += "endloop\n"
        
        script += f"# Measure SWV: E, I, I_fwd, I_rev, E_begin, E_end, E_step, E_amp, freq\n"
        script += f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}\n"
        script += "\tpck_start\n"
        script += "\tpck_add p\n"
        script += "\tpck_add c\n"
        script += "\tpck_add f\n"
        script += "\tpck_add r\n"
        script += "\tpck_end\n"
        script += "endloop\n"
        script += "on_finished:\n"
        script += "cell_off\n\n"
        
        return script
    
    def update_script_preview(self, script):
        """Update the script preview tab"""
        self.script_text.delete(1.0, tk.END)
        self.script_text.insert(1.0, script)
    
    def copy_script(self):
        """Copy script to clipboard"""
        script = self.script_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(script)
        messagebox.showinfo("Success", "Script copied to clipboard")
    
    def add_cv_to_queue(self):
        """Add CV measurement to queue"""
        script = self.generate_cv_script()
        if script:
            self.add_to_queue("CV", script)
    
    def add_swv_to_queue(self):
        """Add SWV measurement to queue"""
        script = self.generate_swv_script()
        if script:
            self.add_to_queue("SWV", script)
    
    def add_to_queue(self, technique, script):
        """Add measurement to queue and save script"""
        # Create date folder
        date_folder = self.base_path / datetime.now().strftime('%Y-%m-%d')
        date_folder.mkdir(exist_ok=True)
        
        # Generate filename
        existing_files = list(date_folder.glob('*.ms'))
        file_number = len(existing_files) + 1
        filename = f"{file_number:03d}_{technique.lower()}.ms"
        filepath = date_folder / filename
        
        # Save script
        with open(filepath, 'w') as f:
            f.write(script)
        
        # Add to queue
        queue_item = {
            'type': technique,
            'script_path': str(filepath),
            'status': 'pending',
            'details': filename
        }
        
        self.measurement_queue.append(queue_item)
        self.refresh_queue_display()
        messagebox.showinfo("Success", f"{technique} added to queue\nSaved as: {filename}")
    
    def refresh_queue_display(self):
        """Refresh the queue tree view"""
        # Clear existing items
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        
        # Add queue items
        for i, item in enumerate(self.measurement_queue):
            self.queue_tree.insert('', 'end', text=str(i+1),
                                  values=(item['type'], 
                                         item['status'].upper(),
                                         item.get('details', '')))
    
    def run_queue(self):
        """Run all measurements in queue"""
        if not self.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items in queue")
            return
        
        if self.is_running:
            messagebox.showwarning("Already Running", "Queue is already running")
            return
        
        # Start queue execution in separate thread
        self.is_running = True
        self.queue_thread = threading.Thread(target=self.execute_queue_serial)
        self.queue_thread.daemon = True
        self.queue_thread.start()
    
    def execute_queue_serial(self):
        """Execute queue items one by one"""
        run_measurement_path = self.script_dir / "run_measurement_serial.py"
        
        # Process each item in queue
        while self.measurement_queue and self.is_running:
            current = self.measurement_queue[0]
            current['status'] = 'running'
            
            self.root.after(0, self.refresh_queue_display)
            self.root.after(0, self.update_status, f"Running: {current['type']}")
            
            try:
                # Check if pump action or measurement
                if current['type'].startswith('PUMP_'):
                    self.execute_pump_action(current)
                else:
                    # Execute measurement
                    script_path = Path(current['script_path'])
                    if not script_path.is_absolute():
                        script_path = self.script_dir / script_path
                    
                    print(f"\n{'='*60}")
                    print(f"Running measurement: {script_path}")
                    print(f"{'='*60}\n")
                    
                    cmd = [sys.executable, str(run_measurement_path), str(script_path)]
                    process = subprocess.Popen(
                        cmd,
                        stdout=subprocess.PIPE, 
                        stderr=subprocess.STDOUT,
                        text=True,
                        bufsize=1,
                        universal_newlines=True,
                        cwd=str(self.script_dir)
                    )
                    
                    while True:
                        line = process.stdout.readline()
                        if not line:
                            break
                        print(line.rstrip())
                    
                    return_code = process.wait()
                    
                    if return_code == 0:
                        current['status'] = 'completed'
                    else:
                        current['status'] = 'failed'
                
            except Exception as e:
                current['status'] = 'failed'
                print(f"Error: {e}")
            
            # Remove completed item
            self.measurement_queue.pop(0)
            self.root.after(0, self.refresh_queue_display)
            
            time.sleep(1)
        
        self.is_running = False
        self.root.after(0, self.update_status, "Queue Complete")
    
    def stop_queue(self):
        """Stop queue execution"""
        self.is_running = False
        self.update_status("Queue Stopped")
    
    def clear_queue(self):
        """Clear all pending measurements"""
        if self.is_running:
            messagebox.showwarning("Queue Running", "Cannot clear queue while running")
            return
        
        self.measurement_queue = []
        self.refresh_queue_display()
        self.update_status("Queue Cleared")
    
    def save_queue(self):
        """Save queue to JSON file"""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        
        if filepath:
            with open(filepath, 'w') as f:
                json.dump(self.measurement_queue, f, indent=2)
            messagebox.showinfo("Success", "Queue saved successfully")
    
    def load_queue(self):
        """Load queue from JSON file"""
        filepath = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")]
        )
        
        if filepath:
            with open(filepath, 'r') as f:
                self.measurement_queue = json.load(f)
            self.refresh_queue_display()
            messagebox.showinfo("Success", "Queue loaded successfully")
    
    def save_script(self):
        """Save current script to file"""
        if not self.current_script:
            messagebox.showwarning("No Script", "Generate a script first")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".ms",
            filetypes=[("MethodSCRIPT", "*.ms")]
        )
        
        if filepath:
            with open(filepath, 'w') as f:
                f.write(self.current_script)
            messagebox.showinfo("Success", "Script saved")
    
    def update_status(self, message):
        """Update status bar"""
        self.status_label.config(text=f"Status: {message}")
    
    # Pump control methods (if available)
    def refresh_pump_ports(self):
        """Refresh available serial ports"""
        if not PUMP_AVAILABLE:
            return
            
        ports = []
        for port in serial.tools.list_ports.comports():
            port_desc = f"{port.device} - {port.description}"
            ports.append(port_desc)
        
        self.pump_port['values'] = ports
        if ports:
            self.pump_port.set(ports[0])
    
    def connect_pump(self):
        """Connect to pump via serial"""
        if not PUMP_AVAILABLE:
            return
            
        try:
            port_str = self.pump_port.get()
            if not port_str:
                messagebox.showerror("Error", "Please select a serial port")
                return
            
            # Extract just the port name
            port = port_str.split(' - ')[0]
            
            # Get settings
            syringe_ul = int(self.syringe_size.get())
            num_ports = int(self.num_ports.get())
            
            # Connect to pump
            self.pump_com = TecanAPISerial(0, port, 9600)
            self.pump = XCaliburD(self.pump_com, 
                                num_ports=num_ports,
                                syringe_ul=syringe_ul)
            
            self.pump_status.config(text=f"Connected to {port}", foreground="green")
            messagebox.showinfo("Success", "Pump connected successfully")
            
        except Exception as e:
            self.pump_status.config(text="Connection failed", foreground="red")
            messagebox.showerror("Error", f"Failed to connect: {str(e)}")
    
    def init_pump(self):
        """Initialize pump"""
        if not self.pump:
            messagebox.showerror("Error", "Pump not connected")
            return
        
        try:
            self.pump.init()
            messagebox.showinfo("Success", "Pump initialized")
        except Exception as e:
            messagebox.showerror("Error", f"Initialization failed: {str(e)}")
    
    def pump_extract_manual(self):
        """Manual extract"""
        if not self.pump:
            messagebox.showerror("Error", "Pump not connected")
            return
        
        try:
            port = int(self.pump_port_select.get())
            volume = int(self.pump_volume.get())
            speed = int(self.pump_speed.get())
            
            self.pump.setSpeed(speed)
            self.pump.extract(port, volume)
            messagebox.showinfo("Success", f"Extracted {volume} µL from port {port}")
        except Exception as e:
            messagebox.showerror("Error", f"Extract failed: {str(e)}")
    
    def pump_dispense_manual(self):
        """Manual dispense"""
        if not self.pump:
            messagebox.showerror("Error", "Pump not connected")
            return
        
        try:
            port = int(self.pump_port_select.get())
            volume = int(self.pump_volume.get())
            speed = int(self.pump_speed.get())
            
            self.pump.setSpeed(speed)
            self.pump.dispense(port, volume)
            messagebox.showinfo("Success", f"Dispensed {volume} µL to port {port}")
        except Exception as e:
            messagebox.showerror("Error", f"Dispense failed: {str(e)}")
    
    def pump_empty(self):
        """Empty to waste"""
        if not self.pump:
            messagebox.showerror("Error", "Pump not connected")
            return
        
        try:
            self.pump.dispenseToWaste()
            messagebox.showinfo("Success", "Emptied to waste")
        except Exception as e:
            messagebox.showerror("Error", f"Empty failed: {str(e)}")
    
    def pump_change_port(self):
        """Change port"""
        if not self.pump:
            messagebox.showerror("Error", "Pump not connected")
            return
        
        try:
            port = int(self.pump_port_select.get())
            self.pump.changePort(port)
            messagebox.showinfo("Success", f"Changed to port {port}")
        except Exception as e:
            messagebox.showerror("Error", f"Port change failed: {str(e)}")
    
    def add_pump_to_queue(self):
        """Add pump action to queue"""
        if not PUMP_AVAILABLE:
            messagebox.showerror("Error", "Pump library not available")
            return
            
        action = self.pump_action.get()
        port = int(self.queue_port.get())
        param = float(self.queue_param.get())
        
        if action == "Wait":
            details = f"Wait {param} seconds"
            queue_type = "PUMP_WAIT"
        elif action == "Empty to Waste":
            details = f"Empty syringe to waste"
            queue_type = "PUMP_EMPTY"
            port = 9
        else:
            details = f"{action} {param} µL, Port {port}"
            queue_type = f"PUMP_{action.upper()}"
        
        queue_item = {
            'type': queue_type,
            'status': 'pending',
            'details': details,
            'action': action.lower().replace(' ', '_'),
            'port': port,
            'parameter': param
        }
        
        self.measurement_queue.append(queue_item)
        self.refresh_queue_display()
        messagebox.showinfo("Success", f"Added to queue: {details}")
    
    def execute_pump_action(self, action_item):
        """Execute pump action"""
        if not self.pump:
            print("WARNING: Pump not connected, skipping pump action")
            action_item['status'] = 'skipped'
            return
        
        try:
            action = action_item['action']
            port = action_item.get('port', 1)
            param = action_item.get('parameter', 0)
            
            print(f"Executing pump: {action} (port={port}, param={param})")
            
            if action == 'extract':
                self.pump.extract(port, int(param))
                print(f"Extracted {param} µL from port {port}")
            elif action == 'dispense':
                self.pump.dispense(port, int(param))
                print(f"Dispensed {param} µL to port {port}")
            elif action == 'change_port':
                self.pump.changePort(port)
                print(f"Changed to port {port}")
            elif action == 'empty_to_waste':
                self.pump.dispenseToWaste()
                print("Emptied to waste")
            elif action == 'wait':
                print(f"Waiting {param} seconds...")
                time.sleep(param)
            
            action_item['status'] = 'completed'
            
        except Exception as e:
            print(f"Pump action failed: {e}")
            action_item['status'] = 'failed'

def main():
    root = tk.Tk()
    app = ElectrochemGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()