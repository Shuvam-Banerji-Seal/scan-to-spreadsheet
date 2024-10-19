import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import cv2
from pyzbar.pyzbar import decode
import csv
from datetime import datetime
import os
import pandas as pd
import threading
from PIL import Image, ImageTk

class EnhancedQRScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced QR Scanner")
        self.root.geometry("800x600")
        
        # Theme and styling
        style = ttk.Style()
        style.theme_use('clam')
        
        # Variables
        self.spreadsheet_path = None
        self.is_scanning = False
        self.current_camera = 0
        self.available_cameras = self.get_available_cameras()
        self.last_scan = None
        self.duplicate_check = tk.BooleanVar(value=True)
        self.auto_save = tk.BooleanVar(value=True)
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main container
        self.main_container = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel for camera feed and controls
        self.left_panel = ttk.Frame(self.main_container)
        self.main_container.add(self.left_panel, weight=2)
        
        # Camera feed
        self.camera_frame = ttk.LabelFrame(self.left_panel, text="Camera Feed")
        self.camera_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.video_label = ttk.Label(self.camera_frame)
        self.video_label.pack(fill=tk.BOTH, expand=True)
        
        # Camera controls
        self.controls_frame = ttk.Frame(self.left_panel)
        self.controls_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.camera_combo = ttk.Combobox(self.controls_frame, 
                                       values=[f"Camera {i}" for i in self.available_cameras])
        self.camera_combo.set("Camera 0")
        self.camera_combo.pack(side=tk.LEFT, padx=5)
        
        self.scan_button = ttk.Button(self.controls_frame, 
                                    text="Start Scanning", 
                                    command=self.toggle_scanning)
        self.scan_button.pack(side=tk.LEFT, padx=5)
        
        # Right panel for settings and data
        self.right_panel = ttk.Frame(self.main_container)
        self.main_container.add(self.right_panel, weight=1)
        
        # Settings section
        self.settings_frame = ttk.LabelFrame(self.right_panel, text="Settings")
        self.settings_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(self.settings_frame, 
                       text="Check for duplicates", 
                       variable=self.duplicate_check).pack(anchor=tk.W, padx=5, pady=2)
        
        ttk.Checkbutton(self.settings_frame, 
                       text="Auto-save", 
                       variable=self.auto_save).pack(anchor=tk.W, padx=5, pady=2)
        
        # Spreadsheet section
        self.spreadsheet_frame = ttk.LabelFrame(self.right_panel, text="Spreadsheet")
        self.spreadsheet_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.link_button = ttk.Button(self.spreadsheet_frame, 
                                    text="Link Spreadsheet", 
                                    command=self.link_spreadsheet)
        self.link_button.pack(fill=tk.X, padx=5, pady=5)
        
        self.path_label = ttk.Label(self.spreadsheet_frame, 
                                  text="No spreadsheet linked", 
                                  wraplength=200)
        self.path_label.pack(fill=tk.X, padx=5)
        
        # Scan history
        self.history_frame = ttk.LabelFrame(self.right_panel, text="Scan History")
        self.history_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.history_tree = ttk.Treeview(self.history_frame, 
                                       columns=("Time", "Data"),
                                       show="headings")
        self.history_tree.heading("Time", text="Time")
        self.history_tree.heading("Data", text="Data")
        self.history_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Status bar
        self.status_bar = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=2)
        
    def get_available_cameras(self):
        """Check for available cameras"""
        available = []
        for i in range(5):  # Check first 5 possible camera indices
            cap = cv2.VideoCapture(i)
            if cap.isOpened():
                available.append(i)
                cap.release()
        return available if available else [0]
    
    def toggle_scanning(self):
        if not self.is_scanning:
            if not self.spreadsheet_path and self.auto_save.get():
                messagebox.showerror("Error", "Please link a spreadsheet first")
                return
            
            self.is_scanning = True
            self.scan_button.configure(text="Stop Scanning")
            self.camera_combo.configure(state='disabled')
            self.current_camera = int(self.camera_combo.get().split()[-1])
            threading.Thread(target=self.scan_qr, daemon=True).start()
        else:
            self.is_scanning = False
            self.scan_button.configure(text="Start Scanning")
            self.camera_combo.configure(state='normal')
    
    def scan_qr(self):
        cap = cv2.VideoCapture(self.current_camera)
        if not cap.isOpened():
            messagebox.showerror("Error", "Failed to access the camera")
            self.is_scanning = False
            return
        
        while self.is_scanning:
            ret, frame = cap.read()
            if not ret:
                continue
            
            # Convert frame for display
            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
            img = Image.fromarray(cv2image)
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.imgtk = imgtk
            self.video_label.configure(image=imgtk)
            
            # Scan for QR codes
            decoded_objects = decode(frame)
            for obj in decoded_objects:
                qr_data = obj.data.decode("utf-8")
                if self.process_scan(qr_data):
                    # Draw rectangle around detected QR code
                    points = obj.polygon
                    if len(points) > 4:
                        hull = cv2.convexHull(
                            numpy.array([point for point in points], dtype=numpy.float32))
                        points = hull
                    
                    cv2.polylines(frame, [numpy.array(points, dtype=numpy.int32)],
                                True, (0, 255, 0), 3)
            
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
        
        cap.release()
        self.video_label.configure(image='')
        
    def process_scan(self, data):
        """Process scanned QR code data"""
        current_time = datetime.now()
        
        # Check for duplicate if enabled
        if self.duplicate_check.get():
            if data == self.last_scan:
                self.update_status(f"Duplicate scan ignored: {data}")
                return False
        
        self.last_scan = data
        
        # Add to history
        self.history_tree.insert('', 0, values=(current_time.strftime('%H:%M:%S'), data))
        
        # Save to spreadsheet if auto-save is enabled
        if self.auto_save.get():
            self.add_to_spreadsheet(data)
        
        self.update_status(f"Scanned: {data}")
        return True
    
    def link_spreadsheet(self):
        """Link a spreadsheet file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Select or Create Spreadsheet"
        )
        
        if file_path:
            self.spreadsheet_path = file_path
            if not os.path.exists(file_path):
                # Create new spreadsheet with headers
                with open(file_path, 'w', newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(['Timestamp', 'Data'])
            
            self.path_label.configure(text=os.path.basename(file_path))
            self.update_status(f"Linked spreadsheet: {os.path.basename(file_path)}")
    
    def add_to_spreadsheet(self, data):
        """Add scanned data to spreadsheet"""
        if not self.spreadsheet_path:
            return False
        
        try:
            with open(self.spreadsheet_path, 'a', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([datetime.now().strftime('%Y-%m-%d %H:%M:%S'), data])
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save to spreadsheet: {str(e)}")
            return False
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_bar.configure(text=message)

def main():
    root = tk.Tk()
    app = EnhancedQRScannerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
