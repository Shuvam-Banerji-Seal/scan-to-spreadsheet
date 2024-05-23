import tkinter as tk
from tkinter import messagebox, filedialog
import cv2
from pyzbar.pyzbar import decode
import csv

class QRScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Scanner")
        
        self.scan_button = tk.Button(root, text="Scan", command=self.scan_qr)
        self.scan_button.pack()
        
        self.link_button = tk.Button(root, text="Link Spreadsheet", command=self.link_spreadsheet)
        self.link_button.pack()
        
        self.spreadsheet_path = None

    def scan_qr(self):
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            messagebox.showerror("Error", "Failed to access the camera.")
            return

        while True:
            ret, frame = cap.read()
            if not ret:
                messagebox.showerror("Error", "Failed to capture frame.")
                break

            cv2.imshow("QR Scanner", frame)

            decoded_objects = decode(frame)
            if decoded_objects:
                for obj in decoded_objects:
                    qr_data = obj.data.decode("utf-8")
                    success = self.add_to_spreadsheet(qr_data)
                    if success:
                        messagebox.showinfo("Success", "Data successfully added to spreadsheet.")
                    else:
                        messagebox.showerror("Error", "Failed to add data to spreadsheet.")
                break

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()

    def link_spreadsheet(self):
        self.spreadsheet_path = filedialog.askopenfilename(title="Select a spreadsheet file", filetypes=[("CSV files", "*.csv")])
        if not self.spreadsheet_path:
            messagebox.showinfo("Info", "No spreadsheet linked.")
        else:
            messagebox.showinfo("Success", f"Spreadsheet linked: {self.spreadsheet_path}")

    def add_to_spreadsheet(self, data):
        if not self.spreadsheet_path:
            messagebox.showerror("Error", "No spreadsheet linked. Please link a spreadsheet first.")
            return False
        try:
            with open(self.spreadsheet_path, "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([data])
            return True
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while adding data to the spreadsheet: {str(e)}")
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = QRScannerApp(root)
    root.mainloop()
