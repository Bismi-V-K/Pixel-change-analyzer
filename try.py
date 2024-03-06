import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

class PixelChangeAnalyzer:
    def __init__(self):
        self.video_path = ""
        self.changes = []  

    def search_video(self):
        self.video_path = filedialog.askopenfilename(filetypes=[("Video Files", "*.mp4;*.avi")])
        self.video_entry.delete(0, tk.END)
        self.video_entry.insert(0, self.video_path)

    def analyze_video(self):
        if self.video_path == "":
            messagebox.showerror("Error", "Please select a video file.")
            return

        cap = cv2.VideoCapture(self.video_path)
        if not cap.isOpened():
            messagebox.showerror("Error", "Error opening video file.")
            return

        self.changes = []

        prev_frame = None

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            if prev_frame is not None:
                current_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                diff = cv2.absdiff(current_frame, prev_frame)

                threshold = 21 

                total_change = np.sum(diff > threshold)

                self.changes.append(total_change)

            prev_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        cap.release()
        messagebox.showinfo("Analysis Complete", "Pixel change analysis completed.")

    def export_to_excel(self):
        if len(self.changes) == 0:
            messagebox.showerror("Error", "No analysis data available.")
            return

        result_df = pd.DataFrame({"Frame": range(1, len(self.changes) + 1), "Pixel Change": self.changes})

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

        try:
            result_df.to_excel(save_path, index=False)
            messagebox.showinfo("Export Complete", "Data exported to Excel successfully.")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while exporting to Excel:\n\n{str(e)}")

    def create_gui(self):
        window = tk.Tk()
        window.title("Pixel Change Analyzer")

        video_frame = tk.Frame(window)
        video_frame.pack(pady=10)

        tk.Label(video_frame, text="Video File:").pack(side=tk.LEFT)
        self.video_entry = tk.Entry(video_frame, width=50)
        self.video_entry.pack(side=tk.LEFT)
        tk.Button(video_frame, text="Upload", command=self.search_video).pack(side=tk.LEFT)

        analyze_button = tk.Button(window, text="Analyze", command=self.analyze_video)
        analyze_button.pack(pady=10)

        export_button = tk.Button(window, text="Export to Excel", command=self.export_to_excel)
        export_button.pack(pady=10)

        window.mainloop()

analyzer = PixelChangeAnalyzer()
analyzer.create_gui()
