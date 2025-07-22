import fitz  # PyMuPDF
import tkinter as tk
from tkinter import Button, Canvas, Label, messagebox
from PIL import Image, ImageTk
import webbrowser, os


class PDFViewerWidget(tk.Frame):
    def __init__(self, parent, pdf_path, absolute=True, canvas_width=500, canvas_height=620):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.current_page = 0
        self.canvas_width = canvas_width
        self.canvas_height = canvas_height
        try:
            self.doc = fitz.open(pdf_path)
        except Exception as e:
            print(e)
            messagebox.showerror("Error",e)
            return

        # Create canvas
        self.canvas = Canvas(self, width=self.canvas_width, height=self.canvas_height)
        self.canvas.pack(side="top", fill="both", expand=True)

        # Page number label
        self.page_label = Label(self, text=f"Page {self.current_page + 1} / {len(self.doc)}", font=("Arial", 12))
        self.page_label.pack(side="top", pady=5)

        # Navigation buttons
        self.prev_button = Button(self, text="⏪ الصفحة السابقة", font=('Arial',11), command=self.prev_page)
        self.next_button = Button(self, text="⏩ الصفحة التالية", font=('Arial',11), command=self.next_page)
        open_path = self.pdf_path if absolute else f'{os.getcwd()}/{self.pdf_path[1:]}'
        self.open_button = Button(self, text="📤 فتح ", font=('Arial',11), command=lambda: webbrowser.open(open_path))

        self.prev_button.pack(side="left", padx=20, pady=10)
        self.next_button.pack(side="right", padx=20, pady=10)
        self.open_button.pack(side="bottom", padx=20, pady=10)

        # Load and display the first page
        self.display_pdf()

    def display_pdf(self):
        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Maintain aspect ratio while scaling image to fit canvas
        img_ratio = img.width / img.height
        canvas_ratio = self.canvas_width / self.canvas_height

        if img_ratio > canvas_ratio:
            new_width = self.canvas_width
            new_height = int(self.canvas_width / img_ratio)
        else:
            new_height = self.canvas_height
            new_width = int(self.canvas_height * img_ratio)

        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

        img_tk = ImageTk.PhotoImage(img)

        # Center the image within the canvas
        self.canvas.create_image(self.canvas_width//2, self.canvas_height//2, anchor="center", image=img_tk)
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

        self.canvas.image = img_tk  # Prevent garbage collection

        # Update page number label
        self.page_label.config(text=f"Page {self.current_page + 1} / {len(self.doc)}")

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.display_pdf()

    def next_page(self):
        if self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.display_pdf()

    def release_pdf(self):
        """Closes the PDF file to release any locks before removal."""
        if self.doc:
            self.doc.close()  # Close the document safely
            self.doc = None 

    def create_pdf_viewer(parent, pdf_path, absolute=True, width=500, height=620):
        return PDFViewerWidget(parent, pdf_path, absolute, width, height)

# Example usage:
if __name__ == "__main__":
    root = tk.Tk()
    root.title("PDF Viewer with Page Number")

    pdf_file = "your_pdf_file.pdf"  # Replace with your PDF path
    viewer = create_pdf_viewer(root, pdf_file, width=1000, height=800)  # Adjust the size as needed
    viewer.pack(fill="both", expand=True)

    root.mainloop()
