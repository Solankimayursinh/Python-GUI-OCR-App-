  
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pytesseract
import os
import pyperclip # for clipboard copying
from docx import Document # for saving as docx

# Ensure Tesseract is installed and its path is set
pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR/tesseract.exe"  # Adjust to according

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR Application")
        self.root.geometry("800x650")
        self.root.maxsize(800, 650)
        self.root.minsize(800, 650)
        ctk.set_appearance_mode("light")  # Use 'light' for light mode
        ctk.set_default_color_theme("green")  # You can also try 'green', 'blue', etc.

        # Image preview area
        self.image_label = ctk.CTkLabel(self.root, text="No Image Selected", width=350, height=350, fg_color="gray")
        self.image_label.grid(row=0, column=0, padx=20, pady=20, sticky="nw")

        # OCR Text output area
        self.text_area = ctk.CTkTextbox(self.root, width=350, height=350, wrap="word")
        self.text_area.grid(row=0, column=1, padx=20, pady=20, sticky="ne")

        # Browse image button (multiple images)
        self.browse_button = ctk.CTkButton(self.root, text="Browse Images", command=self.browse_images)
        self.browse_button.grid(row=1, column=0, pady=10, padx=10, sticky="w")

        # Centered elements - Language label, dropdown, and OCR button
        self.language_label = ctk.CTkLabel(self.root, text="Select Language:")
        self.language_label.grid(row=2, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

        self.languages = ["eng", "hin", "guj", "san", "tam", "tel"]  # Extend as needed
        self.language_var = ctk.StringVar(value=self.languages[0])
        self.language_dropdown = ctk.CTkOptionMenu(self.root, variable=self.language_var, values=self.languages)
        self.language_dropdown.grid(row=3, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

        # Perform OCR button
        self.ocr_button = ctk.CTkButton(self.root, text="Perform OCR", command=self.perform_ocr)
        self.ocr_button.grid(row=4, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

        # Add Copy to Clipboard button
        self.copy_button = ctk.CTkButton(self.root, text="Copy to Clipboard", command=self.copy_to_clipboard)
        self.copy_button.grid(row=5, column=0, pady=10, padx=10, sticky="w")

        # Add Save as DOCX button
        self.save_button = ctk.CTkButton(self.root, text="Save as DOCX", command=self.save_as_docx)
        self.save_button.grid(row=5, column=1, pady=10, sticky="e")

        self.image_paths = []  # List to store selected image paths

    def browse_images(self):
        # Limit to selecting up to 10 images
        self.image_paths = filedialog.askopenfilenames(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
        
        if self.image_paths:
            if len(self.image_paths) > 10:
                messagebox.showwarning("Warning", "You can select a maximum of 10 images.")
                self.image_paths = self.image_paths[:10]  # Limit to 10 images
            
            # Display the first image from the selection in the image preview
            self.display_image(self.image_paths[0])

    def display_image(self, img_path):
        try:
            image = Image.open(img_path)
            image.thumbnail((300, 300))  # Adjust preview size
            img = ImageTk.PhotoImage(image)
            self.image_label.configure(image=img, text="")
            self.image_label.image = img  # Store a reference to prevent garbage collection
        except Exception as e:
            messagebox.showerror("Error", f"Error displaying image: {str(e)}")

    def perform_ocr(self):
        if self.image_paths:
            # Clear the text area before starting OCR
            self.text_area.delete(1.0, "end")
            
            try:
                language = self.language_var.get()
                for image_path in self.image_paths:
                    # Perform OCR on each image
                    ocr_text = self.extract_text(image_path, language)
                    
                    # Append image file name as a header in the text area
                    image_name = os.path.basename(image_path)
                    self.text_area.insert("end", f"--- OCR result for {image_name} ---\n")
                    self.text_area.insert("end", ocr_text)
                    self.text_area.insert("end", "\n\n")
                
                messagebox.showinfo("Success", "OCR process completed for all images")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred during OCR: {str(e)}")
        else:
            messagebox.showwarning("Warning", "No images selected!")

    def extract_text(self, image_path, lang):
        return pytesseract.image_to_string(Image.open(image_path), lang=lang)

    def copy_to_clipboard(self):
        # Copy the content of the text area to the clipboard
        try:
            ocr_text = self.text_area.get("1.0", "end").strip()
            pyperclip.copy(ocr_text)
            messagebox.showinfo("Success", "Text copied to clipboard.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def save_as_docx(self):
        # Save the OCR results as a .docx file
        try:
            doc = Document()
            ocr_text = self.text_area.get("1.0", "end").strip()
            if ocr_text:
                doc.add_paragraph(ocr_text)
                file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
                if file_path:
                    doc.save(file_path)
                    messagebox.showinfo("Success", f"File saved as {os.path.basename(file_path)}")
            else:
                messagebox.showwarning("Warning", "No OCR text to save.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving: {str(e)}")

# Main application execution
if __name__ == "__main__":
    root = ctk.CTk()
    app = OCRApp(root)
    root.mainloop()
