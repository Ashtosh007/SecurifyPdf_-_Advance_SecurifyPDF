import os
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfWriter, PdfReader, PdfMerger
from pdf2image import convert_from_path
import comtypes.client  # For document conversion to PDF (for Windows)
from PIL import Image

# Initialize main window
root = Tk()
root.title("SecurifyPdf")
root.geometry("440x650")
root.configure(bg="#BFEFFF")  # Light blue Background
Label(root, text="SecurifyPdf", font="Arial 30 bold", bg="#BFEFFF").pack(pady=15)
Label(root, text="🔒 Encrypt Today, Secure Forever! 🔑:").pack(pady=8)
root.resizable(False, False)
# Create a scrollable frame
main_canvas = Canvas(root)
scrollbar = Scrollbar(root, orient=VERTICAL, command=main_canvas.yview)
scrollable_frame = Frame(main_canvas,bg="Light Blue")

scrollable_frame.bind(
    "<Configure>",
    lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
)

main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
main_canvas.configure(yscrollcommand=scrollbar.set)

main_canvas.pack(side=LEFT, fill=BOTH, expand=True)
scrollbar.pack(side=RIGHT, fill=Y)
# Global Variables
filename = None
merge_files = []
split_file = None
output_folder = None


# Function to browse a single PDF
def browse():
    global filename
    filename = filedialog.askopenfilename(title="Select PDF File", filetypes=[('PDF File', '*.pdf')])
    entry1.delete(0, END)
    entry1.insert(END, filename)


# Function to browse multiple PDFs for merging
def browse_multiple():
    global merge_files
    merge_files = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[('PDF File', '*.pdf')])
    entry2.delete(0, END)
    entry2.insert(END, ', '.join(merge_files))


# Function to browse PDF for splitting
def browse_split():
    global split_file
    split_file = filedialog.askopenfilename(title="Select PDF File to Split", filetypes=[('PDF File', '*.pdf')])
    entry3.delete(0, END)
    entry3.insert(END, split_file)


# Function to browse output folder
def browse_output_folder():
    global output_folder
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    entry4.delete(0, END)
    entry4.insert(END, output_folder)


# Function to protect a PDF with a password
def protect_pdf():
    mainfile = entry1.get()
    protectfile = os.path.join(output_folder, "protected.pdf")
    password = entry5.get()

    if not mainfile or not output_folder or not password:
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    try:
        writer = PdfWriter()
        reader = PdfReader(mainfile)

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(password)

        with open(protectfile, "wb") as f:
            writer.write(f)

        messagebox.showinfo("Success", f"PDF protected successfully: {protectfile}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Function to merge multiple PDFs
def merge_pdfs():
    if not merge_files or not output_folder:
        messagebox.showerror("Error", "Please select files and output folder.")
        return

    try:
        merger = PdfMerger()
        for pdf in merge_files:
            merger.append(pdf)

        output = os.path.join(output_folder, 'merged_output.pdf')
        merger.write(output)
        merger.close()

        messagebox.showinfo("Success", f"PDFs merged successfully: {output}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Function to split a PDF
def split_pdf():
    if not split_file or not output_folder:
        messagebox.showerror("Error", "Please select a file and output folder.")
        return

    try:
        pdf = PdfReader(split_file)
        for page_num, page in enumerate(pdf.pages, start=1):
            writer = PdfWriter()
            writer.add_page(page)
            output = os.path.join(output_folder, f'split_page_{page_num}.pdf')

            with open(output, 'wb') as out_file:
                writer.write(out_file)

        messagebox.showinfo("Success", "PDF split successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to convert PDF to images
def pdf_to_images():
    if not filename or not output_folder:
        messagebox.showerror("Error", "Please select a PDF file and output folder.")
        return

    poppler_path = r"F:\FINAL_PROJECT\pythonProject\poppler-24.08.0\Library\bin"  # Update this path

    try:
        images = convert_from_path(filename, poppler_path=poppler_path)
        for i, image in enumerate(images):
            image.save(os.path.join(output_folder, f'page_{i + 1}.png'), 'PNG')

        messagebox.showinfo("Success", "PDF converted to images successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Function to convert multiple images to a PDF
def images_to_pdf():
    global output_folder
    image_files = filedialog.askopenfilenames(
        title="Select Images",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")]
    )

    if not image_files or not output_folder:
        messagebox.showerror("Error", "Please select images and an output folder.")
        return

    try:
        image_list = [Image.open(img).convert("RGB") for img in image_files]
        output_pdf = os.path.join(output_folder, "output.pdf")
        image_list[0].save(output_pdf, save_all=True, append_images=image_list[1:])
        messagebox.showinfo("Success", f"Images converted to PDF successfully: {output_pdf}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Universal function to create a section frame
def create_section(parent, title):
    frame = LabelFrame(parent, text=title, font="Arial 12 bold", fg="#333", padx=10, pady=5, bd=2, relief=GROOVE)
    frame.pack(fill=X, padx=15, pady=8)
    return frame

# Output Folder Selection
output_frame = create_section(scrollable_frame, "Output Folder")
entry4 = Entry(output_frame, width=50)
entry4.pack(side=LEFT, padx=5, pady=5)
Button(output_frame, text="Browse", command=browse_output_folder, bg="#007BFF", fg="white").pack(side=RIGHT)

# PDF Protection Section
protect_frame = create_section(scrollable_frame, "Protect PDF")
Label(protect_frame, text=" File:").pack(pady=2)
entry1 = Entry(protect_frame, width=50)
entry1.pack(pady=5)
Button(protect_frame, text="Browse PDF", command=browse).pack()
Label(protect_frame, text="Enter Password:").pack(pady=2)
entry5 = Entry(protect_frame, width=30, show="*")
entry5.pack(pady=5)
Button(protect_frame, text="Protect PDF", command=protect_pdf, bg="#DC3545", fg="white").pack(pady=5)

# Merge PDFs Section
merge_frame = create_section(scrollable_frame, "Merge PDFs")
entry2 = Entry(merge_frame, width=50)
entry2.pack(pady=5)
Button(merge_frame, text="Browse PDFs", command=browse_multiple).pack()
Button(merge_frame, text="Merge PDFs", command=merge_pdfs, bg="#17A2B8", fg="white").pack(pady=5)

# Split PDF Section
split_frame = create_section(scrollable_frame, "Split PDF")
entry3 = Entry(split_frame, width=50)
entry3.pack(pady=5)
Button(split_frame, text="Browse PDF", command=browse_split).pack()
Button(split_frame, text="Split PDF", command=split_pdf, bg="#FFC107").pack(pady=5)

# Convert PDF to Images Section
pdf_to_img_frame = create_section(scrollable_frame, "Convert PDF to Images")
Button(pdf_to_img_frame, text="Select PDF", command=browse, bg="#28A745", fg="white").pack(pady=5)
Button(pdf_to_img_frame, text="Convert to Images", command=pdf_to_images, bg="#FFC107").pack(pady=5)

# Convert Images to PDF Section
img_to_pdf_frame = create_section(scrollable_frame, "Convert Images to PDF")
Button(img_to_pdf_frame, text="Select Images", command=images_to_pdf, bg="#17A2B8", fg="white").pack(pady=5)


root.mainloop()
