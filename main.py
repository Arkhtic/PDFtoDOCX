import os
from tkinter import Tk, Button, Label, filedialog
from pdf2docx import Converter


def pdf_to_docx(input_folder, output_folder):
    # Loop through PDF files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith('.pdf'):
            # Convert PDF to DOCX
            pdf_path = os.path.join(input_folder, filename)
            docx_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.docx')
            convert_pdf_to_docx(pdf_path, docx_path)

            print(f'{filename} converted to DOCX successfully.')


def convert_pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()


def select_input_folder():
    input_folder = filedialog.askdirectory()
    input_folder_label.config(text=input_folder)


def select_output_folder():
    output_folder = filedialog.askdirectory()
    output_folder_label.config(text=output_folder)


def convert_to_docx():
    input_folder = input_folder_label.cget("text")
    output_folder = output_folder_label.cget("text")
    pdf_to_docx(input_folder, output_folder)
    status_label.config(text="Conversion complete.")


# Create GUI window
root = Tk()
root.title("PDF to DOCX Converter")

# Labels
input_folder_label = Label(root, text="Select input folder")
input_folder_label.grid(row=0, column=0)

output_folder_label = Label(root, text="Select output folder")
output_folder_label.grid(row=1, column=0)

status_label = Label(root, text="")
status_label.grid(row=3, column=0)

# Buttons
input_folder_button = Button(root, text="Browse", command=select_input_folder)
input_folder_button.grid(row=0, column=1)

output_folder_button = Button(root, text="Browse", command=select_output_folder)
output_folder_button.grid(row=1, column=1)

convert_button = Button(root, text="Convert", command=convert_to_docx)
convert_button.grid(row=2, column=0, columnspan=2)

# Run the GUI
root.mainloop()
