import tkinter as tk
import PyPDF2
import docx
from PIL import Image, ImageTk
from tkinter.filedialog import askopenfile, asksaveasfilename

root = tk.Tk()
root.title("PDF to Word Doc")
# All code must be placed between these two lines of code [root and root.mainloop()]

canvas = tk.Canvas(root, width=600, height=300)
canvas.grid(columnspan=3, rowspan=4)

# Logo
logo = Image.open('logo.jpg')
logo = ImageTk.PhotoImage(logo)
logo_label = tk.Label(image=logo)
logo_label.image = logo
logo_label.grid(column=1, row=0)

# instructions
instructions = tk.Label(root, text="Select a PDF file on your computer to extract all it's text", font="Raleway")
instructions.grid(columnspan=3, column=0, row=1)


def open_file():
    browse_text.set("loading...")
    file = askopenfile(parent=root, mode='rb', title="Choose a file", filetype=[('Pdf file', '*.pdf')])
    if file:
        read_pdf = PyPDF2.PdfFileReader(file)
        page = read_pdf.getPage(0)
        page_content = page.extractText()

        #text box | height = lines of text | width = character count
        text_box.delete('1.0', tk.END)
        text_box.insert(1.0, page_content)

        # Set the browse button text back to browse
        browse_text.set("Browse")


def select_directory():
    save_text.set("saving...")
    file_directory = asksaveasfilename(defaultextension=".docx")
    save_file(file_directory)

def save_file(file_directory):
    doc = docx.Document()
    pdf_text = str(text_box.get(1.0, tk.END))
    
    doc.add_paragraph(pdf_text)
    doc.save(file_directory)

    save_text.set("Save")


# browse buttons
browse_text = tk.StringVar()
browse_btn = tk.Button(root, textvariable=browse_text, command=lambda:open_file(), font="Raleway", bg="#20bebe", fg="white", height=2, width=10)
browse_text.set("Browse")
browse_btn.grid(column=1, row=2)

# save button
save_text = tk.StringVar()
save_btn = tk.Button(root, textvariable=save_text, command=lambda:select_directory(), font="Raleway", bg="#20bebe", fg="white", height=2, width=10, pady=4)
save_text.set("Save")
save_btn.grid(column=1, row=4)

# text box
text_box = tk.Text(root, height=10, width=50, padx=15, pady=15)
text_box.insert(1.0, "...")
text_box.tag_configure("center", justify="center")
text_box.tag_add("center", 1.0, "end")
text_box.grid(column=1, row=3)

canvas = tk.Canvas(root, width=600, height=250)
canvas.grid(columnspan=3)

root.mainloop()

#add a save to word doc?
