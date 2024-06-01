import pdfplumber
import re
import tkinter as tk
from tkinter import filedialog
import logging
from operator import itemgetter

print("PDF Outline to Markdown Conversion Version 2.02")
print("updated June 1 2024 by abbeyjulian")
print("---------------------------------")


class Filenames:
    def __init__(self):
        self.filenames = []


def check_bboxes(word, table_bbox):
    """
    Check whether word is inside a table bbox.
    """
    l = word['x0'], word['top'], word['x1'], word['bottom']
    r = table_bbox
    return l[0] > r[0] and l[1] > r[1] and l[2] < r[2] and l[3] < r[3]


# for error reporting
logging.basicConfig(filename='pdf_to_markdown.log', level=logging.ERROR, format='%(asctime)s %(levelname)s %(name)s %(message)s')
logger = logging.getLogger(__name__)


window = tk.Tk()
window.title("Markdown Conversion App")
header1 = tk.Label(text="PDF Outline to Markdown Conversion Version 2.02")
header2 = tk.Label(text="updated June 1 2024 by abbeyjulian")
header3 = tk.Label(text="---------------------------------")
header1.pack()
header2.pack()
header3.pack()


def browseFiles():
    filenames.filenames = filedialog.askopenfilenames(title="Select file(s)",
                                          filetypes=(("PDF files",
                                                      "*.pdf"),
                                                     ("DOCX files",
                                                      "*.docx")))

    # Change label contents
    label_file_explorer.configure(text="{} file(s) selected to convert.".format(len(filenames.filenames)))


# get file to convert
filenames = Filenames()
label_file_explorer = tk.Label(text="Select file(s) to convert:")
button_explore = tk.Button(window, text="Browse Files", command=browseFiles)
label_file_explorer.pack()
button_explore.pack()

# get header information
header_label = tk.Label(text="Number of lines in header: ")
header_entry = tk.Entry()
header_label.pack()
header_entry.pack()

# get footer information
footer_label = tk.Label(text="Number of lines in footer: ")
footer_entry = tk.Entry()
footer_label.pack()
footer_entry.pack()

# get outline numberings
styles_label = tk.Label(text="What is the style of each outline level? Give as a comma-separated list. \nUse 'X' for numbers, 'Y' for uppercase letters, 'y' for lowercase letters, 'R' for uppercase roman numerals, and 'r' for lowercase roman numerals. \n(for example: Y., (X), y., Part X)")
styles_entry = tk.Entry()
styles_label.pack()
styles_entry.pack()

# how many levels use markdown headings
headings_label = tk.Label(text="How many levels should use markdown headings (using #, ##, etc.)? ")
headings_entry = tk.Entry()
headings_label.pack()
headings_entry.pack()

def handle_click(event):
    success = []
    fail = []
    for i in range(len(filenames.filenames)):
        try:
            print("Converting...{}/{}".format(i, len(filenames.filenames)))
            filepath = filenames.filenames[i]
            filetype = filepath.split(".")[-1]
            print("file type is:", filetype)
            filename = "".join(filepath.split(".")[:-1])
            print("file name is:", filename)
            pdf = pdfplumber.open(filepath)
            header_length = int(header_entry.get())
            footer_length = int(footer_entry.get())
            styles = styles_entry.get().split(",")
            styles = [s.strip() for s in styles]
            number_levels = len(styles)
            # set for REGEX matching
            styles = [s.replace("X", r'\d+') for s in styles]
            styles = [s.replace("Y", r'[A-Z]+') for s in styles]
            styles = [s.replace("y", r'[a-z]+') for s in styles]
            styles = [s.replace("R", r'[A-Z]+') for s in styles]
            styles = [s.replace("r", r'[a-z]+') for s in styles]
            styles = [s.replace(".", r'\.') for s in styles]
            styles = [s.replace(")", r'\)') for s in styles]
            styles = [s.replace("(", r'\(') for s in styles]
            styles = [s.strip() for s in styles]
            styles = [s+" " for s in styles]
            number_headings = int(headings_entry.get())
            current_level = 0
            md_text = ""
            for page in pdf.pages:
                tables = page.find_tables()
                table_bboxes = [i.bbox for i in tables]
                tables = [{'table': i.extract(), 'top': i.bbox[1]} for i in tables]
                non_table_words = [word for word in page.extract_words() if
                                   not any([check_bboxes(word, table_bbox) for table_bbox in table_bboxes])]
                iter_list = pdfplumber.utils.cluster_objects(non_table_words + tables, itemgetter('top'), tolerance=5)
                iter_list = iter_list[header_length:-footer_length]
                for cluster in iter_list:
                    if 'text' in cluster[0]:
                        t = ' '.join([i['text'] for i in cluster])
                        matched = False
                        # same level
                        if re.match(styles[current_level], t):
                            md_text += "\n"
                            if current_level < number_headings:
                                md_text += ("#" * (current_level + 1))
                                md_text += " "
                            elif current_level >= number_headings + 1:
                                md_text += ("\t" * (current_level - number_headings))
                            else:
                                md_text += ("\n" + "\t" * (current_level - number_headings))
                            md_text += t
                            matched = True
                        # up level(s)
                        elif current_level - 1 >= 0:
                            cl = current_level
                            while cl - 1 >= 0:
                                if re.match(styles[cl - 1], t):
                                    md_text += "\n"
                                    current_level = cl - 1
                                    if current_level < number_headings:
                                        md_text += ("#" * (current_level + 1))
                                        md_text += " "
                                    elif current_level >= number_headings + 1:
                                        md_text += ("\t" * (current_level - number_headings))
                                    else:
                                        md_text += ("\n" + "\t" * (current_level - number_headings))
                                    md_text += t
                                    matched = True
                                    break
                                cl -= 1
                        # down a level
                        if not matched and current_level + 1 < number_levels and re.match(styles[current_level + 1], t):
                            md_text += "\n"
                            current_level += 1
                            if current_level < number_headings:
                                md_text += ("#" * (current_level + 1))
                                md_text += " "
                            elif current_level >= number_headings + 1:
                                md_text += ("\t" * (current_level - number_headings))
                            else:
                                md_text += ("\n" + "\t" * (current_level - number_headings))
                            md_text += t
                            matched = True
                        # continuation of paragraph
                        if not matched:
                            md_text += " "
                            md_text += t
                    elif 'table' in cluster[0]:
                        formatted_table = ""
                        header = True
                        _table = cluster[0]['table']
                        for row in _table:
                            row = ["" if r is None else r for r in row]
                            row = [s.replace("\n", " ") for s in row]
                            formatted_table += ("| " + " | ".join(row) + " |\n")
                            if header:
                                formatted_table += ("| " + ("---|")*len(row)+"\n")
                                header = False
                        formatted_table = "\n\n" + formatted_table  # preprend newlines
                        md_text += formatted_table
            # save to markdown file
            with open(filename + ".md", 'w') as file:
                file.write(md_text)
            print("Saved to ", filename, ".md")
            success.append(i)
        except Exception as e:
            logger.log(logging.ERROR, "error with file {}: ".format(filenames.filenames[i]))
            logger.exception(e)
            fail.append(i)
            continue

    popup = tk.Toplevel()
    confirmation = tk.Label(popup, text="Conversion complete.")
    confirmation.pack()
    successful_count = tk.Label(popup, text="{}/{} files successfully converted.".format(len(success), len(filenames.filenames)))
    successful_count.pack()
    if len(success) < len(filenames.filenames):
        errorlist = "The following files encountered a conversion error: \n"
        for i in fail:
            errorlist += (filenames.filenames[i] + "\n")
        errorlist += "See pdf_to_markdown.log for error details."
        errorlist_label = tk.Label(popup, text=errorlist)
        errorlist_label.pack()
    ok_button = tk.Button(popup, text="OK", command=popup.destroy)
    ok_button.pack()
    popup.mainloop()


framea = tk.Frame(relief=tk.RAISED)
framea.pack()
submit = tk.Button(master=framea, text="Convert")
submit.pack()
submit.bind("<Button-1>", handle_click)

window.mainloop()
