import pdfplumber
import re

print("PDF Outline to Markdown Conversion")
print("updated May 2024 by abbeyjulian")
print("---------------------------------")

# get file to convert
filepath = input("Filepath of file to be converted: ")
filetype = filepath.split(".")[-1]
print("file type is:", filetype)
filename = "".join(filepath.split(".")[:-1])
print("file name is:", filename)
pdf = pdfplumber.open(filepath)

# get header information
has_header = input("Does the file have a header? (Y/N) ")
if has_header.lower() == 'y':
    header_length = int(input("How many lines is the header? "))
else:
    header_length = 0

# get footer information
has_footer = input("Does the file have a footer? (Y/N) ")
if has_footer.lower() == 'y':
    footer_length = int(input("How many lines is the footer? "))
else:
    footer_length = 0

# get outline numberings
number_levels = int(input("How many levels does the outline have? "))
styles = []
for i in range(1,number_levels+1):
    styles.append(input("What is the style of level {}? \nUse 'X' for numbers, 'Y' for uppercase letters, 'y' for lowercase letters, 'R' for uppercase roman numerals, and 'r' for lowercase roman numerals. \n(for example: Y. or (X) or y. or Part X etc.) ".format(i)))
# set for REGEX matching
styles = [s.replace("X", r'\d') for s in styles]
styles = [s.replace("Y", r'[A-Z]') for s in styles]
styles = [s.replace("y", r'[a-z]') for s in styles]
styles = [s.replace("R", r'[A-Z]+') for s in styles]
styles = [s.replace("r", r'[a-z]+') for s in styles]
styles = [s.replace(".", r'\.') for s in styles]
styles = [s.replace(")", r'\)') for s in styles]
styles = [s.replace("(", r'\(') for s in styles]

# how many levels use markdown headings
number_headings = int(input("How many levels should use markdown headings (using #, ##, etc.)? "))

current_level = 0
md_text = ""
for page in pdf.pages:
    # get text and remove header/footer from page text
    text = page.extract_text()
    text = text.split("\n")[header_length:-footer_length]
    for t in text:
        matched = False
        # same level
        if re.match(styles[current_level], t):
            md_text += "\n"
            if current_level < number_headings:
                md_text += ("#"*(current_level+1))
                md_text += " "
            else:
                md_text += "\n"
            md_text += t
            matched = True
        # up level(s)
        elif current_level-1 >= 0:
            cl = current_level
            while cl-1 >= 0:
                if re.match(styles[cl-1], t):
                    md_text += "\n"
                    current_level = cl-1
                    if current_level < number_headings:
                        md_text += ("#" * (current_level + 1))
                        md_text += " "
                    else:
                        md_text += "\n"
                    md_text += t
                    matched = True
                    break
                cl -= 1
        # down a level
        if not matched and current_level+1 < number_levels and re.match(styles[current_level+1], t):
            md_text += "\n"
            current_level += 1
            if current_level < number_headings:
                md_text += ("#" * (current_level + 1))
                md_text += " "
            else:
                md_text += "\n"
            md_text += t
            matched = True
        # continuation of paragraph
        if not matched:
            md_text += " "
            md_text += t
    # handle tables
    table = page.extract_tables(table_settings={})
    unformatted_table = ""
    formatted_table = ""
    for _table in table:
        unformatted_table = ""
        formatted_table = ""
        header = True
        for row in _table:
            unformatted_table += (" ".join(row) + " ")
            formatted_table += ("| " + " | ".join(row) + " |\n")
            if header:
                formatted_table += ("| " + ("---|")*len(row)+"\n")
                header = False
        unformatted_table = unformatted_table[:-1]  # remove space at end
        formatted_table = "\n\n" + formatted_table  # preprend newlines
        md_text = md_text.replace(unformatted_table, formatted_table)
# save to markdown file
with open(filename + ".md", 'w') as file:
    file.write(md_text)
print("Saved to ", filename, ".md")
