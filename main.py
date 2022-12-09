import os
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from pathlib import Path
import genanki
from PIL import Image

xlsx_file = Path('chem.xlsx')

# parse the xlsx_file
wb = load_workbook(xlsx_file)
sheet = wb.active
image_loader = SheetImageLoader(sheet)

# TODO: find a way to add styling to model
#  .card {
#   font-family: arial;
#   font-size: 20px;
#   text-align: center;
#   color: black;
#   background-color: white;
# }

# models for anki deck
my_model = genanki.Model(
    1607392319,
    'BasicPy',
    fields=[
        {'name': 'Front'},
        {'name': 'Back'}
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Front}}',
            'afmt': '{{FrontSide}}<hr id="answer">{{Back}}',
        },
    ])

# anki deck and package
my_deck = genanki.Deck(1177382377, 'Chemistry in Everyday Life')
my_package = genanki.Package(my_deck)

# images folder
if not os.path.exists("images"):
    os.makedirs("images")

# the actual cards (notes)
img_count = 0
for i in range(2, 53):
    if image_loader.image_in(f'C{i}'):
        image = image_loader.get(f'C{i}')
        rgb_im = image.convert('RGB')
        outfile = f'images/{img_count}.jpg'
        rgb_im.save(outfile)
        my_package.media_files.append(outfile)
        if sheet[f'D{i}'].value:
            my_note = genanki.Note(
                model=my_model,
                fields=[sheet[f'B{i}'].value, f'<img src="{img_count}.jpg"><br>' + sheet[f'D{i}'].value])
            img_count += 1
            my_deck.add_note(my_note)
        else:
            my_note = genanki.Note(
                model=my_model,
                fields=[sheet[f'B{i}'].value, f'<img src="{img_count}.jpg">'])
            img_count += 1
            my_deck.add_note(my_note)
    else:
        if sheet[f'D{i}'].value:
            my_note = genanki.Note(
                model=my_model,
                fields=[sheet[f'B{i}'].value, sheet[f'D{i}'].value])
            my_deck.add_note(my_note)
        else:
            my_note = genanki.Note(
                model=my_model,
                fields=[sheet[f'B{i}'].value,''])
            my_deck.add_note(my_note)

# export apkg
my_package.write_to_file('output.apkg')
