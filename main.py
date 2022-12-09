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

# models for anki deck
img_txt_model = genanki.Model(
    1607392319,
    'Simple Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Answer'},
        {'name': 'Structure'}
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Question}}',
            'afmt': '{{FrontSide}}<hr id="answer">{{MyMedia}}<br>{{Answer}}',
        },
    ])
img_model = genanki.Model(
    2112975270,
    'Simple Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Structure'}
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Question}}',
            'afmt': '{{FrontSide}}<hr id="answer">{{MyMedia}}',
        },
    ])
txt_model = genanki.Model(
    1360356766,
    'Simple Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Answer'}
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Question}}',
            'afmt': '{{FrontSide}}<hr id="answer">{{Answer}}',
        },
    ])
no_model = genanki.Model(
    1471681582,
    'Simple Model',
    fields=[
        {'name': 'Question'}
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Question}}',
            'afmt': '{{FrontSide}}<hr id="answer">',
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
for i in range(1, 53):
    if image_loader.image_in(f'C{i}'):
        image = image_loader.get(f'C{i}')
        rgb_im = image.convert('RGB')
        outfile = f'images/{img_count}.jpg'
        rgb_im.save(outfile)
        my_package.media_files.append(outfile)
        if sheet[f'D{i}'].value:
            my_note = genanki.Note(
                model=img_txt_model,
                fields=[sheet[f'B{i}'].value, sheet[f'D{i}'].value, f'<img src="{img_count}.jpg">'])
            img_count += 1
            my_deck.add_note(my_note)
        else:
            my_note = genanki.Note(
                model=img_model,
                fields=[sheet[f'B{i}'].value, f'<img src="{img_count}.jpg">'])
            img_count += 1
            my_deck.add_note(my_note)
    else:
        if sheet[f'D{i}'].value:
            my_note = genanki.Note(
                model=txt_model,
                fields=[sheet[f'B{i}'].value, sheet[f'D{i}'].value])
            my_deck.add_note(my_note)
        else:
            my_note = genanki.Note(
                model=no_model,
                fields=[sheet[f'B{i}'].value])
            my_deck.add_note(my_note)

# export apkg
my_package.write_to_file('output.apkg')
