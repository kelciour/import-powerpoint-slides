import os
import re

import aqt

from aqt import mw
from aqt.qt import *

from . import form as ui_form

from bs4 import BeautifulSoup

def innerHTML(tag):
    for img in tag.find_all('img'):
        if img.has_attr('alt'):
            del img['alt']
    txt = tag.decode_contents()
    txt = re.sub(r'\s+', ' ', txt)
    txt = re.sub(r'^(\s|<br ?/?>)+|(\s|<br ?/?>)+$', '', txt)
    txt = txt.strip()
    return txt

MODEL_NAME = 'PowerPoint Slides'

def create_new_model():
    model = mw.col.models.new(MODEL_NAME)
    mw.col.models.addField(model, mw.col.models.new_field("Front Text"))
    mw.col.models.addField(model, mw.col.models.new_field("Front Screenshot"))
    mw.col.models.addField(model, mw.col.models.new_field("Back Text"))
    mw.col.models.addField(model, mw.col.models.new_field("Back Screenshot"))
    t = mw.col.models.new_template("Card 1")
    t['qfmt'] = """
<div>{{Front Screenshot}}</div>
<div>{{Front Text}}</div>
""".strip()
    t['afmt'] = """
{{FrontSide}}

<hr id=answer>

<div>{{Back Screenshot}}</div>
<div>{{Back Text}}</div>
""".strip()
    model['css'] =  """
.card {
 font-family: arial;
 font-size: 20px;
 text-align: center;
 color: black;
 background-color: white;
}
""".strip()
    mw.col.models.addTemplate(model, t)
    mw.col.models.add(model)
    return model

def maybe_add_fields():
    m = mw.col.models.by_name(MODEL_NAME)
    fieldnames = mw.col.models.field_names(m)
    nflds = []
    for fld in ["Front Text", "Front Screenshot", "Back Text", "Back Screenshot"]:
        if fld not in fieldnames:
            nflds.append(fld)
    if nflds:
        mw.progress.start()
        for fld in nflds:
            fm = mw.col.models.new_field(fld)
            mw.col.models.addField(m, fm)
        mw.col.models.save(m)
        mw.progress.finish()
    return m

def main():
    config = mw.addonManager.getConfig(__name__)

    dirkey = f"ImportFromPowerPointDirectory"
    dir = aqt.mw.pm.profile.get(dirkey, "")

    fname, _ = QFileDialog.getOpenFileName(filter = "HTML Document (*.htm)", directory=dir)

    dir = os.path.dirname(fname)
    aqt.mw.pm.profile[dirkey] = dir

    if not fname:
        return

    diag = QDialog()
    form = ui_form.Ui_Dialog()
    form.setupUi(diag)

    def showFileDialog(dir):
        fname, _ = QFileDialog.getOpenFileName(filter = "HTML Document (*.htm)", directory=dir)
        if fname:
            form.lineEdit.setText(fname)
            dir = os.path.dirname(fname)
            aqt.mw.pm.profile[dirkey] = dir
        else:
            form.lineEdit.setText('')

    form.fileButton.clicked.connect(lambda: showFileDialog(dir))
    form.lineEdit.setText(fname)
    form.skipFirstSlide.setChecked(config["skip first slide"])
    if config["one slide per note"]:
        form.radioOneSlide.setChecked(True)
    else:
        form.radioTwoSlides.setChecked(True)

    if not diag.exec():
        return

    fname = form.lineEdit.text()

    if not fname:
        return

    config["skip first slide"] = form.skipFirstSlide.isChecked()
    config["one slide per note"] = form.radioOneSlide.isChecked()
    mw.addonManager.writeConfig(__name__, config)

    if not mw.col.models.by_name(MODEL_NAME):
        model = create_new_model()
    else:
        model = maybe_add_fields()

    base_folder = os.path.dirname(fname)
    pptx_filename = os.path.splitext(os.path.basename(fname))[0]
    with open(fname, 'r', encoding='utf-8') as f_htm:
        soup = BeautifulSoup(f_htm, 'html.parser')
    divs = soup.find('table').find_all('div')
    cards = []
    for div in divs:
        img_filename = div.find('a')['href']
        text_filename = img_filename.replace('img', 'text')
        img_filepath = os.path.join(base_folder, img_filename)
        text_filepath = os.path.join(base_folder, text_filename)
        with open(img_filepath, 'r', encoding='utf-8') as f_img, open(text_filepath, 'r', encoding='utf-8') as f_text:
            soup_img = BeautifulSoup(f_img, 'html.parser')
            center = soup_img.body.find('center')
            center.extract()

            center = soup_img.body.find('center')
            center.extract()

            for img in center.find_all('img'):
                img_path = os.path.join(base_folder, img['src'])
                media_filename = pptx_filename + ' ' + img['src']
                media_filename = media_filename.replace(' ', '_')
                with open(img_path, "rb") as file_media:
                    img_src = mw.col.media.write_data(media_filename, file_media.read())
                img['src'] = img_src
            data_img = innerHTML(center)

            soup_text = BeautifulSoup(f_text, 'html.parser')
            center = soup_text.body.find('center')
            center.extract()

            for tag in soup_text.body.find_all('h2'):
                tag.name = 'h3'
            for tag in soup_text.body.find_all('h1'):
                tag.name = 'h2'

            for font in soup_text.body.find_all('font'):
                if font['color'] == "#FFFFFF":
                    font['color'] = "#000000"

            data_text = innerHTML(soup_text.body)
            cards.append({
                "text": data_text,
                "img": data_img,
            })

    mw.col.models.set_current(model)
    did = mw.col.decks.id(pptx_filename)
    mw.col.decks.select(did)
    deck = mw.col.decks.get(did)
    deck['mid'] = model['id']
    mw.col.decks.save(deck)

    if config["skip first slide"]:
        cards = cards[1:]
    
    step = 1
    if not config["one slide per note"]:
        step = 2

    mw.progress.start(immediate=True)
    for i in range(0, len(cards), step):
        note = mw.col.newNote(forDeck=False)
        note.note_type()['did'] = did
        c = cards[i]
        note['Front Text'] = c['text']
        note['Front Screenshot'] = c['img']
        if step == 2 and i + 1 < len(cards):
            c = cards[i+1]
            note['Back Text'] = c['text']
            note['Back Screenshot'] = c['img']
        mw.col.addNote(note)
        mw.app.processEvents()
    mw.progress.finish()
    
    mw.reset()

action = QAction("Import Slides", mw)
action.triggered.connect(main)
mw.form.menuTools.addAction(action)
