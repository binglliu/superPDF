#!/usr/bin/python

# Enable executing scripts without python launcher:
# For Windows 10, run following command in windows command.
#   assoc .py=Python.File
#   ftype Python.File="c:\Program Files\Python36\python.exe" "%1" %*

from wand.image import Image
from PIL import Image as PI
import pyocr
import pyocr.builders
import io
import difflib
import time
import datetime
import sys
import os
import shutil
import yaml
import codecs
import warnings

from PyPDF2 import PdfFileWriter, PdfFileReader

SAMPLES_FOLDER = 'samples'
LANG_ENGLISH = 'eng'

def write(str):
    sys.stdout.write(str)

class Document:
    name = None
    pages = None


class Page:
    page_no = 0
    sample = None
    doc = None

class Settings:

    def __init__(self):
        self.docs = []
        self.pages = []

    def load(self, dir, data):

        # search for samples folder in sample folder.
        doc = Document()
        doc.name = data['name']
        doc.pages = self.load_pages(dir, doc, data['pages'])
        self.append(doc)

    def append(self, doc):
        self.docs.append(doc)
        for page in doc.pages:
            self.pages.append(page)

    def load_pages(self, dir, doc, data):
        pages = []
        for page_no in data:
            page = Page()
            page.doc = doc
            page.page_no = page_no
            page.sample = self.load_sample(dir, data[page_no])
            pages.append(page)
        return pages

    def load_sample(self, dir, sample_filename):
        sample_filename = os.path.join(dir, sample_filename)

        # This works only on python 3
        #with open(sample_filename, 'r', encoding='utf8') as f:
        #    return f.read()

        # This may work on both 2 and 3.
        with codecs.open(sample_filename, 'r', 'utf-8') as f:
            return f.read()


class Match:
    def __init__(self, index, page, ratio):
        self.index = index
        self.page = page
        self.ratio = ratio

# OCR pdf and sort it based on templates.
class SuperPDF:

    settings = Settings()


    def __init__(self):
        tools = pyocr.get_available_tools()
        if len(tools) == 0:
            print("No OCR tool found, install tesseract?")
            sys.exit(1)
        self.tool = tools[0]
        # self.tool.get_available_languages()[0]
        self.lang = LANG_ENGLISH
        self.builder = pyocr.builders.TextBuilder()


    # Load settings from 'samples' folder.
    def load_settings(self):

        dir = os.path.dirname(__file__)
        dir = os.path.join(dir, SAMPLES_FOLDER)

        items = os.listdir(dir)
        for item in items:
            sub_dir = os.path.join(dir, item)
            if os.path.isdir(sub_dir):
                yaml_filename = os.path.join(sub_dir, "doc.yaml")
                if os.path.exists(yaml_filename):
                    #print(yaml_filename)
                    with open(yaml_filename) as file:
                        data = yaml.safe_load(file)
                        self.settings.load(sub_dir, data)


    def save_sample_file(self, output_dir, filename, text):
        filename = os.path.join(output_dir, filename)
        with open(filename, 'wb') as text_file:
            text_file.write(text.encode('utf8'))


    def create_sample_yaml(self, output_dir, doc_name, pages):
        data = {}
        data['name'] = doc_name
        data['created-on'] = datetime.datetime.now()
        data['pages'] = pages
        yaml_filename = os.path.join(output_dir, 'doc.yaml')
        #print (yaml_filename)
        with open(yaml_filename, 'w') as outfile:
            yaml.dump(data, outfile, default_flow_style=False)


    def create_sample_page_filename(self, doc_name, page_no):
        filename = create_filename(doc_name)
        filename = '{page:04d}.txt'.format(name=filename, page=page_no)
        return filename


    # OCR pdf file, extract text from each page and save to seperate text files.
    def sample(self, doc_name, filename, output_dir):


        write('Saving samples to "{dir}".\n'.format(dir=output_dir))

        pages = {}

        with Image(filename=filename, resolution=300) as file:

            page_no = 0

            for pdf_page in file.sequence:
                page_no = page_no + 1

                start = time.time()
                write('\tPage {pageno} ...... '.format(pageno=page_no))

                text = self.ocr(pdf_page)

                write('{t:.2f} seconds.\n'.format(t=time.time() - start))

                # Save sample text.
                page_filename = self.create_sample_page_filename(doc_name, page_no)
                pages[page_no] = page_filename
                self.save_sample_file(output_dir, page_filename, text)

        self.create_sample_yaml(output_dir, doc_name, pages)


    def process(self, filename):

        write('OCR {file} ...\n'.format(file=filename))

        matched = []
        # open pdf, it will convert pdf to images with 300dpi resolution.
        with Image(filename=filename, resolution=300) as file:
            index = 0
            for pdf_page in file.sequence:

                start = time.time()
                write('\tPage {pageno} ...... '.format(pageno=index+1))

                text = self.ocr(pdf_page)

                write('{t:.2f} seconds.'.format(t=time.time() - start))

                m = self.match(matched, index, text)
                if m:
                    write(' Match found. (Ratio={ratio:0.4f})'.format(ratio=m.ratio))
                write('\n')

                index = index + 1


        #split files.
        self.split(filename, matched)

    # OCR page image and return text in utf8 encoding.
    def ocr(self, page):

        with Image(page) as image:

            with image.convert('jpeg') as converted:
                blob = converted.make_blob('jpeg')
                text = self.tool.image_to_string(PI.open(io.BytesIO(blob)), lang=self.lang, builder=self.builder)

        return text


    def match(self, matched, index, text):

        last_ratio = 0
        match = None

        for page in self.settings.pages:

            ratio = difflib.SequenceMatcher(None, text, page.sample).ratio()

            # Max ration is 1.
            if ratio > 0.7:
                if ratio > last_ratio:
                    match = Match(index, page, ratio)
                    last_ratio = ratio
                    #matched.append(match)
                    #print ("Match {doc}, Page {no}. Ratio={r}".format(doc=page.doc.name, no=page.page_no, r=ratio))

        if match:
            matched.append(match)
        return match


    def split(self, filename, matched):

        output_dir = create_process_output_foldername(filename)

        # Sort by doc name and page no.
        matched = sorted(matched, key = lambda x: (x.page.doc.name, x.page.page_no))

        reader = PdfFileReader(open(filename, 'rb'))
        doc = None
        writer = None

        for m in matched:

            if not doc == m.page.doc:
                # New document, save and close current document.
                self.save_pdf(output_dir, writer, doc)

                # Create a new document to write.
                writer = PdfFileWriter()
                doc = m.page.doc

            writer.addPage(reader.getPage(m.index))

        self.save_pdf(output_dir, writer, doc)

    def save_pdf(self, output_dir, writer, doc):

        if writer:

            if not os.path.exists(output_dir):
                os.mkdir(output_dir)

            doc_filename = os.path.join(output_dir, doc.name + '.pdf')
            with open(doc_filename, "wb") as file_out:
                writer.write(file_out)

# lowercase replace ' ' with '_'
def create_filename(doc_name):
    name = doc_name.lower()
    name = '_'.join(name.split(' '))
    return name

# Create output folder for importing. located in 'samples' folder in the same folder as this script file.
def create_import_output_folder(doc_name, filename):
    dir = os.path.dirname(__file__)
    dir = os.path.join(dir, SAMPLES_FOLDER)
    output_dir = create_filename(doc_name)
    output_dir = os.path.join(dir, output_dir)
    output_dir = os.path.abspath(output_dir)
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.mkdir(output_dir)
    return output_dir

# Folder is located in the same folder as the target file.
def create_sample_output_folder(doc_name, filename):
    output_dir = create_filename(doc_name)
    dir_name = os.path.dirname(os.path.realpath(filename))
    output_dir = os.path.join(dir_name, output_dir)
    output_dir = os.path.abspath(output_dir)
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.mkdir(output_dir)
    return output_dir

# In same folder as the target file.
def create_process_output_foldername(filename):
    base_name = os.path.basename(filename)
    base_name = base_name.split('.')[0]
    dir_name = os.path.dirname(os.path.realpath(filename))
    output_dir = os.path.join(dir_name, base_name)
    output_dir = os.path.abspath(output_dir)
    return output_dir

def do_process(filename):
    sp = SuperPDF()
    sp.load_settings()
    sp.process(filename)

def do_sample(doc_name, filename):
    sp = SuperPDF()
    output_dir = create_sample_output_folder(doc_name, filename)
    sp.sample(doc_name, filename, output_dir)

def do_import(doc_name, filename):
    sp = SuperPDF()
    output_dir = create_import_output_folder(doc_name, filename)
    sp.sample(doc_name, filename, output_dir)

def main():

    # turn off warnings.
    warnings.simplefilter('ignore')

    count = len(sys.argv)
    if count == 2:
        filename = sys.argv[1]
        do_process(filename)
        return

    elif count == 3:
        action = sys.argv[1]
        filename = sys.argv[2]

        if action == 'sample':
            name = os.path.basename(filename)
            name = name.split('.')[0]
            do_sample(name, filename)
            return

        elif action == 'import':
            name = os.path.basename(filename)
            name = name.split('.')[0]
            do_import(name, filename)
            return


    # print out usage info:
    print('Usage:')
    print ('    python superpdf.py <filename>')
    print ('    python superpdf.py sample "<filename>"')
    print ('    python superpdf.py import "<filename>"')


def test1():
    do_process("./test/Note.pdf")

main()
