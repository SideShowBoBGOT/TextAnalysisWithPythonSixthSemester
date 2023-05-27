import os
from typing import Final
from enum import Enum

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from webdriver_manager.chrome import ChromeDriverManager

notebook_path = '/home/choleraplague/university/DataScience/TextAnalysisWithPython/TextAnalysis/Lab7/'
notebook_name = 'lab7'
out_name = u'Панченко_Сергій_ІП-11_7.odt'
front_path = u'front.odt'
back_path = u'back.odt'


def driver_factory() -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--window-size=850,1440")
    options.add_experimental_option('detach', True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                              options=options)
    return driver

class JpHtmlClasses(Enum):
    NONE = 'None'
    CELL = 'jp-Cell'
    MARKDOWN = 'jp-MarkdownCell'
    RENDERED_MARKDOWN = 'jp-RenderedMarkdown'
    CODE = 'jp-CodeCell'
    MATHJAX_PREVIEW = 'MathJax_Preview'
    

class StylesEnum(Enum):
    STANDARD = 'Standard'
    HEADING_ONE = 'Heading 1'
    HEADING_TWO = 'Heading 2'
    PICTURE_NAME = 'Picture Name'
    PICTURE = 'Picture'
    CODE = 'Preformatted Text'

class NoteBookParser():
    
    def __init__(self, driver: webdriver.Chrome) -> None:
        self.driver: webdriver.Chrome = driver
        self._doc_elements: list[tuple[StylesEnum, str]] = []
        self._picture_index: int = 0
        self._default_picture_name: Final[str] = 'picture_'
        self._suffix_png = '.png'
        

    def prepare_parse(self) -> None:
        self._doc_elements = []
        self._picture_index = 0

    def gen_picture_name(self) -> str:
        name = self._default_picture_name 
        name += str(self._picture_index) + self._suffix_png
        self._picture_index += 1
        return name
    
    def convert_to_html(self, path: str, notebook_name: str) -> None:
        os.system(f'jupyter nbconvert --to html {path + notebook_name}.ipynb')
    
    def parse(self, path: str, notebook_name: str) -> list:
        self.prepare_parse()
        self.convert_to_html(path, notebook_name)
        self.driver.get(url='file://' + path + notebook_name + '.html')
        html_cells = self.driver.find_elements(By.CLASS_NAME,
                                               JpHtmlClasses.CELL.value)
        for cell in html_cells:
            self.parse_cell(cell)
        return self._doc_elements
          
            
    def parse_cell(self, cell: WebElement) -> None:
        class_name = cell.get_attribute('class')
        class_name_parts = class_name.split(' ')
        if JpHtmlClasses.MARKDOWN.value in class_name_parts:
            self.parse_markdown(cell)
        elif JpHtmlClasses.CODE.value in class_name_parts:
            self.parse_code(cell)
    
    
    def parse_markdown(self, cell: WebElement) -> None:
        rendered = cell.find_element(By.CLASS_NAME,
                                     JpHtmlClasses.RENDERED_MARKDOWN.value)
        if cell.find_elements(By.CLASS_NAME,
                              JpHtmlClasses.MATHJAX_PREVIEW.value):
            return self.parse_mathjax(rendered)
        self.parse_rendered(rendered)


    def tag_as_word_parts(self, tag_attr: str) -> StylesEnum:
        match tag_attr:
            case 'h1':
                return StylesEnum.HEADING_ONE
            case 'h2':
                return StylesEnum.HEADING_TWO
            case 'h3':
                return StylesEnum.STANDARD
            case 'p':
                return StylesEnum.PICTURE_NAME
        
        return StylesEnum.STANDARD


    def parse_rendered(self, cell: WebElement) -> None:
        if cell.find_elements(By.TAG_NAME, 'img'):
            return self.parse_picture(cell)
        heading = cell.find_element(By.XPATH, '*')
        word_part = self.tag_as_word_parts(heading.tag_name)
        self._doc_elements.append((word_part, heading.text))


    def parse_picture(self, cell: WebElement) -> None:
        name = self.gen_picture_name()
        self._doc_elements.append((StylesEnum.PICTURE, name))
        cell.screenshot(name)
    
    
    def parse_mathjax(self, cell: WebElement) -> None:
        self.parse_picture(cell)
        
    
    def parse_code(self, cell: WebElement) -> None:
        self.parse_picture(cell)
    

from odf import style, text, draw
from odf.style import Style
from odf.text import P
from odf.opendocument import OpenDocumentText, OpenDocument, load

from PIL import Image

class WordConverter:
    def __init__(self, word_elements: list[tuple[StylesEnum, str]],
                 notebook_path: str, front_path: str, back_path: str,
                 out_name: str) -> None:
        
        self._word_elements = word_elements
        self._notebook_path = notebook_path
        self._front_path = front_path
        self._back_path = back_path
        self._out_name = out_name
        self._document = None
        self._pixel_cm_ratio = 53.54
        self._prefix_elements = u'elements'
        self._prefix_code = u'code'
        self._doc_name_elements = self._prefix_elements + self._out_name
        self._doc_name_code = self._prefix_code + self._out_name
        
    
    def convert(self) -> None:
        self._write_elements()
        self._write_code()
        os.system(f'ooo_cat {self._doc_name_elements} {self._doc_name_code} -o {self._out_name}')
        for name in [self._doc_name_elements, self._doc_name_code]:
            os.system(f'rm {name}') 
        

    def _write_elements(self) -> None:
        doc: OpenDocument = load(self._front_path)
        for part, text in self._word_elements:
            match part:
                case StylesEnum.PICTURE:
                    self._add_picture_paragraph(doc, text)
                case _:
                    self._add_text_paragraph(doc, part, text)
        doc.save(self._doc_name_elements)
    
    def _write_code(self) -> None:
        os.system(f'jupyter nbconvert {self._notebook_path}.ipynb --to python')
        lines = open(self._notebook_path + '.py').readlines()
        doc: OpenDocument = load(self._back_path)
        for line in list(filter(lambda x: not x.startswith('#') and x.strip() != '', lines)):
            self._add_text_paragraph(doc, StylesEnum.CODE, line)
        doc.save(self._doc_name_code)
        os.system(f'rm {self._notebook_path}.py')     

    
    def _add_text_paragraph(self, doc: OpenDocument, part: StylesEnum, text: str) -> None:
        doc.text.addElement(P(stylename=part.value, text=text))
    
    
    def _add_picture_paragraph(self, doc: OpenDocument, name: str) -> None:
        p = P(stylename=StylesEnum.PICTURE.value)
        doc.text.addElement(p)
        
        image = Image.open(name)
        width, height = image.size
        width /= self._pixel_cm_ratio
        height /= self._pixel_cm_ratio
        width, height = round(width, 2), round(height, 2)
        
        photoframe = draw.Frame(width=f'{width}cm', height=f'{height}cm',
                                anchortype='character')
        href = doc.addPicture(name)
        photoframe.addElement(draw.Image(href=href))
        p.addElement(photoframe)
        
    
        

driver = driver_factory()
parser = NoteBookParser(driver)
word_elements = parser.parse(notebook_path, notebook_name)
WordConverter(word_elements, notebook_path + notebook_name,
              front_path, back_path, out_name).convert()
os.system('rm picture_*.png')

    


    
    


        


        
