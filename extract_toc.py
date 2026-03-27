import zipfile
import xml.etree.ElementTree as ET
import sys

def extract_text(path):
    ns_w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    try:
        with zipfile.ZipFile(path) as docx:
            tree = ET.parse(docx.open('word/document.xml'))
            root = tree.getroot()
            return '\n'.join(''.join(node.text for node in p.iter(f'{ns_w}t') if node.text) for p in root.iter(f'{ns_w}p') if list(p.iter(f'{ns_w}t')))
    except Exception as e:
        return str(e)

if __name__ == '__main__':
    print(extract_text(sys.argv[1]))
