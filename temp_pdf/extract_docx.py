import zipfile
import sys
import xml.dom.minidom

def dump_xml(path):
    with zipfile.ZipFile(path) as z:
        xml_content = z.read('word/document.xml')
        dom = xml.dom.minidom.parseString(xml_content)
        with open('temp_pdf/doc_xml.txt', 'w', encoding='utf-8') as f:
            f.write(dom.toprettyxml(indent='  '))

if __name__ == '__main__':
    dump_xml(sys.argv[1])
