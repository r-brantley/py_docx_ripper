#modules used for extracting information from word / xml documents
import os
import re
import subprocess
import xml.dom.minidom
import zipfile

#Set root for ripper
ripbase = "/Users/ryanbrantley/Library/CloudStorage/OneDrive-SharedLibraries-OneMedNet/Compliance - Word Docs/"

print(ripbase)


### test tree bash run for loops later
    #bash_cmd = "tree "+ripbase
    #process = subprocess.run(bash_cmd, shell=True, check=True, executable='/bin/bash')

## POC for Word docx drill down into XML content
doc = ripbase+"0XX Corporate WIP/SOP-015 Business Continuity/SOP-015 Business Continuity.docx"

docz = zipfile.ZipFile(doc, mode="r")
docz.namelist()

ripxml = xml.dom.minidom.parseString(docz.read('word/document.xml')).toprettyxml(indent='  ')
ripxml = str(docz.read('word/document.xml'))
docz.close

raw = re.findall('<w:t(?:>| xml:space="preserve">)(.*?)<.+?>', ripxml)

print(raw)