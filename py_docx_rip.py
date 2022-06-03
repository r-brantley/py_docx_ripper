#modules used for extracting information from word / xml documents
import os
import re
import xml.dom.minidom
import zipfile
from atlassian import confluence

# Set root for ripper
riproot = "/Users/ryanbrantley/Library/CloudStorage/OneDrive-SharedLibraries-OneMedNet/Compliance - Word Docs/"

# is it is a dotfile
def dotfile_file(file):
    if raw := re.match('^[.][\S]+$', file) is not None:
        return(True)

# is it docx file extension?
def docx_file(file):
    if raw := re.search('^[^.]+$|\.(?=docx)[^. \n\r]', file) is not None:
        return(True)

# function to get the path to the file to define breadcrumb taxonomy
def doc_tax(path):
    slices = []
    cookie = path.split("/")
    if (len(cookie) >= 1 and range(len(cookie))):
        for c in range(len(cookie)):
            slices.append(cookie[c])
    return(slices)

# pipe the XML content section of docx to string
def str_xml(file):
    docx = zipfile.ZipFile(file, mode="r")
    ripxml = xml.dom.minidom.parseString(docx.read('word/document.xml'))
    ripxml = str(docx.read('word/document.xml'))
    docx.close
    return ripxml
    
# pull the relevant groups and sections out of the docx XML
def nsx(doc):
        out = []
        scan = re.findall(r'(<w[^>]+(?:/>|>)(?:.*?)</w:t>)',doc)
        style_scan = re.compile(r'<w:pStyle w:val=\"([\S]+)\"')
        content_scan = re.compile(r'(?:<w:t(?: xml.*?\">|>))(.*?)</w:t>')

        for w in range(len(scan)):
            if style_scan.findall(scan[w]):    
                dx_style = style_scan.findall(scan[w])[0]
            else:
                dx_style = "default"

            if content_scan.findall(scan[w]):
                content = content_scan.findall(scan[w])[0]
                cdict = {
                    "style": dx_style,
                    "body": content
                }
                out.append(cdict)
                
        return out

def tree_printer(riproot):
    payload = []

    for root, dirs, files in os.walk(riproot): #walk the dir specified for docs
        for f in range(len(files)):
            file = files[f]

            if not dotfile_file(file): #skip dotfiles
                taxroot = root.split(riproot)[1] #capture taxonomy root of SOP
                
                if docx_file(file):
                    content = str_xml(root+"/"+file)
                    content = nsx(content)
                    rip_out = {
                        'file': file,
                        'root': root,
                        'taxonomy': taxroot,
                        'content': content
                    }
                    payload.append(rip_out)
                
                if not docx_file(file): #handle non-docx files that are not dotfiles
                    pass

    return payload

def uq(alist):
    out = []
    alist = (list(alist))
    for x in range(len(alist)):
       for y in range(len(alist[x]["content"])):
            if alist[x]["content"][y]["style"] not in out:
                out.append(str(alist[x]["content"][y]["style"]))
    return out


allofit = tree_printer(riproot)

print(*uq(allofit), sep="\n")