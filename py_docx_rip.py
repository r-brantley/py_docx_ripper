#modules used for extracting information from word / xml documents
import os
import re
import xml.dom.minidom
import zipfile

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
    cookie = path.split("/")
    if (len(cookie) >= 1 and range(len(cookie))):
        slices = []
        for c in range(len(cookie)):
            crumb = cookie[c]
            slices.append(crumb)
        return(slices)

# pipe the XML content section of docx to string
def str_xml(file):
    docx = zipfile.ZipFile(file, mode="r")
    ripxml = xml.dom.minidom.parseString(docx.read('word/document.xml'))
    ripxml = str(docx.read('word/document.xml'))
    docx.close
    return ripxml
    
# pull the relevant groups and sections out of the docx XML
def docx_content(docx):
    docx_sect = []
    praw = re.findall(r"<w:pStyle w:val=\"([\S]+)\"\/>[.\s\S]*?(<w:(?:t|t xml:[.\S]+)>(?:[.\s\S]+?)<\/w:p>)", docx)
    for row in praw:
        tag = row[0]
        for val in row:
            nest = re.findall(r"(?:<w:(?:t|t xml:[.\S]+)>([.\s\S]+?)<\/w:t>)",val)
            s = ' '.join(nest)
        docx_sect.append([tag, s])
    return docx_sect

def tree_printer(riproot):
    payload = []
    for root, dirs, files in os.walk(riproot):
        for f in range(len(files)):
            file = files[f]
            if not dotfile_file(file): ### skip dotfiles
                taxroot = root.split(riproot)[1] #capture raw taxonomy root of SOP
                if docx_file(file): ### what do we do to docx files??
                    content = str_xml(root+"/"+file)
                    content = docx_content(content)
                    rip_out = {
                        'file': file,
                        'taxonomy': taxroot,
                        'content': content
                    }
                    payload.append(rip_out)
                    print("BING! BONG!")
                if not docx_file(file):
                    print("DO THE OTHER FILE HANDLER FOR:"+file)
                    #do pdf and vsd attachment
    
    return payload

files = tree_printer(riproot)
print(files)