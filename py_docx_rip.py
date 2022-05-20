#modules used for extracting information from word / xml documents
import os, zipfile, re, subprocess
import xml.dom.minidom

#Set root for ripper
ripbase = "/Users/ryanbrantley/Library/CloudStorage/OneDrive-SharedLibraries-OneMedNet/Compliance - Word Docs"

print(ripbase)

bash_cmd = "tree "+ripbase
process = subprocess.run(bash_cmd, shell=True, check=True, executable='/bin/bash')