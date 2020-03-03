import os
import tempfile
import shutil
import zipfile


def createNewDocxFromOld(originalDocx, xmlContent, newFilename):
    tmpDir = tempfile.mkdtemp()
    zip = zipfile.ZipFile(open(originalDocx, "rb"))
    zip.extractall(tmpDir)
    with open(os.path.join(tmpDir, "word/document.xml"), "wb") as f:
        f.write(xmlContent)
    filenames = zip.namelist()
    zipCopyFilename = newFilename
    with zipfile.ZipFile(zipCopyFilename, "w") as docx:
        for filename in filenames:
            docx.write(os.path.join(tmpDir, filename), filename)
    shutil.rmtree(tmpDir)

def get_word_xml(docx_file):
    with open(docx_file, mode='rb') as f:
        zip = zipfile.ZipFile(f)
        xml_content = zip.read('word/document.xml')
        xml_content = xml_content.decode('utf-8')
    return xml_content
