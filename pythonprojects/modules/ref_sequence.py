import os
import tempfile
import shutil
import zipfile
import re
from sys import argv

def createNewDocxFromOld(originalDocx, xmlContent, newFilename):
    tmpDir = tempfile.mkdtemp()
    zip = zipfile.ZipFile(open(originalDocx, "rb"))
    zip.extractall(tmpDir)
    with open(os.path.join(tmpDir, "word/document.xml"), "wb") as f:
        f.write(xmlContent)
    filenames = zip.namelist()
    zipCopyFilename = newFilename
    with zipfile.ZipFile(zipCopyFilename, "w", zipfile.ZIP_DEFLATED) as docx:
        for filename in filenames:
            docx.write(os.path.join(tmpDir, filename), filename)
    shutil.rmtree(tmpDir)

def get_word_xml(docx_file):
    with open(docx_file, mode='rb') as f:
        zip = zipfile.ZipFile(f)
        xml_content = zip.read('word/document.xml')
        xml_content = xml_content.decode('utf-8')
    return xml_content

def refreorder(docx_file):
    file = open(docx_file.lower().replace(".docx", "-report.xml"), "r")
    content = file.read()
    file.close()
    file_xml = get_word_xml(docx_file)
    file_xml = re.sub(r'(<w:p\W)', r"\n\1", file_xml)
    file_xml = re.sub(r'&lt;bib id="bib.*?&gt;', lambda m: re.sub(r"<[^<>]+>", "", m.group(0)), file_xml)
    file_xml = re.sub(r'&lt;number.*?/number&gt;', lambda m: re.sub(r"<[^<>]+>", "", m.group(0)), file_xml)
    contents = re.findall('<ref key="[^<>]+">(?:\n|\r|.)*?</ref>', content)
    for cont in contents:
        matches = re.findall(r'<matched>(.+?)</matched>', cont)
        con = re.search(r'<ref key="([^<>]+)">', cont)
        for match in matches:
            file_xml = file_xml.replace('w:anchor="' + match + '"', 'w:anchor="' + con.group(1) + '"')
            file_xml = file_xml.replace('w:name="' + match + '"', 'w:name="deletedbib"')
            file_xml = re.sub("<w:hyperlink.*?</w:hyperlink>", lambda match1: match1.group(0).replace("<w:t>" + match.replace("bib","") + "</w:t>",
                                                   "<w:t>" + con.group(1).replace("bib","") + "</w:t>") if 'w:anchor="bib' in match1.group(0) else match1.group(0), file_xml)
            file_xml = re.sub(r"<w:p[^a-zA-Z].*?</w:p>", lambda match1:"" if '&lt;bib id="' + match + '"' in match1.group(0) else match1.group(0), file_xml)

    base = 0
    previous = 0
    bookmarks = re.findall(r'w:name="bib([^<> ]*?)"', file_xml)
    for bookmark in bookmarks:
        match = re.match(r"^([0-9]+)([a-z]*?)$", bookmark)
        if match is not None:
            a = int(match.group(1))
            if not match.group(2).isalpha():
                previous = base + 100
                base = base + 100
            else:
                previous = previous + 1
        convert_number_to_string = lambda number:str(number // 100) if number % 100 == 0 else str(number // 100) + str(chr((number % 100) + 96))
        if not (match.group(0) == convert_number_to_string(previous)):
            file_xml = file_xml.replace('w:name="bib' + match.group(0) + '"',
                                        'w:name="bib' + convert_number_to_string(previous) + '"')
            file_xml = file_xml.replace('w:anchor="bib' + match.group(0) + '"',
                                        'w:anchor="bib' + convert_number_to_string(previous) + '"')
            file_xml = file_xml.replace('&lt;bib id="bib' + match.group(0) + '"',
                                        '&lt;bib id="bib' + convert_number_to_string(previous) + '"')
            file_xml = re.sub("<w:hyperlink.*?</w:hyperlink>", lambda match1:match1.group(0).replace("<w:t>" + match.group(0) + "</w:t>",
                                                   "<w:t>" + convert_number_to_string(previous) + "</w:t>") if 'w:anchor="bib' in match1.group(0) else match1.group(0), file_xml)

    matches = set(re.findall("bib([0-9]+)b", file_xml))
    single_matches = set(re.findall("bib([0-9]+)a", file_xml))
    for match in single_matches.difference(matches):
        file_xml = file_xml.replace('w:name="bib' + match + 'a"', 'w:name="bib' + match + '"')
        file_xml = file_xml.replace('w:anchor="bib' + match + 'a"', 'w:anchor="bib' + match + '"')
        file_xml = re.sub("<w:hyperlink.*?</w:hyperlink>", lambda match1: match1.group(0).replace("<w:t>" + match + "a</w:t>",
                                               "<w:t>" + match + "</w:t>") if 'w:anchor="bib' in match1.group(0) else match1.group(0), file_xml)
        file_xml = re.sub(r"<w:p[^a-zA-Z].*?</w:p>", lambda match1: "" if '&lt;bib id="bib' + match + '"' in match1.group(0) else match1.group(0).replace('&lt;bib id="bib' + match + 'a"', '&lt;bib id="bib' + match + '"') if '&lt;bib id="bib' + match + 'a"' in match1.group(0) else match1.group(0), file_xml)
    file_xml = re.sub('&lt;bib id="bib([0-9]+)([a-z]*?)"(.*?&gt;&lt;number&gt;)\W*?\w+\W*?(&lt;/number&gt;)',
                      lambda m: '&lt;bib id="bib'+m.group(1) +m.group(2)+'"'+m.group(3) + '[' + m.group(1) + ']'+m.group(4) if (
                                  m.group(1) + m.group(2)).isdigit() else '&lt;bib id="bib'+m.group(1)+ m.group(2)+'"' +m.group(3)  + m.group(2) + ')'+m.group(4) , file_xml)
    createNewDocxFromOld(docx_file, file_xml.encode(), docx_file.lower().replace(".docx", "-output.docx"))
    print("process done")

if __name__=="__main__":
    try:
        refreorder(argv[1]) if len(argv)>1 else print("No File exists")
    except Exception as e:
        print(e)