from lxml import etree
from lxml import html
from pylatexenc.latexencode import UnicodeToLatexEncoder, UnicodeToLatexConversionRule, RULE_REGEX
from pylatexenc.latexencode import unicode_to_latex
from pylatexenc.latexwalker import LatexWalker, LatexEnvironmentNode, LatexMacroNode, LatexCharsNode, LatexMathNode, LatexGroupNode, LatexCommentNode, LatexSpecialsNode
from pylatexenc.latex2text import LatexNodes2Text
from pymsgbox import alert
from bs4 import BeautifulSoup as bs
import os
import re
import traceback
import shutil
import win32com.client
import zipfile
import chardet
import time
import configparser
from pymsgbox import alert

curDir = os.getcwd()

global serverIP
serverIP=''
if os.path.isfile(os.path.join(curDir,'ServerDetails.exe')):
    serverIP=os.popen(os.path.join(curDir,'ServerDetails.exe')).read().strip()
    if (os.path.isfile(r"\\"+str(serverIP)+r"\License\license.txt")):
        chklicense = open(r"\\"+str(serverIP)+r"\License\license.txt", 'r').read()
        if (chklicense != 'Active'):
            alert(text='Please contact the tech support!', title='expired', button='OK')
            exit()
    else:
        alert(text="Please check the \"internet\" or \"VPN\" connection", title='Expired', button='OK')
        exit()
else:
    alert(text="\"ServerDetails.exe\" file missing...", title='Missing', button='OK')
    exit()

# Create a ConfigParser object
config = configparser.ConfigParser()

# Create a ConfigParser object with case sensitivity
config.optionxform = str  # Preserve the original case

# Read the INI file

if os.path.isfile('//192.168.7.5/SoftwareTools/Journals/Config/LaTeXUnicode.ini'):
    pass
else:
    alert(text="\"LaTeXUnicode.ini\" file missing...", title='Missing', button='OK')
    exit()

config.read('//192.168.7.5/SoftwareTools/Journals/Config/LaTeXUnicode.ini')

ChapterNum = "1"

def TagReplacement(arg1, arg2, arg3):

    arg3 = arg3.split(",")

    for each_command in arg1.xpath(arg2):

        if each_command.text:
            each_command.text = arg3[0] + each_command.text 
        else:
            each_command.text = arg3[0]

        if each_command.tail:
            each_command.tail = arg3[1] + each_command.tail 
        else:
            each_command.tail = arg3[1] 

        each_command.tag = "del-strip-tag"

def TrackConversion(htmlFileCnt,htmlpath):

    '''html file convert into latex content'''

    try:

        testFileName = os.path.join(os.path.split(htmlpath)[0], "out.tex")

        

        htmlFileCnt = htmlFileCnt.encode()

        htmlCnt = html.fromstring(htmlFileCnt)

        # Find the <head> element
        head_element = htmlCnt.xpath("//head")

        # Remove the <head> element and its contents
        if head_element:
            parent = head_element[0].getparent()
            if parent is not None:
                parent.remove(head_element[0])

        # Find all <p> elements
        for p_element in htmlCnt.xpath("//p"):
            # Check if the <p> element is empty (no text, no children)
            if not p_element.text and len(p_element) == 0:
                parent = p_element.getparent()
                if parent is not None:
                    parent.remove(p_element)

        listDict = {}

        # Append list Dictionary through config file
        for each_section in config.sections():
            if each_section == "ListStyles":
                for each_key, each_val in config.items(each_section):
                    listDict[each_key] = each_val

        listCount = 1

        closingList = []

        firstListCount = ""

        for idx, each_para in enumerate(htmlCnt.xpath("//p"), 1):
            if r"class" in each_para.attrib:
                listType = each_para.get(r'class')
                
                if listType in listDict:
                    if listCount == 1:
                        firstListCount = listType
                        if each_para.text:
                            each_para.text = listDict[listType].split(",")[0] + "/custom-new-line/\\item " + each_para.text
                        else:
                            each_para.text = listDict[listType].split(",")[0] + "/custom-new-line/\\item "

                        closingList.append(listDict[listType].split(",")[1])
                        
                    elif listType != firstListCount:
                        firstListCount = listType
                        if each_para.text:
                            each_para.text = listDict[listType].split(",")[0] + "/custom-new-line/\\item " + each_para.text
                        else:
                            each_para.text = listDict[listType].split(",")[0] + "/custom-new-line/\\item "
                        closingList.append(listDict[listType].split(",")[1])
                    else:
                        if each_para.text:
                            each_para.text = "/custom-new-line/\\item " + each_para.text
                        else:
                            each_para.text = "/custom-new-line/\\item "

                    if each_para.getnext().get("class") not in listDict:
                        if each_para.tail is not None:
                            each_para.tail = each_para.tail + "\n\n".join(eachEnd for eachEnd in reversed(closingList)) + "\n\n"
                        else:
                            each_para.tail = "\n\n".join(eachEnd for eachEnd in reversed(closingList)) + "\n\n"
                        listCount = 0
                        closingList.clear()
                    
                    listCount = listCount + 1


        # Append list Dictionary through config file
        for each_section in config.sections():
            if each_section == "CommonStyles":
                for each_key, each_val in config.items(each_section):
                    if "/" in each_key:
                        each_key_parts = each_key.split("/")
                        xpath_expr = f"//{each_key_parts[0]}[@{each_key_parts[1]}='{each_key_parts[2]}']"
                        TagReplacement(htmlCnt, xpath_expr, each_val)
                    else:
                        xpath_expr = f"//{each_key}"
                        TagReplacement(htmlCnt, xpath_expr, each_val)
        
        #delete style span element
        for each_span in htmlCnt.xpath("//span"):
            if r'style' in each_span.attrib:
                if r"Courier" in  each_span.get("style"):
                    if each_span.text is not None:
                        each_span.text = r"\texttt{" + each_span.text
                    else:
                        each_span.text = r"\texttt{"

                    if each_span.tail is not None:
                        each_span.tail = each_span.tail + r"}"
                    else:
                        each_span.tail =  r"}"
                    
                    each_span.tag = "del-strip-tag"

        #delete style span element
        for each_span in htmlCnt.xpath("//span[@class='MsoCommentReference']"):
            each_span.tag = "del-comment-span"

        # Figure Styles
        for findPstyle in htmlCnt.xpath("//p[@class='FigCaption']"):
            

            ImageTop = r"\begin{figure}[!tbhp]"
            ImageTag = findPstyle.getprevious()

            ToolTip = findPstyle.getnext()
            ToolTipText = ""
            
            # Have a problem
            if ToolTip is not None:
                for tooltip in ToolTip.findall(".//tr/td"):
                    ToolTipText = tooltip.xpath("string()")
                    ToolTip.tag = "del-table"

            FigCaptionText = findPstyle.xpath('string()')
            FigCaptionText = re.sub(r"\\textbf\{(Figure|Fig.|Fig|Figure.)\s*([0-9\.]+)(.*?)\}(\s*)", r"\\label{\g<1>:\g<2>}", FigCaptionText, flags= re.S)
            FigCaptionText = r"\caption{" + FigCaptionText + "}"

            if ImageTag is not None:
                if r"class" in ImageTag.attrib and ImageTag.get("class") == "Image":
                    if ImageTag.find(".//img") is not None:
                        if r"src" in ImageTag.find(".//img").attrib:
                            ImageTag.text = ImageTop + "/custom-new-line/" + r"\includegraphics{" + ImageTag.find(".//img").get("src") + r"}/custom-new-line/" + FigCaptionText + "/custom-new-line/" + r"\alttext{" + ToolTipText.strip() + r"}/custom-new-line/" + r"\end{figure}" + "\n\n" 
                            ImageTag.find(".//img").tag = "del-strip-tag"
                            ImageTag.tag = "del-strip-tag"
                    else:
                        ImageTag.clear()
                        ImageTag.text = ImageTop + "/custom-new-line/" + r"\includegraphics{example-image-a}" + "/custom-new-line/" + FigCaptionText + "/custom-new-line/" + r"\alttext{" + ToolTipText.strip() + r"}/custom-new-line/" + r"\end{figure}" + "\n\n"
                        ImageTag.tag = "del-strip-tag"
            
            findPstyle.tag = "del-p"


        # Table Styles
        for findPstyle in htmlCnt.xpath("//p[@class='Tablecaption']"):

            TableCnt = ""
            colcount = ""

            TableTop = r"\begin{table}[!tbhp]" + "/custom-new-line/"
            TableCnt += TableTop
            
            TableDivTag = findPstyle.getnext()

            ToolTipText = ""            

            TableCaptionText = findPstyle.xpath('string()')
            TableCaptionText = re.sub(r"\\textbf\{(Table|Tab.|Tab|Table.)\s*([0-9\.]+)(.*?)\}(\s*)", r"\\label{\g<1>:\g<2>}", TableCaptionText, flags= re.S)
            TableCaptionText = r"\caption{" + TableCaptionText + "}"

            # Tooltip text capture
            if findPstyle.getnext().tag == "div":
                
                if TableDivTag.getnext().tag == "table":

                    for tooltip in TableDivTag.getnext().findall(".//tr/td"):
                        ToolTipText = tooltip.xpath("string()")
                        TableDivTag.getnext().getnext().tag = "del-table"
                
                elif TableDivTag.getnext().getnext().tag == "table":

                    for tooltip in TableDivTag.getnext().getnext().findall(".//tr/td"):
                        ToolTipText = tooltip.xpath("string()")
                        TableDivTag.getnext().getnext().tag = "del-table"


                TableTag = TableDivTag.find(".//table")

                tabRow = 1

                TableBody = ""

                for eachRow in TableTag.findall(".//tr"):

                    if tabRow == 1:
                        colcount = len(eachRow.findall(".//td"))

                    InnerTDCount = 1
                    for eachEntry in eachRow.findall(".//td"):
                        eachEntryCnt = str(eachEntry.xpath("string()"))
                        eachEntryCnt = eachEntryCnt.strip()
                        TableBody += eachEntryCnt
                        if InnerTDCount == colcount:
                            TableBody += r"\\" + "/custom-new-line/"
                        else:
                            TableBody += " " + " /tableAmbersand/ " + " " 

                        InnerTDCount = InnerTDCount + 1
                    
                    tabRow = tabRow + 1
                
            findPstyle.clear()
            findPstyle.text = TableTop + TableCaptionText + "/custom-new-line/" + r"\begin{tabular*}{\textwidth}{" + str("l" * colcount) + "}/custom-new-line/" + TableBody + r"\end{tabular*}/custom-new-line/" + r"\end{table}" + "/custom-new-line//custom-new-line/"

            TableTag.tag = "del-table"
            findPstyle.tag = "del-strip-p"


        # Individual Table
        IndividualTabRow = 1

        for eachTable in htmlCnt.findall(".//table"):
            IndividualTableBody = ""
            for eachRow in eachTable.findall(".//tr"):

                if IndividualTabRow == 1:
                    colcount = len(eachRow.findall(".//td"))

                InnerTDCount = 1
                for eachEntry in eachRow.findall(".//td"):
                    eachEntryCnt = str(eachEntry.xpath("string()"))
                    eachEntryCnt = eachEntryCnt.strip()
                    IndividualTableBody += eachEntryCnt
                    if InnerTDCount == colcount:
                        IndividualTableBody += r"\\" + "/custom-new-line/"
                    else:
                        IndividualTableBody += " " + " /tableAmbersand/ " + " " 

        
                    InnerTDCount = InnerTDCount + 1
            
                IndividualTabRow = IndividualTabRow + 1

            # print(html.tostring(eachTable))    
            eachTable.clear()
            eachTable.text =  r"\begin{table}[!tpbh]/custom-new-line/{\begin{tabular*}{\textwidth}{" + str("l" * colcount) + "}/custom-new-line/" + IndividualTableBody + "CSK" + r"\end{tabular*}}{}/custom-new-line/" + r"\end{table}" + "/custom-new-line//custom-new-line/"
            # print(html.tostring(eachTable))
            # eachTable.tag = "del-table"

       

        # footnote Processing
        for eachCommentSpan in htmlCnt.xpath("//a"):
                if r"name" in eachCommentSpan.attrib:
                    if r"ft" in eachCommentSpan.get("name"):
                        searchID = eachCommentSpan.get("name")
                        searchTerm = r".//a[@href=" + r'"' + "#" + searchID + r'"]'

                        if htmlCnt.find(searchTerm) is not None:
                            if htmlCnt.find(searchTerm).getparent().tag == "p":
                        
                                paraTag = htmlCnt.find(searchTerm).getparent()
                                
                                htmlCnt.find(searchTerm).tag = "del-strip-tag"
                                eachCommentSpan.insert(0, paraTag)
                                
                                paraTag.tag = "del-strip-tag"

        # Delete the footnote reference number
        for eachspan in htmlCnt.xpath(".//span[@class='MsoFootnoteReference']"):
            eachspan.drop_tree()
        
        #href link Process
        escape_char = {r"&":r"\&", 
                       r"%":r"\%",
                       r"$":r"\$", 
                       r"#":r"\#", 
                       r"_":r"\_",
                       r"_": r"\_",
                       r"{":r"\{",
                       r"}":r"\}",
                       r"~":r"\~",
                       r"^":r"\^",
                       r"\\": r"\\backslash"}
        
        for aLink in htmlCnt.xpath("//a"):

            linkStore = str(aLink.xpath("string()"))

            for key,val in escape_char.items():
                if key in linkStore:
                    linkStore = re.sub(key, val, linkStore, flags=re.S)
                else:
                    pass

            if r"href" in aLink.attrib:
                if aLink.text:
                    aLink.text = r"\href{" + aLink.get("href") + r"}{" + linkStore + "}"
                    aLink.tag = "del-strip-tag"
                else:
                    if r"style" in aLink.attrib:
                        if r"footnote" in aLink.get("style"):
                            pass
                        else:
                            linkStore = aLink.xpath("string()")

                            for key,val in escape_char.items():
                                if key in linkStore:
                                    linkStore = re.sub(key, val, linkStore, flags=re.S)
                                else:
                                    pass

                            aLink.text = r"\href{" + aLink.get("href") + r"}{" + linkStore + "}"
                            aLink.tag = "del-strip-tag"
            
            elif r"style" in aLink.attrib:
                aLink.tag = "del-strip-tag"

        # Find all comment nodes and remove them
        for comment in htmlCnt.xpath("//comment()"):
            parent = comment.getparent()
            if parent is not None:
                parent.remove(comment)

        # Find all Reference Nodes
        refCount = 1

        for each_section in config.sections():
            if each_section == "RefStyles":
                for each_key, each_val in config.items(each_section):
                    if r"/" in each_key:
                        each_key = each_key.split("/")
                        each_val = each_val.split(",")

                        totRefCount = len(htmlCnt.xpath("//" + each_key[0] + "[@" +  each_key[1] + "='" + each_key[2] + "']"))

                        for eachParaRef in htmlCnt.xpath("//" + each_key[0] + "[@" +  each_key[1] + "='" + each_key[2] + "']"):

                            if int(refCount) < 10:
                                refCountPadding = "000" + str(refCount)
                            elif int(refCount) > 10 and int(refCount) < 100:
                                refCountPadding = "00" + str(refCount)
                            elif int(refCount) > 100 and int(refCount) < 1000:
                                refCountPadding = "0" + str(refCount)
                            else:
                                pass

                            if eachParaRef.text:

                                if refCount == 1:
                                    eachParaRef.text = r"\begin{thebibliography}{99}" + "\n\n" + each_val[0] + "Chap:" + ChapterNum + ":" + str(refCountPadding) + each_val[1] + eachParaRef.text
                                else:
                                    eachParaRef.text = each_val[0] + "Chap:" + ChapterNum + ":" + str(refCountPadding) + each_val[1] + eachParaRef.text

                            else:


                                if refCount == 1:
                                    eachParaRef.text = r"\begin{thebibliography}{99}" + "\n\n" + each_val[0] + str(refCountPadding) + each_val[1]
                                else:
                                    eachParaRef.text = each_val[0] + str(refCountPadding) + each_val[1]

                            if refCount == totRefCount:
                                if eachParaRef is not None:
                                    if eachParaRef.tail is not None:
                                        eachParaRef.tail = "\n\n" + r"\end{thebibliography}" + eachParaRef.tail
                                    else:
                                        eachParaRef.tail = "\n\n" + r"\end{thebibliography}"

                            eachParaRef.tag = "del-strip-tag"

                            refCount = refCount + 1

        for each_span in htmlCnt.xpath("//span"):
            each_span.tag = "del-strip-span"

        for each_footnote in htmlCnt.xpath("//a"):
            if r"style" in each_footnote.attrib:
                if r"mso-footnote" in each_footnote.get("style"):
                    
                    if each_footnote.text is not None:
                        each_footnote.text = r"\footnote{" + each_footnote.text
                    else:
                        each_footnote.text = r"\footnote{" 


                    if each_footnote.tail is not None:
                        each_footnote.tail = r"}" + each_footnote.tail
                    else:
                        each_footnote.tail = r"}" 

                    each_footnote.tag = "del-footnote-strip-a"

        # Equation conversion
        for tooltip in htmlCnt.xpath("//p[@class='Equation']"):
            mathcnt = tooltip.xpath("string()")
            mathcnt = mathcnt.replace(r"&", r"/mathambersand/")
            mathcnt = mathcnt.replace("\n", r"/mathnewline/")
            tooltip.clear()
            tooltip.text = mathcnt


        for all_div in htmlCnt.xpath("//div"):
            
            if r"style" in all_div.attrib:
                if r"footnote-list" in all_div.attrib["style"]:
                    all_div.tag = "del-div"

                if r"comment-list" in all_div.attrib["style"]:
                    all_div.tag = "del-div"

            if r"class" in all_div.attrib:
                if r"WordSection" in all_div.attrib["class"]:
                    all_div.tag = "del-strip-div"

        for eachshapes in htmlCnt.xpath("//shape"):
            if eachshapes.find(".//imagedata") is not None:
                if r"src" in eachshapes.find(".//imagedata").attrib:
                    if eachshapes.find(".//imagedata").text is not None:
                        eachshapes.find(".//imagedata").text = r"\includegraphics{" + eachshapes.find(".//imagedata").get("src") + r"}" + eachshapes.find(".//imagedata").text
                    else:
                        eachshapes.find(".//imagedata").text = r"\includegraphics{" + eachshapes.find(".//imagedata").get("src") + r"}"

                    eachshapes.find(".//imagedata").tag = "del-strip-tag"

            eachshapes.tag = "del-strip-tag"



        tags_to_remove = ["del-table", "del-p", "del-comment-span", "del-span", "del-div"]

        tags_to_strip = ["del-strip-span", "del-footnote-strip-a", "del-strip-p", "del-strip-div", "del-strip-tag"]
    
        for tag in tags_to_strip:
            for element in htmlCnt.xpath('//' + tag):
                element.drop_tag()

        for tag in tags_to_remove:
            for element in htmlCnt.xpath('//' + tag):
                element.drop_tree()

        # all tags stripped in body content
        for element in htmlCnt.xpath('//html/body/*'):
            element.drop_tag()

        for each_section in config.sections():
            if each_section == "texpreamble":
                for each_key, each_val in config.items(each_section):
                    if each_key == "preambletop":
                        
                        if htmlCnt.find(".//body") is not None:
                            if htmlCnt.find(".//body").text is not None:
                                htmlCnt.find(".//body").text = each_val + htmlCnt.find(".//body").text
                            else:
                                htmlCnt.find(".//body").text = each_val

                    if each_key == "preamblebottom":
                        
                        if htmlCnt.find(".//body") is not None:
                            if htmlCnt.find(".//body").tail:
                                htmlCnt.find(".//body").tail = each_val + htmlCnt.find(".//body").tail
                            else:
                                htmlCnt.find(".//body").tail = each_val
                    
                            htmlCnt.find(".//body").drop_tag()


        with open(htmlpath, "w", encoding="utf-8") as f1:
            
            tot_cnt = html.tostring(htmlCnt, pretty_print=True, method="html", encoding="utf-8").decode()
            # htmlCnt = etree.tostring(htmlCnt, encoding="utf-8").decode()
            # tot_cnt = htmlCnt
            
            # Build the conversion rules from the INI file
            conversion_rules = [
                (re.compile(each_key), each_val) 
                for each_section in config.sections() if each_section == "LatexUnicodes" 
                for each_key, each_val in config.items(each_section)
            ]

            # Create the UnicodeToLatexEncoder with the custom conversion rules
            u = UnicodeToLatexEncoder(
                conversion_rules=[
                    UnicodeToLatexConversionRule(rule_type=RULE_REGEX, rule=conversion_rules),
                    'defaults'
                ]
            )

            tot_cnt = ''.join(list(map(lambda x:  u.unicode_to_latex(x) if ord(x) >= 127 else x, tot_cnt)))

            findCnt = re.findall(r"((\\chapter|\\section|\\subsection|\\subsubsection|\\paragraph|\\subparagraph)\{([0-9\.]+|[A-Za-z]+)\s*)", tot_cnt,flags=re.S)
            
            for eachsec in findCnt:
                secFlag = ""
                secCnt = ""
                if re.search(r"([0-9\.]+)", eachsec[0]):
                    secCnt = eachsec[1] + secFlag + r"{"
                    tot_cnt = tot_cnt.replace(eachsec[0],secCnt)
                else:
                    secFlag = "*"
                    secCnt = eachsec[1] + secFlag + r"{" + eachsec[2]
                    tot_cnt = tot_cnt.replace(eachsec[0],secCnt)

            tot_cnt = tot_cnt.replace(r"%",r"\%")

            # Replace single newline characters with a space, but leave double newlines intact
            tot_cnt = re.sub(r'(?<!\n)\n(?!\n)', ' ', tot_cnt)
            
            tot_cnt = re.sub(r'&amp;', r' \& ', tot_cnt)
            
            tot_cnt = re.sub(r'/custom-new-line/', '\n', tot_cnt)
            tot_cnt = re.sub(r'/mathambersand/', r'&', tot_cnt)
            tot_cnt = re.sub(r'\\item\s*(\(|)([A-Za-z0-9\.]+)(\)|)(\s*|~{1,})', r'\\item ', tot_cnt)
            tot_cnt = re.sub(r'/preamble-percent/', r'%', tot_cnt)
            
            tot_cnt = tot_cnt.replace(r"/tableAmbersand/",r"&")
            tot_cnt = tot_cnt.replace(r"/mathnewline/","\n")
            tot_cnt = tot_cnt.replace(r"$",r"\$")
            tot_cnt = tot_cnt.replace(r"&gt;",r">")
            tot_cnt = tot_cnt.replace(r"&lt;",r"<")

            tot_cnt = re.sub(r'<(/html|html)([^<>]+|)>', r'', tot_cnt)
            
            f1.write(tot_cnt)
        
    except Exception as err:
         print(str(err) + traceback.format_exc())



def DocConversion(docpath):

    output_format = os.path.join(str(os.path.split(docpath)[0]), str(os.path.splitext(os.path.split(docpath)[1])[0]) + r".html") 
   
    doc = win32com.client.GetObject(docpath)
    doc.SaveAs (FileName=output_format, FileFormat=8)
    doc.Close ()

    if os.path.isfile(output_format):

        # alert(text="HTML file converted successfully. ", title="Message Box", button="Ok")
    
        with open(output_format, "rb") as f1:
            binary_content = f1.read()

        # Detect encoding (often Windows-1252 or similar)
        try:
            decoded_content = binary_content.decode('windows-1252')  # Adjust based on actual encoding
        except UnicodeDecodeError:
            decoded_content = binary_content.decode('latin1')

        # Remove unwanted newline characters (if any) that may have been introduced
        # decoded_content = decoded_content.replace('\r\n', '   ')

        # Ensure the HTML has the correct UTF-8 meta tag
        if '<meta charset="UTF-8">' not in decoded_content:
            decoded_content = decoded_content.replace('<head>', '<head>\n<meta charset="UTF-8">')

        with open(output_format, "w", encoding="utf-8", newline='') as f1:
            f1.write(decoded_content)
        
        with open(output_format, "r", encoding="utf-8", errors='ignore') as f1:
            htmlcnt = f1.read()
            htmlPath = os.path.join(os.path.split(output_format)[0], os.path.splitext(os.path.split(output_format)[1])[0] + "_out.tex")
            alert(text="\"Track PDF LaTeX file created.\" The file placed following path " + htmlPath, title='Info', button='OK')
            conversionStart = TrackConversion(htmlcnt,htmlPath)

# Proof PDF TeX Generation Start

def detect_encoding(file_path):
   with open(file_path, 'rb') as f:
      raw_data = f.read()
      result = chardet.detect(raw_data)
      return result['encoding']


def read_file(file_path):
   encoding = detect_encoding(file_path)
   with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
      return [f.read(), encoding]

def LatexWalkerIntialization(inCnt, treePath, encodingVlaue):
        
    try:

        walker = LatexWalker(inCnt, tolerant_parsing=False)

        # Get all nodes from the document
        (nodes, pos, len_) = walker.get_latex_nodes(pos=0)

        collect_cnt = ""
        for node in nodes:
            collect_cnt += str(node)
            collect_cnt += "\n\n"

        with open(treePath, "w", encoding=encodingVlaue) as f1:
            f1.write(collect_cnt)

        return nodes

    except Exception as err:
        print('LaTeX Walker Validation Error ' + str(err) + str(traceback.format_exc()))
        return 'LaTeX Walker Validation Error ' + str(err) + str(traceback.format_exc())


def ProofPDFConversion(texpath):
    
    fileCnt = read_file(texpath)
    treePath = os.path.splitext(texpath)[0] + r"_tree.txt"
    proofTeXPath = os.path.splitext(texpath)[0] + r"_final.tex"

    fileCnt[0] = fileCnt[0].replace(r"\begin{document}", "/StartDocument/")
    fileCnt[0] = fileCnt[0].replace(r"\end{document}", "/EndDocument/")

    latexCnt = fileCnt[0]

    walkerNodes = LatexWalkerIntialization(fileCnt[0], treePath, fileCnt[1])
    
    with open(treePath, "r", encoding=fileCnt[1], errors="ignore") as f1:
        treeCnt = f1.read()
        
    pattern = r"LatexMacroNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), macroname='(DIFadd|DIFdel)',"

    matches = re.findall(pattern, treeCnt, flags=re.S)

    delCount = 1
    for eachMatch in reversed(matches):
        if eachMatch[3] == "DIFdel":
            st = int(eachMatch[1])
            en = int(eachMatch[1]) + int(eachMatch[2])
            str_cnt = latexCnt[st: en]

            searchGroup = re.search(r"LatexGroupNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=" + str(en) + r", len=([0-9]+),", treeCnt, flags=re.S)

            # Check if the search was successful
            if searchGroup:
                # Extract the first group
                startGroupNode = en
                endGroupNode = startGroupNode + int(searchGroup.group(2))

                grb_cnt = latexCnt[startGroupNode: endGroupNode]

                latexCnt = latexCnt[:st] + latexCnt[endGroupNode:]

        elif eachMatch[3] == "DIFadd":
            st = int(eachMatch[1])
            en = int(eachMatch[1]) + int(eachMatch[2])
            str_cnt = latexCnt[st: en]

            searchGroup = re.search(r"LatexGroupNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=" + str(en) + r", len=([0-9]+),", treeCnt, flags=re.S)

            # Check if the search was successful
            if searchGroup:
                # Extract the first group
                startGroupNode = en
                endGroupNode = startGroupNode + int(searchGroup.group(2))

                grb_cnt = latexCnt[startGroupNode: endGroupNode]
                grb_cnt = grb_cnt[1:]
                grb_cnt = grb_cnt[:-1]

                latexCnt = latexCnt[:st] + grb_cnt + latexCnt[endGroupNode:]

    with open(proofTeXPath, "w", encoding=fileCnt[1], errors="ignore") as f1:
        
        latexCnt = latexCnt.replace(r"/StartDocument/", r"\begin{document}")
        latexCnt = latexCnt.replace(r"/EndDocument/", r"\end{document}")

        if os.path.isfile(treePath):
            os.remove(treePath)

        alert(text="\"Track PDF LaTeX file created.\" The file placed following path " + proofTeXPath, title='Info', button='OK')

        f1.write(latexCnt)
            

if __name__ == "__main__":

    try:
        User_Input_One = input("Are you want to create trackpdf or proofpdf: ")

        if User_Input_One == "trackpdf":
            User_Input = os.path.abspath(input("Enter the DOC file path with file name and extension: "))
            docConvert = DocConversion(User_Input)
        else:
            User_Input = os.path.abspath(input("Enter the TeX file path with file name and extension: "))
            docConvert = ProofPDFConversion(User_Input)
    
    except Exception as err:
        alert(text="\"Main funtion error.\"" + str(err) + traceback.format_exc(), title='Info', button='OK')

    # User_Input = os.path.abspath(input("Enter the TeX file path with file name and extension: "))
    # docConvert = ProofPDFConversion(User_Input)

    # User_Input = os.path.abspath(input("Enter the DOC file path with file name and extension: "))
    # docConvert = DocConversion(User_Input)

