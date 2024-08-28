from lxml import etree
from lxml import html
from pylatexenc.latexencode import unicode_to_latex
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
from pylatexenc.latexencode import UnicodeToLatexEncoder, UnicodeToLatexConversionRule, RULE_REGEX
import configparser
# import wx
# from wx.core import EVT_SLIDER

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

        each_command.drop_tag()

def TrackConversion(htmlFileCnt,htmlpath):

    '''html file convert into latex content'''

    try:

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
        previousList = ""
        for idx, each_para in enumerate(htmlCnt.xpath("//p")):
            if r"class" in each_para.attrib:
                listType = each_para.get(r'class')
                if listType in listDict:
                    previousList = str(listDict[listType].split(",")[1]) + "\n\n"
                    if listCount == 1:
                        if each_para.text:
                            each_para.text = listDict[listType].split(",")[0] + "\n\\item " + each_para.text
                        else:
                            each_para.text = listDict[listType].split(",")[0] + "\n\\item "
                    else:
                        if each_para.text:
                            each_para.text =  "\\item " + each_para.text
                        else:
                            each_para.text = "\\item "

                    listCount = listCount + 1
                else:
                    listCount = 1
                    endList = htmlCnt.xpath("//p")[idx-1]
                    if r'class' in endList.attrib:
                        listType = endList.get(r'class')
                        if listType in listDict:
                            if endList.tail:
                                endList.tail = endList.tail + previousList
                            else:
                                endList.tail = previousList

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
                        each_span.text = r"\texttt{" + each_span.text.strip()
                    else:
                        each_span.text = r"\texttt{"

                    if each_span.tail is not None:
                        each_span.tail = each_span.tail + r"}"
                    else:
                        each_span.tail =  r"}"
                    
                    each_span.drop_tag()

        #delete style span element
        for each_span in htmlCnt.xpath("//span"):
            each_span.drop_tag()

        # Figure Styles
        for findPstyle in htmlCnt.xpath("//p[@class='FigCaption']"):

            ImageTop = r"\begin{figure}[!tbhp]"
            ImageTag = findPstyle.getprevious()

            ToolTip = findPstyle.getnext()
            ToolTipText = ""
            for tooltip in ToolTip.findall(".//tr/td"):
                ToolTipText = tooltip.xpath("string()")

            FigCaptionText = findPstyle.xpath('string()')
            FigCaptionText = re.sub(r"\\textbf\{(Figure|Fig.|Fig|Figure.)\s*([0-9\.]+)(.*?)\}(\s*)", r"\\label{\g<1>:\g<2>}", FigCaptionText, flags= re.S)
            FigCaptionText = r"\caption{" + FigCaptionText + "}"
            

            if r"class" in ImageTag.attrib and ImageTag.get("class") == "Image":
                if ImageTag.find(".//img") is not None:
                    if r"src" in ImageTag.find(".//img").attrib:
                        ImageTag.text = ImageTop + "\n" + r"\tooltip{\includegraphics{" + ImageTag.find(".//img").get("src") + r"}}{" + ToolTipText + "}\n" + FigCaptionText + "\n" + r"\end{figure}" + "\n\n" 
                        ImageTag.find(".//img").drop_tag()
                        ImageTag.drop_tag()

            ToolTip.tag = "del-table"
            findPstyle.tag = "del-p"


        # Table Styles
        for findPstyle in htmlCnt.xpath("//p[@class='Tablecaption']"):

            TableCnt = ""
            colcount = ""

            TableTop = r"\begin{table}[!tbhp]" + "\n"
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
                            TableBody += r"\\" + "\n"
                        else:
                            TableBody += " " + " /tableAmbersand/ " + " " 

                        InnerTDCount = InnerTDCount + 1
                    
                    tabRow = tabRow + 1
                
            findPstyle.clear()
            findPstyle.text = TableTop + TableCaptionText + "\n" + r"\tooltip{\begin{tabular*}{\textwidth}{" + str("l" * colcount) + "}\n" + TableBody + r"\end{tabular*}}{" + ToolTipText + "}\n" + r"\end{table}" + "\n\n"

            TableTag.tag = "del-table"

        # footnote Processing
        for eachCommentSpan in htmlCnt.xpath("//a"):
                if r"name" in eachCommentSpan.attrib:
                    if r"ft" in eachCommentSpan.get("name"):
                        searchID = eachCommentSpan.get("name")
                        searchTerm = r".//a[@href=" + r'"' + "#" + searchID + r'"]'

                        if htmlCnt.find(searchTerm).getparent() is not None:
                            if htmlCnt.find(searchTerm).getparent().tag == "p":
                        
                                paraTag = htmlCnt.find(searchTerm).getparent()
                                
                                htmlCnt.find(searchTerm).drop_tag()
                                eachCommentSpan.insert(0, paraTag)
                                
                                for eachspan in eachCommentSpan.xpath("//span[@class='MsoFootnoteReference']"):
                                    eachspan.text = ""
                                    eachspan.tag = "del-span"
                                paraTag.drop_tag()

                                
                                



        # Author Queries
        # for eachCommentSpan in htmlCnt.xpath("//span[@class='MsoCommentReference']"):
        #     if eachCommentSpan.find(".//a") is not None:
        #         if r"name" in eachCommentSpan.find(".//a").attrib:
        #             searchID = eachCommentSpan.find(".//a").get("name")
                    
        #             # print('".//div/a[@name=' + "'" + searchID + "']" + '"')
        #             searchTerm = r".//a[@href=" + r'"#' + searchID + r'"]'

        #             if htmlCnt.find(searchTerm).getparent().tag == "span":
        #                 htmlCnt.find(searchTerm).getparent().drop_tag()
                    
        #                 paraTag = htmlCnt.find(searchTerm).getparent() 
        #                 htmlCnt.find(searchTerm).drop_tag()
        #                 if paraTag.text is not None:
        #                     paraTag.text = r"\AQ{" + paraTag.text.replace("\n", " ") + r"}"
        #                     # Replace eachCommentSpan with paraTag
        #                     eachCommentSpan.getparent().replace(eachCommentSpan, paraTag)
        #                     paraTag.drop_tag()

        #href link Process
        for aLink in htmlCnt.xpath("//a"):
            if r"href" in aLink.attrib:
                if aLink.text:
                    aLink.text = r"\href{" + aLink.get("href") + r"}{" + aLink.text + r"}"
                    aLink.drop_tag()
            elif r"style" in aLink.attrib:
                aLink.drop_tag()

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
                                eachParaRef.text = each_val[0] + "Chap:" + ChapterNum + ":" + str(refCountPadding) + each_val[1] + eachParaRef.text
                            else:
                                eachParaRef.text = each_val[0] + str(refCountPadding) + each_val[1]

                            eachParaRef.drop_tag()

                            refCount = refCount + 1

        # print(html.tostring(htmlCnt, pretty_print=True, method="html", encoding="utf-8").decode())

        # temp_cnt = html.tostring(htmlCnt, pretty_print=True, method="html", encoding="utf-8").decode()
        # temp_cnt = temp_cnt.encode()
        # etreeCnt = etree.fromstring(htmlCnt)

        # print(etreeCnt)
            

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
            tot_cnt = tot_cnt.replace(r"/tableAmbersand/",r"&")
            
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
            htmlPath = os.path.join(os.path.split(output_format)[0], os.path.splitext(os.path.split(output_format)[1])[0] + "_out.html")
            conversionStart = TrackConversion(htmlcnt,htmlPath)

            
if __name__ == "__main__":

    User_Input = os.path.abspath(input("Enter the DOC file path with file name and extension: "))

    docConvert = DocConversion(User_Input)