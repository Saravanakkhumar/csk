import re
import os
from pathlib import Path
from pymsgbox import alert
from bs4 import BeautifulSoup as bs
from lxml import etree
import win32com.client
import traceback
from sys import exit
from pylatexenc.latexencode import unicode_to_latex
from pylatexenc.latexwalker import LatexWalker, LatexEnvironmentNode, LatexMacroNode, LatexCharsNode, LatexMathNode, LatexGroupNode, LatexCommentNode
import gc
import time
import configparser
import unidecode
from pylatexenc.latexencode import UnicodeToLatexEncoder, UnicodeToLatexConversionRule, RULE_REGEX
import chardet
from functools import reduce

# Optimize garbage collection
gc.collect()

missing_Index = []

def replace_with_pandas(big_string, old, new):
    # Convert string to a pandas Series
    series = pd.Series([big_string])
    # Use pandas replace method
    return series.str.replace(old, new).iloc[0]

def roman_to_int(s):
    roman_values = {
        'I': 1,
        'V': 5,
        'X': 10,
        'L': 50,
        'C': 100,
        'D': 500,
        'M': 1000
    }

    total = 0
    prev_value = 0

    for char in reversed(s):
        value = roman_values[char]
        if value < prev_value:
            total -= value
        else:
            total += value
        prev_value = value

    return total

# Create a ConfigParser object
config = configparser.ConfigParser()


# Read the INI file
if os.path.isfile('//192.168.7.5/SoftwareTools/Journals/Config/IndexSorting.ini'):
    pass
else:
    alert(text="\"IndexSorting.ini\" file missing...", title='Missing', button='OK')
    exit()

config.read('//192.168.7.5/SoftwareTools/Journals/Config/IndexSorting.ini')




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

def get_child_nodes(env_node):
    if hasattr(env_node, 'nodelist'):
        return env_node.nodelist
    return []

def IDSequence(FileCnt,treePath):

    try:
        # fout = os.path.splitext(FileCnt)[0] + r"_node_list.tex"
        # ftex = os.path.splitext(FileCnt)[0] + r"_out.tex" 
        # ffinal = os.path.splitext(FileCnt)[0] + r"_out_final.tex" 

        ltx_cnt = FileCnt

        walker = LatexWalker(ltx_cnt)

        # Get all nodes from the document
        nodes, pos, _ = walker.get_latex_nodes()

        Global_Dict = {}

        collect_cnt = ""

        for node in nodes:
            collect_cnt += str(node)
            collect_cnt += "\n"

        # with open(treePath, "w", encoding="utf-8") as f1:
        #     f1.write(collect_cnt)
                        

        check_command_cnt = ["part", "paragraph", "addparasec", "addchap", "addsec", "addsubsec", "footnotetext",  "AuthorAnd", "aulist"]#"tabref", "figref", "secref", "algref"

        double_group_cnt = ["addtocontents", "href", "markboth"]

        tot_cnt = ""
        for nidx, node in enumerate(nodes):
            tot_cnt += str(node) + "\n\n\n"

            str_cnt = ""


            # all_command_cnt, check_command_cnt
            if isinstance(node, LatexMacroNode):

                if hasattr(node,'macroname'):
                    nodeName = str(node.macroname)


                    if any(eachCmd == nodeName for eachCmd in check_command_cnt):
                        starting_node = node.pos

                        GroupFlag = True
                        currentNode = nidx

                        while GroupFlag:
                            currentNode = currentNode + 1
                            if isinstance(nodes[currentNode], LatexGroupNode):
                                str_cnt = ltx_cnt[starting_node : nodes[currentNode].pos+nodes[currentNode].len]
                                IDValue = nodeName + "Cnt-" + str(node.pos)
                                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                                Global_Dict["<" + revIDValue + ">"] = str_cnt
                                GroupFlag = False

                    # double_group_cnt
                    if any(eachCmd == nodeName for eachCmd in double_group_cnt):
                        
                        starting_node = node.pos
                        if isinstance(nodes[nidx + 2], LatexGroupNode):
                            str_cnt = ltx_cnt[starting_node : nodes[nidx+2].pos+nodes[nidx+2].len]
                            IDValue = nodeName + "Cnt-" + str(node.pos)
                            revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                            Global_Dict["<" + revIDValue + ">"] = str_cnt

                    # Book 12 error
                    # if nodeName == "makeatletter":

                    #     starting_node = node.pos
                        
                    #     GroupFlag = True
                    #     currentNode = nidx

                    #     while GroupFlag:
                    #         currentNode = currentNode + 1
                    #         if isinstance(nodes[currentNode], LatexMacroNode):
                    #             if nodes[currentNode].macroname == "makeatother":
                    #                 str_cnt = ltx_cnt[starting_node : nodes[currentNode].pos+nodes[currentNode].len]
                    #                 Global_Dict["<" + nodeName + "Cnt-" + str(node.pos) + ">"] = str_cnt
                    #                 GroupFlag = False


        # "chapter", "section", "subsection", "subsubsection",  "subparagraph", "footnote", "url" ref", "eqref", "cite", "citealt", "citet", "citep", "citealt", "citealp", "citeauthor", "citeyear", "citeyearpar", "citetext", "citenum", "footnote", "url", "Cref"

        TeXCommandPattern = r"LatexMacroNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), macroname='(chapter|section|subsection|subsubsection|subparagraph|footnote)',"
        
        TeXCommandMatches = re.findall(TeXCommandPattern, collect_cnt, flags=re.S)
        
        if len(TeXCommandMatches) != 0:
            for eachMatch in reversed(TeXCommandMatches):
                st = int(eachMatch[1])
                en = int(eachMatch[1]) + int(eachMatch[2])
                str_cnt = ltx_cnt[st: en]
                IDValue = eachMatch[3] + "Cnt-" + str(eachMatch[1])
                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                Global_Dict["<" + revIDValue + ">"] = str_cnt


        # PreEnvironmentPattern
        PreEnvironmentPattern = r"LatexEnvironmentNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), environmentname='(landscape|tablehere)',"
        
        PreEnvironmentMatches = re.findall(PreEnvironmentPattern, collect_cnt, flags=re.S)
        
        if len(PreEnvironmentMatches) != 0:
            for eachMatch in reversed(PreEnvironmentMatches):
                st = int(eachMatch[1])
                en = int(eachMatch[1]) + int(eachMatch[2])
                str_cnt = ltx_cnt[st: en]
                IDValue = eachMatch[3] + "Cnt-" + str(eachMatch[1])
                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                Global_Dict["<" + revIDValue + ">"] = str_cnt



        # EnvironmentPattern
        EnvironmentPattern = r"LatexEnvironmentNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), environmentname='(figure|figure\*|table|table\*|algorithm|math|displaymath|eqnarray|eqnarray\*|align|align\*|flalign|flalign\*|multline|multline\*|gather|gather\*|subequation|equation\*|equation|thebibliography|verbatim)',"
        
        EnvironmentMatches = re.findall(EnvironmentPattern, collect_cnt, flags=re.S)
        
        if len(EnvironmentMatches) != 0:
            for eachMatch in reversed(EnvironmentMatches):
                st = int(eachMatch[1])
                en = int(eachMatch[1]) + int(eachMatch[2])
                str_cnt = ltx_cnt[st: en]
                IDValue = eachMatch[3] + "Cnt-" + str(eachMatch[1])
                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                Global_Dict["<" + revIDValue + ">"] = str_cnt


        # PostEnvironmentPattern
        PostEnvironmentPattern = r"LatexEnvironmentNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), environmentname='(tabular)',"
        
        PostEnvironmentMatches = re.findall(PostEnvironmentPattern, collect_cnt, flags=re.S)
        
        if len(PostEnvironmentMatches) != 0:
            for eachMatch in reversed(PostEnvironmentMatches):
                st = int(eachMatch[1])
                en = int(eachMatch[1]) + int(eachMatch[2])
                str_cnt = ltx_cnt[st: en]
                IDValue = eachMatch[3] + "Cnt-" + str(eachMatch[1])
                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                Global_Dict["<" + revIDValue + ">"] = str_cnt


        # Inline and normal display equation replacement
        MathInlineDisplayPattern = r"LatexMathNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), displaytype='(inline|display)',"


        MathInlineDisplayMatches = re.findall(MathInlineDisplayPattern, collect_cnt, flags=re.S)
        
        if len(MathInlineDisplayMatches) != 0:
            for eachMatch in reversed(MathInlineDisplayMatches):
                st = int(eachMatch[1])
                en = int(eachMatch[1]) + int(eachMatch[2])
                str_cnt = ltx_cnt[st: en]
                IDValue = eachMatch[3] + "Cnt-" + str(eachMatch[1])
                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                Global_Dict["<" + revIDValue + ">"] = str_cnt


        # Latex comment line replacement
        LatexCommentPattern = r"LatexCommentNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), comment='(.*?)', comment\_post\_space='(.*?)'\)"

        LatexCommentMatches = re.findall(LatexCommentPattern, collect_cnt, flags=re.S)
        
        if len(LatexCommentMatches) != 0:
            for eachMatch in reversed(LatexCommentMatches):
                st = int(eachMatch[1])
                en = int(eachMatch[1]) + int(eachMatch[2])
                str_cnt = ltx_cnt[st: en]
                IDValue = "LatexCommentCnt-" + str(eachMatch[1])
                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                Global_Dict["<" + revIDValue + ">"] = str_cnt


        PostTeXCommandPattern = r"LatexMacroNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), macroname='(label|url|ref|eqref|cite|citealt|citet|citep|citealt|citealp|citeauthor|citeyear|citeyearpar|citetext|citenum|footnote|url|Cref)',"
        
        PostTeXCommandMatches = re.findall(PostTeXCommandPattern, collect_cnt, flags=re.S)
        
        if len(PostTeXCommandMatches) != 0:
            for eachMatch in reversed(PostTeXCommandMatches):
                st = int(eachMatch[1])
                en = int(eachMatch[1]) + int(eachMatch[2])
                str_cnt = ltx_cnt[st: en]
                IDValue = eachMatch[3] + "Cnt-" + str(eachMatch[1])
                revIDValue = "|".join(IDValue[eachchar] for eachchar in range(0,len(IDValue)))
                Global_Dict["<" + revIDValue + ">"] = str_cnt


        if len(Global_Dict) != 0:
            for key,val in Global_Dict.items():
                if r"LatexComment" in key or r"makeatletter" in key:
                    ltx_cnt = ltx_cnt.replace(val, key + "\n", 1)
                else:
                    ltx_cnt = ltx_cnt.replace(val, key, 1)

        return [ltx_cnt, Global_Dict]

    except Exception as err:
        print("ID sequence error line no. 165" + str(err) + str(traceback.format_exc()))


variables = {}

def find_latex_comments_advanced(latex_content):
    # Regex pattern to ignore escaped '%' and find actual comments
    pattern = r"(?<!\\)%.*?\n"
    comments = re.findall(pattern, latex_content)
    return comments


def extract_paragraphs_from_latex(file_path):

    try:


        pathObject = os.path.split(file_path)[0]
        fileObject = os.path.split(file_path)[1]
        

        skip_element = [r"\bookseries", r"\begin{copyrightenv}", r"\mainmatter"] # Some processed files may skip this object
        
        if (r"_out.tex" in fileObject) or (r"out_final.tex" in fileObject):
            pass
        else:
            with open(os.path.join(pathObject, fileObject), 'r', encoding='latin-1') as file:
                latex_source = file.read()


                if any(skip_text in latex_source for skip_text in skip_element):
                    pass
                else:
                    print(str(os.path.join(pathObject, fileObject)) + " ... page info are under processing.")
                    
                    idseqtree = str(os.path.join(pathObject, fileObject))
                    idseqtreePath = os.path.splitext(idseqtree)[0] + r"_tree.txt"

                    latex_to_ID = IDSequence(latex_source, idseqtreePath)

                    with open(os.path.join(pathObject, str(os.path.splitext(fileObject)[0]) + "_out.tex"), "w", encoding='latin-1') as f1:
                        
                        try:
                            # Define a regex pattern for matching paragraphs in LaTeX
                            pattern = re.compile(r'(.+?\n{2,})', re.S)

                            patternone = re.compile(r'(\\pageinfoStart\{\})', re.S)

                            # Use the regex pattern to find paragraphs in the LaTeX source
                        
                            if latex_to_ID is not None:
                                paragraphs = re.findall(pattern, latex_to_ID[0])
                            
                                i = 1
                                if len(paragraphs) != 0:
                                    for eachParas in paragraphs:
                                        eachPara = eachParas
                                        eachPara = re.sub(r"(\\noindent\s*(.+?)\s+|\\item\s+(.+?)\s+|\s+([A-Za-z]+)\s+)", r"\g<1>\\pageinfoStart{}", eachPara, count=1)
                                        # eachPara = r"\pageinfoStart{"+ str(i) + r"}" + eachPara
                                        i = i + 1
                                        latex_to_ID[0] = latex_to_ID[0].replace(eachParas, eachPara)


                                pageinfo = re.findall(patternone, latex_to_ID[0])
                                j=1
                                if len(pageinfo) != 0:
                                    for eachpageinfo in pageinfo:
                                        latex_to_ID[0] = latex_to_ID[0].replace(eachpageinfo, r"\pageinfoStart{" + str(j) + "}", 1)
                                        j = j + 1


                                f1.write(latex_to_ID[0])

                        except Exception as err:
                            print("paragraph id sequence line no. 214" + str(err) + str(traceback.format_exc()))
                        

                        

                    with open(os.path.join(pathObject, str(os.path.splitext(fileObject)[0]) + "_out.tex"), "r", encoding='latin-1') as f1:
                        tex_cnt = f1.read()

                    with open(os.path.join(pathObject, str(os.path.splitext(fileObject)[0]) + "_out_final.tex"), "w", encoding='latin-1') as f2:
                        
                        # Comment lines are replaced
                        if latex_to_ID is not None:
                            for key,val in latex_to_ID[1].items():
                                if r"LatexCommentCnt" in key:
                                    tex_cnt = tex_cnt.replace(key + "\n", val, 1)
                                else:
                                    tex_cnt = tex_cnt.replace(key, val, 1)                        
                            f2.write(tex_cnt)
                        else:
                            pass
            
            return "Converted successfully."

    except Exception as err:
        print(file_path, "extract file paragraph err line no. 242", str(err) + str(traceback.format_exc()))

def DocRead(docpath):

    doc_format = str(os.path.splitext(docpath)[1])
        
    # if doc_format == r".doc":
    #     return alert(text="Could you please change the format doc to docx. ", title="Message Box", button="Ok")
    
    output_format = os.path.join(str(os.path.split(docpath)[0]), str(os.path.splitext(os.path.split(docpath)[1])[0]) + r".html") 
   
    doc = win32com.client.GetObject(docpath)
    doc.SaveAs (FileName=output_format, FileFormat=8)
    doc.Close ()

    if os.path.isfile(output_format):


        alert(text="HTML file converted successfully. ", title="Message Box", button="Ok")
    
        with open(output_format, "rb") as f1:
            binary_content = f1.read()

        # Detect encoding (often Windows-1252 or similar)
        try:
            decoded_content = binary_content.decode('windows-1252')  # Adjust based on actual encoding
        except UnicodeDecodeError:
            decoded_content = binary_content.decode('latin1')

        # Ensure the HTML has the correct UTF-8 meta tag
        if '<meta charset="UTF-8">' not in decoded_content:
            decoded_content = decoded_content.replace('<head>', '<head>\n<meta charset="UTF-8">')


        with open(output_format, "w", encoding="utf-8") as f1:
            f1.write(decoded_content)
        
        inHTML = decoded_content

        
        inHTML = inHTML.encode()

        soup = bs(inHTML, features="lxml")

        if len(soup.find_all("p")) != 0:

            texcmd_change = {"b": [r"\textbf{", "}"] , "i":  [r"\textit{", "}"], "emph":  [r"\emph{", "}"]}

            for key, val in texcmd_change.items():
                if len(soup.find_all(key)) != 0:
                    for eachFormat in soup.find_all(key):
                        eachFormat.insert(0, val[0])
                        eachFormat.append("}")

        if len(soup.find_all("span")) != 0:
            for each_cnt in soup.find_all("span"):

                # Extract the text from the element
                element_text = each_cnt.get_text()

                # Replace the element with its text
                each_cnt.replace_with(element_text)

        for new_line in soup.find_all("p"):

            if r"style" in new_line.attrs:
                del new_line["style"]
            
            new_line_cnt = str(new_line.text)
            new_line_cnt = re.sub("\n+", " ", new_line_cnt, flags=re.S)
            new_line.string = new_line_cnt

        # Index 1 Placement with sorting in out.html file Starting

        index_1_sorting = {}

        index_1_id = 1
        for index_1 in soup.find_all("p", {"class": "MsoIndex1"}):
            
            index_1_cnt = index_1.text

            index_1_sorting_cnt = ""
    
            sortingText = index_1_cnt
            patternCntOne = r""
                    
            for each_section in config.sections():
                if each_section == "IgnoreWords":
                    for each_key, each_val in config.items(each_section):
                        patternCntOne = each_val

            matchOne = re.search(r"(^" + patternCntOne  + r"(.*)$)"   , sortingText, flags=re.S)

            matchTwo = re.search(r"(^(\\textit\{[^\{\}]+\}|\\textbf\{[^\{\}]+\}))", sortingText, flags=re.S)

            # matchThree = re.search(r"(^(Â“|Â‘|``|`|\"|[0-9]+|\\([a-zA-Z]+)\{[a-zA-Z]+\})(.*)$)", sortingText, flags=re.S)
            matchThree = re.search(r"(^([^a-zA-Z])(.*)$)", sortingText, flags=re.S)


            if matchOne:

                sortCnt = matchOne.group(3)
                sortCnt = sortCnt.strip()
                sortCnt = unidecode.unidecode(sortCnt)
                index_1_sorting_cnt = sortCnt

            elif matchTwo:

                sortCnt = matchTwo.group(1)
                sortCnt = re.sub(r"\\textit\{([^\{\}]+)\}", r"\g<1>", sortCnt, flags=re.S)
                sortCnt = re.sub(r"\\textbf\{([^\{\}]+)\}", r"\g<1>", sortCnt, flags=re.S)
                sortCnt = sortCnt.strip()
                sortCnt = unidecode.unidecode(sortCnt)
                index_1_sorting_cnt = sortCnt

            elif matchThree:
                sortCnt = matchThree.group(1)
                sortCnt = re.sub(r"(^(Â“|Â‘|``|`|\"|[0-9]+|)(.*)$)", r"\g<3>", sortCnt, flags=re.S)
                sortCnt = re.sub("(Â”|Â’|)", "", sortCnt, flags=re.S)
                sortCnt = re.sub(r"(^(\$(.*?)\$))", "", sortCnt, flags = re.S)
                sortCnt = sortCnt.strip()
                sortCnt = unidecode.unidecode(sortCnt)
                index_1_sorting_cnt = sortCnt

            # index_1_sorting_cnt = re.sub("Â–", "--", index_1_sorting_cnt, flags=re.S)
            index_1_sorting_cnt = re.sub(r",\s+([0-9]+)–([0-9]+)", "", index_1_sorting_cnt, flags=re.S)
            index_1_sorting_cnt = re.sub(r",\s+([0-9]+)", "", index_1_sorting_cnt, flags=re.S)
            index_1_sorting_cnt = re.sub(r"\n+", " ", index_1_sorting_cnt, flags=re.S)

            # index_1_cnt = re.sub("Â–", "--", index_1_cnt, flags=re.S)
            index_1_cnt = re.sub(r",\s+([0-9]+)–([0-9]+)", "", index_1_cnt, flags=re.S)
            index_1_cnt = re.sub(r",\s+([0-9]+)", "", index_1_cnt, flags=re.S)
            index_1_cnt = re.sub(r"\n+", " ", index_1_cnt, flags=re.S)

            if len(index_1_sorting_cnt):
                index_1_sorting_cnt = index_1_sorting_cnt + r"@"
                index_1_cnt = index_1_sorting_cnt + index_1_cnt
                index_1_sorting["SortLevel-1-" + str(index_1_id)] = index_1_cnt
                index_1["id"] = "SortLevel-1-" + str(index_1_id)
            else:
                index_1_sorting_cnt = ""

            index_1_id = index_1_id + 1

        if len(index_1_sorting) != 0:
            for key, val in index_1_sorting.items():
                if soup.find("p", {"id": key}) is not None:
                    soup.find("p", {"id": key}).string = val

        # Index 1 Placement with sorting in out.html file Ending

        # Index 2 Placement with sorting in out.html file Starting

        index_2_sorting = {}

        index_2_id = 1
        for index_2 in soup.find_all("p", {"class": "MsoIndex2"}):
            
            index_1_cnt = str(index_2.find_previous("p", {"class": "MsoIndex1"}).text)

            index_2_sorting_cnt = ""
            
            if r"class" in index_2.find_previous("p").attrs:
                if index_2.find_previous("p").get("class")[0] == "MsoIndex1":
                    pass
                elif index_2.find_previous("p").get("class")[0] == "MsoIndex2":

                    sortingText = index_2.text
                    
                    patternCntOne = r""
                    
                    for each_section in config.sections():
                        if each_section == "IgnoreWords":
                            for each_key, each_val in config.items(each_section):
                                patternCntOne = each_val

                    matchOne = re.search(r"(^" + patternCntOne  + r"(.*)$)"   , sortingText, flags=re.S)

                    matchTwo = re.search(r"(^(\\textit\{[^\{\}]+\}|\\textbf\{[^\{\}]+\}))", sortingText, flags=re.S)

                    matchThree = re.search(r"(^([^a-zA-Z])(.*)$)", sortingText, flags=re.S)

                    if matchOne:
                        sortCnt = matchOne.group(3)
                        sortCnt = sortCnt.strip()
                        sortCnt = unidecode.unidecode(sortCnt)
                        index_2_sorting_cnt = sortCnt

                    elif matchTwo:

                        sortCnt = matchTwo.group(1)
                        sortCnt = re.sub(r"\\textit\{([^\{\}]+)\}", r"\g<1>", sortCnt, flags=re.S)
                        sortCnt = re.sub(r"\\textbf\{([^\{\}]+)\}", r"\g<1>", sortCnt, flags=re.S)
                        sortCnt = sortCnt.strip()
                        sortCnt = unidecode.unidecode(sortCnt)
                        index_2_sorting_cnt = sortCnt

                    elif matchThree:

                        sortCnt = matchThree.group(1)
                        sortCnt = re.sub(r"(^(Â“|Â‘|``|`|\"|[0-9]+|)(.*)$)", r"\g<3>", sortCnt, flags=re.S)
                        sortCnt = re.sub("(Â”|Â’|)", "", sortCnt, flags=re.S)
                        sortCnt = re.sub(r"(^(\$(.*?)\$))", "", sortCnt, flags = re.S)
                        sortCnt = sortCnt.strip()
                        sortCnt = unidecode.unidecode(sortCnt)
                        index_2_sorting_cnt = sortCnt

                    # index_2_sorting_cnt = re.sub("Â–", "--", index_2_sorting_cnt, flags=re.S)
                    index_2_sorting_cnt = re.sub(r",\s+([0-9]+)–([0-9]+)", "", index_2_sorting_cnt, flags=re.S)
                    index_2_sorting_cnt = re.sub(r",\s+([0-9]+)", "", index_2_sorting_cnt, flags=re.S)
                    index_2_sorting_cnt = re.sub(r",\s+([0-9]+)", "", index_2_sorting_cnt, flags=re.S)
                    index_2_sorting_cnt = re.sub(r"\n+", " ", index_2_sorting_cnt, flags=re.S)

                    if index_2_sorting_cnt:
                        index_2_sorting_cnt = index_2_sorting_cnt + r"@"
                        index_2_sorting_cnt = re.sub("-([0-9]+)@", "@", index_2_sorting_cnt, flags=re.S)
                    else:
                        pass
                    
            
            # index_1_cnt = re.sub("31–36", "--", index_1_cnt, flags=re.S)
            index_1_cnt = re.sub(r",\s+([0-9]+)–([0-9]+)", "", index_1_cnt, flags=re.S)
            index_1_cnt = re.sub(r",\s+([0-9]+)", "", index_1_cnt, flags=re.S)
            index_1_cnt = re.sub(r"\n+", " ", index_1_cnt, flags=re.S)
            
            if len(index_2_sorting_cnt) != 0:
                index_2_sorting["SortLevel-2-" + str(index_2_id)] = index_1_cnt + "!" + index_2_sorting_cnt + index_2.text
                index_2["id"] = "SortLevel-2-" + str(index_2_id)
            else:
                index_2_sorting_cnt = ""

            index_2.string = index_1_cnt + "!" + index_2.text

            index_2_id = index_2_id + 1

        if len(index_2_sorting) != 0:
            for key, val in index_2_sorting.items():
                if soup.find("p", {"id": key}) is not None:
                    soup.find("p", {"id": key}).string = val

        # Index 2 Placement with sorting in out.html file Ending

        for index_3 in soup.find_all("p", {"class": "MsoIndex3"}):
            
            index_2_cnt = str(index_3.find_previous("p", {"class": "MsoIndex2"}).text)
            # index_2_cnt = re.sub( "Â–", "--", index_2_cnt, flags=re.S)
            index_2_cnt = re.sub(r",\s+([0-9]+)–([0-9]+)", "", index_2_cnt, flags=re.S)
            index_2_cnt = re.sub(r",\s+([0-9]+)", "", index_2_cnt, flags=re.S)
            index_2_cnt = re.sub(r"\n+", " ", index_2_cnt, flags=re.S)
            
            index_3.string = index_2_cnt + "!" + index_3.text


        with open(os.path.join(os.path.split(docpath)[0], "out.html"), "w", encoding="utf-8", errors="ignore") as f1:

            
            if len(soup.find_all("p")) != 0:
                for eachpara in soup.find_all("p"):
                    paraCnt = eachpara.text
                    paraCnt = re.sub(r"\s+", r" ",paraCnt, flags=re.S)
                    eachpara.string = paraCnt
            index_cnt = str(soup)
            
            # index_cnt = index_cnt.encode()
            # index_cnt = index_cnt.decode("utf-8")

            # result = chardet.detect(index_cnt)
            # encodingCnt = result['encoding']

            # Build the conversion rules from the INI file
            # conversion_rules = [
            #     (re.compile(each_key), each_val) 
            #     for each_section in config.sections() if each_section == "LatexUnicodes" 
            #     for each_key, each_val in config.items(each_section)
            # ]

            # Create the UnicodeToLatexEncoder with the custom conversion rules
            # u = UnicodeToLatexEncoder(
            #     conversion_rules=[
            #         UnicodeToLatexConversionRule(rule_type=RULE_REGEX, rule=conversion_rules),
            #         'defaults'
            #     ]
            # )

            # index_cnt = ''.join(list(map(lambda x:  u.unicode_to_latex(x) if ord(x) >= 127 else x, index_cnt)))

            index_cnt = re.sub( "Â", "", index_cnt, flags=re.S)
            index_cnt = re.sub( "𝛿", r"$\\delta$", index_cnt, flags=re.S)
            index_cnt = re.sub( "𝜀", r"$\\epsilon$", index_cnt, flags=re.S)
            index_cnt = re.sub( "𝛽", r"$\\beta$", index_cnt, flags=re.S)
            index_cnt = re.sub( "𝜎", r"$\\sigma$", index_cnt, flags=re.S)
            index_cnt = re.sub( "ï", r"\"{\\i}", index_cnt, flags=re.S)

            # for each_section in config.sections():
            #     if each_section == "JunkChars":
            #         for each_key, each_val in config.items(each_section):
            #             print(each_key)
            #             index_cnt = re.sub(each_key, each_val, index_cnt, flags=re.S)


            index_cnt = ''.join(list(map(lambda x:  unicode_to_latex(x) if ord(x) >= 127 else x, index_cnt)))

            pattern = r"([0-9, {\\textendash}]+{\\textendash}[0-9, {\\textendash}]+|[0-9, ]+)</p>"
            para_cnt = re.findall(r"(<p([^<>]+)>([^><]+)<\/p>)", index_cnt, flags=re.S)

            split_cnt = ""

            for each_para in para_cnt:

                if each_para[0] is not None:
                    para_cnt = each_para[0]
                    para_cnt = para_cnt.replace("</p>", "")

                    cnt_split = re.split(r",(?=\s*\d)", para_cnt, flags=re.DOTALL|re.MULTILINE)
                    
                    index_organizer = {}

                    for each_split_cnt in cnt_split:
                        each_split_cnt = each_split_cnt.strip()
                        if r"class" in each_split_cnt:
                            split_cnt += each_split_cnt
                        elif "textendash" in each_split_cnt:
                            num_split = re.split(r"{\\textendash}", each_split_cnt, flags=re.DOTALL|re.MULTILINE)
                            start_range = "\n<page>" + num_split[0] + r"\(" + "</page>"
                            end_range = "\n<page>" + num_split[1] + r"\)" + "</page>"
                            split_cnt += start_range 
                            split_cnt += end_range
                        elif any(char.isdigit() for char in each_split_cnt):
                            split_cnt +=  "\n<page>" + each_split_cnt + "</page>"
                        elif any(char.isalpha() for char in each_split_cnt):
                            split_cnt += ", " + each_split_cnt
                    split_cnt += "</p>"
                    split_cnt += "\n\n"
                        
            f1.write(split_cnt)


def check_file_with_extension(extension):
    files_in_directory = os.listdir('.')
    for file_name in files_in_directory:
        if file_name.endswith(extension):
            return file_name
    return False


def pageinfocnt(file_pathA):

    extension = '.paginfo'
    os.chdir(file_pathA)
    check_page_info = check_file_with_extension(extension)
    fopen = open(check_page_info, "r")
    page_info_list = fopen.readlines()
    return page_info_list


def IndexImplementOnTeX(file_pathA,docpathA,pageInfoCnt,variablesIn):
    
    try:
        
        verify_index_cnt = []

        out_html_path = os.path.join(os.path.split(docpathA)[0], r"out.html")

        Log_file_cnt = []

        pageinfostartIndex = []

        if os.path.isfile(out_html_path):
            with open(out_html_path, "r", encoding="utf-8") as f1:
                html_cnt = f1.read()
                html_cnt = html_cnt.encode()

            soup = bs(html_cnt, features="lxml")

            if len(soup.find_all("p")) != 0:
                i = 1
                for each_para in soup.find_all("p"):
                    each_para_text_cnt = str(each_para.contents[0])
                    if len(each_para_text_cnt) != 0:
                        each_para_text_cnt = each_para_text_cnt.strip()
                    if len(each_para.find_all("page")) != 0:
                        for each_page in each_para.find_all("page"):
                            each_page_text_cnt = str(each_page.contents[0])
                            page_Search = str(each_page_text_cnt)


                            check_number = ""
                            text_cnt = ""

                            if r"\(" in page_Search:
                                text_cnt = "|("
                                page_Search = page_Search.replace(r"\(", "")
                                check_number = page_Search
                            elif r"\)" in page_Search:
                                text_cnt = "|)"
                                page_Search = page_Search.replace(r"\)", "")
                                check_number = page_Search
                            else:
                                text_cnt = ""
                                check_number = page_Search

                            page_Search = str(check_number)

                            if len(page_Search) != 0:
                                page_Search = page_Search.strip()
                                # print("\n")
                                Log_file_cnt.append(i)
                                verify_index_cnt.append("S.No: " + str(i) + ", Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt))
                                print("S.No: " + str(i) + ", Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt))
                                i = i + 1

                                if page_Search.isdigit():

                                    if any("-PST:" + page_Search in each_cnt for each_cnt in pageInfoCnt):
                                        pattern = r"(((.*)\.tex)-PID:([0-9]+)-PST:(" + page_Search + ")\n)"
                                    elif any("-PST:" + str(int(page_Search) - 1) in each_cnt for each_cnt in pageInfoCnt):
                                        pattern = r"(((.*)\.tex)-PID:([0-9]+)-PST:(" + str(int(page_Search) - 1) + ")\n)"
                                    elif any("-PST:" + str(int(page_Search) + 1) in each_cnt for each_cnt in pageInfoCnt):
                                        pattern = r"(((.*)\.tex)-PID:([0-9]+)-PST:(" + str(int(page_Search) + 1) + ")\n)"
                                    elif any("-PST:" + str(int(page_Search) - 2) in each_cnt for each_cnt in pageInfoCnt):
                                        pattern = r"(((.*)\.tex)-PID:([0-9]+)-PST:(" + str(int(page_Search) - 2) + ")\n)"
                                    elif any("-PST:" + str(int(page_Search) + 2) in each_cnt for each_cnt in pageInfoCnt):
                                        pattern = r"(((.*)\.tex)-PID:([0-9]+)-PST:(" + str(int(page_Search) + 2) + ")\n)"

                                    # List to store all occurrences
                                    all_occurrences = []

                                    # Loop through each string in the list
                                    for s in pageInfoCnt:
                                        # # Find all occurrences of the pattern in the current string
                                        occurrences = re.findall(pattern, s)
                                        # # Extend the list of all occurrences with the occurrences from the current string
                                        all_occurrences.extend(occurrences)

                                    # verify_index_cnt.append(all_occurrences)
                                    verify_index_cnt.append("\n")

                                    if len(all_occurrences) != 0:

                                        search_term = ""
                                        if r"!" in each_para_text_cnt:
                                            split_word = each_para_text_cnt.split("!")
                                            search_term = split_word[-1]
                                        else:
                                            search_term = each_para_text_cnt

                                        tex_read = open(all_occurrences[0][1], "r", encoding="latin-1").read()

                                        differentiate_cnt = ""
                                        for each_char in range(0, len(each_para_text_cnt)):
                                            if r" " in each_para_text_cnt[each_char]:
                                                differentiate_cnt += each_para_text_cnt[each_char]
                                            else:
                                                differentiate_cnt += each_para_text_cnt[each_char] + "/idx/"

                                        index_cnt = "/idx/i/idx/n/idx/d/idx/e/idx/x/idx/{" + differentiate_cnt + text_cnt + "}"

                                        # if r"\pageinfoStart{" + str(int(all_occurrences[0][3]) - 1) + r"}" in tex_read:
                                        #     text_cnt_pattern = re.search(r"pageinfoStart\{" + str(int(all_occurrences[0][3]) - 1) + r"\}" + r"(.*)" + r"pageinfoStart\{" + str(int(all_occurrences[-1][3])) + r"\}", tex_read, flags=re.S)
                                        # else:
                                        
                                        text_cnt_pattern = re.search(r"pageinfoStart\{" + all_occurrences[0][3] + r"\}" + r"(.*)" + r"pageinfoStart\{" + str(int(all_occurrences[-1][3])) + r"\}", tex_read, flags=re.S)

                                        if text_cnt_pattern:

                                            if search_term in text_cnt_pattern.group(1):
                                                old_cnt = text_cnt_pattern.group(1)
                                                new_cnt = old_cnt.replace(search_term, index_cnt + search_term, 1)
                                                tex_read = tex_read.replace(old_cnt, new_cnt, 1)
                                            else:
                                                deletewords = [r"and", r"in", r"of", r"was", r"on", r"to", r",", r"by", r"with", r"as", r"for", r"from", r"or"]

                                                # search_term = reduce(lambda term, word: re.sub(r"\s+", " ", re.sub(word, "", term)), deletewords, search_term)                                                
                                                search_term = search_term.split(" ")

                                                # Remove words in deletewords from search_term
                                                filtered_search_term = [word for word in search_term if word not in deletewords]
                                                # filtered_search_term = [word for word in search_term if word not in deletewords and not re.search(r"\d+", word)]

                                                if any(word in text_cnt_pattern.group(1) for word in filtered_search_term):
                                                    for word in filtered_search_term:
                                                        firstcheck = word + " "
                                                        secondcheck = " " + word 
                                                        old_cnt = text_cnt_pattern.group(1)
                                                        if firstcheck in text_cnt_pattern.group(1):
                                                            new_cnt = old_cnt.replace(firstcheck, index_cnt + firstcheck, 1)
                                                            tex_read = tex_read.replace(old_cnt, new_cnt, 1)
                                                            break
                                                        elif secondcheck in text_cnt_pattern.group(1):
                                                            new_cnt = old_cnt.replace(secondcheck, index_cnt + secondcheck, 1)
                                                            tex_read = tex_read.replace(old_cnt, new_cnt, 1)
                                                            break
                                                        else:
                                                            pass
                                                else:
                                                    tex_read = tex_read.replace(r"\pageinfoStart{" + all_occurrences[0][3] + r"}", index_cnt + r"\pageinfoStart{" + all_occurrences[0][3] + r"}", 1)
                                                    pageinfostartIndex.append(" Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt) + ", File: " + os.path.split(all_occurrences[0][1])[1] + "\n")
                                                        

                                            with open(all_occurrences[0][1], "w", encoding="latin-1") as file:
                                                file.write(tex_read)
                                        else:
                                            # print(all_occurrences[0][1], index_cnt.replace("/idx/", ""), r"\pageinfoStart{" + all_occurrences[0][3] + r"}")
                                            tex_read = tex_read.replace(r"\pageinfoStart{" + all_occurrences[0][3] + r"}", index_cnt + r"\pageinfoStart{" + all_occurrences[0][3] + r"}") 

                                            pageinfostartIndex.append(" Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt) + ", File: " + os.path.split(all_occurrences[0][1])[1] + "\n")

                                            with open(all_occurrences[0][1], "w", encoding="latin-1") as file:
                                                file.write(tex_read)

                                    else:
                                        missing_Index.append("S.No: " + str(i) + ", Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt))
                                else:

                                    # Roman numeral process start
                                    
                                    if any("-PST:" + page_Search in each_cnt for each_cnt in pageInfoCnt):
                                        pattern = r"(((.*)\.tex)-PID:([0-9]+)-PST:(" + page_Search + ")\n)"

                                    # List to store all occurrences
                                    all_occurrences = []

                                    # Loop through each string in the list
                                    for s in pageInfoCnt:
                                        # # Find all occurrences of the pattern in the current string
                                        occurrences = re.findall(pattern, s)
                                        # # Extend the list of all occurrences with the occurrences from the current string
                                        all_occurrences.extend(occurrences)

                                    if len(all_occurrences) != 0:

                                        search_term = ""
                                        if r"!" in each_para_text_cnt:
                                            split_word = each_para_text_cnt.split("!")
                                            search_term = split_word[-1]
                                        else:
                                            search_term = each_para_text_cnt

                                        tex_read = open(all_occurrences[0][1], "r", encoding="latin-1").read()

                                        differentiate_cnt = ""
                                        for each_char in range(0, len(each_para_text_cnt)):
                                            if r" " in each_para_text_cnt[each_char]:
                                                differentiate_cnt += each_para_text_cnt[each_char]
                                            else:
                                                differentiate_cnt += each_para_text_cnt[each_char] + "/idx/"

                                        index_cnt = "/idx/i/idx/n/idx/d/idx/e/idx/x/idx/{" + differentiate_cnt + text_cnt + "}"

                                        text_cnt_pattern = re.search(r"(pageinfoStart\{" + all_occurrences[0][3] + r"\}" + r"(.*)" + r"pageinfoStart\{" + str(int(all_occurrences[-1][3])) + r"\}" + ")", tex_read, flags=re.S)

                                        if text_cnt_pattern:

                                            if search_term in text_cnt_pattern.group(1):
                                                old_cnt = text_cnt_pattern.group(1)
                                                new_cnt = old_cnt.replace(search_term, index_cnt + search_term, 1)
                                                tex_read = tex_read.replace(old_cnt, new_cnt, 1)
                                            else:
                                                
                                                tex_read = tex_read.replace(r"\pageinfoStart{" + all_occurrences[0][3] + r"}", index_cnt + r"\pageinfoStart{" + all_occurrences[0][3] + r"}") 

                                                pageinfostartIndex.append(" Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt) + ", File: " + os.path.split(all_occurrences[0][1])[1] + "\n")

                                            with open(all_occurrences[0][1], "w", encoding="latin-1") as file:
                                                file.write(tex_read)
                                        else:
                                            # print(all_occurrences[0][1], index_cnt.replace("/idx/", ""), r"\pageinfoStart{" + all_occurrences[0][3] + r"}")
                                            tex_read = tex_read.replace(r"\pageinfoStart{" + all_occurrences[0][3] + r"}", index_cnt + r"\pageinfoStart{" + all_occurrences[0][3] + r"}") 

                                            pageinfostartIndex.append(" Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt) + ", File: " + os.path.split(all_occurrences[0][1])[1] + "\n")

                                            with open(all_occurrences[0][1], "w", encoding="latin-1") as file:
                                                file.write(tex_read)
                                    else:
                                        missing_Index.append("S.No: " + str(i) + ", Page No: " + str(page_Search) + ", Index Term: " + str(each_para_text_cnt))
        

        pageinfopath = os.path.join(file_pathA, "pageInfoPath.txt")
        with open(pageinfopath, "w", encoding="utf-8") as f1:
            
            collect_cnt = ""
            count = 1
            for eachIndex in pageinfostartIndex:
                collect_cnt += "S. No: " + str(count) + eachIndex
                collect_cnt += "\n"
                count = count + 1
            
            f1.write(collect_cnt)


        # active code end

                                # for each_occurrence in all_occurrences:
                                #     print(each_occurrence[1], each_occurrence[2], each_occurrence[3], each_occurrence[4])
                                    

                                #     tex_read = open(each_occurrence[1], "r", encoding="latin-1").read()

                                #     check_PID = "\\pageinfoStart{" + each_occurrence[3] + "} "

                                #     grep_ID = r"pageinfoStart\{" + str(int(each_occurrence[3]) + 4) + r"\}"

                                #     text_cnt_pattern = re.search(r"(pageinfoStart\{" + each_occurrence[3] + r"\}" + r"(.*)" + grep_ID + ")", tex_read, flags=re.S)

        #                             if check_PID in tex_read:
        #                                 text_cnt = text_cnt.replace(r"\(", "|()")
        #                                 text_cnt = text_cnt.replace(r"\)", "|)")
                                        
        #                                 Check_Word = each_para_text_cnt

        #                                 differentiate_cnt = ""
        #                                 for each_char in range(0, len(each_para_text_cnt)):
        #                                     if r" " in each_para_text_cnt[each_char]:
        #                                         differentiate_cnt += each_para_text_cnt[each_char]
        #                                     else:
        #                                         differentiate_cnt += each_para_text_cnt[each_char] + "/idx/"

        #                                 index_cnt = "/idx/i/idx/n/idx/d/idx/e/idx/x/idx/{" + differentiate_cnt + text_cnt + "}"
                                        
        #                                 if temp_cnt == 1:
        #                                     if grep_ID:
        #                                         Log_file_cnt.append(check_PID + " to " + "\\pageinfoStart{" + str(int(each_occurrence[3]) + 2) + "}")
                                                
        #                                     Log_file_cnt.append(each_occurrence[1] + ", " + index_cnt.replace(r"/idx/", "") + ", " + check_PID)

        #                                     Log_file_cnt.append(Check_Word)

        #                                     if text_cnt_pattern is not None:

        #                                         check_count = 1
        #                                         if Check_Word in text_cnt_pattern.group(1):
        #                                             replace_cnt = text_cnt_pattern.group(1)
        #                                             old_value = text_cnt_pattern.group(1)
        #                                             new_value = ""

        #                                             print(each_occurrence[1], index_cnt.replace(r"/idx/", ""), check_PID, "Second Print")

        #                                             replace_cnt = replace_cnt.replace("pageinfoStart", "p/idx/a/idx/g/idx/e/idx/i/idx/n/idx/f/idx/o/idx/S/idx/t/idx/a/idx/r/idx/t")

        #                                             replace_cnt = replace_cnt.replace(Check_Word, index_cnt + Check_Word, 1)

        #                                             replace_cnt = replace_cnt.replace("p/idx/a/idx/g/idx/e/idx/i/idx/n/idx/f/idx/o/idx/S/idx/t/idx/a/idx/r/idx/t", "pageinfoStart")

        #                                             new_value = replace_cnt

        #                                             tex_read = tex_read.replace(old_value, new_value, 1)
        #                                             check_count = check_count + 1

        #                                                     # print(text_cnt_pattern.group(1))
        #                                                     # new_string = ''.join('new' if substring == 'old' else substring for substring in old_string.split('old'))
        #                                                     # print(text_cnt_pattern.group(1))
        #                                                     # print("\n\n")
        #                                                     # print(replace_cnt)
        #                                                     # tex_read = re.sub(re.escape(text_cnt_pattern.group(1)), re.escape(replace_cnt), tex_read, count=1, flags=re.S)

        #                                                     # tex_read = tex_read[:start_index] + new_value + tex_read[start_index:]
        #                                                     # tex_read = tex_read.replace(tex_read[start_index: end_index], "CSK")
        #                                                     # tex_read = re.sub(re.escape(old_value), re.escape(new_value), tex_read, flags=re.DOTALL | re.M)

        #                                             # tex_read = tex_read.replace(old_value, new_value, 1)
        #                                             # old_value = text_cnt_pattern.group(1)
        #                                             # new_value = replace_cnt
        #                                             # start_index = tex_read.find(old_value)
        #                                             # end_index = start_index + len(old_value)

        #                                             # print("+++++++++++++++++++++++++++++")
        #                                             # print("\n")
                                                    
        #                                             # print(tex_read[start_index: end_index])
                                                    
        #                                             # print("\n")
        #                                             # print("+++++++++++++++++++++++++++++")


        #                                         else:
        #                                                 print(each_occurrence[1], index_cnt.replace(r"/idx/", ""), check_PID, "Third Print")
        #                                                 tex_read = tex_read.replace(check_PID, index_cnt + check_PID, 1)
        #                                                 check_count = check_count + 1

        #                                     else:
        #                                         print(each_occurrence[1], index_cnt.replace(r"/idx/", ""), check_PID, "Fourth Print")
        #                                         # tex_read = tex_read.replace(check_PID, index_cnt + check_PID, 1)
        #                                 else:
        #                                     break
                                        
        #                             temp_cnt = temp_cnt + 1

        #                             fopen = open(os.path.join(file_pathA, each_occurrence[1]), "w", encoding="utf-8")
        #                             fopen.write(tex_read)
        #                             fopen.close()
        #                             break

            # with open(os.path.join(file_pathA, "Index_log.txt"), "w", encoding="utf-8") as fcnt:
            #     for each_cnt in Log_file_cnt:
            #         fcnt.write(str(each_cnt))
            #         fcnt.write("\n")
            #         fcnt.write("\n")                                

            with open(os.path.join(file_pathA, "Index_content.txt"), "w", encoding="utf-8") as fcnt:
                for each_cnt in verify_index_cnt:
                    fcnt.write(str(each_cnt))
                    fcnt.write("\n")

            with open(os.path.join(file_pathA, "MissingIndex.txt"), "w", encoding="utf-8") as fcnt:
                for each_cnt in missing_Index:
                    fcnt.write(str(each_cnt))
                    fcnt.write("\n")


        for fname,cnts in variablesIn.items():

            with open(fname, "r", encoding="latin-1") as f1:
                final_cnt = f1.read()

                for key,val in cnts.items():
                    if r"LatexCommentCnt" in key:
                        final_cnt = final_cnt.replace(key + "\n", val, 1)
                    else:
                        final_cnt = final_cnt.replace(key, val, 1)

                with open(fname, "w", encoding="latin-1") as f2:
                    f2.write(final_cnt)
             

    except Exception as err:
        print("Index Implementation line no. 746 " + str(traceback.format_exc()))                                    


def pagenumInsertion(file_pathA,pageinfocnt):

    print("Page num Insertion Processing")

    if len(pageinfocnt) != 0:
        for eachCnt in pageinfocnt:

            cntSplit = eachCnt.split("-")
            paraID = cntSplit[1].replace("PID:", "")
            pageNumber = cntSplit[2].replace("PST:", "").replace("\n", "")
            openFilePath = os.path.join(file_pathA, cntSplit[0])
            
            fread = open(openFilePath, "r", encoding="latin-1").read()
            totCnt = fread
            totCnt = totCnt.replace("\\pageinfoStart{" + paraID + "}", "\\pageinfoStart{" + paraID + "," + pageNumber + "}") 

            fwrite = open(openFilePath, "w", encoding="latin-1")
            fwrite.write(totCnt)
            fwrite.close()
    
    time.sleep(1)
    print("Page num Insertion Completed...")

if __name__ == "__main__":

    try:

        file_path = input("Enter the TeX Package root path: ")
        docpath = input("Enter the DocX path with filename: ")

        # file_path = r"d:\accessbility\2024_Goals\02_Inserting_Index_Term\round_1\Book_22\test"
        # docpath = r"d:\accessbility\2024_Goals\02_Inserting_Index_Term\round_1\Book_12\Working\Index\ACM_Seneviratne_SubjectIndex_22May23.doc"


        # file_path = r'd:\accessbility\2024_Goals\02_Inserting_Index_Term\TeX'  # Replace with the actual path to your LaTeX file
        # docpath = r'd:\accessbility\2024_Goals\02_Inserting_Index_Term\test\ACM_TaliaSubjectIndex09Sep23.doc'  # Replace with the actual path to your LaTeX file

        # file_paths = [os.path.join(folder_path, file) for folder_path, _, files in os.walk(file_path) for file in files]
        file_paths = []

        for root, dirs, files in os.walk(file_path):
            # Process files in the current directory
            for file in files:
                # Process the file here
                file_paths.append(os.path.join(root, file))

            # Skip subfolders
            break

        for file in file_paths:
            if file.endswith("out.tex"):# bookseries, {copyrightenv}, Author's Biography}
                os.remove(file)
            elif file.endswith("out_final.tex"):
                os.remove(file)
            elif file.endswith(".aux"):
                os.remove(file)
            elif file.endswith(".txt"):
                os.remove(file)

        for file in file_paths:
            if file.endswith(".tex"):# bookseries, {copyrightenv}, Author's Biography}
                extract_paragraphs_from_latex(file)

                    
        alert(text="Para ID Completed. Please load the mother tex final.tex files and generate the pageinfo.", title="Message Box", button="Ok")

        DocRead(docpath)

        page_info_cnt = pageinfocnt(file_path)

        for each_file in [os.path.join(file_path,file) for file in os.listdir(file_path) if file.endswith('_final.tex')]:
            with open(each_file, "r", encoding="latin-1") as f1:# bookseries, {copyrightenv}, Author's Biography}
                TeXCnt = f1.read()
                IDSequencetreePath = os.path.splitext(each_file)[0] + r".txt"
                LaTeXCnt = IDSequence(TeXCnt,IDSequencetreePath)
                variables[os.path.join(file_path, each_file)] = LaTeXCnt[1]
                with open(each_file, "w", encoding="latin-1") as f2:
                    revisedTeXCnt = str(LaTeXCnt[0])
                    f2.write(revisedTeXCnt)
            print(each_file + " saved")

        IndexImplementOnTeX(file_path,docpath,page_info_cnt,variables)

        time.sleep(1)

        print("Index IDX replacement started...")

        for each_file in [os.path.join(file_path,file) for file in os.listdir(file_path) if file.endswith('_final.tex')]:
            with open(each_file, "r", encoding="latin-1") as f1:
                cnt = f1.read() 
                cnt = cnt.replace(r"/idx/i/idx/n/idx/d/idx/e/idx/x/idx/", "\\index")
                cnt = re.sub("/idx/", "", cnt)
                
                with open(each_file, "w", encoding="latin-1") as f1:
                    f1.write(cnt)

        print("Index IDX replacement Ended...")

        pagenumInsertion(file_path,page_info_cnt)       
    
    except Exception as err:
        print("Main function line no. 812" + str(err) + str(traceback.format_exc()))        
        
