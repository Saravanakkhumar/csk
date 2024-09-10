from pylatexenc.latexencode import unicode_to_latex
from pylatexenc.latexwalker import LatexWalker, LatexEnvironmentNode, LatexMacroNode, LatexCharsNode, LatexMathNode, LatexGroupNode, LatexCommentNode
from pylatexenc.latex2text import LatexNodes2Text
import re
import os
import chardet
import sys
# from pymsgbox import alert
import traceback

# curDir = os.getcwd()
# global serverIP
# serverIP=''
# if os.path.isfile(os.path.join(curDir,'ServerDetails.exe')):
#     serverIP=os.popen(os.path.join(curDir,'ServerDetails.exe')).read().strip()
#     if (os.path.isfile(r"\\"+str(serverIP)+r"\License\license.txt")):
#         chklicense = open(r"\\"+str(serverIP)+r"\License\license.txt", 'r').read()
#         if (chklicense != 'Active'):
#             alert(text='Please contact the tech support!', title='expired', button='OK')
#             exit()
#     else:
#         alert(text="Please check the \"internet\" or \"VPN\" connection", title='Expired', button='OK')
#         exit()
# else:
#     alert(text="\"ServerDetails.exe\" file missing...", title='Missing', button='OK')
#     exit()

# fin = input("Enter the original author file: ")
# bibcheckIn = input("Enter the bibcheck file path: ")

labelOrder = []


def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        return result['encoding']

def read_file(file_path):
    encoding = detect_encoding(file_path)
    with open(file_path, 'r', encoding=encoding) as f:
        return [f.read(), encoding]

def OriginalLabelCnt(originalfin):

    try:
        
        readFile = read_file(originalfin)
        texCnt = readFile[0]

        if r"{thebibliography}" in texCnt:
            pass
        else:
            print('File not converted. Bibliography not pasted in the original file. Please copy bbl and paste it to the Original LaTeX Source:\n')
            # sys.exit()
            

        texCnt = re.sub(r"\\begin{document}", "/StartDocument/", texCnt, flags=re.S)
        texCnt = re.sub(r"\\end{document}", "/EndDocument/", texCnt, flags=re.S)

        texCnt = re.sub(r"\\begin{thebibliography}{([^{}]+)}", "/bibStart/", texCnt, flags=re.S)
        texCnt = re.sub(r"\\end{thebibliography}", "/bibEnd/", texCnt, flags=re.S)
        
        texCnt = re.sub(r"\\bibitem(\[([^\[\]]+)\]|)(\{*)([^\{\}]+)(\}*)(\n|)", r"\\bibitem\g<1>{\g<4>}\n", texCnt, flags=re.MULTILINE)

        # texCnt = re.sub(r"\\bibitem(\[([^\[\]]+)\]|)(\{(.*)\})", r"\\bibitem\g<1>\g<2>\n", texCnt, flags=re.S)


        walker = LatexWalker(texCnt, tolerant_parsing=False)

        # Get all nodes from the document
        (nodes, pos, len_) = walker.get_latex_nodes(pos=0)

        collect_cnt = ""
        count = 0
        for node in nodes:
            collect_cnt += str(count) + " " + str(node)
            collect_cnt += "\n\n"
            count = count + 1

        for idx, eachNode in enumerate(nodes, start=0):
            
            currentNode = idx
            if isinstance(eachNode, LatexMacroNode): 
                if eachNode.macroname == "bibitem":

                    GroupFlag = True
                    
                    currentNode = currentNode + 1
                    if isinstance(nodes[currentNode], LatexCharsNode):
                        while GroupFlag:
                            if isinstance(nodes[currentNode], LatexCharsNode):
                                if hasattr(nodes[currentNode], "chars"):
                                    if r"]" in nodes[currentNode].chars:
                                        if isinstance(nodes[currentNode + 1], LatexGroupNode):
                                            if hasattr(nodes[currentNode + 1], "nodelist"):
                                                for eachNode in nodes[currentNode + 1].nodelist:
                                                    labelOrder.append(eachNode.chars)
                                        GroupFlag = False
                                    else:
                                        GroupFlag = True
                            else:
                                GroupFlag = True
                            currentNode = currentNode + 1
                    
                    elif isinstance(nodes[currentNode], LatexGroupNode):
                        if hasattr(nodes[currentNode], "nodelist"):
                            for eachNode in nodes[currentNode].nodelist:
                                labelOrder.append(eachNode.chars)

        treePath = os.path.splitext(originalfin)[0] + "_tree.txt"
        with open(treePath, "w", encoding=readFile[1]) as file:
            file.write(collect_cnt)
        return labelOrder
    except Exception as err:
        print("Original Label content error " + str(err) + " " + str(traceback.format_exc()))



def arrangeLabelCnt(bibcheckIn, getLabel, inputfile):
    
    try:

        orderedReference = []
        MissingLabel = []

        readFile = read_file(bibcheckIn)
        texCnt = readFile[0]
        texCnt = re.sub(r"\\begin{document}", "/StartDocument/", texCnt, flags=re.S)
        texCnt = re.sub(r"\\end{document}", "/EndDocument/", texCnt, flags=re.S)

        texCnt = re.sub(r"\\begin{thebibliography}(\s*|){([^{}]+)}", "/bibStart/", texCnt, flags=re.S)
        texCnt = re.sub(r"\\end{thebibliography}", "/bibEnd/", texCnt, flags=re.S)

        #\\bibitem(\[([^\[\]]+)\]|)(\{(.*)\})\n

        walker = LatexWalker(texCnt, tolerant_parsing=False)

        # Get all nodes from the document
        (nodes, pos, len_) = walker.get_latex_nodes(pos=0)

        collect_cnt = ""
        for node in nodes:
            collect_cnt += str(node)
            collect_cnt += "\n\n"

        for eachLabel in getLabel:
            if eachLabel in texCnt:
                pattern = r"(\\bibitem(\[([^\[\]]+)\]|){" + eachLabel + r"}(.*?))(\\bibitem{([^{}]+)}\n|/bibEnd/|\\bibitem(\[([^\[\]]+)\]){([^{}]+)}\n)"
            
                matches = re.search(pattern, texCnt, flags=re.S)

                if matches:
                    if matches[1]:
                        orderedReference.append(matches[1].strip())
            else:
                
                readFile = read_file(inputfile)

                originalTeXCnt = readFile[0]

                pattern = r"(\\bibitem(\[([^\[\]]+)\]|){" + eachLabel + r"}(.*?))(\\bibitem{([^{}]+)}\n|/bibEnd/|\\bibitem(\[([^\[\]]+)\]){([^{}]+)}\n)"
                
                matches = re.search(pattern, originalTeXCnt, flags=re.S)

                if matches:
                    if matches[1]:
                        orderedReference.append(matches[1].strip())

                MissingLabel.append(eachLabel)

           

        startNode = int()
        endNode = int()

        for idx, eachNode in enumerate(nodes, start=0):
        
            currentNode = idx
            if isinstance(eachNode, LatexMacroNode): 
                if eachNode.macroname == "bibitem":

                    GroupFlag = True
                    startNode = eachNode.pos

                    while GroupFlag:
                        currentNode = currentNode + 1
                        if isinstance(nodes[currentNode], LatexCharsNode):
                            idCnt = str(nodes[currentNode].chars)

                            if r"bibEnd" in idCnt:
                                endNode = nodes[currentNode].pos - 1
                                GroupFlag = False
                        else:
                            GroupFlag = True

                    break

        
        ReplaceCnt = texCnt[startNode:endNode]

        orderedReferenceCnt = "\n\n\n".join(orderedReference)            

        texCnt = texCnt.replace(r"/bibStart/", "\\begin{thebibliography}{99}")
        texCnt = texCnt.replace(r"/bibEnd/", "\\end{thebibliography}")
        texCnt = texCnt.replace(r"/StartDocument/", "\\begin{document}")
        texCnt = texCnt.replace(r"/EndDocument/", "\\end{document}")
        texCnt = texCnt.replace(ReplaceCnt, "<orderedReferenceCnt>", 1)

        if r"<orderedReferenceCnt>s" in texCnt:
            texCnt = texCnt.replace("<orderedReferenceCnt>s", orderedReferenceCnt, 1)
        else:
            texCnt = texCnt.replace("<orderedReferenceCnt>", orderedReferenceCnt, 1)

        store_file = os.path.splitext(bibcheckIn)[0] + "_converted.tex"

        missing_label_file = os.path.splitext(bibcheckIn)[0] + "_converted_missing_label.txt"

        # with open(store_file, "w", encoding=readFile[1]) as file:
        with open(inputfile, "w", encoding=readFile[1], errors='ignore') as file:
            file.write(texCnt)

        with open(missing_label_file, "w", encoding=readFile[1], errors='ignore') as file:
            writeLabel = "\n\n".join(MissingLabel)
            file.write(writeLabel)

        # if os.path.isfile(bibcheckIn):
        #     print('Bib Checked file converted successfully. The file placed the following path:\n' + store_file, title='Message', button='OK')

        # else:
        #     alert(text='File not converted. Please recheck the programme inputs:\n' + store_file, title='Message', button='OK')

    except Exception as err:
        print("Arrange Label content error. " + str(err) + " " + str(traceback.format_exc()))


# getLabel = OriginalLabelCnt(fin)
# arrangeLabel = arrangeLabelCnt(bibcheckIn, getLabel, fin)



