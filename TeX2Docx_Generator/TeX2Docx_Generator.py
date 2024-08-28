import wx,os,pymsgbox
from pymsgbox import *
from wx.core import EVT_SLIDER
import os
import re
import pymsgbox
import time
import traceback
from pylatexenc.latexencode import unicode_to_latex
from pylatexenc.latexwalker import LatexWalker, LatexEnvironmentNode, LatexMacroNode, LatexCharsNode, LatexMathNode, LatexGroupNode, LatexCommentNode
import chardet
import configparser
import subprocess
import shutil
# MENU_FILE_EXIT = wx.NewId()
# DRAG_SOURCE    = wx.NewId()

cfg_content = r'''%\Preamble{xhtml,no-cut,fn-in,chapter-filename,early_,early^,NoFonts,-css,2,mathml,mathjax}
\Preamble{xhtml,no-cut,fn-in,early_,early^,NoFonts,-css,mathml,mathjax}
\newtoks\eqtoks 
\Configure{VERSION}{}
%\Configure{DOCTYPE}{\HCode{<!DOCTYPE html>\Hnewline}}
%\Configure{HTML}{\HCode{<html>\Hnewline}}{\HCode{\Hnewline</html>}}
\Configure{DOCTYPE}{\HCode{<?xml version="1.0" encoding="utf-8"?>}}
\Configure{HTML}{\HCode{<html lang="en" xml:lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">\Hnewline}}{\HCode{\Hnewline</html>}}
%\Configure{xmlns}{epub}{}
%\Configure{@HEAD}{\HCode{<head>\Hnewline}}{\HCode{\Hnewline</head>}}
%\Configure{@HEAD}{\HCode{<title>\jobname}}{\HCode{\Hnewline</title>}}
%\Configure{@HEAD}{\HCode{<meta charset="UTF-8" />\Hnewline}}
%\Configure{@HEAD}{\HCode{<meta name="generator" content="TeX4ht (http://www.cse.ohio-state.edu/\string~gurari/TeX4ht/)" />\Hnewline}}
\Configure{@HEAD}{}
\Configure{@HEAD}{\HCode{<link rel="stylesheet" type="text/css" href="../css/stylesheet.css"/>\Hnewline<script type="text/javascript" src="http://cdn.mathjax.org/mathjax/latest/MathJax.js?config=TeX-AMS-MML_HTMLorMML"></script>}}
%\Configure{@HEAD}{\HCode{<style type="text/css">\Hnewline
%    .MathJax_MathML {text-indent: 0;}\Hnewline
%  </style>\Hnewline}}
%\Configure{MathjaxSource}{https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml-full.js}

\renewcommand*{\^}[1]{\sp{#1}}
\renewcommand*{\_}[1]{\sb{#1}}

\Css{.theorem {font-style: italic;}}
\Css{.Ch-Author span{font-size:0.8rem;}}
\Css{body{font-family: MJXc-TeX-main-Rw,  MJXc-TeX-main-Iw,  MJXc-TeX-main-Bw, sans-serif;}}

\ExplSyntaxOn
\cs_new_protected:Npn \alteqtoks #1
{
  \tl_set:Nx \l_tmpa_tl {\detokenize{#1}}
  % % replace < > and & with xml entities
  \regex_replace_all:nnN { \x{26} } { &amp; } \l_tmpa_tl
  \regex_replace_all:nnN { \x{3C} } { &lt; } \l_tmpa_tl
  \regex_replace_all:nnN { \x{3E} } { &gt; } \l_tmpa_tl
  % replace \par command with blank lines
  \regex_replace_all:nnN { \x{5C}par\b } {\x{A}\x{A}} \l_tmpa_tl
  \tl_set:Nx \eqtoks{ \l_tmpa_tl }
  
  %\HCode{\l_tmpb_tl}
}
\ExplSyntaxOff

%\def\AltMath#1${\alteqtoks{#1}% 
%   #1\HCode{</mrow><annotation encoding="application/x-tex">\eqtoks</annotation>}$} 
%\Configure{$}{\Configure{@math}{display="inline"}\DviMath\HCode{<semantics><mrow>}}{\HCode{</semantics>}\EndDviMath}{\expandafter\AltMath} 

%\long\def\AltDisplay#1\]{\alteqtoks{#1}#1\HCode{</mrow><annotation encoding="application/x-tex">\eqtoks</annotation></semantics>}\]}
%\Configure{[]}{\Configure{@math}{display="block"}\DviMath$$\DisplayMathtrue\HCode{<semantics><mrow>}\AltDisplay}{$$\EndDviMath}

%\newcommand\eqannotate[1]{\alteqtoks{#1}\HCode{<semantics><mrow>}#1\HCode{</mrow><annotation encoding="application/x-tex">\eqtoks</annotation></semantics>}}
% environment support
\newcommand\VerbMathToks[2]{%
  \alteqtoks{\begin{#2}
    #1
  \end{#2}}%
  \ifvmode\IgnorePar\fi\EndP\Configure{@math}{display="block"}\DviMath\DisplayMathtrue\HCode{<semantics><mrow>}
  \begin{old#2}
    #1
  \end{old#2}
  \HCode{</mrow><annotation encoding="application/x-tex">}
  \HCode{\eqtoks}
  \HCode{</annotation></semantics>}
  \EndDviMath
}
\ExplSyntaxOn
\newcommand\VerbMath[1]{%
  \cs_if_exist:cTF{#1}{
    \expandafter\let\csname old#1\expandafter\endcsname\csname #1\endcsname
    \expandafter\let\csname endold#1\expandafter\endcsname\csname end#1\endcsname
    \RenewDocumentEnvironment{#1}{+!b}{%
      \NoFonts\expandafter\VerbMathToks\expandafter{##1}{#1}\EndNoFonts%
    }{}
  }{}%
}

\ExplSyntaxOff
%\Configure{mathbf}{\HCode{<b>}\NoFonts}{\HCode{</b>}\EndNoFonts}
%\Configure{mathrm}{\HCode{<roman>}\NoFonts}{\HCode{</roman>}\EndNoFonts}
\Configure{textit}{\HCode{<i>}\NoFonts}{\HCode{</i>}\EndNoFonts}
\Configure{textbf}{\HCode{<b>}\NoFonts}{\HCode{</b>}\EndNoFonts}
%\Configure{NormalFont}{STIX-Regular.otf}
%\Configure{ItalicFont}{STIX-Italic.otf}
%\Configure{BoldFont}{STIX-Bold.otf}
%\Configure{BoldItalicFont}{STIX-BoldItalic.otf}
%\ConfigureEnv{algorithm}{\Picture*{}}{\EndPicture}{}{}
%\Configure{Picture}{.svg}
\ConfigureEnv{copyrightenv}
{\ifvmode\IgnorePar\fi\EndP\HCode{<section type="copyright-page">\Hnewline}\IgnoreRule}
{\ifvmode\IgnorePar\fi\EndP\HCode{</section>\Hnewline}\EndIgnoreRule}{}{}



\catcode`\:=11
\def\fx:pt#1xxx!*?: {%
   \expandafter\ifx \csname big:#1:\endcsname\relax
         \expandafter\gHAssign\csname big:#1:\endcsname  0  \fi
   \expandafter\gHAdvance\csname big:#1:\endcsname  1
   \edef\big:fn{-#1-\csname big:#1:\endcsname}}
\catcode`\:=12
\chardef\%=`\%


\begin{document}
\VerbMath{alignat}
%\renewcommand\eqref[1]{\NoFonts\HChar{92}eqref\{\detokenize{#1}\}\EndNoFonts}
\EndPreamble'''

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

# Read the INI file
if os.path.isfile(r'd:\DailyTask\11-07-2024\MathEnvironment.ini'):
    pass
else:
    alert(text="\"MathEnvironment.ini\" file missing...", title='Missing', button='OK')
    exit()

config.read(r'd:\DailyTask\11-07-2024\MathEnvironment.ini')

def detect_encoding(file_path):
   with open(file_path, 'rb') as f:
      raw_data = f.read()
      result = chardet.detect(raw_data)
      return result['encoding']

def read_file(file_path):
   encoding = detect_encoding(file_path)
   with open(file_path, 'r', encoding=encoding) as f:
      return [f.read(), encoding]


def LatexWalkerIntialization(inCnt, treePath, encoding):

        
    try:

        walker = LatexWalker(inCnt, tolerant_parsing=False)

        # Get all nodes from the document
        (nodes, pos, len_) = walker.get_latex_nodes(pos=0)

        collect_cnt = ""
        for node in nodes:
            collect_cnt += str(node)
            collect_cnt += "\n\n"

        with open(treePath, "w", encoding=encoding) as f1:
            f1.write(collect_cnt)

        return nodes

    except Exception as err:
        print('LaTeX Walker Validation Error ' + str(err) + str(traceback.format_exc()))
        return 'LaTeX Walker Validation Error ' + str(err) + str(traceback.format_exc())

Global_Dict = {}

def run_command(command, filePath):
    try:
        result = subprocess.run(command, check=True, capture_output=True, text=True)
        print(f"Command '{' '.join(command)}' executed successfully.")
        print("Output:", result.stdout)
        print("Errors:", result.stderr)
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")
        print("Output:", e.stdout)
        print("Errors:", e.stderr)

def eqprocess(inpath):
   

   try:

      readFile = read_file(inpath)

      treePath = os.path.splitext(inpath)[0] + "_tree.txt"

      writefile = os.path.splitext(inpath)[0] + "_out.tex"

      texCnt = readFile[0]
      texCnt = re.sub(r"\\begin{document}", "/StartDocument/", texCnt, flags=re.S)
      texCnt = re.sub(r"\\end{document}", "/EndDocument/", texCnt, flags=re.S)
      texCnt = re.sub(r"\\chapter\[([^\[\]]+)\]", r"\\chapter", texCnt, flags=re.S)
      texCnt = re.sub(r"\\title\[([^\[\]]+)\]", r"\\title", texCnt, flags=re.S)
      
      latexWalker = LatexWalkerIntialization(texCnt, treePath, readFile[1])

      inlineMathFormulas = {}

      if os.path.isfile(treePath):
         
         with open(treePath, "r", encoding=readFile[1]) as file:
            collect_cnt = file.read()


         #Pramble Cnt
         prambleCnt = ""
         chapterCnt = ""

         begingDocumentPattern = r"LatexCharsNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), chars=\'(.*?)\'\)"

         begingDocumentMatches = re.findall(begingDocumentPattern, collect_cnt, flags=re.S)
         
         if len(begingDocumentMatches) != 0:
            for eachMatch in reversed(begingDocumentMatches):
               if r"StartDocument" in eachMatch[3]:
                  en = int(eachMatch[1]) + int(eachMatch[2])
                  str_cnt = texCnt[0: en]
                  prambleCnt = str_cnt
               

         ChapterPattern = r"LatexMacroNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), macroname='(chapter|title)'"

         ChapterPatternMatches = re.findall(ChapterPattern, collect_cnt, flags=re.S)

         if len(ChapterPatternMatches) != 0:
            for eachMatch in ChapterPatternMatches:
               st = int(eachMatch[1])
               en = int(eachMatch[1]) + int(eachMatch[2])

               str_cnt = texCnt[st: en]
               if eachMatch[3]:
                  if r"title" in eachMatch[3]:
                     chapterCnt = str_cnt
                     chapterCnt += "\n" + r"\maketitle"
                     chapterCnt += "\n\n"
                  else:
                     
                     chapterCnt += r"\mainmatter" + "\n\n"
                     chapterCnt += r"\setcounter{chapter}{1}" + "\n"
                     chapterCnt += str_cnt
                     chapterCnt += "\n\n"
               break
            
         # Inline and normal display equation replacement
         MathPattern = r"(LatexMathNode|LatexEnvironmentNode)\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), (displaytype='(inline|display)',|environmentname='(" + config['MathEnvironmentList']['environ'] + ")',)"

         MathPatternMatches = re.findall(MathPattern, collect_cnt, flags=re.S)
         
         if len(MathPattern) != 0:

            i = 1
            for eachMatch in MathPatternMatches:
               st = int(eachMatch[2])
               en = int(eachMatch[2]) + int(eachMatch[3])
               str_cnt = texCnt[st: en]
               if eachMatch[0] == "LatexMathNode":
                  Global_Dict["EQCnt-" + str(i)] = [str_cnt, st, en, "inline"]
               elif eachMatch[0] == "LatexEnvironmentNode":
                  Global_Dict["EQCnt-" + str(i)] = [str_cnt, st, en, "display"]
               i = i + 1

         # EnvironmentPattern
         # EnvironmentPattern = r"LatexEnvironmentNode\(parsing_state=\<parsing state ([0-9]+)\>, pos=([0-9]+), len=([0-9]+), environmentname='(" + config['MathEnvironmentList']['environ'] + ")',"
         
         # EnvironmentMatches = re.findall(EnvironmentPattern, collect_cnt, flags=re.S)
         
         # if len(EnvironmentMatches) != 0:
         #    j = 1
         #    for eachMatch in reversed(EnvironmentMatches):
         #       st = int(eachMatch[1])
         #       en = int(eachMatch[1]) + int(eachMatch[2])
         #       str_cnt = texCnt[st: en]
         #       Global_Dict["DisplayEqCnt-" + str(j)] = str_cnt
         #       texCnt = texCnt[:st] + "\\begin{center}\n\\includegraphics{DisplayEqCnt-" + str(j) + ".png}\n\\end{center}" + texCnt[en:]
         #       j = j + 1

         mathFormulaFile = os.path.splitext(inpath)[0] + "-mathFormula.tex"

         with open(mathFormulaFile, "w", encoding=readFile[1]) as file:
            
            math_collect = ""

            
            math_collect += prambleCnt.replace("/StartDocument/", "")

            math_collect += r"\usepackage[active,tightpage]{preview}"
            math_collect += "\n\\scrollmode\n\n"
            math_collect += "\\begin{document}\n\n"
            # math_collect += chapterCnt

            for key,val in Global_Dict.items():
               math_collect += "%" + key
               math_collect += "\n"
               math_collect += r"\begin{preview}" + "\n" + val[0] + "\n" + r"\end{preview}"
               math_collect += "\n\n"

            math_collect += "\\end{document}"
            file.write(math_collect)   

         alert(text='Please update the tex!', title='expired', button='OK')

         command = ["lualatex", "--interaction=nonstopmode", mathFormulaFile]
         os.chdir(os.path.split(mathFormulaFile)[0])
         eqlog = os.path.splitext(mathFormulaFile)[0] + r"_eq-log.log"

         try:
            result = subprocess.run(command, check=True, capture_output=True, text=True)
            result1 = subprocess.run(command, check=True, capture_output=True, text=True)

            print("Output:", result1.stdout)
            
            print("PDF compiled successfully.")
               
         except subprocess.CalledProcessError as e:
            print(f"An error occurred: {e}")
            print("Output:", e.stdout)
            print("Errors check:", e.stderr)
                    
         totEQ = len(Global_Dict)
         outPDF = os.path.splitext(mathFormulaFile)[0] + ".pdf" 
         extractPath = os.path.split(mathFormulaFile)[0]

         i = 1

         for key,val in Global_Dict.items():
               print(r"Processing the following equation " + key + " out of " + str(totEQ) + "....")

               # Original Tiff
               # os.system(r"gswin64c.exe -sDEVICE=pdfwrite -dBATCH -dNOPAUSE -dSAFER -sOutputFile=" + extractPath + "\\" + key  + ".pdf -dFirstPage=" + str(i) + " -dLastPage=" + str(i) + " " + outPDF)
               # os.system(r"d:\some\gs\gs9.19\bin\gswin32c.exe -q -dBATCH -dNOPAUSE -sDEVICE=bmpgray -r1200 -dEmbedAllFonts=true -sOutputFile=" + extractPath + "\\" + key  + ".bmp -dTextAlphaBits=4 -dDisplayFormat=4 -dGraphicsAlphaBits=4 " + extractPath + "\\" + key  + ".pdf")
               # os.system(r"d:\some\convert.exe -trim " + extractPath + "\\" + key  + ".bmp " + extractPath + "\\" + key  + ".bmp")
               # os.system(r"d:\some\convert.exe " + extractPath + "\\" + key  + ".bmp " + extractPath + "\\" + key  + ".tiff")
               

               os.system(r"gswin64c.exe -sDEVICE=pdfwrite -dBATCH -dNOPAUSE -dSAFER -sOutputFile=" + extractPath + "\\" + key  + ".eps -dFirstPage=" + str(i) + " -dLastPage=" + str(i) + " " + outPDF)
               # os.system(r"d:\some\gs\gs9.19\bin\gswin32c.exe -q -dBATCH -dNOPAUSE -sDEVICE=bmpgray -r1200 -dEmbedAllFonts=true -sOutputFile=" + extractPath + "\\" + key  + ".png -dTextAlphaBits=4 -dDisplayFormat=4 -dGraphicsAlphaBits=4 " + extractPath + "\\" + key  + ".pdf")
               # os.system(r"d:\some\convert.exe -trim " + extractPath + "\\" + key  + ".png " + extractPath + "\\" + key  + ".png")
               # os.system(r"d:\some\convert.exe " + extractPath + "\\" + key  + ".png " + extractPath + "\\" + key  + ".eps")

               # if os.path.isfile(extractPath + "\\" + key + ".png"):
               #    os.remove(extractPath + "\\" + key + ".png")
               
               if os.path.isfile(extractPath + "\\" + key + ".pdf"):
                  os.remove(extractPath + "\\" + key + ".pdf")
      
               i = i + 1

         for key,val in reversed(Global_Dict.items()):
            if r"inline" in val[3]:
              texCnt = texCnt[:val[1]] + r"\includegraphics{" +  key + r".png}" + texCnt[val[2]:]
            if r"display" in val[3]:
               texCnt = texCnt[:val[1]] + r"\begin{center}" + "\n" + r"\includegraphics{" +  key + r".png}" + "\n" + r"\end{center}" + texCnt[val[2]:]

         with open(writefile, "w", encoding=readFile[1]) as file:
            
            texCnt = re.sub("/StartDocument/", r"\\begin{document}", texCnt, flags=re.S)
            texCnt = re.sub("/EndDocument/", r"\\end{document}", texCnt, flags=re.S)

            file.write(texCnt)

      return writefile      
   except Exception as err:
      print('File contains eqprocess error!' + str(err) +  str(traceback.format_exc()))
      alert(text='File contains eqprocess error!' + str(err) +  str(traceback.format_exc()), title='Error Message', button='OK')
      

def tex2docx(rootin):

   try:

      file_path = os.path.abspath(rootin)

      if os.path.isfile("//192.168.7.5/proof/ACM_Testing/pdmr-acm.cfg"):
         pass
      else:
         return alert(text="\"License Expired. Please contact Technical team.\".", title='Missing', button='OK')

      readFile = read_file(file_path)
      
      tex_cnt = readFile[0]
      tex_cnt = re.sub(r"(\n\s*|\n|\\\\\n\s*)\\end{(array|matrix|bmatrix|pmatrix)}", r"\\\\\n\\end{\g<2>}", tex_cnt, flags=re.S)

      with open(file_path, "w", encoding=readFile[1]) as fCnt:
         fCnt.write(tex_cnt)

      # with open(r"d:\pdmr-acm.cfg", "w") as cfgCnt:
      #    cfgCnt.write(cfg_content)

      path_split = os.path.splitext(file_path)
      change_dir = os.path.split(file_path)[0]
      os.chdir(change_dir)
      Tex_File_Name = os.path.splitext(os.path.split(file_path)[1])[0]

      os.system("@echo off")
      os.system("del *.aux")
      os.system("del *.xhtml")
      os.system("del *.epub")
      os.system("del *.opf")
      os.system("del .tmp")
      os.system("del .ncx")
      os.system("del .lg")
      os.system("del .idv")
      os.system("del .xref")
      os.system("del .log")
      os.system("del .idx")
      os.system("del .dvi")
      os.system("del .4tc")
      os.system("del .4ct")
      os.system("del .tex~")
      os.system("latexmk -C")
      os.system("make4ht -m clean " + Tex_File_Name)
      os.system("tex4ebook -a debug -c " + r"'//192.168.7.5/proof/ACM_Testing/pdmr-acm.cfg'" + " -f epub3+tidy -l " + Tex_File_Name + ".tex")

      time.sleep(15)

      # if os.path.isfile(r"d:\pdmr-acm.cfg"):
      #    os.remove(r"d:\pdmr-acm.cfg")


      epub_path = os.path.join(Tex_File_Name + r"-epub3", "oebps", Tex_File_Name + r".xhtml")
      # if os.path.isfile(epub_path):
      #    alert(text='EPUB file generated. Please check the following file ' + epub_path + ' is available in this path.', title="TeX2Docx Conversion", button='Ok')
      # else:
      #    alert(text='EPUB file not generated in the following path ' + epub_path + '. Please check the error. ', title="TeX2Docx Conversion", button='Ok')


      if os.path.isfile(epub_path):
         file_cnt = ""
         readEpubFile = read_file(epub_path)
         
         file_cnt = readEpubFile[0]
         file_cnt = ''.join(list(map(lambda x: '&#x' + (hex(ord(x)).lstrip('0x').zfill(4).upper() + ';') if ord(x) >= 127 else x, file_cnt)))

         file_cnt = file_cnt.replace(r"<?xml version='1.0' encoding='utf-8' ?>", "")
         file_cnt = file_cnt.replace(r"<p class='noindent'>&#x00A0; ", r"<p class='noindent'>")

         with open(Tex_File_Name + ".html", "w", encoding=readEpubFile[1]) as htmlfile:
            htmlfile.write(file_cnt)
         

         if os.path.isfile(Tex_File_Name + ".html"):

               
            refereceDocPath = r"//192.168.7.5/SoftwareTools/Journals/Config/reference.docx"

            if os.path.isfile(refereceDocPath):
               shutil.copy(refereceDocPath, os.path.split(rootin)[0])
               os.system('pandoc -s ' + Tex_File_Name + ".html --mathml " +  '-o ' + Tex_File_Name + r".docx --reference-doc=reference.docx")
               time.sleep(10)
            else:
               os.system('pandoc -s ' + Tex_File_Name + ".html --mathml " +  '-o ' + Tex_File_Name + r".docx")
               time.sleep(10)


      if os.path.isfile(Tex_File_Name + ".log"):
         with open(Tex_File_Name + ".log", "r", encoding="latin-1") as logfile:
            log_cnt = logfile.read()
            log_cnt = log_cnt.replace(r"d:/pdmr-acm.cfg", "")
            log_cnt = log_cnt.replace(r"//192.168.7.5/proof/ACM_Testing/pdmr-acm.cfg", "")
            with open(Tex_File_Name + ".log", "w", encoding="latin-1") as logfile:
               logfile.write(log_cnt)


      if os.path.isfile(Tex_File_Name + ".docx"):
         alert(text='DOCX file successfully generated in the same path. Please check.', title="TeX2Docx Conversion", button='Ok')
      else:
         alert(text='File contains a error. DOCX file not generated. Please check.', title="TeX2Docx Conversion", button='Ok')

   except Exception as err:
      alert(text='TeX2Docx conversion funtion error ' + str(err) + str(traceback.format_exc()), title="TeX2Docx Conversion", button='Ok')


# Define Text Drop Target class
class TextDropTarget(wx.TextDropTarget):
   """ This object implements Drop Target functionality for Text """
   def __init__(self, obj):
      """ Initialize the Drop Target, passing in the Object Reference to
          indicate what should receive the dropped text """
      # Initialize the wx.TextDropTarget Object
      wx.TextDropTarget.__init__(self)
      # Store the Object Reference for dropped text
      self.obj = obj

   def OnDropText(self, x, y, data):
      """ Implement Text Drop """
      # When text is dropped, write it into the object specified
      self.obj.WriteText(data + '\n\n')

# Define File Drop Target class
class FileDropTarget(wx.FileDropTarget):
   """ This object implements Drop Target functionality for Files """
   def __init__(self, obj):
      """ Initialize the Drop Target, passing in the Object Reference to
          indicate what should receive the dropped files """
      # Initialize the wxFileDropTarget Object
      wx.FileDropTarget.__init__(self)
      # Store the Object Reference for dropped files
      self.obj = obj
      self.obj.WriteText('The tool is ready now. You can able to do Drag & Drop the TeX file over here...')
      

   def OnDropFiles(self, x, y, filenames):
      """ Implement File Drop """
      # For Demo purposes, this function appends a list of the files dropped at the end of the widget's text
      # Move Insertion Point to the end of the widget's text
      self.obj.SetInsertionPointEnd()
      # append a list of the file names dropped
      # self.obj.WriteText("%d file(s) dropped at %d, %d:\n" % (len(filenames), x, y))
      # for file in filenames:
      #    self.obj.WriteText(file + '\n')
      
      self.obj.Clear()
      self.obj.WriteText(filenames[0] + '\n')
      # if filenames[0].endswith('.pdf'):
      #    self.obj.WriteText(filenames[0] + '\n')
      # else:
      #    self.obj.WriteText('Please choose the valid PDF file\n')


class MainWindow(wx.Frame):
   """ This window displays the GUI Widgets. """
   def __init__(self,parent,id,title):
      # wx.Frame.__init__(self,parent, wx.ID_ANY, title, size = (500,265), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
      wx.Frame.__init__(self,parent, wx.ID_ANY, title, size = (550,165), style=(wx.DEFAULT_FRAME_STYLE|wx.MAXIMIZE) & ~ (wx.RESIZE_BORDER|wx.MAXIMIZE_BOX|wx.MINIMIZE_BOX))
      self.SetBackgroundColour(wx.WHITE)
      
      # List1 = ['None', 'Irish CGT', 'Capital Tax Acts', 'Irish Income Tax','Taxation of Companies 2021','Arthur Cox Employment Law Yearbook']
      # self.Jnum = wx.RadioBox(self, label = 'Choose Book Title', pos = (10,10), choices = List1, majorDimension = 2)
      JLists = ['CBM', 'UCLOE']
        #   self.Jnum = wx.ComboBox(self, pos = (300,95), choices = JLists)
        #   self.toFindTextLabel = wx.StaticText(self, -1, 'ISBN-No: ', pos=(98,99))
        #   self.FindText = toFindText = wx.TextCtrl(self, -1, pos = (150,96), size=(100, 20), style=wx.TE_PROCESS_ENTER)
        #   self.bookNameLabel = wx.StaticText(self, -1, 'BookName: ', pos=(260,99))
        #   self.bookName = toFindText = wx.TextCtrl(self, -1, pos = (325,96), size=(120, 20), style=wx.TE_PROCESS_ENTER)
      self.conlink = wx.Button(self, 1, "Convert", pos = (455,95))

      vbox = wx.BoxSizer(wx.VERTICAL)
      hbox = wx.BoxSizer(wx.HORIZONTAL)

      # vbox.Add(self.Jnum, 0, wx.LEFT|wx.ALL, 1) #, 1, flag=wx.EXPAND
        #   vbox.Add(self.bookNameLabel, 0, wx.LEFT|wx.ALL, 1) #, 1, flag=wx.EXPAND
        #   vbox.Add(self.bookName, 0, wx.LEFT|wx.ALL, 1) #, 1, flag=wx.EXPAND
        #   vbox.Add(self.toFindTextLabel, 0, wx.LEFT|wx.ALL, 1) #, 1, flag=wx.EXPAND
        #   vbox.Add(self.FindText, 0, wx.LEFT|wx.ALL, 1) #, 1, flag=wx.EXPAND
      hbox.Add(self.conlink, 0, wx.RIGHT|wx.ALL, 1) #, 1, flag=wx.EXPAND
      
      # self.Jnum.Bind(wx.EVT_BUTTON,self.Conversion)
        #   self.bookNameLabel.Bind(wx.EVT_BUTTON,self.Conversion)
        #   self.bookName.Bind(wx.EVT_BUTTON,self.Conversion)
        #   self.toFindTextLabel.Bind(wx.EVT_BUTTON,self.Conversion)
        #   self.FindText.Bind(wx.EVT_BUTTON,self.Conversion)
      self.conlink.Bind(wx.EVT_BUTTON, self.Conversion)


      # Define a Text Control to receive Dropped Files
      # Label the control
      lbl1 = wx.StaticText(self, -1, "PDMR @ 2024", pos=(10,100))
      lbl1.SetForegroundColour((25,25,25))
      #wx.StaticText(self, -1, "File Drop Target (from off application only)", (370, 261))
      # Create a read-only Text Control
      self.text3 = wx.TextCtrl(self, -1, "", pos=(10,10), size=(516,80), style = wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY)
      # Make this control a File Drop Target
      # Create a File Drop Target object
      dt3 = FileDropTarget(self.text3)
      # Link the Drop Target Object to the Text Control
      self.text3.SetDropTarget(dt3)

      # Display the Window
      self.Show(True)
      self.Maximize(False)
      

      icon = wx.Icon()
      # iconfile=r'\\192.168.7.5\License\Blooms-Irish\favicon.ico'
      iconfile=r"\\"+str(serverIP)+r'\License\PDMRLOGO\PDMR_Logo_300dp.ico'
      # iconfile=r'C:\Users\ws230\Downloads\PDMR Logo_300dp.ico'
      icon.CopyFromBitmap(wx.Bitmap(iconfile, wx.BITMAP_TYPE_ANY))
      self.SetIcon(icon)


   def CloseWindow(self, event):
      """ Close the Window """
      self.Close()


   def Conversion(self,e):
      # bookName=self.Jnum.GetStringSelection()
    #   bookName=self.bookName.GetValue()
    #   isbn=self.FindText.GetValue()
      
      filePath = self.text3.GetValue().strip('\n')

      rootIn = filePath

      con_start_eq = eqprocess(rootIn)


      if len(con_start_eq) != 0:
         if os.path.isfile(con_start_eq):
            con_start = tex2docx(con_start_eq)
         else:
           alert(text='Equation process not completed. Out file not generated.', title="Error Message", button='Ok') 

      self.Close()
      
    #   if not os.path.isdir(filePath):
    #      alert(text='The tool is ready now. You can Drag & Drop the PDF file over here...', title="ACM Books - XHTML to XML Conversion", button='Ok')


   def OnDragInit(self, event):
      """ Begin a Drag Operation """
      # Create a Text Data Object, which holds the text that is to be dragged
      tdo = wx.PyTextDataObject(self.text.GetStringSelection())
      # Create a Drop Source Object, which enables the Drag operation
      tds = wx.DropSource(self.text)
      # Associate the Data to be dragged with the Drop Source Object
      tds.SetData(tdo)
      # Initiate the Drag Operation
      tds.DoDragDrop(True)


class MyApp(wx.App):
   """ Define the Drag and Drop Example Application """
   def OnInit(self):
      """ Initialize the Application """
      # Declare the Main Application Window
      frame = MainWindow(None, -1, "TeX2Docx Generation Process.")
      # Show the Application as the top window
      self.SetTopWindow(frame)
      return True


# # Declare the Application and start the Main Loop
if (os.path.isfile(r"\\192.168.7.5\License\license.txt")):
   chklicense = open(r"\\192.168.7.5\License\license.txt", 'r').read()
   if (chklicense != 'Active'):
       alert(text='Please contact the tech support!', title='expired', button='OK')
       exit()
else:
    alert(text="Please check the \"internet\" or \"VPN\" connection", title='expired', button='OK')
    exit()


if __name__ == "__main__":
   # app = MyApp(0)
   # app.MainLoop()
   User_Input = os.path.abspath(input("Enter the tex file path with file name and extension: "))
   con_start_eq = eqprocess(User_Input)
   con_start = tex2docx(con_start_eq)

