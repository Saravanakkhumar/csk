%-----------------------------------------------------------------------------
%
%               LaTeX Class/Style File
%
% Name:         acm-book.cls
%
% Purpose:      A LaTeX 2e class file for the Association for Computing
%               Machinery book series design. The goal is to have authors'
%               manuscripts look something like the final design.
%
% Author:       Paul C. Anagnostopoulos
%               Windfall Software
%               978 371-2316
%               paul [at] windfall.com
%
% Created:      paul@windfall.com
%
% Revisions:    See end of file.
%
% Copyright (c) 2019 by the Association for Computing Machinery
%
%-----------------------------------------------------------------------------

%http://tex.loria.fr/ctan-doc/macros/latex/doc/html/fntguide/node27.html

\NeedsTeXFormat{LaTeX2e}[1995/12/01]
\ProvidesClass{acm-book-v2}[2019/2/19 2.1 ACM book series design]
\RequirePackage{iftex}
\ifPDFTeX\ClassError{acm-book-v2}{You could not use PDFLaTeX.\MessageBreak
                                  Use LuaLaTeX}{Please Press "Ctrl+Z" to return\MessageBreak
                                  and Re-compile in LuaLaTeX.}
\fi

\ifxetex\ClassError{acm-book-2}{You could not use XeLaTeX.\MessageBreak
                                Use LuaLaTeX}{Please Press "Ctrl+Z" to return\MessageBreak
                                and Re-compile in LuaLaTeX.}
\fi

\def\thickspace{\quad}
\RequirePackage{etoolbox}
\RequirePackage{fixltx2e}
\RequirePackage{ragged2e}
\RequirePackage{cmap}
%\usepackage[noTeX]{mmap}
%\RequirePackage[LY1]{fontenc}
%\RequirePackage{textcomp}
%\RequirePackage{nfssext}
\RequirePackage{graphicx}%
\ifluatex\else\RequirePackage[tbtags]{amsmath}\fi
\RequirePackage{amsthm}
\RequirePackage{array}
\RequirePackage{booktabs}
\RequirePackage{dcolumn}
\newcolumntype{P}[1]{>{\raggedright\let\\\@arraycr}p{#1}}
\newcolumntype{d}[1]{D{.}{.}{#1}} 
\RequirePackage{url}
\urlstyle{same}
\RequirePackage{floatpag}
\RequirePackage{stfloats}
\RequirePackage{xkeyval}
\ifluatex\else\RequirePackage[psamsfonts]{amssymb,amsfonts}\fi
\RequirePackage{enumerate}
\RequirePackage[figuresright]{rotating}
\RequirePackage{color}
\RequirePackage{colortbl}%
\RequirePackage[table]{xcolor}%
\RequirePackage{threeparttable}
\RequirePackage{multicol}
\RequirePackage{alltt}
%\RequirePackage{fixcolumn}
\RequirePackage{framed}
\RequirePackage{twoopt}
\RequirePackage{marginnote}
\RequirePackage{tipa}
\RequirePackage{environ}
\RequirePackage{leading}
\RequirePackage{type1cm}
\RequirePackage{noindentafter}


%%%
\setlength{\cmidrulewidth}{0.5\p@}
\definecolor{tabmidrule}{cmyk}{0,0,0,0}
\arrayrulecolor{tabmidrule}


% The following few pages contain LaTeX programming extensions adapted
% from the ZzTeX macro package.

%                       Token Hackery
%                       ----- -------

\def \@expandaftertwice {\expandafter\expandafter\expandafter}
\def \@expandafterthrice {\expandafter\expandafter\expandafter\expandafter
                          \expandafter\expandafter\expandafter}

% This macro discards the next token.

\def \@discardtok #1{}%                                  token

% This macro removes the `pt' following a dimension.

{\catcode `\p = 12 \catcode `\t = 12

\gdef \@remover #1pt{#1}

} % \catcode

% This macro extracts the contents of a macro and returns it as plain text.
% Usage: \expandafter\@defof \meaning\macro\@mark

\def \@defof #1:->#2\@mark{#2}

%                       Control Sequence Names
%                       ------- -------- -----


\def \@name #1{%                                        {\tokens}
  \csname \expandafter\@discardtok \string#1\endcsname}

\def \@withname #1#2{%                                  {\command}{\tokens}
  \expandafter#1\csname \expandafter\@discardtok \string#2\endcsname}

%                       Flags (Booleans)
%                       ----- ----------

% The boolean literals \@true and \@false are appropriate for use with
% the \if command, which tests the codes of the next two characters.

\def \@true {TT}
\def \@false {FL}

\def \@setflag #1=#2{\edef #1{#2}}%              \flag = boolean

%                       IF and Predicates
%                       -- --- ----------

% A "predicate" is a macro that returns \@true or \@false as its value.
% Such values are suitable for use with the \if conditional.  For example:
%
%   \if \@oddp{\x} <then-clause> \else <else-clause> \fi

% A predicate can be used with \@setflag as follows:
%
%   \@setflag \flag = {<predicate>}

% Here are the predicates for TeX's repertoire of conditional
% commands.  These might be more appropriately interspersed with
% other definitions in this module, but what the heck.
% Some additional "obvious" predicates are defined.

\def \@eqlp   #1#2{\ifnum #1 = #2\@true \else \@false \fi}
\def \@neqlp  #1#2{\ifnum #1 = #2\@false \else \@true \fi}
\def \@lssp   #1#2{\ifnum #1 < #2\@true \else \@false \fi}
\def \@gtrp   #1#2{\ifnum #1 > #2\@true \else \@false \fi}
\def \@zerop  #1{\ifnum #1 = 0\@true \else \@false \fi}
\def \@onep   #1{\ifnum #1 = 1\@true \else \@false \fi}
\def \@posp   #1{\ifnum #1 > 0\@true \else \@false \fi}
\def \@negp   #1{\ifnum #1 < 0\@true \else \@false \fi}
\def \@oddp   #1{\ifodd #1\@true \else \@false \fi}
\def \@evenp  #1{\ifodd #1\@false \else \@true \fi}
\def \@rangep #1#2#3{\if \@orp{\@lssp{#1}{#2}}{\@gtrp{#1}{#3}}\@false \else
                                                          \@true \fi}
\def \@tensp  #1{\@rangep{#1}{10}{19}}

\def \@dimeqlp   #1#2{\ifdim #1 = #2\@true \else \@false \fi}
\def \@dimneqlp  #1#2{\ifdim #1 = #2\@false \else \@true \fi}
\def \@dimlssp   #1#2{\ifdim #1 < #2\@true \else \@false \fi}
\def \@dimgtrp   #1#2{\ifdim #1 > #2\@true \else \@false \fi}
\def \@dimzerop  #1{\ifdim #1 = 0pt\@true \else \@false \fi}
\def \@dimposp   #1{\ifdim #1 > 0pt\@true \else \@false \fi}
\def \@dimnegp   #1{\ifdim #1 < 0pt\@true \else \@false \fi}

\def \@vmodep     {\ifvmode \@true \else \@false \fi}
\def \@hmodep     {\ifhmode \@true \else \@false \fi}
\def \@mathmodep  {\ifmmode \@true \else \@false \fi}
\def \@textmodep  {\ifmmode \@false \else \@true \fi}
\def \@innermodep {\ifinner \@true \else \@false \fi}

\long\def \@codeeqlp #1#2{\if #1#2\@true \else \@false \fi}

\long\def \@cateqlp #1#2{\ifcat #1#2\@true \else \@false \fi}

\long\def \@tokeqlp  #1#2{\ifx #1#2\@true \else \@false \fi}
\long\def \@xtokeqlp #1#2{\expandafter\ifx #1#2\@true \else \@false \fi}

\long\def \@definedp #1{%
  \expandafter\ifx \csname \expandafter\@discardtok \string#1\endcsname
                   \relax \@false \else \@true \fi}

\long\def \@undefinedp #1{%
  \expandafter\ifx \csname \expandafter\@discardtok \string#1\endcsname
                   \relax \@true \else \@false \fi}

\def \@emptydefp #1{\ifx #1\@empty \@true \else \@false \fi}%       {\name}

\let \@emptylistp = \@emptydefp

\long\def \@emptyargp #1{%                               {#n}
  \@empargp #1\@empargq\@mark}
\long\def \@empargp #1#2\@mark{%
  \ifx #1\@empargq \@true \else \@false \fi}
\def \@empargq {\@empargq}

\def \@emptytoksp #1{%                                   {\tokenreg}
  \expandafter\@emptoksp \the#1\@mark}

\long\def \@emptoksp #1\@mark{\@emptyargp{#1}}

\def \@voidboxp #1{\ifvoid #1\@true \else \@false \fi}
\def \@hboxp #1{\ifhbox #1\@true \else \@false \fi}
\def \@vboxp #1{\ifvbox #1\@true \else \@false \fi}

\def \@eofp #1{\ifeof #1\@true \else \@false \fi}


% Flags can also be used as predicates, as in:
%
%   \if \flaga <then-clause> \else <else-clause> \fi


% Now here we have predicates for the common logical operators.

\def \@notp #1{\if #1\@false \else \@true \fi}

\def \@andp #1#2{\if #1%
                  \if #2\@true \else \@false \fi
                \else
                  \@false
                \fi}

\def \@orp #1#2{\if #1%
                 \@true
               \else
                 \if #2\@true \else \@false \fi
               \fi}

\def \@xorp #1#2{\if #1%
                  \if #2\@false \else \@true \fi
                \else
                  \if #2\@true \else \@false \fi
                \fi}

\def \stringeql #1#2#3{%                {\return-flag}{string1}{string2}
  \edef \@zstra {#2}%
  \edef \@zstrb {#3}%
  \edef #1{\ifx \@zstra\@zstrb \@true \else \@false \fi}}

%                       Arithmetic
%                       ----------

\def \@increment #1{\advance #1 by 1\relax}%             {\count}

\def \@decrement #1{\advance #1 by -1\relax}%            {\count}

%                       New Expression Evaluation
%                       --- ---------- ----------


% These macros provide a rudimentary expression evaluation facility.
% The accumulator can be a count, dimen, or skip register. The expression
% is in the form {operand0,op1,operand1,op2,operand2,...}.
%
% WARNING: The \register cannot be used in the expression except possibly
%          as the first operand.


\newdimen{\@calct}


\def \@calculate #1=#2{%                        \register={expression}
  \let \@accum = #1%
  \@calca =,#2,;,,\zmark}

\def \@gcalculate #1=#2{%                       \register={expression}
  \let \@accum = #1%
  \@calca =,#2,;,,\zmark
  \global \@accum = \@accum}

\def \@calca #1,#2,#3\zmark{%
  \def \znext{\zcalca #3\zmark}%
       \if #1=\@accum = #2%                             % Assignment
  \else\if #1+\advance \@accum #2%                      % Addition
  \else\if #1-\advance \@accum -#2%                     % Subtraction
  \else\if #1*\multiply \@accum #2%                     % Multiplication
  \else\if #1/\divide \@accum #2%                       % Division
  \else\if #1N\if \dimlssp{#2}{\@accum}\@accum = #2\fi  % Minimum
  \else\if #1P\divide \@accum 100                       % Percentage
              \expandafter\@accum #2\@accum
  \else\if #1R\@calct = #2%                             % Ratio
              \divide \@calct by 4096
              \divide \@accum by \@calct
              \multiply \@accum by 16\relax
  \else\if #1S\@accum = \thefactor#2\@accum             % Scale
  \else\if #1X\if \dimgtrp{#2}{\@accum}\@accum = #2\fi  % Maximum
  \else\if #1;\let \znext = \relax                      % Done
  \else
    \ClassError{operator}{Invalid operator `#1' in expression}%
  \fi\fi\fi\fi\fi\fi\fi\fi\fi\fi\fi
  \znext}

%                       Options
%                       -------


\@setflag \@authoryear = \@true
\@setflag \@explicitsize = \@false
\@setflag \@mathtime = \@false
\@setflag \@natbib = \@true
\@setflag \@ninepoint = \@true
\newcount{\@numheaddepth} \@numheaddepth = 3
\@setflag \@preprint = \@false
\@setflag \@tenpoint = \@false
\@setflag \@times = \@false

% Note that all the dangerous book class options are trapped.

\DeclareOption{9pt}{\@setflag \@ninepoint = \@true
                    \@setflag \@explicitsize = \@true}

\DeclareOption{10pt}{\PassOptionsToClass{10pt}{book}%
                     \@setflag \@ninepoint = \@false
                     \@setflag \@tenpoint = \@true
                     \@setflag \@explicitsize = \@true}

\DeclareOption{11pt}{\PassOptionsToClass{11pt}{book}%
                     \@setflag \@ninepoint = \@false
                     \@setflag \@explicitsize = \@true}

\DeclareOption{12pt}{\@unsupportedoption{12pt}}

\DeclareOption{a4paper}{\@unsupportedoption{a4paper}}

\DeclareOption{a5paper}{\@unsupportedoption{a5paper}}

\DeclareOption{authoryear}{\@setflag \@authoryear = \@true}

\DeclareOption{b5paper}{\@unsupportedoption{b5paper}}

\DeclareOption{cm}{\@setflag \@times = \@false}

\DeclareOption{computermodern}{\@setflag \@times = \@false}

\DeclareOption{executivepaper}{\@unsupportedoption{executivepaper}}

\DeclareOption{landscape}{\@unsupportedoption{landscape}}

\DeclareOption{legalpaper}{\@unsupportedoption{legalpaper}}

\DeclareOption{letterpaper}{\@unsupportedoption{letterpaper}}

\DeclareOption{mathtime}{\@setflag \@mathtime = \@true}

\DeclareOption{natbib}{\@setflag \@natbib = \@true}

\DeclareOption{nonatbib}{\@setflag \@natbib = \@false}

\DeclareOption{notitlepage}{\@unsupportedoption{notitlepage}}

\DeclareOption{numberedpars}{\@numheaddepth = 4}

\DeclareOption{numbers}{\@setflag \@authoryear = \@false}

\DeclareOption{preprint}{\@setflag \@preprint = \@true}

\DeclareOption{times}{\@setflag \@times = \@true}

\DeclareOption{titlepage}{\@unsupportedoption{titlepage}}

\DeclareOption*{\PassOptionsToClass{\CurrentOption}{book}}

\ExecuteOptions{10pt,indentedstyle,times}
\@setflag \@explicitsize = \@false
\ProcessOptions

\PassOptionsToClass{twoside,onecolumn}{book}
\LoadClass{book}

\def \@unsupportedoption #1{%
  \ClassError{proc}{The standard '#1' option is not supported.}}

% This can be used with the 'reprint' option to get the final folios.

\def \setpagenumber #1{%
  \setcounter{page}{#1}}

\AtEndDocument{\label{sigplanconf@finalpage}}

%                       Utilities
%                       ---------


\newcommand{\setvspace}[2]{%
  #1 = #2
  \advance #1 by -1\parskip}

%                       Document Parameters
%                       -------- ----------


% Page:

\newdimen{\trimwidth}
\newdimen{\trimheight}

\setlength{\trimwidth}{7.5in}
\setlength{\trimheight}{9.25in}

\setlength{\hoffset}{.5in}
\setlength{\voffset}{0in}

%\setlength{\topmargin}{2.9pc}
\setlength{\topmargin}{2.25pc}
\setlength{\headheight}{7pt}
\setlength{\headsep}{28.5pt}

\setlength{\evensidemargin}{.71in}
\setlength{\oddsidemargin}{.98in}

% Text area:

\newdimen{\standardtextwidth}
\setlength{\standardtextwidth}{30pc}

\setlength{\textwidth}{30pc}

\setlength{\topskip}{8pt}
\setlength{\columnsep}{2pc}
\setlength{\textheight}{45pc}

% Running foot:
\setlength{\footskip}{30pt}

% Paragraphs:

\setlength{\parskip}{0pt}
\setlength{\parindent}{14pt}

\setlength{\lineskip}{.5pt}
\setlength{\lineskiplimit}{\lineskip}

\frenchspacing
\pretolerance = 400
\tolerance = \pretolerance
\setlength{\emergencystretch}{8pt}
\clubpenalty = 10000
\widowpenalty = 10000
\setlength{\hfuzz}{.5pt}

% Page break penalties
\@lowpenalty   51
\@medpenalty  151
\@highpenalty 301

% Set badness and tolerance
\vbadness=9999
\pretolerance=-1
\tolerance=9999

% Don't allow consecutive hyphens and hyphens at the bottom of the page.
\doublehyphendemerits 100000
\finalhyphendemerits  1000000


% Disable page breaks before equations, allow pagebreaks after
% equations and discourage widow lines before equations.
\displaywidowpenalty 500
\predisplaypenalty   10000
\postdisplaypenalty  0

% Allow breaking the page in the middle of a paragraph
\interlinepenalty 0

% Disallow breaking the page after a hyphenated line
\brokenpenalty 10000

% Hyphenation; don't split words into less than three characters
\lefthyphenmin=3
\righthyphenmin=3


\setlength{\overfullrule}{0pt}

% Standard vertical spaces:

\newskip{\standardvspace}
\setvspace{\standardvspace}{5pt plus 1pt minus .5pt}

% Margin paragraphs:

\setlength{\marginparwidth}{36pt}
\setlength{\marginparsep}{2pt}
\setlength{\marginparpush}{8pt}


\setlength{\skip\footins}{8pt plus 3pt minus 1pt}
\setlength{\footnotesep}{10.1pt}

\renewcommand{\footnoterule}{%
  \begingroup\color{titlecolor}\hrule width .265\columnwidth height .33pt depth 0pt\endgroup\vskip0.8pt}

%\long\def\@makefnmark{%
%\hbox {{\normalfont \@thefnmark.\space}}}%

\long\def\@makefnmark{%
\hbox {{\normalfont\color{black}\textsuperscript{\@thefnmark}}}}%


%\makeatletter\show\@makefntext\makeatother
 \renewcommand{\@makefntext}[1]{%
   \noindent\@thefnmark.\space{#1}}

\def\footnotetextrelax{\let\@makefntext\relax}

%\renewcommand{\@makefntext}[1]{#1}

% Floats:

\setcounter{topnumber}{5}
\setcounter{bottomnumber}{5}
\setcounter{totalnumber}{10}

\renewcommand{\fps@figure}{tbph}
\renewcommand{\fps@table}{tbph}
\renewcommand{\topfraction}{0.93}
\renewcommand{\bottomfraction}{.93}
\renewcommand{\textfraction}{0.07}
\renewcommand{\floatpagefraction}{0.89}

\setcounter{dbltopnumber}{4}

\renewcommand{\dbltopfraction}{\topfraction}
\renewcommand{\dblfloatpagefraction}{\floatpagefraction}

\setlength\floatsep    {12\p@ \@plus 2\p@ \@minus 2\p@}
\setlength\textfloatsep{14\p@ \@plus 2\p@ \@minus 4\p@}
\setlength\intextsep   {18\p@ \@plus 2\p@ \@minus 2\p@}

\setlength\dblfloatsep    {16\p@ \@plus 2\p@ \@minus 2\p@}
\setlength\dbltextfloatsep{14\p@ \@plus 4\p@ \@minus 2\p@}

% \setlength{\floatsep}{28pt plus 4pt minus 2pt}
% \setlength{\textfloatsep}{28pt plus 4pt minus 3pt}
% \setlength{\intextsep}{10pt plus 4pt minus 3pt}
%
% \setlength{\dblfloatsep}{28pt plus 4pt minus 2pt}
% \setlength{\dbltextfloatsep}{20pt plus 4pt minus 3pt}




% Miscellaneous:

\errorcontextlines = 5

\DeclareTextCommandDefault{\NG}{\begingroup\fontencoding{T1}\selectfont\char141\endgroup}
\DeclareTextCommandDefault{\ng}{\begingroup\fontencoding{T1}\selectfont\char173\endgroup}

\DeclareFontEncoding{TS1}{}{}
\newcommand{\textlongs}{{%
\fontencoding{TS1}\fontfamily{lmr}\selectfont\char115%
}}

%\RequirePackage{lmodern-v1}%

%%%%%%%%%%%%Math Fixing%%%%%%%%%%%%%%%%%%%%%%%%%%
\newcommand\hmmax{0}
\newcommand\bmmax{0}

\usepackage{unnumberedtotoc}

\ifluatex\else

\usepackage{libertine}
\usepackage{libertinust1math}

\DeclareSymbolFont{Skfont}   {OT1}{lmr} {m}{n}%SK
\SetSymbolFont{Skfont}{bold}{OT1}{lmr} {bx}{n}

\DeclareSymbolFont{Skfontone}{OMS}{lmsy}{m}{n}%SK
\SetSymbolFont{Skfontone}{bold}{OMS}{lmsy} {b}{n}

\DeclareSymbolFont{Skfonttwo}{OML}{lmm}{m}{it}
\SetSymbolFont{Skfonttwo}{bold}{OML}{lmm}{b}{it}

\DeclareMathSymbol{=}{\mathrel}{Skfont}{"3D}
\DeclareMathSymbol{+}{\mathbin}{Skfont}{"2B}

%\DeclareSymbolFont{yhlargesymbols}{OMX}{yhex}{m}{n}
%\DeclareMathAccent{\widetilde}{\mathord}{yhlargesymbols}{"65}
%\DeclareMathAccent{\widehat}{\mathord}{yhlargesymbols}{"62}

\DeclareMathAccent{\acute}{\mathalpha}{Skfont}{"13}
\DeclareMathAccent{\tilde}{\mathalpha}{Skfont}{"7E}
\DeclareMathAccent{\hat}{\mathalpha}{Skfont}{"5E}
\DeclareMathAccent{\bar}{\mathalpha}{Skfont}{"16}
\DeclareMathAccent{\ddot}{\mathalpha}{Skfont}{"7F}
\DeclareMathAccent{\vec}{\mathord}{letters}{"7E}


\DeclareMathSymbol{\otimes}{\mathbin}{Skfontone}{"0A}
%\DeclareMathSymbol{\ell}{\mathord}{Skfontone}{"60}
\DeclareMathSymbol{\ldotp}{\mathpunct}{Skfonttwo}{"3A}
\DeclareMathSymbol{\alpha}{\mathord}{Skfonttwo}{"0B}
\DeclareMathSymbol{\beta}{\mathord}{Skfonttwo}{"0C}
\DeclareMathSymbol{\gamma}{\mathord}{Skfonttwo}{"0D}
\DeclareMathSymbol{\delta}{\mathord}{Skfonttwo}{"0E}
\DeclareMathSymbol{\epsilon}{\mathord}{Skfonttwo}{"0F}
\DeclareMathSymbol{\zeta}{\mathord}{Skfonttwo}{"10}
\DeclareMathSymbol{\eta}{\mathord}{Skfonttwo}{"11}
\DeclareMathSymbol{\theta}{\mathord}{Skfonttwo}{"12}
\DeclareMathSymbol{\iota}{\mathord}{Skfonttwo}{"13}
\DeclareMathSymbol{\kappa}{\mathord}{Skfonttwo}{"14}
\DeclareMathSymbol{\lambda}{\mathord}{Skfonttwo}{"15}
\DeclareMathSymbol{\mu}{\mathord}{Skfonttwo}{"16}
\DeclareMathSymbol{\nu}{\mathord}{Skfonttwo}{"17}
\DeclareMathSymbol{\xi}{\mathord}{Skfonttwo}{"18}
\DeclareMathSymbol{\pi}{\mathord}{Skfonttwo}{"19}
\DeclareMathSymbol{\rho}{\mathord}{Skfonttwo}{"1A}
\DeclareMathSymbol{\sigma}{\mathord}{Skfonttwo}{"1B}
\DeclareMathSymbol{\tau}{\mathord}{Skfonttwo}{"1C}
\DeclareMathSymbol{\upsilon}{\mathord}{Skfonttwo}{"1D}
\DeclareMathSymbol{\phi}{\mathord}{Skfonttwo}{"1E}
\DeclareMathSymbol{\chi}{\mathord}{Skfonttwo}{"1F}
\DeclareMathSymbol{\psi}{\mathord}{Skfonttwo}{"20}
\DeclareMathSymbol{\omega}{\mathord}{Skfonttwo}{"21}
\DeclareMathSymbol{\varepsilon}{\mathord}{Skfonttwo}{"22}
\DeclareMathSymbol{\vartheta}{\mathord}{Skfonttwo}{"23}
\DeclareMathSymbol{\varpi}{\mathord}{Skfonttwo}{"24}
\DeclareMathSymbol{\varrho}{\mathord}{Skfonttwo}{"25}
\DeclareMathSymbol{\varsigma}{\mathord}{Skfonttwo}{"26}
\DeclareMathSymbol{\varphi}{\mathord}{Skfonttwo}{"27}

\DeclareMathSymbol{\Sigma}{\mathalpha}{Skfont}{"06}

\DeclareMathSymbol{\partial}{\mathord}{Skfonttwo}{"40}

\DeclareMathSymbol{\Omega}{\mathalpha}{Skfont}{"0A}

\DeclareMathSymbol{\not}{\mathrel}{Skfontone}{"36}
\DeclareMathSymbol{\times}{\mathbin}{Skfontone}{"02}
\DeclareMathSymbol{\pm}{\mathbin}{Skfontone}{"06}
\DeclareMathSymbol{\times}{\mathbin}{Skfontone}{"02}
\DeclareMathSymbol{\nearrow}{\mathrel}{Skfontone}{"25}
\DeclareMathSymbol{\searrow}{\mathrel}{Skfontone}{"26}
\DeclareMathSymbol{\nwarrow}{\mathrel}{Skfontone}{"2D}
\DeclareMathSymbol{\swarrow}{\mathrel}{Skfontone}{"2E}
\DeclareMathSymbol{\leq}{\mathrel}{Skfontone}{"14}
\let\le=\leq
\DeclareMathSymbol{\geq}{\mathrel}{Skfontone}{"15}
\let\ge=\geq
\DeclareMathSymbol{-}{\mathbin}{Skfontone}{"00}

\DeclareMathSymbol{\geqq}{\mathrel}{AMSa}{61}
\DeclareMathSymbol{\leqq}{\mathrel}{AMSa}{53}
\DeclareMathSymbol{\nsupseteqq}{\mathrel}{AMSb}{35}
\DeclareMathSymbol{\varkappa}{\mathord}{AMSb}{123}

\DeclareMathSymbol{<}{\mathrel}{Skfonttwo}{"3C}
\DeclareMathSymbol{>}{\mathrel}{Skfonttwo}{"3E}
\DeclareMathSymbol{\simeq}{\mathrel}{Skfontone}{"27}
\DeclareMathSymbol{\equiv}{\mathrel}{Skfontone}{"11}
\DeclareMathSymbol{\land}{\mathbin}{Skfontone}{"5E}
\DeclareMathSymbol{\lor}{\mathbin}{Skfontone}{"5F}
\DeclareMathSymbol{\prec}{\mathrel}{Skfontone}{"1E}
\DeclareMathSymbol{\preceq}{\mathrel}{Skfontone}{"16}
\DeclareMathSymbol{\cap}{\mathbin}{Skfontone}{"5C}
\DeclareMathSymbol{\sqsubseteq}{\mathrel}{Skfontone}{"76}
\DeclareMathSymbol{\supseteq}{\mathrel}{Skfontone}{"13}
\DeclareMathSymbol{\subseteq}{\mathrel}{Skfontone}{"12}
\DeclareMathSymbol{\subset}{\mathrel}{Skfontone}{"1A}
\DeclareMathSymbol{\oplus}{\mathbin}{Skfontone}{"08}
\DeclareMathSymbol{\Rightarrow}{\mathrel}{Skfontone}{"29}

\DeclareMathSymbol{\exists}{\mathord}{Skfontone}{"39}
\DeclareRobustCommand
  \doteq{\buildrel\textstyle.\over=}
\DeclareMathSymbol{\circ}{\mathbin}{Skfontone}{"0E}

\DeclareMathSymbol{\sim}{\mathrel}{Skfontone}{"18}
\DeclareMathSymbol{\div}{\mathbin}{Skfontone}{"04}
\DeclareMathSymbol{\approx}{\mathrel}{Skfontone}{"19}
\DeclareMathSymbol{\bullet}{\mathbin}{Skfontone}{"0F}
\DeclareMathSymbol{\emptyset}{\mathord}{Skfontone}{"3B}
\DeclareMathSymbol{\forall}{\mathord}{Skfontone}{"38}
\DeclareMathSymbol{\Leftrightarrow}{\mathrel}{Skfontone}{"2C}
\DeclareMathSymbol{\cup}{\mathbin}{Skfontone}{"5B}
\DeclareMathSymbol{*}{\mathbin}{Skfontone}{"03} % \ast


\DeclareMathSymbol{\triangleq}{\mathrel}{AMSa}{"2C}
\DeclareMathSymbol{\wedge}{\mathbin}{Skfontone}{"5E}
\DeclareMathSymbol{\supset}{\mathrel}{Skfontone}{"1B}
\DeclareMathSymbol{\subseteq}{\mathrel}{Skfontone}{"12}
\DeclareMathSymbol{\vdash}{\mathrel}{Skfontone}{"60}
\DeclareMathSymbol{\cup}{\mathbin}{Skfontone}{"5B}
\DeclareMathSymbol{\vee}{\mathbin}{Skfontone}{"5F}
\DeclareMathSymbol{\neg}{\mathord}{Skfontone}{"3A}
\DeclareMathSymbol{\prec}{\mathrel}{Skfontone}{"1E}
\DeclareMathSymbol{\subset}{\mathrel}{Skfontone}{"1A}
\DeclareMathSymbol{\cap}{\mathbin}{Skfontone}{"5C}
\DeclareMathSymbol{\bot}{\mathord}{Skfontone}{"3F}
\DeclareMathSymbol{\dagger}{\mathbin}{Skfontone}{"79}
\DeclareMathSymbol{\gg}{\mathrel}{Skfontone}{"1D}
\DeclareMathSymbol{\ll}{\mathrel}{Skfontone}{"1C}
\DeclareMathSymbol{\oplus}{\mathbin}{Skfontone}{"08}
\DeclareMathSymbol{\subsetneq}{\mathrel}{AMSb}{"28}

\DeclareSymbolFont{AMSb}{U}{msb}{m}{n}
\DeclareSymbolFontAlphabet{\mathbb}{AMSb}
\DeclareSymbolFontAlphabet{\mathcal}{Skfontone}

\DeclareSymbolFont{letters}     {OML}{cmm} {m}{it}
\SetSymbolFont{letters}  {bold}{OML}{cmm} {b}{it}
\DeclareMathSymbol{\ell}{\mathord}{letters}{96}



%\DeclareSymbolFont{yhlargesymbols}{OMX}{yhex}{m}{n}
%\DeclareMathAccent{\widetilde}{\mathord}{yhlargesymbols}{"65}
%\DeclareMathAccent{\widehat}{\mathord}{yhlargesymbols}{"62}

\renewcommand{\ttdefault}{lmtt}
%\DeclareMathAlphabet{\mathtt}{OT1}{lmtt}{m}{n}
%\SetMathAlphabet{\mathtt}{bold}{OT1}{lmtt}{b}{n}

\SetMathAlphabet{\mathtt}{normal}{OT1}{lmtt}{m}{n}
\SetMathAlphabet{\mathtt}{bold}  {OT1}{lmtt}{m}{n}

\renewcommand{\neq}{\not=}

%\renewcommand*{\mathellipsis}{%
%  \mathinner{\ldotp\ldotp\ldotp}%
%}
%
%\let\dots\ldots

\DeclareRobustCommand
  \ldots{\mathinner{\ldotp\ldotp\ldotp}}
\DeclareRobustCommand
  \cdots{\mathinner{\cdotp\cdotp\cdotp}}
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
\fi

\ifluatex\else
\RequirePackage[lining,proportional]{ArnhemPro}
\RequirePackage[lining]{Flama}\fi

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

\ifluatex

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Font Configuration %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
\RequirePackage{libertinus-otf}
%\RequirePackage[no-math]{fontspec}
\defaultfontfeatures{Ligatures={TeX,NoCommon}}%

\newfontfamily{\ArnhemProSemiBold}{ArnhemPro-SemiBold}
\newfontfamily{\FlamaBlack}{Flama-Black}

\setmainfont{ArnhemPro-Blond}[ 
%  RawFeature     = {\libertinus@figurealign;\libertinus@figurestyle},
%  Scale          = \libertinusRM@scale ,
  UprightFont    = ArnhemPro-Blond,
  ItalicFont     = ArnhemPro-BlondItalic,
  ItalicFeatures = {SmallCapsFont = ArnhemPro-BlondItalic},
  SlantedFont    = ArnhemPro-Blond,
  SlantedFeatures= {FakeSlant=0.2},
  BoldFont       = ArnhemPro-Bold,
  BoldFeatures   = {SmallCapsFont = ArnhemPro-Bold},
  BoldItalicFont = ArnhemPro-BoldItalic,
  BoldItalicFeatures = {SmallCapsFont = ArnhemPro-BoldItalic},
  BoldSlantedFont= ArnhemPro-SemiBold,
  BoldSlantedFeatures= {FakeSlant=0.2, SmallCapsFont = ArnhemPro-SemiBold},
  SmallCapsFont  = ArnhemPro-Blond,
%  SmallCapsFeatures={RawFeature=+smcp},
  SmallCapsFeatures={Letters=SmallCaps}, 
  Extension      = .otf  
]


\setsansfont{Flama-Basic}[%
  UprightFont    = Flama-Basic,
  ItalicFont     = Flama-BasicItalic,
  %ItalicFeatures = {SmallCapsFont = Flama-BasicItalic},
  SlantedFont    = Flama-Basic,
  SlantedFeatures= {FakeSlant=0.2},
  BoldFont       = Flama-Bold,
  %BoldFeatures   = {SmallCapsFont = Flama-Bold},
  BoldItalicFont = Flama-BoldItalic,
  %BoldItalicFeatures = {SmallCapsFont = Flama-BoldItalic},
  Extension=.ttf]

\setmonofont[Mapping=tex-text]{Latin Modern Mono}

%%%\fontfamily{lmtt}\fontseries{b}\selectfont

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

\else

\fi

\RequirePackage{acmfonts}%
\RequirePackage{titletoc}
\RequirePackage{acmtoc}
\RequirePackage[utf8]{inputenc}

%\DeclareSymbolFont{largesymbols}{OMX}{cmex}{m}{n}

%                       Fonts
%                       -----

% \if \@times
%   \usepackage{times}%
% %%%  \renewcommand{\rmdefault}{ptm}%
%   \if \@mathtime
%     \usepackage[mtbold,noTS1]{mathtime}%
%   \else
%     \usepackage{mathptm}%
%   \fi
% \else
%   \relax
% \fi

\renewcommand{\normalsize}{%
  \@setfontsize{\normalsize}{10pt}{14pt}%
  \setlength{\abovedisplayskip}{12pt plus 3pt minus 4pt}%
  \setlength{\belowdisplayskip}{\abovedisplayskip}%
  \setlength{\abovedisplayshortskip}{12pt plus 3pt minus 4pt}%
  \setlength{\belowdisplayshortskip}{\abovedisplayshortskip}}

\renewcommand{\tiny}{\@setfontsize{\tiny}{5pt}{6pt}}

\renewcommand{\scriptsize}{\@setfontsize{\scriptsize}{7pt}{8pt}}

\newcommand{\tabnotesize}{\@setfontsize{\scriptsize}{8pt}{10pt}}

\renewcommand{\small}{%
  \@setfontsize{\small}{9pt}{12pt}%
  \setlength{\abovedisplayskip}{11pt plus 2pt minus 3pt}%
  \setlength{\belowdisplayskip}{\abovedisplayskip}%
  \setlength{\abovedisplayshortskip}{6pt plus 2pt minus 3pt}%
  \setlength{\belowdisplayshortskip}{\abovedisplayshortskip}}

\renewcommand{\footnotesize}{%
  \@setfontsize{\footnotesize}{8.5pt}{11.9pt}%
  \setlength{\abovedisplayskip}{10.5pt plus 2pt minus 4pt}%
  \setlength{\belowdisplayskip}{\abovedisplayskip}%
  \setlength{\abovedisplayshortskip}{5pt plus 2pt minus 2pt}%
  \setlength{\belowdisplayshortskip}{\abovedisplayshortskip}}

\renewcommand{\large}{\@setfontsize{\large}{12pt}{15pt}}

\renewcommand{\Large}{\@setfontsize{\Large}{14pt}{18pt}}

\renewcommand{\LARGE}{\@setfontsize{\LARGE}{24pt}{30pt}\raggedright}

%\renewcommand{\LARGE}{\@setfontsize{\LARGE}{18pt}{20pt}}

\newcommand{\partnamefont}{\@setfontsize{\partnamefont}{24pt}{26pt}}

\renewcommand{\huge}{\@setfontsize{\huge}{20pt}{25pt}}

\renewcommand{\Huge}{\@setfontsize{\Huge}{29pt}{32pt}\raggedright}

\newcommand{\parttitlefont}{\@setfontsize{\parttitlefont}{30pt}{35pt}\raggedright}

\newcommand{\monster}{\@setfontsize{\monster}{362pt}{362pt}}

%\newcommand{\partmonster}{\@setfontsize{\partmonster}{389pt}{389pt}}

\newcommand{\partmonster}{\@setfontsize{\partmonster}{279pt}{279pt}} %%%CODE NATION

\newcommand{\dedicationfont}{\@setfontsize{\dedicationfont}{12pt}{20pt}}

%\newcommand{\chapterauthorfont}{\@setfontsize{\chapterauthorfont}{14pt}{18.10pt}}%%%SK
\newcommand{\chapterauthorfont}{\@setfontsize{\chapterauthorfont}{12pt}{14pt}}%%%SK

%                       Bibliography
%                       ------------






\if \@natbib

\if \@authoryear
  \typeout{Using natbib package with 'authoryear' citation style.}
  \usepackage[authoryear,sort,square]{natbib}
  \bibpunct{[}{]}{,}{a}{}{,}    % Change fences to brackets;
                                % citation separator to semicolon;
                                % eliminate comma between author and year.
%  \bibliographystyle{acm-book}  % Used to be 'abbrvnat'.
  \let \cite = \citep
\else
  \typeout{Using natbib package with 'numbers' citation style.}
  \usepackage[numbers,sort&compress,square]{natbib}
\fi
\setlength{\bibsep}{3pt plus .5pt minus .25pt}

\fi

\def \bibfont {\small\raggedright}

%%% SK %%%
\renewenvironment{thebibliography}[1]
     {\addsec{References}\label{sec:bib\thechapter}%
     %\addtocontents{toc}{\protect\contentsline {chapter}{{Bibliography}}{\rmfamily\bfseries\thepage}{page.\thepage}}
    %\markright{References}
      \list{\@biblabel{\@arabic\c@enumiv}}%
           {\settowidth\labelwidth{\@biblabel{#1}}%
            \leftmargin\labelwidth
            \setlength{\leftmargin}{9\p@}%
     \setlength{\listparindent}{\parindent}%
     \setlength{\itemindent}{-9.5\p@}%
     \setlength{\topsep}{\medskipamount}%
     \setlength{\labelsep}{\z@}%
     \setlength{\labelwidth}{\z@}%
     \setlength{\itemsep}{\z@}
            \advance\leftmargin\labelsep
            %\itemsep = 2pt plus 1pt minus .5pt
            \@openbib@code
            \usecounter{enumiv}%
            \let\p@enumiv\@empty
            \renewcommand\theenumiv{\@arabic\c@enumiv}}%
      \bibfont
      \clubpenalty4000
      \@clubpenalty \clubpenalty
      \widowpenalty4000%
      \sfcode`\.\@m}
     {\def\@noitemerr
       {\@latex@warning{Empty `thebibliography' environment}}%
      \endlist}






%                       Enunciations
%                       ------------




\def\@begintheorem#1#2[#3]{%
  \deferred@thm@head{\the\thm@headfont \thm@indent
    \@ifempty{#1}{\let\thmname\@gobble}{\let\thmname\@iden}%
    \@ifempty{#2}{\let\thmnumber\@gobble}{\let\thmnumber\@iden}%
    \@ifempty{#3}{\let\thmnote\@gobble}{\let\thmnote\@iden}%
    \llap{\hbox to 5pc{\hss \textcolor{titlecolor}{\sans\bfseries\fontsize{8pt}{8pt}\selectfont\thm@swap\swappedhead\thmhead{#1}{#2}{#3}\hskip8\p@}}}%
%    \the\thm@headpunct
    \thmheadnl % possibly a newline.
    %\hskip\thm@headsep
  }%
  \ignorespaces}




 \def \@opargbegintheorem #1#2#3{%               {name}{number}{title}
   \trivlist
   \item[%
     \hskip\labelsep \textsc{#1\ #2}%
     \if \@notp{\@emptyargp{#3}}\nut (#3).\fi]%
   \itshape\selectfont
   \ignorespaces}

%                       Figures
%                       -------


\@setflag \@caprule = \@true


\newcommand{\nocaptionrule}{%
  \@setflag \@caprule = \@false}

\long\def \@makecaption #1#2{%
  \stringeql{\@figurep}{\@captype}{figure}%
  \if \@figurep
    \addvspace{6pt}
    \if \@caprule
      %\hrule width \hsize height .33pt
      \vspace{2.3pt}
    \fi
    {\softraggedright
     \noindent
     \llap{\hbox to 3pc{\hss \textcolor{titlecolor}{\sans\bfseries\fontsize{8pt}{8pt}\selectfont #1}\hspace{10.5pt}}}%
     \rmfamily\fontsize{9.5pt}{11.5pt}\selectfont#2\par}%
  \else
    {\softraggedright
     \noindent
     \llap{\hbox to 3pc{\hss \textcolor{titlecolor}{\sans\bfseries\fontsize{8pt}{8pt}\selectfont#1}\hspace{10.5pt}}}%
     \rmfamily\bfseries\fontsize{9.5pt}{11.5pt}\selectfont#2\par}%
    \addvspace{6pt}%
  \fi}

\def\caption{%
   \ifx\@captype\@undefined
     \@latex@error{\noexpand\caption outside float}\@ehd
     \expandafter\@gobble
   \else
     \refstepcounter\@captype
     \expandafter\@firstofone
   \fi
   {\@dblarg{\@caption\@captype}}%
}

%                       Hierarchy
%                       ---------


\setcounter{secnumdepth}{\@numheaddepth}

\def\fulltitle#1{%
	\thispagestyle{empty}%
	\noindent%
	\vbox to 0\p@{\vspace*{4pc}%
	\bfseries\sffamily\fontsize{24}{28}\selectfont\raggedright\color{titlecolor}#1}\vfill%
	\afterpage{\blankpage}
}

\newcount\bookseries@a@count%
\newcount\bookseries@b@count%
\newcount\bookseries@c@count%

\newcount\bookaffil@a@count%
\newcount\bookaffil@b@count%
\newcount\bookaffil@c@count%

\def\bookseries#1{%
	  \begingroup\expandafter\protected@xdef\csname bookseries\the\bookseries@a@count\endcsname{#1}%
	  \global\advance\bookseries@a@count by \@ne\endgroup%
	}

\def\bookaffil#1{%
	  \begingroup\expandafter\protected@xdef\csname bookaffil\the\bookaffil@a@count\endcsname{#1}%
	  \global\advance\bookaffil@a@count by \@ne\endgroup%
	  }


\def\printbookseries{\begingroup%
\pagestyle{empty}%
\RaggedRight\small ACM Books is a new series of high-quality 
books for the computer science community,{\break} published by 
ACM. ACM Books publications are widely distributed in both 
print and{\break} digital formats through booksellers and to libraries 
(and library consortia) and individual ACM members via the ACM 
Digital Library platform.\hfill\par\vspace{0.8\baselineskip}
  \bookseries@b@count\bookseries@a@count%
%  \advance\bookseries@b@count by -\@ne%
  \bookaffil@b@count\bookaffil@a@count%
  \@whilenum\bookseries@c@count<\bookseries@b@count\do{%
  \@whilenum\bookaffil@c@count<\bookaffil@b@count\do{%
   \begingroup\noindent\normalsize\color{titlecolor}\RaggedRight\csname bookseries\the\bookseries@c@count\endcsname\endgroup\unskip\newline%
    \global\advance\bookseries@c@count by \@ne%
     \begingroup\noindent\small\RaggedRight\csname bookaffil\the\bookaffil@c@count\endcsname\unskip\endgroup\newline\vspace{0.5\baselineskip}%
      \global\advance\bookaffil@c@count by \@ne%
  }%
  }%
  \global\bookseries@a@count=0\relax%
  \global\bookaffil@a@count=0\relax%
%  \let\cleardoublepage\clearpage%
\thispagestyle{empty}%
\endgroup}	%

% \def\printbookaffil{%
%   \bookaffil@b@count\bookaffil@a@count%
%   \wlog{ check bookaffil\the\bookaffil@a@count}
%   \@whilenum\bookaffil@c@count<\bookaffil@b@count\do{%
%     \noindent\csname bookaffil\the\bookaffil@c@count\endcsname\unskip\par%
%     \global\advance\bookaffil@c@count by \@ne%
%   }%
%   \global\bookaffil@a@count=0\relax
% }

 \def\editorinchief#1{\begingroup%
	\clearpage
 	\thispagestyle{empty}%
 	\vbox to 27.5\p@{\vspace*{1pt}%
 	\begingroup\bfseries\sffamily\fontsize{20}{22}\selectfont\noindent\color{titlecolor}ACM Books\endgroup}%
	\begingroup\hfill\break\rmfamily\fontsize{10}{12}\selectfont\color{titlecolor}Editors in Chief\endgroup\hfill\par\noindent#1\hfill\par\vspace*{0.5\baselineskip}}%\vspace*{-0.5\baselineskip}

 \def\foundingeditor#1{\begingroup%
	\begingroup\hfill\break\rmfamily\fontsize{10}{12}\selectfont\color{titlecolor}Founding Editor\endgroup\hfill\par\noindent#1\hfill\par\vspace*{0.5\baselineskip}}

\newenvironment{copyrightenv}
	{%
	\thispagestyle{empty}%
	%\pagebreak%
	\begingroup%
	\parskip0.5\baselineskip
	\RaggedRight\small%
	}{
%\pagebreak
\endgroup}


\newcommand{\EticalLabel}[1]{\gdef\@EticalLabel{#1}}

\newcommand{\subchapterhead}[1]{\gdef\@subchapterhead{#1}}

\def\@makechapterhead#1{\begingroup%
	\thispagestyle{empty}%
	\vspace*{-179.5pt}%
	\noindent%
	\llap{\vbox to 258pt{%
		\vss%\fontspec{Flama-Basic}\sffamily\bfseries
		\hspace*{21.8pc}\hbox{\ifluatex\FlamaBlack\else\sffamily\fontseries{black}\fontshape{n}\fi\monster\selectfont\textcolor{chapternumcolor}{\thechapter}}%\thechapter\FlamaBlackRegular
		\vskip -12pc}}%
		\sffamily\bfseries\Huge\color{titlecolor}#1\par\endgroup%
  \vskip 32pt%
}

    
%%Ethical Chapter head%%
\def\EthicalHeading{\let\@makechapterhead\@makeethicalchapterhead}

\def\subchapterheading{\let\@makechapterhead\@makesubchapterheading}

\def\@makeethicalchapterhead#1{\begingroup%
	\thispagestyle{empty}%
	\vspace*{-196.5pt}%
    %\vspace*{-179.5pt}%
	\noindent%
	\llap{\vbox to 258pt{%
		\vss
		\hspace*{20.8pc}\hbox{\sffamily\fontseries{black}\fontshape{n}\monster\selectfont\textcolor{chapternumcolor}{\thechapter}}%\thechapter\FlamaBlackRegular
		\vskip -12pc}}%
        \bgroup\sffamily\fontseries{n}\fontsize{15}{18}\selectfont\@EticalLabel\egroup\vspace*{26pt}\par
		\sffamily\bfseries\Huge\color{titlecolor}#1\par\endgroup%
  \vskip 32pt}


\def\@makesubchapterheading#1{\begingroup%
	\thispagestyle{empty}%
	\vspace*{-196.5pt}%
    %\vspace*{-179.5pt}%
	\noindent%
	\llap{\vbox to 258pt{%
		\vss
		\hspace*{20.8pc}\hbox{\sffamily\fontseries{black}\fontshape{n}\monster\selectfont\textcolor{chapternumcolor}{\thechapter}}%\thechapter\FlamaBlackRegular
		\vskip -12pc}}%
        \bgroup\sffamily\fontseries{n}\fontsize{15}{18}\selectfont\@subchapterhead\egroup\vspace*{26pt}\par
		\sffamily\bfseries\Huge\color{titlecolor}#1\par\endgroup%
  \vskip 32pt}

\def\@makeschapterhead#1{%
  \thispagestyle{empty}
   \pdfbookmark[0]{#1}{page.\thepage}%
    \vspace*{85pt}%
     \noindent
  {\color{titlecolor}\huge \sffamily\bfseries\RaggedRight #1}\par
  \vskip 27.5pt%
}

%%%%Author Macro%%%%
\let\@aulink\@empty
\newcommand{\aulink}[1]{\bgroup\textsuperscript{\fontsize{7\p@}{7\p@}\selectfont#1}\egroup}

\newcommand{\affiltext}[1]{\bgroup\fontsize{10\p@}{12\p@}\selectfont\space(#1)\egroup}

\newcount\aulist@a@count%
\newcount\aulist@b@count%
\newcount\aulist@c@count%

\RequirePackage{xifthen}
\NewDocumentCommand{\aulist}{O{}m}{%
\ifthenelse{\isempty{#1}}
{\expandafter\protected@xdef\csname aulist\the\aulist@a@count\endcsname{#2}%
    \global\advance\aulist@a@count by \@ne%
    \wlog{aulist\the\aulist@a@count}}%
{\expandafter\protected@xdef\csname aulist\the\aulist@a@count\endcsname{\hyperref[#1]{\color{titlecolor}#2}}%
    \global\advance\aulist@a@count by \@ne%
    \wlog{aulist\the\aulist@a@count}}%
}%

%\def\aulist#1{%
%  \expandafter\protected@xdef\csname aulist\the\aulist@a@count\endcsname{#1}%
%  \global\advance\aulist@a@count by \@ne%
%  \wlog{aulist\the\aulist@a@count}
%}%

\def\AuthorAnd{ and } 
\def\AuthorCommaSep{,}

\def\print@aulist{%
  \aulist@b@count\aulist@a@count%
   \ifnum\aulist@a@count=1
    \csname aulist\the\aulist@c@count\endcsname%
     \else
       \global\advance\aulist@b@count by -\@ne
        \@whilenum\aulist@c@count<\aulist@b@count\do{%
          \csname aulist\the\aulist@c@count\endcsname%
           \global\advance\aulist@c@count by \@ne%
            \ifnum\aulist@c@count=\aulist@b@count%
              \else%
 	        \AuthorCommaSep\space%\hfill\break%
            \fi%
  }%
  \ifnum\aulist@a@count=1%
   \else%
        \AuthorAnd
	 %\AuthorCommaSep\par
\csname aulist\the\aulist@c@count\endcsname%%
     \fi%
    \fi%
\global\aulist@a@count=0
\global\aulist@b@count=0
\global\aulist@c@count=0
}%


\newcount\affil@a@count%
\newcount\affil@b@count%
\newcount\affil@c@count%

\def\affils#1{%
  %\showthe\affil@a@count%
  \expandafter\protected@xdef\csname affil\the\affil@a@count\endcsname{#1}%
  \wlog{ affil\the\affil@a@count}
  \global\advance\affil@a@count by \@ne%
}%

\def\print@affil{%
  \affil@b@count\affil@a@count%
  \global\advance\affil@b@count by -\@ne
  \@whilenum\affil@c@count<\affil@b@count\do{%
    \csname affil\the\affil@c@count\endcsname;\space%
    \global\advance\affil@c@count by \@ne%
  }%
  \csname affil\the\affil@c@count\endcsname%
}%

%%%%%%%%%%%%%%%%%%%%%%
\newskip{\@aulistbelowskip}
\setvspace{\@aulistbelowskip}{10pt plus 5pt minus 1pt}

\def\chapterauthor{%
  \vbox to \z@{\vskip-29\p@
   \parbox{\textwidth}{\color{titlecolor}\RaggedRight\chapterauthorfont\rmfamily\bfseries\noindent\print@aulist\hfill\break}}
    \vskip\@aulistbelowskip
  }

\NoIndentAfterCmd{\chapterauthor}


\newskip{\@sectionaboveskip}
\setvspace{\@sectionaboveskip}{13pt plus 5pt minus 1pt}

\newskip{\@sectionbelowskip}
\setlength{\@sectionbelowskip}{1pt}%
%\setlength{\@sectionbelowskip}{13pt plus 5pt minus 1pt}%

\renewcommand{\section}{%
  \@startsection{section}%
    {2}%
    {\z@}%
    {-\@sectionaboveskip}%
    {\@sectionbelowskip}%
    {\sans\bfseries\large\raggedright}}

\newskip{\@subsectionaboveskip}
\setvspace{\@subsectionaboveskip}{13pt plus 3pt minus .5pt}

\newskip{\@subsectionbelowskip}
\setlength{\@subsectionbelowskip}{.01pt}%

\renewcommand{\subsection}{%%\hspace*{-1.75pt}\strut\vspace{-0.87\baselineskip}%
  \@startsection%
    {subsection}%
    {3}%
    {0pt}%
    {-\@subsectionaboveskip}%
    {\@subsectionbelowskip}%
    {\sans\bfseries\fontsize{10.75pt}{12.75pt}\selectfont\raggedright}}

\renewcommand{\subsubsection}{%
  \@startsection%
    {subsubsection}%
    {3}%
    {0pt}%
    {-\@subsectionaboveskip}
    {\@subsectionbelowskip}%
    {\sans\bfseries\raggedright\normalsize}}

\newskip{\@paragraphaboveskip}
\setvspace{\@paragraphaboveskip}{6pt plus 2pt minus 2pt}

\renewcommand{\paragraph}{%
  \@startsection%
    {paragraph}%
    {4}%
    {0pt}%
    {\@paragraphaboveskip}
    {-1em}%
    {\normalsize \bfseries \if \@times \itshape \fi}}

\def\@sect#1#2#3#4#5#6[#7]#8{%
  \ifnum #2>\c@secnumdepth
    %\let\@svsec\@empty
    \refstepcounter{#1}%
    \protected@edef\@svsec{\@seccntformat{#1}\relax}%
  \else
    \refstepcounter{#1}%
    \protected@edef\@svsec{\@seccntformat{#1}\relax}%
  \fi
  \@tempskipa #5\relax
  \ifdim \@tempskipa>\z@
    \begingroup
      #6{%
%%%%%%Section Triggering%%%%%%%
\if \@eqlp{#2}{2}
\@hangfrom{\raisebox{-15pt}[0pt][0pt]{\llap{\hbox to 3pc{%
                     \if \@eqlp{#2}{2}\fontsize{29pt}{31pt}\selectfont\unskip\else\fontsize{12pt}{12pt}\selectfont\fi%\LARGE
                     \hss\@svsec\hspace*{1.3pt}}}}}%
\else%
\@hangfrom{\llap{\hbox to 3pc{%
                     \if \@eqlp{#2}{2}\fontsize{29pt}{31pt}\selectfont\unskip\else\fontsize{12pt}{12pt}\selectfont\fi%\LARGE
                     \hss \@svsec\unskip}}}%
\fi
          \interlinepenalty \@M #8\@@par}%
%%%%%%Section Triggering%%%%%%%
    \endgroup
    \csname #1mark\endcsname{#7}%
    \addcontentsline{toc}{#1}{%
      \ifnum #2>\c@secnumdepth \else
        \protect\numberline{\csname the#1\endcsname}%
      \fi
      #7}%
  \else
    \def\@svsechd{%
      #6{\hskip #3\relax
      \@svsec #8}%
      \csname #1mark\endcsname{#7}%
      \addcontentsline{toc}{#1}{%
        \ifnum #2>\c@secnumdepth \else
          \protect\numberline{\csname the#1\endcsname}%
        \fi
        #7}}%
  \fi
  \@xsect{#5}}

%                       Lists
%                       -----


\setlength{\leftmargini}{26.5pt}
\setlength\leftmarginii{29pt}
\setlength\leftmarginiii{29pt}
\setlength\leftmarginiv{29pt}
\setlength{\labelsep}{5.5pt}

\setlength{\topsep}{\standardvspace}
\setlength{\itemsep}{1pt}
\setlength{\parsep}{3pt}

\renewcommand{\labelitemi}{{\fontsize{12pt}{12pt}\selectfont$\bullet$}\hskip5.5\p@}
%%%\renewcommand{\labelitemi}{{\small \centeroncapheight{\textbullet}}}
\renewcommand{\labelitemii}{\centeroncapheight{\rule{3.5pt}{3.5pt}\hskip5.5\p@}}
\renewcommand{\labelitemiii}{$-$}
\renewcommand{\labelitemiv}{{\Large \textperiodcentered}}

\renewcommand{\@listi}{%
  \leftmargin = \leftmargini
  \listparindent = 0pt}
%%%  \itemsep = 1pt
%%%  \parsep = 3pt}
%%%  \listparindent = \parindent}

\let \@listI = \@listi

\renewcommand{\@listii}{%
  \leftmargin = \leftmarginii
  \topsep = 1pt
  \labelwidth = \leftmarginii
  \advance \labelwidth by -\labelsep
  \listparindent = \parindent}

\renewcommand{\@listiii}{%
  \leftmargin = \leftmarginiii
  \labelwidth = \leftmarginiii
  \advance \labelwidth by -\labelsep
  \listparindent = \parindent}

\renewcommand{\@listiv}{%
  \leftmargin = \leftmarginiv
  \labelwidth = \leftmarginiv
  \advance \labelwidth by -\labelsep
  \listparindent = \parindent}

%                       Mathematics
%                       -----------


\def \theequation {\arabic{chapter}.\arabic{equation}}

%                       Miscellaneous
%                       -------------


\newcommand{\balancecolumns}{%
  \vfill\eject
  \global\@colht = \textheight
  \global\ht\@cclv = \textheight}

\newcommand{\nut}{\hspace{.5em}}

\newcommand{\softraggedright}{%
  \let \\ = \@centercr
  \leftskip = 0pt
  \rightskip = 0pt plus 12pt}

%                       Program Code
%                       ------- ----


\newcommand{\mono}[1]{%
  {\@tempdima = \fontdimen2\font
   \texttt{\spaceskip = 1.1\@tempdima #1}}}

%                       Running Heads and Feet
%                       ------- ----- --- ----


\def \ps@headings {%
  \def \@evenhead {\hskip-88\p@\@setfolio{\thepage}\hspace*{11.3pt}\fontsize{9.5pt}{9.5pt}\selectfont\leftmark\hfill}%
  \let \@evenfoot = \@empty
  \def \@oddhead {\hfill\fontsize{9.5pt}{9.5pt}\selectfont\LastMark{2e-right}\hspace*{10.5pt}\@setfolio{\thepage}}%
  \let \@oddfoot = \@empty
  \let \@mkboth = \markboth
  \def \chaptermark ##1{%
    \markboth{%
      \ifnum \c@secnumdepth >\m@ne
        \if@mainmatter
          \@chapapp\ \thechapter\kern7pt %\quad
        \fi
      \fi
      \emph{##1}}{}}%
  \def \sectionmark ##1{%
    \markright{%
      \ifnum \c@secnumdepth >\z@
        %\raisebox{15pt}{
	\thesection%}
	\nut
      \fi
      ##1}}}
\ps@headings

\let\@ChapterMonth\@empty
\let\@ChapterVolume\@empty
\let\@ChapterNumber\@empty

\newcommand{\ChapterMonth}[1]{\gdef\@ChapterMonth{#1}}
\newcommand{\ChapterVolume}[1]{\gdef\@ChapterVolume{Volume\space#1}}
\newcommand{\ChapterNumber}[1]{\gdef\@ChapterNumber{Number\space#1}}


\def \ps@customheading {%
  \def \@evenhead {\hskip-88\p@\@setfolio{\thepage}\hspace*{11.3pt}\fontsize{9.5pt}{9.5pt}\selectfont\leftmark\hfill}%
  \def \@evenfoot {\hskip-88\p@{\fontsize{9.5pt}{9.5pt}\selectfont Communications of ACM}\hfill{\fontsize{9.5pt}{9.5pt}\selectfont \ifx\@ChapterMonth\@empty\else\@ChapterMonth\fi\quad\ifx\@ChapterVolume\@empty\else\@ChapterVolume\fi\quad\ifx\@ChapterNumber\@empty\else\@ChapterNumber\fi}}
  \def \@oddhead {\hfill\fontsize{9.5pt}{9.5pt}\selectfont\rightmark\hspace*{10.5pt}\@setfolio{\thepage}}%
  \def \@oddfoot {\hskip-88\p@{\fontsize{9.5pt}{9.5pt}\selectfont \ifx\@ChapterMonth\@empty\else\@ChapterMonth\fi\quad\ifx\@ChapterVolume\@empty\else\@ChapterVolume\fi\quad\ifx\@ChapterNumber\@empty\else\@ChapterNumber\fi}\hfill{\fontsize{9.5pt}{9.5pt}\selectfont Communications of ACM}}%
  \let \@mkboth = \markboth
  \def \chaptermark ##1{%
    \markboth{%
      \ifnum \c@secnumdepth >\m@ne
        \if@mainmatter
          \@chapapp\ \thechapter\kern7pt %\quad
        \fi
      \fi
      \emph{##1}}{}}%
  \def \sectionmark ##1{%
    \markright{%
      \ifnum \c@secnumdepth >\z@
        %\raisebox{15pt}{
	\thesection%}
	\nut
      \fi
      ##1}}}


\def \ps@plain {%
  \let \@evenhead = \@empty
  \def \@evenfoot {\hfill \@setfolio{\thepage}\hfill}%
  \let \@oddhead = \@empty
  \def \@oddfoot {\hfill \@setfolio{\thepage}\hfill}%
  \let \@mkboth = \markboth}

\def \@setfolio #1{%
  {\sans\bfseries\small #1}}

%                       Sidebars
%                       --------


\newcounter{sidebarnumber}[chapter]
\newenvironment{sidebar}{%
 \bgroup\par\noindent\color{titlecolor}\rule{\hsize}{.5pt}\egroup\par
  \raisebox{0.5\p@}[\z@][\z@]{%
   \hskip-20\p@\makebox[\z@][r]{\bgroup%
    \hbox to 6pc{\hfill%
     \sans\bfseries\fontsize{8pt}{8pt}\selectfont%
      \noindent\refstepcounter{sidebarnumber}\color{titlecolor}Sidebar{\space}\thechapter.\arabic{sidebarnumber}}
     \egroup}
  }
  \nobreak%
  \noindent}%
{%
  \par
  \nobreak
  \vspace{-4\p@}
  \bgroup\noindent\color{titlecolor}\rule{\hsize}{.5pt}\egroup\par
  \vspace{6\p@ plus3\p@ minus3\p@}}

%                       Special Characters
%                       ------- ----------


\DeclareRobustCommand{\euro}{%
  \protect{\rlap{=}}{\sf \kern .1em C}}

%                       Utilities
%                       ---------


\newcommand{\centeroncapheight}[1]{%
  {\setbox\@tempboxa = \hbox{#1}%
   \@measurecapheight{\@tempdima}%         % Calculate ht(CAP) - ht(text)
   \advance \@tempdima by -\ht\@tempboxa   %           ------------------
   \divide \@tempdima by 2                 %                   2
   \raise \@tempdima \box\@tempboxa}}

\newbox{\@measbox}

\def \@measurecapheight #1{%                            {\dimen}
  \setbox\@measbox = \hbox{ABCDEFGHIJKLMNOPQRSTUVWXYZ}%
  #1 = \ht\@measbox}

\long\def \@titlefootnote #1#2{%
  \insert\footins{%
    \reset@font\footnotesize
    \interlinepenalty\interfootnotelinepenalty
    \splittopskip\footnotesep
    \splitmaxdepth \dp\strutbox \floatingpenalty \@MM
    \hsize\columnwidth \@parboxrestore
%%%    \protected@edef\@currentlabel{%
%%%       \csname p@footnote\endcsname\@thefnmark}%
    \color@begingroup
      \def \@makefnmark {\textsuperscript{\color{black}#1}}%
      \@makefntext{%
        \rule\z@\footnotesep\ignorespaces#2\@finalstrut\strutbox}%
    \color@endgroup}}

%                       LaTeX Modifications
%                       ----- -------------



\def \cleardoublepage {%
  \clearpage
  \ifodd \c@page \else
    \thispagestyle{empty}
    \hbox{}%
    \clearpage
  \fi}

\def \@seccntformat #1{%
  \llap{\hfill\color{titlecolor}\@name{\the#1}%
  \@expandaftertwice\@seccntformata \csname the#1\endcsname.\@mark\kern10\p@%
  }}

\def \@seccntformata #1.#2\@mark{%
  \if \@emptyargp{#2}\fi}


%%%%SK Edits
%\renewcommand{\thesection}{\raisebox{-15pt}{\thechapter.\arabic{section}}}
%\renewcommand{\thesubsection}{\raisebox{15pt}{\thesection}.\arabic{subsection}}

\newcommand{\ExternalLink}[2]{\hyperref[#1]{\textcolor{titlecolor}{#2}}}

\RequirePackage{hyperref}%

\hypersetup{%
colorlinks=true,%
allcolors=black,%
breaklinks=true,%
linktoc=page,%
linkcolor=titlecolor,%
urlcolor=titlecolor,%
citecolor=titlecolor,%
bookmarks=true,%
bookmarksopen=true,%
bookmarksdepth=5,%
bookmarksnumbered=true,%
psdextra,%
unicode
}
\RequirePackage{bookmark}

\expandafter\def\expandafter\UrlBreaks\expandafter{\UrlBreaks%  save the current one
\do\:\do\-\do\.\do\@\do\\\do\/\do\!\do\_\do\|\do\;\do\>\do\]\do\)\do\,\do\?\do\'\do+\do\=\do\#\do\a\do\b\do\c\do\d\do\e\do\f\do\g\do\h\do\i\do\j\do\k\do\l\do\m\do\n\do\o\do\p\do\q\do\r\do\s\do\t\do\u\do\v\do\w\do\x\do\y\do\z\do\A\do\B\do\C\do\D\do\E\do\F\do\G\do\H\do\I\do\J\do\K\do\L\do\M\do\N\do\O\do\P\do\Q\do\R\do\S\do\T\do\U\do\V\do\W\do\X\do\Y\do\Z}%
\RequirePackage{breakurl}
\RequirePackage{epstopdf}%

%\def\affil#1{{\gdef\@affil{\rmfamily\itseries\fontsize{10pt}{12pt}\selectfont#1}}

%%%%affil%%%%
\newcount\affil@a@count%
\newcount\affil@b@count%
\newcount\affil@c@count%

\newcommand{\affil}[1]{%
  \expandafter\protected@xdef\csname affil\the\affil@a@count\endcsname{#1}%
  \global\advance\affil@a@count by \@ne%
}%

\def\print@affil{%
  \affil@b@count\affil@a@count%
  \advance\affil@b@count by -\@ne%
  \@whilenum\affil@c@count<\affil@b@count\do{%
    \unskip\csname affil\the\affil@c@count\endcsname\\%
    \global\advance\affil@c@count by \@ne%
  }%
  \csname affil\the\affil@c@count\endcsname%
}%


%\newcommand{\affil}[1]{\gdef\@affil{\bgroup#1\egroup}}

\newcommand{\acmtag}[1]{\gdef\@acmtag{\bgroup\rmfamily\itshape\normalsize ACM Books\space#1\egroup\hfill}}

\newcommand{\halftitle}[1]{\gdef\@halftitle{\sans\itshape\fontsize{16pt}{22pt}\selectfont#1}}

\newenvironment{dedication}{\vspace*{149\p@}%
	\thispagestyle{empty}%
	\noindent\rmfamily\itshape\dedicationfont\RaggedRight}{}

%\newcommand{\dedcation}[1]{\gdef\@dedcation{\vspace*{149\p@}\noindent\rmfamily\itshape\dedicationfont#1}\@dedcation\thispagestyle{empty}}

%\newcommand{\dedcation}[1]{\gdef\@dedcation{\vspace*{149\p@}\noindent\rmfamily\itshape\dedicationfont#1}\@dedcation\thispagestyle{empty}}

\if@compatibility
\renewenvironment{titlepage}
    {%
      \cleardoublepage
      \if@twocolumn
        \@restonecoltrue\onecolumn
      \else
        \@restonecolfalse\newpage
      \fi
      \thispagestyle{empty}%
      %\setcounter{page}\z@
    }%
    {\if@restonecol\twocolumn \else \newpage \fi
    }
\else
\renewenvironment{titlepage}
    {%
      \cleardoublepage
      \if@twocolumn
        \@restonecoltrue\onecolumn
      \else
        \@restonecolfalse\newpage
      \fi
      \thispagestyle{empty}%
      %\setcounter{page}\@ne
    }%
    {\if@restonecol\twocolumn \else \newpage \fi
     \if@twoside\else
       % \setcounter{page}\@ne
     \fi
    }
\fi

%%%%author macro%%%%
\newcount\author@a@count%
\newcount\author@b@count%
\newcount\author@c@count%

\renewcommand{\author}[1]{%
  \expandafter\protected@xdef\csname author\the\author@a@count\endcsname{#1}%
  \global\advance\author@a@count by \@ne%
}%

 \def\print@author{%
   \author@b@count\author@a@count%
   \advance\author@b@count by -\@ne%
   \@whilenum\author@c@count<\author@b@count\do{%
     \unskip\csname author\the\author@c@count\endcsname\\%
     \global\advance\author@c@count by \@ne%
   }%
   \csname author\the\author@c@count\endcsname%
 }%

 \def\printauthor#1{\gdef\print@author{#1}}%


\newcommand\affiliation[1]{\bgroup\rmfamily\itshape\fontsize{10pt}{12pt}\selectfont \textcolor{black}{#1}\egroup}

\gdef\copyoftitle{}

  \renewcommand\maketitle{\begin{titlepage}%
  \let\footnotesize\small
  \let\footnoterule\relax
  \let \footnote \thanks
  \bgroup\null%\vfil
   \vskip 15\p@
    \color{titlecolor}\sans\bfseries\LARGE\@title\let\copyoftitle\@title\par
    \phantomsection
    \pdfbookmark[0]{\@title}{DocStart}        
     \egroup%
     \vskip 11.5\p@%
      \ifx\@halftitle\undefined\else\bgroup\RaggedRight\sans\bfseries\@halftitle\par\egroup
      \vskip 35.5\p@\fi%
    %{%\large
     %\lineskip .75em%
      %\begin{tabular}[t]{@{\hspace*{-14pt}}l}%
        \bgroup\raggedright\fontsize{16pt}{18pt}\selectfont\rmfamily\textcolor{titlecolor}{\print@author}\egroup
        %\bgroup\color{titlecolor}\raggedright\fontsize{16pt}{18pt}\selectfont\rmfamily\bfseries\print@author\egroup
      %\end{tabular}\par
      %\vskip 8.5\p@%
      %{\noindent\ifx\affil\undefined\else\bgroup\rmfamily\itshape\fontsize{10pt}{12pt}\selectfont\print@affil\egroup\fi\par}%}%
\vfill 
\raggedright\ifx\@acmtag\undefined\else\@acmtag\fi\par\vskip 
12\p@ 
\noindent%\includegraphics{g:/ACM_books/logo/acm-logo.pdf}%\includegraphics[scale=0.80]{g:/ACM_books/logo/acm-logo.pdf} 
\tooltipLOGO{\includegraphics{acm-logo.pdf}}{ACM Book 
\#54}[-80pt,-100pt]%\includegraphics[scale=0.80]{g:/ACM_books/logo/acm-logo.pdf}%%%online 
%\includegraphics{acm-logo.eps}%\includegraphics[scale=0.80]{g:/ACM_books/logo/acm-logo.pdf}%%%print 
  \@thanks
  \vfil\null
%  \pagenumbering{roman}
%  \afterpage{\blankpage}
  \end{titlepage}%
  \setcounter{footnote}{0}%
  \global\let\thanks\relax
  \global\let\maketitle\relax
  \global\let\@thanks\@empty
  \global\let\@author\@empty
  \global\let\@date\@empty
  \global\let\@title\@empty
  \global\let\title\relax
  \global\let\author\relax
  \global\let\date\relax
  \global\let\and\relax}

%\hypersetup{
%pdftitle={},%
%%pdfauthor={},%
%%pdfkeywords={},%
%%pdfsubject={\@abstract@meta}%
%}

%%%pdfdisplaydoctitle=true,  %%%%%%Sriniv
%%%pdfstartview={FitH},  %%%%%%Sriniv


%%%%Please check the part
%%%%Please check the part

%\renewcommand{\partname}{\textbf{\textsf{P{\kern4.8pt}A{\kern4.6pt}R{\kern4.7pt}T}}}
\newcommand{\partnameE}{\textbf{\textsf{P{\kern4.7pt}A{\kern4.7pt}R{\kern4.7pt}T}}} %%Book 8


%\renewcommand\part{\cleardoublepage%
%   \if@noskipsec \leavevmode \fi
%   \par
%   \addvspace{4ex}%
%   \@afterindentfalse
%   \secdef\@part\@spart}

   \def\@part[#1]#2{\begingroup%
	\thispagestyle{empty}
        \ifnum \c@secnumdepth >\m@ne
          \refstepcounter{part}%
          \addcontentsline{toc}{part}{\protect\numberline{\thepart}#1}%
        \else
          \addcontentsline{toc}{part}{#1}%
        \fi
       \begingroup
  %      \vspace*{40pt}
         \ifnum \c@secnumdepth >\m@ne
 		\vbox to \z@{\vspace*{-22.7pc}%
		\hbox to \z@{\llap{\vbox to 258pt{\hspace*{21.91pc}%%
   		%\vss%
     		\hbox{\partmonster\rmfamily\bfseries\textcolor{chapternumcolor}{\thepart}}}}%
     		\hspace*{37pt}\hbox to \z@{%
     		\llap{\raisebox{218pt}[\z@][\z@]{%
     		\textcolor{partcolor}{\fontfamily{Flama}\fontseries{sb}\fontsize{24pt}{25pt}\selectfont\partnameE}}}}
     		}}\endgroup
         \fi%\vspace*{-160pt}
 	\parindent \z@ %\raggedright                    % Commented \raggedright
        \interlinepenalty \@M
        \raisebox{189.5pt}[\z@][\z@]{\hspace*{0pc}\vbox to \z@{\sans\bfseries\parttitlefont#2}}\par%
        \markboth{}{}%
       \nobreak
       %\vskip 12ex
       \@afterheading%
       %\let\cleardoublepage\clearpage
       \endgroup}

%  \def\@part[#1]#2{%
%       \ifnum \c@secnumdepth >-2\relax
%         \refstepcounter{part}%
%         \addcontentsline{toc}{part}{\thepart\hspace{1em}#1}%
%       \else
%         \addcontentsline{toc}{part}{#1}%
%       \fi
%       \markboth{}{}%
%       \begingroup%\centering
%        \thispagestyle{empty}%
%        \parindent \z@
%        \interlinepenalty \@M
%          \ifnum \c@secnumdepth >-2\relax
% %         \vspace*{-145pt}%
%          \llap{\vbox to 258pt{%
%   		\hspace*{22pc}\hbox{\partmonster\rmfamily\bfseries\textcolor{chapternumcolor}{\thepart}}%
%   		\hspace*{-8.2pc}\hbox{%
%   		\raisebox{185pt}[0pt][0pt]{%
%   		\textcolor{partcolor}{\fontsize{24pt}{25pt}\selectfont\FlamaBlackRegular\partname}}}
%   		}}%
%        \fi
%        \vspace*{-168pt}%
%        \sans\parttitlefont#2\par\vfil\endgroup%
%   %     \@endpart
%        \let\cleardoublepage\clearpage
%        \vspace*{5.8pc}
%       }

\newcommand{\tablehead}[1]{{\fontsize{9.5pt}{11.5pt}\selectfont#1}}

\RequirePackage{afterpage}
\RequirePackage{makeidx}

\newcommand\blankpage{%
    \null
    \thispagestyle{empty}%
%    \addtocounter{page}{-1}%
    \newpage}

%%%%%%%%%%%Ligatures Control%%%%%%%%%%%%%%%%%%%%
\RequirePackage{microtype}
\microtypesetup{expansion=false}
\microtypesetup{protrusion=false}
\microtypesetup{spacing=false}
\microtypesetup{kerning=false}
\microtypesetup{tracking=false}
\DisableLigatures[f]{encoding = *, family = * }
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
\RequirePackage[pdftex]{eforms}%
\newcounter{infoLineNum}
\setcounter{infoLineNum}{0}
\newcommand{\infoInput}[2][2.5in]{%
  \stepcounter{infoLineNum}%
  \fbox{\makebox[150pt][l]{%
    \kern 4 pt
    \raisebox{.75ex}
      {
\textField[\Ff\FfMultiline\W0\BC{}\BG{}\TU{#2}]{name\theinfoLineNum}{#1}{20bp}

}%
  }}
    %\hrulefill
}
\def\radioButtonDefaults
{%
    \W{2.9}\S{S}\BC{0 0 0}\F{\FPrint}%
}%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%%%%%Query Definition Begin%%%%
%\newif\if@xmlconv \@xmlconvfalse
%\newif\if@setqueries \@setqueriestrue
%\RequirePackage{longtable}
%
%\RequirePackage{Queries}%
%
%\def\multicols#1{\col@number#1\relax
%  \let\old@mn@@@marginnote\@mn@@@marginnote
%  \long\def\@mn@@@marginnote[##1]##2[##3]{\begingroup\docolaction{\ifodd\value{page}\reversemarginpar\else\fi}{}{\ifodd\value{page}\else\reversemarginpar\fi}\old@mn@@@marginnote[{##1}]{##2}[{##3}]\endgroup}%
%  \ifnum\col@number<\tw@
%     \PackageWarning{multicol}%
%      {Using `\number\col@number'
%       columns doesn't seem a good idea.^^J
%       I therefore use two columns instead}%
%     \col@number\tw@ \fi
%  \ifnum\col@number>10
%     \PackageError{multicol}%
%      {Too many columns}%
%      {Current implementation doesn't
%       support more than 10 columns.%
%       \MessageBreak
%       I therefore use 10 columns instead}%
%     \col@number10 \fi
%     \ifx\@footnotetext\mult@footnotetext\else
%       \let\orig@footnotetext\@footnotetext
%       \let\@footnotetext\mult@footnotetext
%     \fi
%  \@ifnextchar[\mult@cols{\mult@cols[]}}
%\newcommand\queriesoff{%
%  \global\@setqueriesfalse%
%  %\let\query\@gobble%
%  %\let\theendnotes\relax%
%}
%\newcommand\querieson{\global\@setqueriestrue}
%\queriesoff
%
%  \AtBeginDocument{%
%   \typesetqueries[Q1\querylabelpunct]
% }
% \AtEndDocument{\AtEndDocument{%
%
%   \if@filesw
%     \immediate\closeout\query@out
%   \fi
% }}
%%%%%Query Definition End%%%%

\RequirePackage{wrapfig}
\RequirePackage{quoting}
%\RequirePackage[linesnumbered,noline]{algorithm2e}%,linesnumberednoline,lined,boxed,algochapter
%\RequirePackage{algpseudocode}
%\RequirePackage{algcompatible}
\RequirePackage{listings}
%\RequirePackage{CJKutf8}
\RequirePackage{book-2-custom}%%%%custom sty based on book styles
%\RequirePackage{trackchanges-off}

\AtBeginDocument{\setcounter{secnumdepth}{3}}

\renewcommand{\thesubsubsection}{\fontsize{10pt}{10pt}\selectfont\thesubsection.\arabic{subsubsection}}%\sffamily

\renewenvironment{quote}{\begin{quoting}[leftmargin=14pt,rightmargin=14pt,indentfirst=false]%%
	\setlength{\parindent}{\z@}%
	}{\end{quoting}}

%\RequirePackage{epigraph}

\RequirePackage{upgreek}

%%%% Srini
\RequirePackage{tcolorbox}
\tcbuselibrary{breakable,skins}
\RequirePackage{fnpos}
\RequirePackage{multirow}
\def\tablenotesize{\@setfontsize\tablenotesize{8\p@}{10\p@}}
\def\tabnotesalign{\leavevmode\kern-\scriptspace\kern-\labelsep}
%%%%%%%%%%%%%%%

\RequirePackage{pdfcomment}

%%Table notes alignment%%%
\gdef\tablenotsalign{\appto\TPTnoteSettings{\leftmargin\z@\rightmargin\z@}}
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

\newcommand{\epigraph}[2]{\begin{flushleft}
	\parindent\z@%
	\RaggedRight%
	\itshape%
	#1\unskip%
	\vspace*{-0.3\baselineskip}
	\end{flushleft}
%	\begin{minipage}{\textwidth}
	\bgroup\rmfamily\fontsize{9pt}{11pt}\selectfont#2\egroup\medskip\par
%	\end{minipage}
	}

\NoIndentAfterCmd{\epigraph}

%tabnotes
\def\TNmakelabel#1{%
  \gdef\@tempa{#1}%
  \ifx\@tempa\@empty%
 \else%
	\textsuperscript{#1}%
  \fi%
}%

\newenvironment{tabnotes}{%
  \vspace*{8pt}%- Adjust the spacing according to the style
  \list{}{%
\topsep\z@%
   \partopsep\z@
	\itemsep\z@%
	\parsep\z@%
	\itemindent\z@%
	\setlength\leftmargin{\z@}%
	\setlength\labelwidth{\z@}%
	\setlength\labelsep{\z@}%
	\rightskip\z@%
	\leftskip\z@%
%   \let\textbf\tablefootboldfont%
%   \let\textit\tablefootfont%
	\let\makelabel\TNmakelabel%
	\raggedright\rmfamily\tabnotesize\noindent\ignorespaces\unskip%
	\nobreakspace
  }%
}{\endlist}%

\newcommand{\Hbar}{\begingroup\fontencoding{T1}\selectfont\char141\endgroup}
\renewcommand{\NG}{\begingroup\fontencoding{T1}\selectfont\char141\endgroup}
\renewcommand{\ng}{\begingroup\fontencoding{T1}\selectfont\char173\endgroup}
\renewcommand{\ae}{\begingroup\fontencoding{T1}\selectfont\char230\endgroup}
\renewcommand{\AE}{\begingroup\fontencoding{T1}\selectfont\char198\endgroup}
\renewcommand{\O}{\begingroup\fontencoding{T1}\selectfont\char216\endgroup}
\renewcommand{\o}{\begingroup\fontencoding{T1}\selectfont\char248\endgroup}


\newcommand{\unnumtoc}[1]{\addtocontents{toc}{\protect\contentsline{chapter}{\color{titlecolor}#1}{\rmfamily\bfseries\thepage}{page.\thepage}}}%

\newcommand{\markleft}[1]{#1}%
\let\markleft\chaptermark
\usepackage[euler]{textgreek}
\newenvironment{tablehere}[1]
{\hfill\break\noindent\begin{minipage}{#1}\vspace{.5\baselineskip \@plus 2\p@ \@minus 2\p@}\parindent\z@\def\@captype{table}}
{\vspace{.5\baselineskip \@plus 2\p@ \@minus 2\p@}\end{minipage}}

%\RequirePackage{custom-tooltip}
\RequirePackage[all]{nowidow}
\DeclareRobustCommand
  \models{\mathrel{|}\joinrel\Relbar}

\newbox\artbox
\newcommand{\acmbookart}[2]{\global\setbox\artbox\hbox{\includegraphics[#1]{#2}}%
			    \ifdim\wd\artbox>360pt%
			     \hskip-98pt\hbox to 458pt{\hfill\includegraphics[#1]{#2}}%
			      \else%
			       \includegraphics[#1]{#2}\fi}

\RequirePackage{lscape}



%                       Revision History
%                       -------- -------


%  Date         Person  Ver.    Change
%  ----         ------  ----    ------

%  2014/02/06   PCA     0.1     Begin work on class file.

%  2014/02/15   PCA     0.1     Send first samples to Diane.

%  2014/02/24   PCA     1.0     Release initial version.

%  2014/02/25   PCA     1.1     Use natbib in author/date style as the
%                               default bibliography package.

%  2014/03/05   PCA     1.2     Change text leading to 14 points.
%                               Change citation fences to brackets.

%  2016/4/4     PCA     1.3     Use 'abbrvnat' instead of 'plainnat'
%                               biblio style.

%  2016/5/13    PCA     1.4     Use new 'mcp-acm' biblio style.

%  2019/1/10    PCA     2.0     Rename this class to acm-book.cls.
%                               Rename biblio style file to acm-book.bst.

%  2019/2/19    PCA     2.1     Add sidebar environment.

%\NewDocumentCommand{\tooltipLOGO}{%
  %ssssO{\ifdefined\@linkcolor\@linkcolor\else blue\fi}mO{yellow!20}mO{0pt,0pt}%
%}{{%
%  \leavevmode%
%  \IfBooleanT{#2}{%
%    %for variants with two and more stars, put tip box on a PDF Layer (OCG)
%    \ocgbase@new@ocg{tipOCG.\thetcnt}{%
%      /Print<</PrintState/OFF>>/Export<</ExportState/OFF>>%
%    }{false}%
%    \xdef\tpTipOcg{\ocgbase@last@ocg}%
%    %prevent simultaneous visibility of multiple non-draggable tooltips
%    \ocgbase@add@ocg@to@radiobtn@grp{tool@tips}{\ocgbase@last@ocg}%
%  }%
%  \tpPdfLink{%
%    \IfBooleanTF{#4}{%
%      /Subtype/Link/Border[0 0 0]/A <</S/SetOCGState/State [/Toggle \tpTipOcg]>>
%    }{%
%      /Subtype/Screen%
%      /AA<<%
%        \IfBooleanTF{#3}{%
%          /E<</S/SetOCGState/State [/Toggle \tpTipOcg]>>%
%        }{%
%          \IfBooleanTF{#2}{%
%            /E<</S/SetOCGState/State [/ON \tpTipOcg]>>%
%            /X<</S/SetOCGState/State [/OFF \tpTipOcg]>>%
%          }{
%            \IfBooleanTF{#1}{%
%              /E<</S/JavaScript/JS(%
%                var fd=this.getField('tip.\thetcnt');%
%                if(typeof(click\thetcnt)=='undefined'){%
%                  var click\thetcnt=false;%
%                  var fdor\thetcnt=fd.rect;var dragging\thetcnt=false;%
%                }%
%                if(fd.display==display.hidden){%
%                  fd.delay=true;fd.display=display.visible;fd.delay=false;%
%                }else{%
%                  if(!click\thetcnt&&!dragging\thetcnt){fd.display=display.hidden;}%
%                  if(!dragging\thetcnt){click\thetcnt=false;}%
%                }%
%                this.dirty=false;%
%              )>>%
%            }{%
%              /E<</S/JavaScript/JS(%
%                var fd=this.getField('tip.\thetcnt');%
%                if(typeof(click\thetcnt)=='undefined'){%
%                  var click\thetcnt=false;%
%                  var fdor\thetcnt=fd.rect;var dragging\thetcnt=false;%
%                }%
%                if(fd.display==display.hidden){%
%                  fd.delay=true;fd.display=display.visible;fd.delay=false;%
%                }%
%               this.dirty=false;%
%              )>>%
%              /X<</S/JavaScript/JS(%
%                if(!click\thetcnt&&!dragging\thetcnt){fd.display=display.hidden;}%
%                if(!dragging\thetcnt){click\thetcnt=false;}%
%                this.dirty=false;%
%              )>>%
%            }%
%            /U<</S/JavaScript/JS(click\thetcnt=true;this.dirty=false;)>>%
%            /PC<</S/JavaScript/JS (%
%              var fd=this.getField('tip.\thetcnt');%
%              try{fd.rect=fdor\thetcnt;}catch(e){}%
%              fd.display=display.hidden;this.dirty=false;%
%            )>>%
%            /PO<</S/JavaScript/JS(this.dirty=false;)>>%
%          }%
%        }%
%      >>%
%    }%
%  }{{\color{#5}#6}}%
%  \sbox\tiptext{%
%    \IfBooleanT{#2}{%
%      \ocgbase@oc@bdc{\tpTipOcg}\ocgbase@open@stack@push{\tpTipOcg}}%
%    %\fcolorbox{black}{#7}{#8}%
%    \tcbox[colframe=black,colback=#7,size=fbox,arc=0ex,
%		sharp corners=southwest]{\parbox{0.18\textwidth}{\raggedright #8}}%
%    \IfBooleanT{#2}{\ocgbase@oc@emc\ocgbase@open@stack@pop\tpNull}%
%  }%
%  \cListSet\tpOffsets{#9}%
%  \edef\twd{\the\wd\tiptext}%
%  \edef\tht{\the\ht\tiptext}%
%  \edef\tdp{\the\dp\tiptext}%
%  \tipshift=12pt%
%  \IfBooleanTF{#2}{%
%    %OCG-based (that is, all non-draggable) boxes should not extend beyond the
%    %current column as they may get overlaid by text in the neighbouring column
%    \setlength\whatsleft{\linegoal}%
%  }{%
%    \measureremainder{\whatsleft}%
%  }%
%  \ifdim\whatsleft<\dimexpr\twd+\cListItem\tpOffsets{1}\relax%
%    \setlength\tipshift{\whatsleft-\twd-\cListItem\tpOffsets{1}}\fi%
%  \IfBooleanF{#2}{\tpPdfXform{\tiptext}}%
%  \raisebox{\heightof{#6}+\tdp+\cListItem\tpOffsets{2}}[0pt][0pt]{%
%    \makebox[0pt][l]{\hspace{\dimexpr\tipshift+\cListItem\tpOffsets{1}\relax}%
%    \IfBooleanTF{#2}{\usebox{\tiptext}}{%
%      \tpPdfAnnot{\twd}{\tht}{\tdp}{%
%        /Subtype/Widget/FT/Btn/T (tip.\thetcnt)%
%        /AP<</N \tpPdfLastXform>>%
%        /MK<</TP 1/I \tpPdfLastXform/IF<</S/A/FB true/A [0.0 0.0]>>>>%
%        /Ff 65536/F 3%
%        /AA <<%
%          /U <<%
%            /S/JavaScript/JS(%
%              var fd=event.target;%
%              var mX=this.mouseX;var mY=this.mouseY;%
%              var drag=function(){%
%                var nX=this.mouseX;var nY=this.mouseY;%
%                var dX=nX-mX;var dY=nY-mY;%
%                var fdr=fd.rect;%
%                fdr[0]+=dX;fdr[1]+=dY;fdr[2]+=dX;fdr[3]+=dY;%
%                fd.rect=fdr;mX=nX;mY=nY;%
%              };%
%%              if(!dragging\thetcnt){%
%%                dragging\thetcnt=true;Int=app.setInterval("drag()",1);%
%%              }%
%%              else{app.clearInterval(Int);dragging\thetcnt=false;}%
%              this.dirty=false;%
%            )%
%          >>%
%        >>%
%      }%
%      \tpAppendToFields{\tpPdfLastAnn}%
%    }%
%  }}%
%  \stepcounter{tcnt}%
%}}
\newenvironment{dijkstra}	
	{\begingroup%
		\fontfamily{Dijkstra-TLF}\fontencoding{LY1}\fontseries{m}\fontshape{n}\selectfont%
	}%
	{%
	\endgroup%
	}

%%%%%%%%%% Index See and Seealso dot seperator %%%%%%%%%%%%%%%%

\makeatletter
      \def\HyInd@@@wrindex#1#2#3\\{%
        \def\Hy@temp@A{#2}%
        \ifcase0\ifx\Hy@temp@A\HyInd@ParenLeft 1\fi
                \ifx\Hy@temp@A\HyInd@ParenRight 1\fi
                \relax
          \protected@write\@indexfile{}{%
            \string\indexentry{%
              %#1|hyperindexformat{/#2#3}%
              #1|#2#3%Srinivasan - 20Oct2022
            }{\thepage}%
          }%
        \else
          \ifx\\#3\\%
            \ifx\Hy@temp@A\HyInd@ParenRight
              \HyInd@DefKey{#1}%
              \ltx@IfUndefined{HyInd@(\HyInd@key)}{%
                \let\Hy@temp\ltx@empty
              }{%
                \expandafter\let\expandafter\Hy@temp
                \csname HyInd@(\HyInd@key)\endcsname
              }%
              \protected@write\@indexfile{}{%
                 \string\indexentry{#1|#2\Hy@temp}{\thepage}%
              }%
            \else
              \protected@write\@indexfile{}{%
                 \string\indexentry{#1|#2hyperpage}{\thepage}%
              }%
              \HyInd@DefKey{#1}%
              \expandafter
              \gdef\csname HyInd@(\HyInd@key)\endcsname{%
                hyperpage%
              }%
            \fi
          \else
            \protected@write\@indexfile{}{%
               \string\indexentry{%
                 #1|#2hyperindexformat{/#3}%
               }{\thepage}%
            }%
            \ifx\Hy@temp@A\HyInd@ParenLeft
              \HyInd@DefKey{#1}%
              \expandafter
              \gdef\csname HyInd@(\HyInd@key)\endcsname{%
                hyperindexformat{/#3}%
              }%
            \fi
          \fi
        \fi
      }
\makeatother

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%%%%%%%%%%%%%%%%Index page set in Odd page%%%%%%%%%%%%%%%%%%%%%%%%%%%%
\makeatletter
\renewcommand\printindex{
\IfFileExists{\jobname.mst}{\ifodd \c@page%
    \@input@{\jobname.ind}%
    \ifodd\c@page
    \else%
        \afterpage{\blankpage}
    \fi%
\else
    \afterpage{\blankpage}
    \@input@{\jobname.ind}
    \ifodd\c@page
    \else%
        \afterpage{\blankpage}
    \fi%
\fi%
}
{\ClassError{acm-book-2}{!MISSING \jobname.mst FILE MISSING}{Please copy the MST file from some other book.}}
}
\makeatother

%%%%%%%%%%%% If Subsection immediately followed section space given

%\makeatletter
%\def\@afterheading{%
%  \@nobreakfalse
%  \everypar{%
%    \if@nobreak
%      \@nobreakfalse
%      \clubpenalty \@M
%      \if@afterindent \else
%        {\setbox\z@\lastbox
%         %%%\hrule\vspace{1ex}%<---
%        \noindent}%
%      \fi
%    \else
%      \clubpenalty \@clubpenalty
%      \hspace*{-17pt}
%      \everypar{}%
%    \fi}
%}
%\makeatother

\def\dummytext{\phantom{0}\vspace*{-1.5em}\nolinebreak}

%\newcommand{\pageinfoStart}[1]{\marginpar{\textcolor{red}{\CurrentFile-PID:#1-PST:\thepage}}}

\newwrite\testdboutput
	\immediate\openout\testdboutput=\jobname.paginfo
	\newcommand{\pageinfoStart}[1]{\immediate\write\testdboutput{\CurrentFile-PID:#1-PST:\thepage}}
