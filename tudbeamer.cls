%
% Praesentationsfolien-Klasse fuer die TU Dresden
% Autoren: Lukas Baron, Klaus Bergmann
%
\def\fileversion{1}
\def\filedate{2018/02/05}
\def\filename{tudbeamer}

\NeedsTeXFormat{LaTeX2e}  
\ProvidesClass{\filename}[\filedate Praesentationsfolien-Klasse fuer die TU Dresden]
\typeout{Class: '\filename' Version \fileversion, Praesentationsfolienvorlage der Technischen Universitaet Dresden}

\RequirePackage{kvoptions}
\SetupKeyvalOptions{%
    family=tud,%
    prefix=tud@%
}

\newif\if@useGerman                             % explizit deutsche Namen verwenden und german-Package einbinden ?
\newif\if@useNoGerman                           % explizit englische Namen verwenden
\DeclareOption{nogerman}{\@useNoGermantrue\@useGermanfalse}
\DeclareOption{german}{\@useGermantrue\@useNoGermanfalse}
\DeclareOption{ngerman}{\@useGermantrue\@useNoGermanfalse}

%\newif\if@ifaBibliographyStyle\@ifaBibliographyStylefalse
%\DeclareOption{bibIfa}{\@ifaBibliographyStyletrue}

\newif\if@useNavBar\@useNavBarfalse
\DeclareOption{navbar}{\@useNavBartrue}

\newif\if@useProgress\@useProgresstrue
\DeclareOption{noprogress}{\@useProgressfalse}

% \newif\if@insertAdditionalLogo\@insertAdditionalLogotrue
% \DeclareVoidOption{noAdditionalLogo}{\@insertAdditionalLogofalse}
% \newif\if@insertAdditionalLogoFooter\@insertAdditionalLogoFootertrue
% \DeclareVoidOption{noAdditionalLogoFooter}{\@insertAdditionalLogoFooterfalse}

\newif\if@sansMath\@sansMathtrue
\DeclareOption{serifmath}{\@sansMathfalse}

\newif\if@useSectionFrame\@useSectionFrametrue
\DeclareOption{nosectionframe}{\@useSectionFramefalse}
\newif\if@useSectionTOCFrame\@useSectionTOCFrametrue
\DeclareOption{noSectiontoc}{\@useSectionTOCFramefalse}
\newif\if@usePartFrame\@usePartFrametrue
\DeclareOption{nopartframe}{\@usePartFramefalse}
\newif\if@usePartTOCFrame\@usePartTOCFrametrue
\DeclareOption{noparttoc}{\@usePartTOCFramefalse}
\newif\if@useTOCFrame\@useTOCFrametrue
\DeclareOption{notoc}{\@useTOCFramefalse}

\newif\if@useSectionNumInFrameTitle\@useSectionNumInFrameTitletrue
\DeclareOption{noframetitlesectionnum}{\@useSectionNumInFrameTitlefalse}

\DeclareStringOption{titlestyle}[tud] %option is tud@titlestyle
\DeclareStringOption{partstyle}[tud] %option is tud@partstyle

\DeclareOption*{\PassOptionsToClass{\CurrentOption}{beamer}}
\ProcessKeyvalOptions*
\ProcessOptions\relax

%
% additional settings
%
\newcommand*\tud@affiliation{}
\newcommand*{\affiliation}[1]{\renewcommand*{\tud@affiliation}{#1}}
\newcommand*{\datecity}[1]{\renewcommand*{\insertdatecity}{#1}}

\newcommand*\tud@additionallogo{}
\newcommand*{\additionallogo}[1]{\renewcommand*{\tud@additionallogo}{#1}}
\newcommand*\tud@additionallogofooter{}
\newcommand*{\additionallogofooter}[1]{\renewcommand*{\tud@additionallogofooter}{#1}}
\newcommand*\tud@designimage{}
\newcommand*{\setdesignimage}[1]{\renewcommand*{\tud@designimage}{#1}}

\def\today{\number\day.\number\month.\number\year}


\PassOptionsToPackage{table,fixpdftex,svgnames,hyperref}{xcolor}

%
% load main class
%
\LoadClass[]{beamer}

%
% language and font packages
%
\usepackage[default,scale=.95]{opensans} %scales the letter spacing, default refers to the printing style for numbers
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage[english, main=ngerman]{babel}
\RequirePackage{csquotes}          % Ergänzungspaket zu Babel für erweiterte Zitierfunktionen

%\DeclareLanguageMapping{ngerman}{ngerman-apa}
%\DeclareLanguageMapping{german}{german-apa}
%\DeclareLanguageMapping{english}{english-apa}

%
% other packages
%
\RequirePackage{pifont}
\RequirePackage{caption}      % Paket zum Einbinden von Captions bei Nicht-Float-Objekten (hauptsächlich für Anhang)
\RequirePackage[	 % Verwendung von biblatex für das Literaturverzeichnis
    style=numeric-comp,%
    backend=biber,  %
	natbib=true,%
    sorting=none%
]{biblatex}

\RequirePackage{zref}
\RequirePackage{graphicx}
\RequirePackage{ifthen}
\RequirePackage{color}
\RequirePackage[bse=none]{pdfcomment}

\RequirePackage{MnSymbol}

\usepackage[absolute,overlay]{textpos}
\setlength{\TPHorizModule}{1cm}
\setlength{\TPVertModule}{\TPHorizModule}
\textblockorigin{0mm}{0mm}

\RequirePackage{wasysym}
\RequirePackage{multimedia}
\RequirePackage{comment}
\RequirePackage{transparent}
\RequirePackage{tikz}
\RequirePackage{pgfplots}
\pgfplotsset{compat=1.15}
\usepgfplotslibrary{patchplots}
\usetikzlibrary{backgrounds}
%\RequirePackage{xcolor}
%\RequirePackage{hyperref} %always require hyperref as the last package


%
% color setup - basic colors
%
% basic colors

\definecolor{white}{gray}{1.00}
\definecolor{gray}{gray}{0.65}
\definecolor{black}{gray}{0.00}

% cd print colors
\definecolor{hks41}{cmyk/RGB}{1,0.7,0.1,0.5/0,48,94}%TU blue primary cd color
\definecolor{hks92}{cmyk/RGB}{0.1,0,0.05,0.65/114,120,121}%TU grey secondary cd color - grey

\definecolor{hks44}{cmyk/RGB}{1,0.5,0,0/0,106,179}%highlight color 1st category cd   - brightblue
\definecolor{cyan}{cmyk/RGB}{1,0,0,0/0,158,224}%highlight color 1st category cd      - cyan

\definecolor{hks65}{cmyk/RGB}{0.65,0,1,0/106,176,35}%highlight color 2nd category cd - bright green
\definecolor{hks57}{cmyk/RGB}{1,0,0.9,0.2/0,125,64}%highlight color 2nd category cd - dark green

\definecolor{hks33}{cmyk/RGB}{0.5,1,0,0/147,16,126}%highlight color 2nd category cd - purple
\definecolor{hks36}{cmyk/RGB}{0.8,0.9,0,0/84,55,138}%highlight color 2nd category cd - violet

\definecolor{hks07}{cmyk/RGB}{0,0.6,1,0/238,127,0}%highlight color 2nd category cd - orange

%web colors
\definecolor{webredbright}{RGB}{221,39,39}
\definecolor{webreddark}{RGB}{181,28,28}

%sins colors
\definecolor{mathnaturebright}{cmyk/RGB}{0.65,0,1,0/106,176,35}
\definecolor{mathnaturedark}{cmyk/RGB}{0.9,0.17,1,0.05/0,133,49}

\definecolor{architectureenvironmentbright}{cmyk/RGB}{0.8,0,0.4,0/0,171,164}
\definecolor{architectureenvironmentdark}{cmyk/RGB}{1,0.1,0.4,0.15/0,129,139}

\definecolor{engineeringbright}{cmyk/RGB}{0.85,0.1,0.5,0/0,161,213}
\definecolor{engineeringdark}{cmyk/RGB}{1,0.2,0.05,0.2/0,119,173}

\definecolor{humansocialbright}{cmyk/RGB}{0.05,0.9,0,0/221,46,135}
\definecolor{humansocialdark}{cmyk/RGB}{0.05,1,0,0.1/204,0,114}

\definecolor{medicinebright}{cmyk/RGB}{0,0.85,0.85,0/230,68,46}
\definecolor{medicinedark}{cmyk/RGB}{0.15,1,1,0/204,7,30}

\definecolor{teachingbright}{cmyk/RGB}{0,0.6,1,0/238,127,0}
\definecolor{teachingdark}{cmyk/RGB}{0.1,0.8,1,0.05/209,77,23}


%default design colors
\colorlet{bgwhite}{white}
\colorlet{bgtud1}{hks41}
\colorlet{bgtud2}{hks44}
    \newcommand{\bgtudposone}{0.14}
    \newcommand{\bgtudpostwo}{1}
    \newcommand{\bgtudangle}{250}
\colorlet{bgdesign}{white}

\colorlet{designbar}{white}
\colorlet{designbarborder}{gray}
\newcommand{\designbaropacity}{0.5}

\colorlet{text}{hks41}
\colorlet{footertext}{hks92}

\colorlet{titletextbgwhite}{hks41}
\colorlet{titletextbgtud}{white}
\colorlet{titletextbgdesign}{white}

\colorlet{titleaffiliationtextbgwhite}{hks41}
\colorlet{titleaffiliationtextbgtud}{white}
\colorlet{titleaffiliationtextbgdesign}{hks92}

\colorlet{titledatecitytextbgwhite}{hks41}
\colorlet{titledatecitytextbgtud}{white}
\colorlet{titledatecitytextbgdesign}{hks92} %TODO?

\colorlet{frametitletext}{hks41}
\colorlet{framesubtitletext}{hks92!125}

\colorlet{blockbg}{hks41!70}

\newcommand{\setbeamercolors}{%
    \setbeamercolor{normal text}{fg=text}%
    \setbeamercolor{structure}{fg=text}%
    \setbeamercolor{alerted text}{fg=hks07}%
    \setbeamercolor{frametitle}{fg=frametitletext}%
    \setbeamercolor{block title}{bg=blockbg,fg=white}%
    \setbeamercolor{block body}{bg=blockbg!80}%
}

\newcommand{\setdefaultcolors}{%%
    %TODO eval class options%
%
    \colorlet{text}{hks41}%%
    \colorlet{frametitletext}{hks41}%%
    \colorlet{framesubtitletext}{hks92}%%
%
    \colorlet{footertext}{hks92}%
    \ifthenelse{\equal{\tud@titlestyle}{tud}}{%
        \colorlet{titletext}{titletextbgtud}%
        \colorlet{titleaffilitationtext}{titleaffiliationtextbgtud}%
        \colorlet{titledatecitytext}{titledatecitytextbgtud}%
    }{%
       \ifthenelse{\equal{\tud@titlestyle}{design}}{%
            \colorlet{titletext}{titletextbgdesign}%
            \colorlet{titleaffilitationtext}{titleaffiliationtextbgdesign}%
            \colorlet{titledatecitytext}{titledatecitytextbgdesign}%
        }{%
            \colorlet{titletext}{titletextbgwhite}%
            \colorlet{titleaffilitationtext}{titleaffiliationtextbgwhite}%
            \colorlet{titledatecitytext}{titledatecitytextbgwhite}%
        }% 
    }%
%
    \colorlet{blockbg}{hks41!70}%
%
    \mode<handout>{%
        \colorlet{text}{black}%
        \colorlet{frametitle}{black}%
        \colorlet{framesubtitletext}{gray}%
        \colorlet{bg}{bgwhite}%
        \colorlet{titlebg}{bgwhite}%
        \colorlet{footerbg}{bgwhite}%
        \colorlet{footertext}{gray}%
        \colorlet{titlebg}{white}%
        \colorlet{blockbg}{gray}%
    }%
    \setbeamercolors%
}

%TODO more to come
\hypersetup{
    linkcolor=text,%
    urlcolor=text,%
    citecolor=hks57,%
    pdfborder=0 0 0%
}
\mode<handout>{
    \hypersetup{
        linkcolor=text,%
        urlcolor=text,%
        citecolor=text,%  
        pdfborder=0 0 0%
    } 
}

%
% bibliography settings
%
\setlength{\bibitemsep}{6pt}
\defbibheading{bibliography}{}
% for Texnic Center
\renewcommand*{\bibliography}[1]{\addbibresource{#1.bib}}
% adaption brackets of 'authoryear-comp' style
\renewcommand*{\citep}[2][]{\mkbibbrackets{\cite[#1]{#2}}}

%
% logo settings TODO where does the logo come from?
%
% \def\logo@front{logo_weiss}
% \def\logo@frontBrightBg{logo_blau}
% \def\logo@default{logo_blau}
% \ifx\pdfoutput\undefined\else\ifx\pdfoutput\relax\else\ifcase\pdfoutput\else\def\logo@front{TU_Logo_SW}\def\logo@default{TU_Logo_SW}\fi\fi\fi
% \def\IfAlogo@front{IfA_logo_weiss}
% \def\IfAlogo@frontBrightBg{IfA_logo_blau}
% \def\IfAlogo@print{IfA_logo_schwarz}
% \let\IfAlogo\IfAlogo@front


%
% font settings
%
\renewcommand*\encodingdefault{OT1}
\renewcommand*\familydefault{fos} %family open sans
\renewcommand*\bfseries{\fontseries{b}\selectfont}
\renewcommand*\shapedefault{n}
\renewcommand*\sfdefault{fos}
\renewcommand*\ttfamily{\not@math@alphabet\ttfamily\mathtt\fontfamily\ttdefault\fontseries\mddefault\selectfont}
%\newcommand*{\dinfamily}{\fontencoding{OT1}\fontfamily{\familydefault}\fontseries{b}\selectfont} %TODO remove

\setbeamerfont{itemize/enumerate subbody}{size=\scriptsize}
\setbeamerfont{itemize/enumerate subsubbody}{size=\scriptsize}
\setbeamerfont{title}{size=\LARGE,series=\bfseries,family=\fontfamily{\familydefault}}
\setbeamerfont{subtitle}{series=\bfseries,family=\sffamily}

\if@sansMath
    \DeclareSymbolFont{fosLetters}{T1}{fos}{l}{sl}
    \SetSymbolFont{fosLetters}{bold}{T1}{fos}{b}{sl}
    \DeclareSymbolFont{fosOperators}{T1}{fos}{l}{n}
    \SetSymbolFont{fosOperators}{bold}{T1}{fos}{b}{n}
    
    \DeclareMathSymbol{a}{\mathalpha}{fosLetters}{`a}
    \DeclareMathSymbol{b}{\mathalpha}{fosLetters}{`b}
    \DeclareMathSymbol{c}{\mathalpha}{fosLetters}{`c}
    \DeclareMathSymbol{d}{\mathalpha}{fosLetters}{`d}
    \DeclareMathSymbol{e}{\mathalpha}{fosLetters}{`e}
    \DeclareMathSymbol{f}{\mathalpha}{fosLetters}{`f}
    \DeclareMathSymbol{g}{\mathalpha}{fosLetters}{`g}
    \DeclareMathSymbol{h}{\mathalpha}{fosLetters}{`h}
    \DeclareMathSymbol{i}{\mathalpha}{fosLetters}{`i}
    \DeclareMathSymbol{j}{\mathalpha}{fosLetters}{`j}
    \DeclareMathSymbol{k}{\mathalpha}{fosLetters}{`k}
    \DeclareMathSymbol{l}{\mathalpha}{fosLetters}{`l}
    \DeclareMathSymbol{m}{\mathalpha}{fosLetters}{`m}
    \DeclareMathSymbol{n}{\mathalpha}{fosLetters}{`n}
    \DeclareMathSymbol{o}{\mathalpha}{fosLetters}{`o}
    \DeclareMathSymbol{p}{\mathalpha}{fosLetters}{`p}
    \DeclareMathSymbol{q}{\mathalpha}{fosLetters}{`q}
    \DeclareMathSymbol{r}{\mathalpha}{fosLetters}{`r}
    \DeclareMathSymbol{s}{\mathalpha}{fosLetters}{`s}
    \DeclareMathSymbol{t}{\mathalpha}{fosLetters}{`t}
    \DeclareMathSymbol{u}{\mathalpha}{fosLetters}{`u}
    \DeclareMathSymbol{v}{\mathalpha}{fosLetters}{`v}
    \DeclareMathSymbol{w}{\mathalpha}{fosLetters}{`w}
    \DeclareMathSymbol{x}{\mathalpha}{fosLetters}{`x}
    \DeclareMathSymbol{y}{\mathalpha}{fosLetters}{`y}
    \DeclareMathSymbol{z}{\mathalpha}{fosLetters}{`z}
    \DeclareMathSymbol{A}{\mathalpha}{fosLetters}{`A}
    \DeclareMathSymbol{B}{\mathalpha}{fosLetters}{`B}
    \DeclareMathSymbol{C}{\mathalpha}{fosLetters}{`C}
    \DeclareMathSymbol{D}{\mathalpha}{fosLetters}{`D}
    \DeclareMathSymbol{E}{\mathalpha}{fosLetters}{`E}
    \DeclareMathSymbol{F}{\mathalpha}{fosLetters}{`F}
    \DeclareMathSymbol{G}{\mathalpha}{fosLetters}{`G}
    \DeclareMathSymbol{H}{\mathalpha}{fosLetters}{`H}
    \DeclareMathSymbol{I}{\mathalpha}{fosLetters}{`I}
    \DeclareMathSymbol{J}{\mathalpha}{fosLetters}{`J}
    \DeclareMathSymbol{K}{\mathalpha}{fosLetters}{`K}
    \DeclareMathSymbol{L}{\mathalpha}{fosLetters}{`L}
    \DeclareMathSymbol{M}{\mathalpha}{fosLetters}{`M}
    \DeclareMathSymbol{N}{\mathalpha}{fosLetters}{`N}
    \DeclareMathSymbol{O}{\mathalpha}{fosLetters}{`O}
    \DeclareMathSymbol{P}{\mathalpha}{fosLetters}{`P}
    \DeclareMathSymbol{Q}{\mathalpha}{fosLetters}{`Q}
    \DeclareMathSymbol{R}{\mathalpha}{fosLetters}{`R}
    \DeclareMathSymbol{S}{\mathalpha}{fosLetters}{`S}
    \DeclareMathSymbol{T}{\mathalpha}{fosLetters}{`T}
    \DeclareMathSymbol{U}{\mathalpha}{fosLetters}{`U}
    \DeclareMathSymbol{V}{\mathalpha}{fosLetters}{`V}
    \DeclareMathSymbol{W}{\mathalpha}{fosLetters}{`W}
    \DeclareMathSymbol{X}{\mathalpha}{fosLetters}{`X}
    \DeclareMathSymbol{Y}{\mathalpha}{fosLetters}{`Y}
    \DeclareMathSymbol{Z}{\mathalpha}{fosLetters}{`Z}

    \DeclareMathSymbol{,}{\mathpunct}{fosLetters}{`,}
    \DeclareMathSymbol{.}{\mathord}{fosLetters}{`.}
    \DeclareMathSymbol{/}{\mathord}{fosLetters}{`/}

    \DeclareMathSymbol{0}\mathalpha{fosOperators}{"30}
    \DeclareMathSymbol{1}\mathalpha{fosOperators}{"31}
    \DeclareMathSymbol{2}\mathalpha{fosOperators}{"32}
    \DeclareMathSymbol{3}\mathalpha{fosOperators}{"33}
    \DeclareMathSymbol{4}\mathalpha{fosOperators}{"34}
    \DeclareMathSymbol{5}\mathalpha{fosOperators}{"35}
    \DeclareMathSymbol{6}\mathalpha{fosOperators}{"36}
    \DeclareMathSymbol{7}\mathalpha{fosOperators}{"37}
    \DeclareMathSymbol{8}\mathalpha{fosOperators}{"38}
    \DeclareMathSymbol{9}\mathalpha{fosOperators}{"39}

    \DeclareSymbolFontAlphabet{\mathsf}{fosOperators}
    \DeclareSymbolFontAlphabet{\mathnormal}{fosLetters}
    \DeclareMathAlphabet\mathbf{T1}{fos}{b}{n}
    \SetMathAlphabet\mathbf{bold}{T1}{fos}{b}{n}
    \DeclareMathAlphabet\mathit{T1}{fos}{l}{sl}
    \SetMathAlphabet\mathit{bold}{T1}{fos}{l}{sl}
    
    \let\operator@font\sf
\fi
\usefonttheme[onlymath]{serif}

%
% language settings
%
% predefined symbols
\newcommand*{\englishnames}{%
    \newcommand*{\@slide}{slide}
    \newcommand*{\@of}{of}
}
\newcommand*{\germannames}{%
    \newcommand*{\@slide}{Folie}
    \newcommand*{\@of}{von}
}
\DeclareRobustCommand\one{1\kern0.1pt} % 1 mit anschliessendem Freiraum, damit sie nicht deutlich kuerzer ist als die anderen Ziffern

%
% length settings
%
\newlength{\tudlogox}
\newlength{\tudlogoy}
\newlength{\tudlogoheight}

\newlength{\tudlogofooterx}
\newlength{\tudlogofootery}
\newlength{\tudlogofooterheight}

\newlength{\designbary}
\newlength{\designbarheight}
\newlength{\designbarborderwidth}

\newlength{\footery}
\newlength{\footerheight}

\newlength{\bodyy} %dynamic
\newlength{\bodyx} %dynamic
\newlength{\bodydefaultx}
\newlength{\bodyheight} %dynamic
\newlength{\bodydefaulty}
\newlength{\bodydefaultheight}
\newlength{\bodywidth}%dynamic
\newlength{\bodydefaultwidth}

\newlength{\textborderleft}
\newlength{\textborderright}

\newlength{\footertextx}
\newlength{\footertexty}
\newlength{\footertextwidth}

\newlength{\footerslidenumx}
\newlength{\footerslidenumy}
\newlength{\footerslidenumwidth}

\newlength{\addlogoxright}
\newlength{\addlogoy}
\newlength{\addlogoheight}
\newlength{\addlogofooterxright}
\newlength{\addlogofootery}
\newlength{\addlogofooterheight}

\newlength{\titley}
\newlength{\subtitley}
\newlength{\authoraffiliationy}

\newlength{\frametitley}
\newlength{\framesubtitley}

\newlength{\parttitley}
\newlength{\templength}

% dimensions are derived from the powerpoint template
%ppt width 33.867cm
%ppt height 19.05cm
\setlength{\tudlogox}{0.0239\beamer@paperwidth}
\setlength{\tudlogoy}{0.0509\beamer@paperheight}
\setlength{\tudlogoheight}{0.0751\beamer@paperheight}

\setlength{\tudlogofooterx}{0.0416\beamer@paperwidth}
\setlength{\tudlogofootery}{0.9239\beamer@paperheight}
\setlength{\tudlogofooterheight}{0.0472\beamer@paperheight}

\setlength{\designbary}{0.1496\beamer@paperheight}
\setlength{\designbarheight}{0.0252\beamer@paperheight}
\setlength{\designbarborderwidth}{1pt}

\setlength{\footery}{0.8929\beamer@paperheight}
\setlength{\footerheight}{0.1071\beamer@paperheight}

\setlength{\footertextx}{0.2392\beamer@paperwidth}
\setlength{\footertexty}{0.9218\beamer@paperheight}
\setlength{\footertextwidth}{0.4251\beamer@paperwidth}

\setlength{\footerslidenumx}{0.7355\beamer@paperwidth}
\setlength{\footerslidenumwidth}{0.0579\beamer@paperwidth}
\setlength{\footerslidenumy}{0.0541\beamer@paperheight}

\setlength{\textborderleft}{0.0718\beamer@paperwidth}
\setlength{\textborderright}{0.0718\beamer@paperwidth}

\setlength{\addlogoxright}{0.9771\beamer@paperwidth}
\setlength{\addlogoy}{0.0509\beamer@paperheight}
\setlength{\addlogoheight}{0.0751\beamer@paperheight}
\setlength{\addlogofooterxright}{0.9632\beamer@paperwidth}
\setlength{\addlogofootery}{0.9239\beamer@paperheight}
\setlength{\addlogofooterheight}{0.0509\beamer@paperheight}

\setlength{\titley}{0.4945\beamer@paperheight}
\setlength{\subtitley}{0.6556\beamer@paperheight}
\setlength{\authoraffiliationy}{0.3533\beamer@paperheight}

\setlength{\frametitley}{0.0504\beamer@paperheight}
\setlength{\framesubtitley}{0.12756\beamer@paperheight}

\setlength{\bodydefaultx}{0.0718\beamer@paperwidth}
\setlength{\bodydefaulty}{0.2157\beamer@paperheight}
\setlength{\bodydefaultwidth}{0.8679\beamer@paperwidth}
\setlength{\bodydefaultheight}{0.6357\beamer@paperheight}

\setlength{\parttitley}{0.494\beamer@paperheight}

\newcommand{\setbeamerdimensions}{
    %\beamersetleftmargin{\bodyx}
    %\beamersetrightmargin{\bodyx}%\beamer@paperwidth-\bodyx-\bodywidth}
}

\newcommand{\setbodydefaultdimensions}{%
    \setlength{\bodyy}{\bodydefaulty}%
    \setlength{\bodyx}{\bodydefaultx}%
    \setlength{\bodyheight}{\bodydefaultheight}%
    \setlength{\bodywidth}{\bodydefaultwidth}%
    %\setbeamerdimensions
}

\newcommand{\setbodynotitledimensions}{%    
    \setlength{\bodyy}{\frametitley}%
    \setlength{\bodyheight}{\bodydefaulty-\frametitley+\bodydefaultheight}%
    %\setbeamerdimensions
}
\newcommand{\setbodynomargindimensions}{%
    \setlength{\bodyx}{0}%
    \setlength{\bodywidth}{\bodydefaultwidth+2\bodydefaultx}%
    %\setbeamerdimensions
}
\newcommand{\setbodytitleframedimensions}{
    \setbodydefaultdimensions
    \setlength{\bodyheight}{\beamer@paperheight-\bodyy}
   %\setbeamerdimensions
}



%
% template settings
%
%% patching the vertical skip between sections in table of contents
\patchcmd{\beamer@sectionintoc}{\vskip1.5em}{\vskip0.5em}{}{}

%\newcommand{}{}


% Templates
\if@useNavBar\else\setbeamertemplate{navigation symbols}{}\fi    % Nav-Leiste aus

\newcounter{backmatterframes}
\setcounter{backmatterframes}{0}


%
% set template commands
%

\newcommand{\tud@logofile}{TU_Logo_SW}

\newcommand{\tud@designframestyle}{tud} %possible values white/tud/design
\newcommand{\setDesignFrameStyle}[1]{\renewcommand{\tud@designframestyle}{#1}}

%use tud@designimage variable for the design frame style


\newcommand*{\insertdatecity}{Dresden}
\newcommand*{\inserttotalpagenumber}{\textbf{??}}

\newcommand{\setAuxilliaryBeamerTemplates}{
    %TODO 

}


\newcommand{\setTUDHeadlineTemplates}{
    \setbeamertemplate{headline}{
        \setlength{\templength}{\beamer@paperwidth/2-\tudlogox}
        \begin{textblock*}{\templength}(\tudlogox,\tudlogoy)
          \color{hks41}\includegraphics[height=\tudlogoheight]{\tud@logofile}
        \end{textblock*}
        \ifx\additionallogo\@empty\else
        \begin{textblock*}{\templength}(\addlogoxright-\templength,\addlogoy)%
          \color{hks41}%
          \raggedleft\includegraphics[height=\addlogoheight]{\tud@additionallogo}
        \end{textblock*}
        \fi 
    }
}
\newcommand{\setTUDFooterTemplates}{
    \setbeamertemplate{footline}{}
}
\newcommand{\setFrameHeadlineTemplates}{
    \setbeamertemplate{headline}{}
    
}
\newcommand{\setFrameFooterTemplates}{
    
}
%drawTikzBackground[mode]{width,height}
\newcommand{\drawTikzColorBackground}[3][tud]{
    \ifthenelse{\equal{#1}{tud}}{%
        \begin{axis}[%
            width=#2,%
            height=#3,%
            scale only axis,%
            domain=0:\number#2,%
            y domain=0:\number#3,%
            view={0}{90},%
            hide axis,%
            colormap= {tud}{color(0cm)=(bgtud1); color(\bgtudposone cm)=(bgtud1); color(\bgtudpostwo cm)=(bgtud2); color(1.0001cm)=(bgtud2)}]%
                %zmin=0,zmax=1,title=,xlabel=,ylabel=,zlabel=,,ymin=0,ymax=10,xmin=0,xmax=10,
            \addplot3[surf,shader=flat,patch type=bilinear,samples=50] {(-.4*x) + (.4) + (y)};%shader=interp% does not work on sumatrapdf
                %,mesh/color input=explicit,draw=\transparent{1},shader=interp,patch,shader=faceted interp
        \end{axis}%
    }{%
        \fill[#1] (0,0) rectangle (#2,#3);%
    }%
}

\newcommand{\setTitleBodyTemplates}{
    \setbodytitleframedimensions%
    %1. background
    \setbeamertemplate{background canvas}{%
        \ifthenelse{\equal{\tud@titlestyle}{design}}{%            
            \begin{textblock*}{\beamer@paperwidth}(o,\designbary)%
                \includegraphics[height=\beamer@paperheight-\designbary,width=\beamer@paperwidth,keepaspectratio=true]{\tud@designimage}
            \end{textblock*}%
        }{%
            \begin{tikzpicture}[overlay,inner frame sep=0,shift={(0cm,-\the\beamer@paperheight+1cm)}]%current page.north west)]
                %\pgfplotset{width=\the\beamer@paperwidth,height=\the\beamer@paperheight};                
                \drawTikzColorBackground[\tud@titlestyle]{\beamer@paperwidth}{\beamer@paperheight-\the\designbary}%
                \pgfresetboundingbox%
            \end{tikzpicture}%
        }%
        \ifthenelse{\equal{\tud@titlestyle}{white}}{%
            \begin{tikzpicture}[overlay,inner frame sep=0,shift={(0cm,1cm)}]%current page.north west)]
                %\begin{scope}%
                    \draw[line width=\the\designbarborderwidth,designbarborder] (0,-\the\designbary)--(\the\beamer@paperwidth,-\the\designbary);%
                    \draw[line width=\the\designbarborderwidth,designbarborder] (0,-\the\designbary-\the\designbarheight)--(\the\beamer@paperwidth,-\the\designbary-\the\designbarheight);%
                    \fill[designbar,opacity=\designbaropacity] (0,-\the\designbary) rectangle (\the\beamer@paperwidth, -\the\designbary-\the\designbarheight);%
                %\end{scope}%
            \end{tikzpicture}%
        }{%            
            \transparent{\designbaropacity}%
            \begin{tikzpicture}[overlay,inner frame sep=0,shift={(0cm,1cm)}]%current page.north west)]
                %\begin{scope}%
                   \fill[designbar,opacity=\designbaropacity] (0,-\the\designbary) rectangle (\the\beamer@paperwidth, -\the\designbary-\the\designbarheight);%
                %\end{scope}%
            \end{tikzpicture}%
            \transparent{1}%  
        }%
    }%
    \setbeamertemplate{title page}{%
        %affiliation
        \begin{textblock*}{\beamer@paperwidth-\textborderleft-\textborderright}(\textborderleft,\authoraffiliationy)%
          \usebeamerfont{subtitle}\color{titleaffilitationtext}\insertauthor\\%
          \tud@affiliation%
        \end{textblock*}%
        %title
        \begin{textblock*}{\beamer@paperwidth-\textborderleft-\textborderright}(\textborderleft,\titley)%
          \usebeamerfont{title}\color{titletext}\inserttitle
        \end{textblock*}%
        %subtitle
        \begin{textblock*}{\beamer@paperwidth-\textborderleft-\textborderright}(\textborderleft,\subtitley)%
          \usebeamerfont{subtitle}\color{titletext}\insertsubtitle\\%
          \vspace{2ex}\usebeamerfont{subtitle}\color{titledatecitytext}\insertdatecity
        \end{textblock*}%
    }%
}
\newcommand{\setDefaultBodyTemplates}{
    \setbeamertemplate{background canvas}{}
    
}
\newcommand{\setSectionFrameBodyTemplates}{
    
}
\newcommand{\setPartFrameBodyTemplates}{
    
}
\newcommand{\setDesignFrameBodyTemplates}{
    
}
\newcommand{\setTitleFrameTemplates}{
    \setTUDHeadlineTemplates
    \setTUDFooterTemplates
    \setTitleBodyTemplates
}
\newcommand{\setDefaultFrameTemplates}{
    \setFrameHeadlineTemplates
    \setFrameFooterTemplates
    \setDefaultBodyTemplates
}
\newcommand{\setDesignFrameTemplates}{    
    \setFrameHeadlineTemplates
    \setFrameFooterTemplates
    \setDefaultBodyTemplates
}
\newcommand{\setPartFrameTemplates}{
    \setFrameHeadlineTemplates
    \setFrameFooterTemplates
    \setPartFrameBodyTemplates    
}
\newcommand{\setSectionFrameTemplates}{
    \setFrameHeadlineTemplates
    \setFrameFooterTemplates
    \setSectionFrameBodyTemplates      
}
\newcommand{\setFinalFrameTemplates}{
    \setTUDHeadlineTemplates
    \setTUDFooterTemplates
    \setDesignFrameBodyTemplates
}

%
% environments for special frames
%
% \newenvironment{backmatterframe}[1][]{
%     \addtocounter{backmatterframes}{1}
%     \ifx#1\empty\begin{frame}\else\begin{frame}[#1]\fi
%     \setbackmattertemplates
% }{
%     \end{frame}\setbeamertemplates
% }

% \newenvironment{regardsframe}[1][]{
%     \ifx#1\empty\begin{frame}\else\begin{frame}[#1]\fi
%     \setregardtemplates
% }{
%     \end{frame}\setbeamertemplates
% }

% \newenvironment{partframe}[1][]{    
%     \ifx#1\empty\begin{frame}\else\begin{frame}[#1]\fi
%     \setregardtemplates
% }{
%     \end{frame}\setbeamertemplates
% }

% \newenvironment{sectionframe}[1][]{
%     \ifx#1\empty\begin{frame}\else\begin{frame}[#1]\fi
%     \setregardtemplates
% }{
%     \end{frame}\setbeamertemplates
% }

% \newenvironment{titleframe}[1][]{
%     \ifx#1\empty\begin{frame}\else\begin{frame}[#1]\fi
%     \setregardtemplates
% }{
%     \end{frame}\setbeamertemplates
% }

%
% defining hooks
%

%\beamertemplatedotitem
%\usesubitemizeitemtemplate{--}

\renewcommand\maketitle{%
    \setTitleFrameTemplates
    \frame{\titlepage}%
    % Kopf-/Fusszeilen fuer restliche Folien
    \setDefaultFrameTemplates% %TODO
}

% % Faktor zur PPT-Vorlage: / 0,1984375      * 5,04
% \setbeamertemplate{title page}{%
%     \color{white}%
%     {\vfill}{\renewcommand\baselinestretch{0.8}\usebeamerfont*{title}
%     \vskip2ex plus1ex minus1ex\MakeUppercase{\inserttitle}}\vfill%
%     {\ifx\insertsubtitle\empty\else\usebeamerfont*{subtitle}\insertsubtitle\vfill\fi%
%     \scriptsize\insertauthor}\vfill\vfill%
% }

% \setbeamertemplate{part page}{%
% 	\color{white}{\vfill}{
% 	\renewcommand\baselinestretch{0.8}\usebeamerfont*{title}
% 	\vskip-4ex plus1ex minus1ex\insertpart}
% }

% \newcommand\settitletemplates{
% 	%\beamertemplateshadingbg{darkblue}{darkblue}
%     \setbeamercolor{normal text}{bg=darkblue}
%     % Kopf-/Fusszeile fuer Titel
%     \setbeamertemplate{headline}{%
%         \vskip6.15mm\color{white}%
%         \setlength{\arrayrulewidth}{0.3pt}%
%         \begin{tabular*}{\paperwidth}[b]{p{\textwidth}}%@{\extracolsep\fill}
%           % TUD-Logo
%             \hspace*{3.0mm}\color{white}\includegraphics[height=7.81mm]{\logo@front}%
%             \if@insertIfALogo\hfill\includegraphics[height=7.5mm,trim=
%             0 8mm 0 0]{\IfAlogo}\hspace*{-2.0mm}\fi%
%             \hspace*{11.76mm}\\[1.2mm]\hline
%             \hspace*{11.76mm}\rule[-0.8mm]{0pt}{2.47mm}%
%             \def\@@dummyComma{}%
%             \rule{0pt}{5.8pt}\textbf{\@einrichtung}\ifx\@einrichtung\@empty\else\def\@@dummyComma{ }\fi%
%             \ifx\@fachrichtung\@empty\else\@@dummyComma\@fachrichtung%
%                 \ifx\@institut\@empty\else\def\@@dummyComma{, }\fi%
%                 \ifx\@professur\@empty\else\def\@@dummyComma{, }\fi%
%             \fi%
%             \ifx\@institut\@empty\else\@@dummyComma\@institut
%                 \ifx\@professur\@empty\else\def\@@dummyComma{, }\fi
%             \fi%
%             \ifx\@professur\@empty\else\@@dummyComma\fi%
%             \@professur\\[-1pt]\hline
%         \end{tabular*}\hspace{-\paperwidth}%
%     }%
%     \setbeamertemplate{footline}{%
%         \parbox[t][9.63mm]{\paperwidth}{%
%             \vspace*{-4.63mm}%
%             \rule{13.86mm}{0pt}\normalsize\color{white}\insertdatecity\ifx\insertdatecity\empty\else\ifx\insertdate\empty\else, \fi\fi\insertdate%
%             \if@ddcfooter\vskip-12.5mm\hspace*{\fill}\includegraphics[height=11.76mm]{DDC-weissf}\hspace*{9.7mm}\fi
%         }
%     }%
% }

% \newcommand\setparttemplates{
% 	%\beamertemplateshadingbackground{darkblue}{darkblue}
%     \setbeamercolor{normal text}{bg=darkblue}
%     % Kopf-/Fusszeile fuer Titel
%     \setbeamertemplate{headline}{%
%         \vskip6.15mm\color{white}%
%         \setlength{\arrayrulewidth}{0.3pt}%
%         \begin{tabular*}{\paperwidth}[b]{p{\textwidth}}%@{\extracolsep\fill}
%           % TUD-Logo
%             \hspace*{3.0mm}\color{white}\includegraphics[height=7.81mm]{\logo@front}%
%             \if@insertIfALogo\hfill\includegraphics[height=7.5mm,trim=
%             0 8mm 0 0]{\IfAlogo}\hspace*{-2.0mm}\fi%
%             \hspace*{11.76mm}\\[1.2mm]\hline
%             \hspace*{11.76mm}\rule[-0.8mm]{0pt}{2.47mm}%
%             \def\@@dummyComma{}%
%             \rule{0pt}{5.8pt}\textbf{\@einrichtung}\ifx\@einrichtung\@empty\else\def\@@dummyComma{ }\fi%
%             \ifx\@fachrichtung\@empty\else\@@dummyComma\@fachrichtung%
%                 \ifx\@institut\@empty\else\def\@@dummyComma{, }\fi%
%                 \ifx\@professur\@empty\else\def\@@dummyComma{, }\fi%
%             \fi%
%             \ifx\@institut\@empty\else\@@dummyComma\@institut
%                 \ifx\@professur\@empty\else\def\@@dummyComma{, }\fi
%             \fi%
%             \ifx\@professur\@empty\else\@@dummyComma\fi%
%             \@professur\\[-1pt]\hline
%         \end{tabular*}\hspace{-\paperwidth}%
%     }%
%     \setbeamertemplate{footline}{%   
%     }%
% }

\newcommand\setbeamertemplates{%	

        %\beamertemplateshadingbackground{white}{white}
    \setbeamercolor{normal text}{bg=white}
    % \if@useHeader%
    %     \setbeamertemplate{headline}{%
    %         \vskip6.15mm\color{gray}
    %         \setlength{\arrayrulewidth}{0.1pt}
    %         \begin{tabular*}{\paperwidth}[b]{l@{\extracolsep\fill}}
    %             \hspace*{5.0mm}\color{text}\includegraphics[height=5.95mm]{\logo@default}
    %             \\[-1mm]\hspace*{-\tabcolsep}\rule{\paperwidth}{.3pt} \rule[-2.47mm]{0pt}{2.47mm}\\[-2.47mm]\hspace*{-\tabcolsep}\rule{\paperwidth}{.3pt} \rule[-2.47mm]{0pt}{2.47mm}
    %         \end{tabular*}
    %     }
    % \else
        \setbeamertemplate{headline}{%
            %\rule{0pt}{14.57mm}
            \rule{0pt}{1cm}%\the\topmarginnoheader}
        }
    % \fi

    \setbeamertemplate{footline}{%
        \color{gray}%
        \parbox[t][7.46mm]{\paperwidth}{%
            \vspace*{1.63mm}%
            \rule{13.86mm}{0pt}%
            \parbox{28.83mm}{\ifx\insertdate\empty\else\insertdate\fi}\rule{1.06mm}{0pt}%TU Dresden\ifx\insertdate\empty\else,\insertdate\fi
           \centerline{\insertshorttitle}\rule{7.46mm}{0pt}%
            % \if@usePageNum%
            %     \let\insertframenumber=\insertpagenumber%
            %     \let\inserttotalframenumber=\inserttotalpagenumber%
            % \fi%
            % \raggedleft\@slide\ \insertframenumber\ \if@noTotalPages\relax\else\@of\ \inserttotalframenumber\fi\rule{13.86mm}{0pt}%
        }
    }

    \setbeamertemplate{frametitle}{%
        \vskip-2.0mm
       % \color{frametitle}\if@useNoFrameTitleSection\else\frame@title@section\fi\insertframetitle\\%
        \vskip0.5ex\color{text}\small\bfseries\insertframesubtitle%
    }
}

% \newcommand\setregardtemplates{
% 	\setbeamercolor{normal text}{bg=darkblue}
% 	\color{white}
%       \setbeamertemplate{headline}{%
%         \vskip6.15mm\color{white}%
%         \setlength{\arrayrulewidth}{0.3pt}%
%         \begin{tabular*}{\paperwidth}[b]{p{\textwidth}}%@{\extracolsep\fill}
%           % TUD-Logo
%             \hspace*{3.0mm}\color{white}\includegraphics[height=7.81mm]{\logo@front}%
%             \if@insertIfALogo\hfill\includegraphics[height=7.5mm,trim=
%             0 8mm 0 0]{\IfAlogo}\hspace*{10.0mm}\fi%
%             \hspace*{11.76mm}\\[-1mm]\hspace*{-\tabcolsep}\rule{\paperwidth}{.3pt}\\[-0.85mm] \rule[-2.47mm]{0pt}{2.47mm}
%             \hspace*{11.76mm}\rule[-0.8mm]{0pt}{2.47mm}%
%             \def\@@dummyComma{}%
%             \rule{0pt}{5.8pt}\textbf{\@einrichtung}\ifx\@einrichtung\@empty\else\def\@@dummyComma{ }\fi%
%             \ifx\@fachrichtung\@empty\else\@@dummyComma\@fachrichtung%
%                 \ifx\@institut\@empty\else\def\@@dummyComma{, }\fi%
%                 \ifx\@professur\@empty\else\def\@@dummyComma{, }\fi%
%             \fi%
%             \ifx\@institut\@empty\else\@@dummyComma\@institut
%                 \ifx\@professur\@empty\else\def\@@dummyComma{, }\fi
%             \fi%
%             \ifx\@professur\@empty\else\@@dummyComma\fi%
%             \@professur\\[-3.5mm]\hspace*{-\tabcolsep}\rule{\paperwidth}{.3pt} \rule[-2.47mm]{0pt}{2.47mm}
%         \end{tabular*}\hspace{-\paperwidth}%
%     }%
%     \setbeamertemplate{footline}{
%         \parbox[t][9.63mm]{\paperwidth}{
%             \vspace*{10.63mm}
%             \rule{13.86mm}{0pt}
%             \if@ddcfooter\vskip-12.5mm\hspace*{\fill}\includegraphics[height=11.76mm]{DDC-weissf}\hspace*{9.7mm}\fi
%         }
%     }
% }

% \newcommand\setbackmattertemplates{
% 	\setbeamercolor{normal text}{bg=white}
% 	\color{text}
%     \if@useHeader%
%         \setbeamertemplate{headline}{%
%             \vskip6.15mm\color{gray}
%             \setlength{\arrayrulewidth}{0.1pt}
%             \begin{tabular*}{\paperwidth}[b]{l@{\extracolsep\fill}}
%                 \hspace*{5.0mm}\color{text}\includegraphics[height=5.95mm]{\logo@default}
%                 \if@useHeaderToc\else
% 	                \if@insertIfALogo\hspace{\headerlogoskip}\raisebox{-.5mm}{\includegraphics[height=6.39mm]{\IfAlogo@frontBrightBg}}
% 	            	\fi
%                 \fi
%                 \\[-1mm]\hspace*{-\tabcolsep}\rule{\paperwidth}{.3pt} \rule[-2.47mm]{0pt}{2.47mm}\\[-2.47mm]\hspace*{-\tabcolsep}\rule{\paperwidth}{.3pt} \rule[-2.47mm]{0pt}{2.47mm}
%             \end{tabular*}
%         }
%     \else
%         \setbeamertemplate{headline}{%
%             %\rule{0pt}{14.57mm}
%             \rule{0pt}{\the\topmarginnoheader}
%         }
%     \fi

%     \setbeamertemplate{footline}{%
%         \color{gray}%
%         \parbox[t][7.46mm]{\paperwidth}{%
%             \vspace*{1.63mm}%
%             \rule{13.86mm}{0pt}%
%             \parbox{28.83mm}{\ifx\insertdate\empty\else\insertdate\fi}\rule{1.06mm}{0pt}%TU Dresden\ifx\insertdate\empty\else,\insertdate\fi
%             \parbox{\footercenterwidth}{\centerline{\insertshorttitle}}\rule{7.46mm}{0pt}%
%             \if@usePageNum%
%                 \let\insertframenumber=\insertpagenumber%
%                 \let\inserttotalframenumber=\inserttotalpagenumber%
%             \fi
%             \parbox{\paperwidth-\footercenterwidth-13.86mm-13.86mm-28.83mm-1.06mm-7.46mm}{\raggedleft\@slide\ \insertframenumber\ \if@noTotalPages\relax\else\@of\ \inserttotalframenumber\fi}\rule{13.86mm}{0pt}%
%         }
%     }

%     \setbeamertemplate{frametitle}{%
%         \vskip-2.0mm
%         \color{frametitle}\if@useNoFrameTitleSection\else\frame@title@section\fi\insertframetitle\\%
%         \vskip0.5ex\color{text}\small\bfseries\insertframesubtitle%
%     }
% }

% frame title
% TODO what is this doing
\gdef\frame@title@section{}%
\def\frametitle{\secdef\tud@frametitlea\tud@frametitles}%
\newcommand<>\tud@frametitlea{\gdef\frame@title@section{\ifnum\c@section>9\else0\fi\arabic{section}\,\,}\alt#1{\@dblarg\beamer@@frametitle}{\beamer@gobbleoptional}}%
\newcommand<>\tud@frametitles{\gdef\frame@title@section{}\alt#1{\@dblarg\beamer@@frametitle}{\beamer@gobbleoptional}}%

% Sprache einstellen
\AtBeginDocument{%
    \ifx\@undefined\if@german%        
        \newif\if@german                            % Nutzung des german-Packages ?
                    %                
        \ifx\@undefined\germanTeX\else\@germantrue\fi       % Wenn \(n)germanTeX definiert ist,
        \ifx\@undefined\ngermanTeX\else\@germantrue\fi      % ist das (n)german-Package eingebunden
        \ifx\@undefined\language\else                       %
            \ifnum0=\language\@germanfalse\else\fi          % \language=0  =>  english
        \fi%
                                                 %
        \if@useGerman\@germantrue\fi                        % Zur expliziten Nutzung der deutschen Texte (das ngerman-Package wurde automatisch eingebunden)
        \if@useNoGerman\@germanfalse\fi                     % Zur expliziten Nutzung der englischen Texte (auch bei vorheriger Einbindung des german-Packages)
        \if@german                                          %
%
            \ifx\@undefined\language\else                   %
                \ifnum2<\language\selectlanguage{german}\fi % deutsch gewuenscht, aber andere Sprache eingestellt ??  =>  deutsch
                \ifnum2>\language\selectlanguage{german}\fi %
            \fi                                             %
            \germannames                                    %
            \def\today{\number\day.\number\month.\number\year}  % Datum im Format DD.MM.YYYY
        \else                                               %
%
            \ifx\@undefined\language\else                   %
                \ifnum0<\language\selectlanguage{english}\fi% englisch gewuenscht, aber andere Sprache eingestellt ??  =>  englisch
            \fi                                             %
            \englishnames                                   %
            \def\today{\number\year/\number\month/\number\day}  % Datum im Format YYYY/MM/DD
        \fi%       
                                          %
    \fi%   
    \setdefaultcolors%
    \setbodydefaultdimensions%
    \setAuxilliaryBeamerTemplates%
    \setbeamertemplates% %TODO
    \scriptsize%
    \makeatother%
}%

\AtEndDocument{%
    %reset total frame counter in order to hide pages behind actual content.
    \addtocounter{framenumber}{-\value{backmatterframes}}%
    \clearpage%
    \beamer@tempcount=\c@page\advance\beamer@tempcount by -1%
    \if@filesw
        \immediate\write\@auxout{\string\@writefile{nav}%
            {\noexpand\headcommand{\noexpand\beamer@partpages{\the\beamer@partstartpage}{\the\beamer@tempcount}}}}%
        \immediate\write\@auxout{\string\@writefile{nav}%
            {\noexpand\headcommand{\noexpand\beamer@subsectionpages{\the\beamer@subsectionstartpage}{\the\beamer@tempcount}}}}%
        \immediate\write\@auxout{\string\@writefile{nav}%
            {\noexpand\headcommand{\noexpand\beamer@sectionpages{\the\beamer@sectionstartpage}{\the\beamer@tempcount}}}}%
        \immediate\write\@auxout{\string\@writefile{nav}%
            {\noexpand\headcommand{\noexpand\beamer@documentpages{\the\beamer@tempcount}}}}
        \immediate\write\@auxout{\string\@writefile{nav}%
            {\noexpand\headcommand{\noexpand\def\noexpand\inserttotalframenumber{\the\c@framenumber}}}}
        \addtocounter{page}{-1}
        \immediate\write\@auxout{\string\@writefile{nav}%
            {\noexpand\headcommand{\noexpand\def\noexpand\inserttotalpagenumber{\thepage}}}}
        \newwrite\tf@nav
        \immediate\openout\tf@nav\jobname.nav\relax
        \newwrite\tf@toc
        \immediate\openout\tf@toc\jobname.toc\relax
        \newwrite\tf@snm
        \immediate\openout\tf@snm\jobname.snm\relax
    \fi%
}


% \newcommand\regardspage[1]{
% 	\setregardtemplates
% 	\frame{#1}%
% 	\setbeamertemplates
% }
% \newcommand\backmatterpage[1]{
% 	\setbackmattertemplates
% 	\frame{#1}%
% 	\setbeamertemplates
% }