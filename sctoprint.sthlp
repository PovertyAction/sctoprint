{smcl}

{title:Title}

{cmd:sctoprint} {hline 2} Produces printable pdf and word document from SurveyCTO or ODK XLSForm.


{title:Syntax}

{p 8 10 2}
{cmd:sctoprint using} {it:help filename} 
[{cmd:,} {it:options}]


{synoptset 23 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:Main}
{synopt:{opt t:itle(string)}}title of the survey {p_end}
{synopt:{opth save(filename)}}file name with folder location to save as{p_end}

{syntab:Options}
{synopt:{opt l:anguage}}languge to use to export if multiple languages used in {cmd:XLSForm}{p_end}
{synopt:{opt pdf}}to export as a {cmd:PDF} file{p_end}
{synopt:{opt word}}to export as a {cmd:Word} file {p_end}
{synopt:{opt replace}}replace existing file {p_end}
{synopt:{opt clear}}clear and reload data in the memory after finishing exporting {p_end}
{synoptline}
{p2colreset}{...}


{title:Description} 

{pstd}
{cmd:sctoprint} exports print friendly questionnaire from XLSForm definition used for SurveyCTO or ODK.



{title:Options}

{phang}
{opt t:itle(string)} title to print as heading for the questionnaire.

{phang}
{opth save(filename)} specifies the location and file name to export the printable questionnaire as.

{phang}
{opt l:anguage} if multiple languages used in XLSForm, specify which language to be used for export. If no language specified in XLSForm, no need to use this option.

{phang}
{opt pdf} exports pdf document.

{phang}
{opt word} exports word document.

{phang}
{opt replace} specifies to overwrite filename, if it exists.

{phang}
{opt clear} save any data in memory before running the program and reload after execution.


{title:Examples} 

{phang}
{com}. sctoprint using "X:/Projects 2020/01_instruments/03_xls/Phase one_v1.xlsx", title("Household Questionnaire") save(X:/Projects 2020/01_instruments/02_print/Phase one_v1") pdf replace clear{p_end}

{phang}
{com}. sctoprint using "X:/Projects 2020/01_instruments/03_xls/Phase one_v1.xlsx", title("Household Questionnaire") save(X:/Projects 2020/01_instruments/02_print/Phase one_v1") word replace clear{p_end}


{title:Remarks}

{pstd}The GitHub repository for {cmd:sctoprint} is {browse "https://github.com/PovertyAction/sctoprint":here}.

{title:Author}

{pstd}Mehrab Ali, Rosemarie Sandino, Shabib Raihan, Christopher Boyer and Chiara Pasquini{p_end}
{pstd}Last updated: January 21, 2020{p_end}
	
