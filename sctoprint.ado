 * ---------------------------------------- *
 * file:    make_word_survey.do               
 * author:  Christopher Boyer and Chiara Pasquini,              
 *          Innovations for Poverty Action
 * date:    2019-03-20                     
 * ---------------------------------------- *
 * outputs: 
 *   @3_word/citizen_survey.docx

 
 /* let's create a nice word version of a questionnaire
	directly from our SurveyCTO coding */
	
 * -------------------------------------------------------------------------- *
 * 1. Prepare SurveyCTO form
 
 *Set locals
 local ctoform "" //enter directory for excel form 
 local wordsurvey "" //enter directory you want to give to word doc
 local title "" //enter title that you want to appear as header of the word doc
 
 * create tempfiles to store survey and choices sheet that will be merged later on
 clear all
 tempfile survey choices
 
 * import the survey sheet
	 import excel using "`ctoform'", clear firstrow
	 drop if labelenglish == "" // get rid of rows with no labels
	 replace labelenglish=subinstr(labelenglish,"$","",.) // remove dollar sign from references to other variables, leave the curly brackets to show it's a reference
	 drop if type=="calculate" // get rid of calculate fields

	*Drop PROJECT SPECIFIC questions you don't want in the word version (these are rows in the survey sheet of the excel form)
	 drop if name=="enum" // enum is the variable where enumerators select their names, I don't want it to appear in the word version
	 
	*Drop PROJECT SPECIFIC variables not needed for the word doc (these are columns in the survey sheet of the excel form) 
	 drop labelJapdhola labelrunyoro labelateso labelluganda labellugbara labellugisu labelluo labellusoga labelrunyankole default disabled readonly calculation repeat_count mediaimage mediaaudio mediavideo note response_note AD-BL
	 
	 * extract list name from select_one and select_multiple questions
	 g list_name = ""
	 replace list_name = word(type, -1) if regexm(type, "(select_one)|(select_multiple)")
	 
	 save "`survey'", replace
 
 * import the choices sheet
	 import excel using "`ctoform'", clear firstrow sheet("choices")
	 drop if list_name==""
	 keep list_name name labelenglish
	 
	 * this is necessary to make sure choice options are sorted properly
	 drop if !regexm(name, "[-0-9]") 
	 destring name, replace
	 
	 * need to reshape the choices sheet so there's one list per row
	 bys list_name (name): g j = _n
	 reshape wide name labelenglish, i(list_name) j(j)
	 
	 * concatenate the value-label pairs
	 forval i = 1/74 {
		qui egen choices`i' = concat(name`i' labelenglish`i'), punct(" ")
	 }

	 save "`choices'", replace
 
 * merge survey and choices sheets together
	 use "`survey'", clear
	 
	 * keep original order
	 g row = _n
	 
	 merge m:1 list_name using "`choices'", keepus(choices1-choices74) keep(1 3)
	 
	 sort row 
 
 * -------------------------------------------------------------------------- *
 
 * 2. Create Word document
 
 putdocx clear 

 *begin document
 putdocx begin, font("Calibri", 9) landscape
 
 putdocx paragraph, /// 
	style(Title) /// 
	halign(center)

 *title	
 putdocx text ("`title'"), bold
 
  *set up of table and put variable name, questions and options in the table
 putdocx paragraph

 gen empty=""
 putdocx table tab = data(name labelenglish empty), headerrow(1) layout(autofitcontents)
 
 putdocx table tab(1, .), addrows(1, before)
 putdocx table tab(1, 1) = ("Variable name"), bold
 putdocx table tab(1, 2) = ("Question"), bold
 putdocx table tab(1, 3) = ("Choices"), bold

 forval i = 2/`=_N' {
	
	forval j = 1/74 {
		if !inlist(choices`j'[`i'-1], ".", "") {
			putdocx table tab(`i', 3) = ("`=choices`j'[`i'-1]'"), append linebreak
		}
	}
 }

*add hints to question cell in italic
 forval i = 2/`=_N' {
	
		if !inlist(hint[`i'-1], ".", "") {
			putdocx table tab(`i', 2) = ("`=empty[`i'-1]'"), append linebreak
			putdocx table tab(`i', 2) = ("HINT: "), append italic	
			putdocx table tab(`i', 2) = ("`=hint[`i'-1]'"), append italic
		}
	}

*add relevance expressions
replace relevance=subinstr(relevance,"$","",.)
replace relevance=subinstr(relevance,"{","",.)
replace relevance=subinstr(relevance,"}","",.)

 forval i = 2/`=_N' {
	
		if !inlist(relevance[`i'-1], ".", "") {
			putdocx table tab(`i', 2) = ("`=empty[`i'-1]'"), append linebreak
			putdocx table tab(`i', 2) = ("Relevant if "), append
			putdocx table tab(`i', 2) = ("`=relevance[`i'-1]'"), append
		}
	}

 putdocx save "`wordsurvey'", replace

 
