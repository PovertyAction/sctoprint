*! version 1.1.0 Mehrab Ali, Shabib Raihan, Christopher Boyer and Chiara Pasquini 22jan2019

cap program drop sctoprint
program  sctoprint

version 15

/*
Things I added recently:
1. Choice list as: "code = label", instead of "code label"
2. Added "Please write down: _________" if it is a text field
3. Added a line for preloaded choice lists: "This list comes from X list from Y file."
4. Added spell out the constraint and relevances. 
5. Removed all html codes
6. pdf or word as option

To do:
Make table for repeat group
Drop or mark disabled
Annex codes
Attachments: image, option list
Carriage return in option list
*/


/** think about:
- create appendix of choice lists / figure out how to deal with long choice lists
- add groups
- field list?
*/

*  ----------------------------------------------------------------------------
*  1. Define syntax                                                            
*  ----------------------------------------------------------------------------

#d ;
	syntax using/,
	[Language(string)]
	save(string)
	Title(string)
	[replace]
	[Clear]
	[pdf] [word]
	;

#d cr

qui {

* PDF or Word option
if "`pdf'" == "" & "`word'" == "" {
	n di as err "Please specify etiher pdf or word option."
	exit 198
	} 

* Check output file
cap confirm file "`save'"
	if _rc & "`replace'" =="" {
		noi disp as err "The file `save' already exists. Please specify replace option."
		exit 602
	}


noi di "This might take few moments. Please wait... (☉.☉)" 


* create tempfiles to store survey and choices sheet that will be merged later on
tempfile survey choices original

save `original', emptyok

* 1. Prepare SurveyCTO form
* import the survey sheet
import excel using "`using'", clear firstrow sheet("survey") all

* Fix label macro 
local label label`language'
cap confirm var hint`language'
if !_rc  local hint hint`language'
if _rc local hint hint

* Fix relevance vs relevant
cap ren relevant relevance


* add repeated message 
gen re="[Repeated in " + name[_n-1] + "] " if type[_n-1]=="begin repeat" & type!="end repeat"
replace re=re[_n-1] if regex(re[_n-1],"Repeated in")==1 & re=="" & type[_n]!="end repeat"

* ensure string
ds hint* relevance constraint*
foreach var in `r(varlist)' {
	cap confirm str `var'
	if _rc cap tostring `var', replace
		if _rc {
			dis as error "`var' must be convertible to a string."
			exit `_rc'
		}

	if ustrregexm(`var', "constraint") replace `var' = usubinstr(constraint, "$", "", .)	
}


replace `label' = usubinstr(`label', `"""', "'", .)
replace `hint' = usubinstr(`hint', `"""', "'", .)
replace relevance = usubinstr(relevance, `"""', "'", .)
replace relevance = usubinstr(relevance, "$", "", .)
replace relevance = "Relevant if " + relevance if !mi(relevance)
replace `hint' = usubinstr(`hint', "$", "", .)
replace appearance = usubinstr(appearance, "$", "", .)
replace name = ustrtrim(ustrtrim(name))


* rename value value
drop if name == "" | ///
	inlist(type, "start", "end", "deviceid", "subscriberid", "simserial", "phonenumber", "text audit", "audio audit") | ///
	inlist(type, "speed violations count", "speed violations list", "calculate", "calculate_here", "comments", "username", "caseid") // get rid of rows with no name
replace `label'=usubinstr(`label',"$","",.) // remove dollar sign from references to other variables, leave the curly brackets to show it's a reference

//drop if ustrregexm(type, "calculate|comment")  // get rid of calculate fields , include calculate_here, and comment

keep type name label* hint* constraint relevance appearance re repeat_count

* extract list name from select_one and select_multiple questions
g list_name = word(type, -1) if ustrregexm(type, "(select_one)|(select_multiple)"), after(type)
gen type_d = type
replace type = usubinstr(word(type, 1), "_", " ", .)
foreach var in `label' `hint' {
	replace `var' = ustrregexra(`var', "<.*?>", " " ) if strpos(`var',"<")	

	* Converting common html entities
	replace `var' = ustrregexra(`var', "&nbsp;", "`=char(09)'	" ) if strpos(`var',"&nbsp;")
	replace `var' = ustrregexra(`var', "&lt;", "<" ) if strpos(`var',"&lt;")
	replace `var' = ustrregexra(`var', "&gt;", ">" ) if strpos(`var',"&gt;")
	replace `var' = ustrregexra(`var', "&amp;", "&" ) if strpos(`var',"&amp;")

	replace `var' = ustrtrim(strtrim(stritrim(`var')))
}





replace `hint' = "appearance: " + appearance + " " + `hint' if !mi(appearance)

save "`survey'", replace

************************************* Choices *******************************

* import the choices sheet, clean, reshape wide and create choice list strings
import excel using "`using'", clear sheet("choices") firstrow allstring
cap ren name value
keep list_name value `label'
* clean label and drop spaces
replace `label' = usubinstr(`label',"(","[",.)
replace `label' = usubinstr(`label',")","]",.) 
replace `label' = usubinstr(`label',"$","",.) 

drop if list_name==""

tempvar numbervalue
destring value, gen(`numbervalue') force
replace `numbervalue' = -1*`numbervalue' if `numbervalue' < 0
bysort list_name (`numbervalue') : gen j = _n  
drop `numbervalue'

reshape wide value `label', i(list_name) j(j)

ds value*
forval i = 1/`: word count `r(varlist)'' {
	gen choices`i' = strtrim(stritrim(value`i' + " = " + `label'`i')) if !mi(value`i')
	//replace choices`i' = "" if choices`i' == "."
}

save "`choices'", replace

************* merge survey and choices sheets together ********************

use "`survey'", clear

* keep original order
g row = _n

merge m:1 list_name using "`choices'", keepus(choice*) keep(1 3)

sort row 

forval x = 1/5 {
	g bucket`x' = ""
}

egen choicecount = rownonmiss(choices*), strok
recode choicecount (. = 0)

* for each variable
forval n = 1/`=_N' {

	loc first 1
	loc start 2

	* create buckets based on size. If fewer than 20, just create four groups with 5 in each
	if choicecount[`n'] > 20 {
			loc end = ceil(choicecount[`n']/4)
			loc counter = ceil(choicecount[`n']/4)
	} // end if

	else {
			loc end = 5
			loc counter = 5
	} // end else 
	

	* put choice lists into buckets
	forval bucket = 1/4 {

		*first value of bucket is the first variable	
		replace bucket`bucket' = choices`first'[`n'] if _n == `n'

		* loop through the rest that should go in the bucket and add them with a linebreak
		forval i = `start'/`end' {
			cap confirm var choices`i'
			if !_rc 	{

				replace bucket`bucket' = bucket`bucket'[`n'] +  char(13) +  choices`i'[`n'] if _n == `n' & choices`i'[`n'] != ""
			
			} //end if error

		} //end forval i = start/end

		loc first = `first' + `counter'
		loc start = `start' + `counter'
		loc end = `end' + `counter'
	
	} // end forval bucket = 1/4


} // forval n = 1/`=_N'
	
* drop extraneous variables
drop choices* _merge

* Add text if choice list comes from attachment
split appearance, p(')
cap confirm appearance2
if !_rc {
	replace bucket1 = "Choice list comes from " + list_name + " list from preloaded " + appearance2 + " file." if regex(appearance, "search")==1
}

* Add specific text if the field type is text
replace bucket1 = "Please write down: __________________________" if type=="text"

* Spell out the constraints and relevances
    *Replace functions with English phrases that have the same meaning.
 foreach var of varlist constraint relevance {
    if "`var'"=="constraint" {
   		local verb = "must be"	
    }
    
    else if "`var'"=="relevance" {
    	local verb = "is"
    }

    replace `var' = usubinstr(`var', "selected(", "", .)
	replace `var' = usubinstr(`var', "},", "} has ", .)
	//replace `var' = ustrregexrf(`var', "([0-9+])+(\))",  ustrregexs(3) + " selected" ) if ustrregexm(`var', "(.)+(,+)+([0-9]+)+(\))")
    replace `var'=usubinstr(`var', "string-length(", "length of ", .) 
    replace `var'=usubinstr(`var', ".", "answer ", .) 
    replace `var'=usubinstr(`var', ">=", " `verb' greater than or equal to ", .) 
    replace `var'=usubinstr(`var', "<=", " `varb' less than or equal to ", .) 
    replace `var'=usubinstr(`var', ">", " `varb' greater than ", .) 
    replace `var'=usubinstr(`var', "<", " `varb' less than ", .) 
    replace `var'=usubinstr(`var', "!=", " does not equal ", .)
    replace `var'=usubinstr(`var', "=", " `varb' equal to ", .)
    replace `var' = usubinstr(`var', "(", "", .) if length(`var') - length(subinstr(`var', ")", "", .))!=length(`var') - length(subinstr(`var', "(", "", .))
    replace `var' = usubinstr(`var', ")", "", .) if length(`var') - length(subinstr(`var', ")", "", .))!=length(`var') - length(subinstr(`var', "(", "", .))
    
    replace `var' = ustrtrim(strtrim(stritrim(`var')))
    replace `var' = usubinstr(`var', word(`var',1),proper(word(`var',1)),.)
    }


* reshape to allow for easy creation of columns
reshape long bucket, i(name list_name `label' `hint' type repeat_count re) j(var)  
sort row var

* create columns
gen col1 = cond(var==1, name, cond(var == 2, type, cond(var == 3, `label', cond(var == 4, `hint', "")))), before(name) 
gen col3 = cond(var == 1, relevance, cond(var == 2, constraint, "")), after(col1)

* add choice buckets
bysort row (var) : replace col1 = bucket[_n - 4] if var == 5 
bysort row (var) : g col2 = bucket[_n - 3] if var == 5, after(col1)
bysort row (var) : replace col3 = bucket[_n - 2] if var == 5
bysort row (var) : gen col4 = bucket[_n - 1] if var == 5, after(col3)


****************************************************************************

***** Create blocks for separate tables *******
* Dropping extra rows of groups
drop if inlist(type, "begin", "end") & row==row[_n-1]

* Generate text for begin group, begin repeat and end of group or repeat 
replace bucket ="("+ name + ") " + `label' + ": "  + relevance  if type_d=="begin group"   // begin group

replace repeat_count = "Can be added as many groups as needed" if mi(repeat_count) & !mi(re[_n+1]) & type_d=="begin"
replace repeat_count = " [Repeat condition: " + usubinstr(repeat_count, "$", "", .) + "]" if !mi(repeat_count) & type_d=="begin repeat"
replace bucket = "(Repeat group "+ name + ") " + `label' + ": "  + relevance  +  repeat_count if type_d=="begin repeat"

replace bucket = "Group "+ `"""' + name + " " + `label'+ `"""' + " ends." if type_d=="end group" // End group
replace bucket = "Repeat Group "+ `"""' + name + " " + `label'+ `"""' + " ends." if type_d=="end repeat" 
* Create block numbers
gen dum = inlist(type, "begin", "end")
gen block = 1 in 1
replace block = block[_n-1] + dum if _n>1

**********************************************************
if "`pdf'" != "" {

	* 2. Create document
	putpdf clear 

	*begin document
	putpdf begin, font("Calibri", 8)

	putpdf paragraph, halign(right)  
	putpdf text ("Created on: `c(current_date)'"), bold

	putpdf paragraph, halign(center)
	putpdf text ("`title'"), bold font(Calibri, 24)
}

if "`word'" != "" {
	* 2. Create document
	putdocx clear 

	*begin document
	putdocx begin, font("Calibri", 8)

	putdocx paragraph, halign(right)  
	putdocx text ("Created on: `c(current_date)'"), bold

	putdocx paragraph, halign(center)
	putdocx text ("`title'"), bold font(Calibri, 24)
}

levelsof block, loc(blocks)

nois _dots 0, title(Printing tables) reps(`r(r)')
foreach x of local blocks {
	nois _dots `x' 0
	preserve
	
	keep if block ==`x'

	if "`pdf'" != "" {

		putpdf paragraph, halign(left)
	
		if inlist(type,"end") & _n==1 {
			putpdf text (bucket[1]), bold font(Calibri, 10, "gray") linebreak(2)
			drop if inlist(type,"end")
		}
		if inlist(type, "begin") & _n==1 {
			putpdf text (bucket[1]), bold underline font(Calibri, 12, "cadetblue")
			drop if inlist(type, "begin")	
		}
	 
		if _N>0 {
			putpdf table outputtable`x' = data(col*), border(all, nil) //border(insideV, nil) border(end, nil) //border(end, nil) border(start, nil)
		
			forval i = 1(5)`=_N' {
				

				if inlist(type[`i'],"end", "begin")==1 putpdf table outputtable`x'(`i', 1), border(start, single)

				* left two rows colspan
				putpdf table outputtable`x'(`i', 1) = ("`=col1[`i']'"), bgcolor(lightgray)  bold
				putpdf table outputtable`x'(`i', 1), colspan(2) 
				putpdf table outputtable`x'(`i', 3) = ("`=col3[`i']'"), bgcolor(lightgray) italic 
				putpdf table outputtable`x'(`i', 3), colspan(2) halign(right) // change to 3

				* right two rows colspan
				putpdf table outputtable`x'(`=`i'+1', 1), colspan(2) 
				putpdf table outputtable`x'(`=`i'+1', 3), colspan(2) halign(right) 

				* question and hint colspan
				putpdf table outputtable`x'(`=`i'+2', 1), colspan(4) font("Calibri", 10) 
				putpdf table outputtable`x'(`=`i'+3', 1), colspan(4) italic font("Calibri", 8, darkslategray)

				* bg and nosplit
				putpdf table outputtable`x'(`i' `=`i'+1', .), bgcolor(lightgray)
				putpdf table outputtable`x'(`=`i'+4', .), nosplit 
				putpdf table outputtable`x'(`i', .),  border(top, single) 
				putpdf table outputtable`x'(`=`i'+4', .),  border(bottom, single)  

			}
		}
	}


	if "`word'" != "" {

		putdocx paragraph, halign(left)
	
		if inlist(type,"end") & _n==1 {
			putdocx text (bucket[1]), bold font(Calibri, 10, "gray") linebreak(2)
			drop if inlist(type,"end")
		}
		if inlist(type, "begin") & _n==1 {
			putdocx text (bucket[1]), bold underline font(Calibri, 12, "cadetblue")
			drop if inlist(type, "begin")	
		}
	 
		if _N>0 {
			putdocx table outputtable`x' = data(col*), border(all, nil) //border(insideV, nil) border(end, nil) //border(end, nil) border(start, nil)
		
			forval i = 1(5)`=_N' {

				* left two rows colspan
				putdocx table outputtable`x'(`i', 1) = ("`=col1[`i']'"),   bold  shading(lightgray)
				putdocx table outputtable`x'(`i', 1), colspan(2) 
				putdocx table outputtable`x'(`i', 3) = ("`=col3[`i']'"),  italic shading(lightgray)
				putdocx table outputtable`x'(`i', 2), colspan(2) halign(right) 

				* right two rows colspan
				putdocx table outputtable`x'(`=`i'+1', 1), colspan(2) 
				putdocx table outputtable`x'(`=`i'+1', 2), colspan(2) halign(right) 

				* question and hint colspan
				putdocx table outputtable`x'(`=`i'+2', 1), colspan(4) font("Calibri", 10) 
				putdocx table outputtable`x'(`=`i'+3', 1), colspan(4) italic font("Calibri", 8, darkslategray)

				* bg and nosplit
				putdocx table outputtable`x'(`i' `=`i'+1', .), shading(lightgray)
				putdocx table outputtable`x'(`=`i'+4', .), nosplit 
				putdocx table outputtable`x'(`i', .),  border(top, single) 
				putdocx table outputtable`x'(`=`i'+4', .),  border(bottom, single)  

			}
		}
	}

	restore
}
	


	if "`pdf'" != "" putpdf save "`save'", replace
	if "`word'" != "" putdocx save "`save'", replace

	if "`clear'" == "" use `original', clear

	if "`pdf'" != "" loc savenew = "`save'"+".pdf"
	if "`word'" != "" loc savenew = "`save'"+".docx"


	noi display `"The print version questionnaire is saved here {browse "`savenew'":`save'}"' 	
	



} //qui bracket

end
