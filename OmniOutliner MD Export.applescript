use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
-- use script "RegexAndStuffLib"

-- bcdav  2021-07-11-12-47
-- https://github.com/bcdavasconcelos/Ooutline-to-PDF
-- with a little help from my friends (https://discourse.omnigroup.com/t/script-to-convert-oo-to-markdown-tex-and-pdf-using-pandoc/31429/15)

property UseSelection : false -- if false → all rows
property IncludeRowNote : true
property ConvertWithPandoc : false
property OpenAfterConversion : true

set PandocPath to "export PATH=/Library/TeX/texbin:$PATH && /usr/local/bin/pandoc"
set PandocDefaults to "-drefs2 -dabntex -dpdf"


set theCompleteText to {}

set theRows to {}

tell application "OmniOutliner"
	
	set theFile to get file of front document
	set thePath to POSIX path of theFile
	
	tell front document
		
		
		if UseSelection then
			set theRows to selected rows
		else
			set theRows to rows
		end if
		
		--		set theRow to item 1 of theRows
		repeat with theRow in theRows
			
			
			set {RowId, RowText, RowNote, RowStyles, RowPrefix, RowSufix} to {"", "", "", "", "", ""} -- clean variables
			
			-- 1. Row styles
			
			set RowId to the id of theRow
			set RowId to "[]{#" & RowId & "}"
			set RowStyles to name of named styles of style of theRow -- retrieve named Styles
			
			
			set {RowPrefix, RowSufix} to my TranslateStylesToMarkdown(RowStyles, RowId) -- retrieve what should come before or/and after the row text 
			
			
			-- 2. Main text with bold, italics and links
			
			set {lsText, lsFont, lsLinks, lsStyles} to {its text, its font, its attribute "link" of style, its name of named styles of styles} of topic's attribute runs of theRow -- necessary data for translating to md
			
			set RowText to my TranslateRtfToMarkdown(lsText, lsFont, lsLinks, lsStyles) -- translate rtf to md
			
			-- 3. Note text with bold, italics and links
			
			set {lsTextofNote, lsFontofNote, lsLinksofNote, lsStylesofNote} to {its text, its font, its attribute "link" of style, its name of named styles of styles} of note's attribute runs of theRow -- necessary data for translating to md
			set RowNote to my TranslateRtfToMarkdown(lsTextofNote, lsFontofNote, lsLinksofNote, lsStylesofNote) -- translate rtf to md
			
			
			set RowText to RowPrefix & RowText & RowSufix & linefeed
			
			if IncludeRowNote then
				if RowNote is not "" then set RowText to RowText & linefeed & RowNote & "  " & linefeed & linefeed
			end if
			
			set theCompleteText to theCompleteText & RowText
			
			
		end repeat
		
		set theCompleteText to my RegexReplace(theCompleteText)
		
		set the clipboard to (theCompleteText as text)
		--	return theCompleteText as text
		
	end tell
end tell


if ConvertWithPandoc then
	set theMDPath to thePath & ".md"
	set thePDFPath to thePath & ".pdf"
	
	do shell script "touch " & quoted form of theMDPath & " && LANG=pt_BR.UTF-8 pbpaste > " & quoted form of theMDPath
	set theSH to PandocPath & space & "-s" & space & quoted form of theMDPath & space & PandocDefaults & space & "-o" & space & quoted form of thePDFPath
	if OpenAfterConversion then set theSH to theSH & "&& open " & quoted form of thePDFPath
	do shell script theSH
	
end if


on fixHomePath(thePath)
	if thePath contains "~/" then
		set thePath to replacetext(thePath, "~/", "$HOME/")
	else
		set HomePath to (POSIX path of (path to home folder))
		set thePath to replacetext(thePath, HomePath, "$HOME/")
	end if
	return thePath
end fixHomePath


on replacetext(theString, old, new)
	set {TID, text item delimiters} to {text item delimiters, old}
	set theStringItems to text items of theString
	set text item delimiters to new
	set theString to theStringItems as text
	set text item delimiters to TID
	return theString
end replacetext


on RegexReplace(theText)
	
	--	set theText to regex change theText search pattern "(-@.+?\\b)" replace template "[$1]"
	--	set theText to regex change theText search pattern "(\\p{Greek}+)" replace template "\\\\grc{$1}"
	
	return theText
end RegexReplace

on TranslateStylesToMarkdown(RowStyles, RowId)
	set {pre, pos} to {RowId, "  " & linefeed}
	
	if RowStyles contains "Heading 1" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "# ", "  " & linefeed}
	else if RowStyles contains "Heading 2" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "## ", "  " & linefeed}
	else if RowStyles contains "Heading 3" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "### ", "  " & linefeed}
	else if RowStyles contains "Heading 4" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "#### ", "  " & linefeed}
	else if RowStyles contains "Heading 5" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "##### ", "  " & linefeed}
	else if RowStyles contains "Heading 6" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "###### ", "  " & linefeed}
	else if RowStyles contains "Heading 7" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "####### ", "  " & linefeed}
	else if RowStyles contains "HeadingNo" then
		set {pre, pos} to {RowId & "  " & linefeed & linefeed & "# ", " {-}  " & linefeed}
	else if RowStyles contains "YAML" then
		set {pre, pos} to {"", ""}
	else if RowStyles contains "Ordered List" then
		set {pre, pos} to {"1. ", ""}
	else if RowStyles contains "Unordered List" then
		set {pre, pos} to {"- ", ""}
	else if RowStyles contains "Paragraph" then
		set {pre, pos} to {linefeed & RowId & "  " & linefeed, "  " & linefeed}
	else if RowStyles contains "YAMLitem" then
		set {pre, pos} to {"", ": |"}
	else if RowStyles contains "YAMLtext" then
		set {pre, pos} to {linefeed & "  ", "  "}
	else if RowStyles contains "Small" then
		set {pre, pos} to {linefeed & "<small>", "</small>"}
	else if RowStyles contains "Blockquote" then
		set {pre, pos} to {linefeed & "> ", " " & RowId & "  " & linefeed}
	else if RowStyles contains "Code" then
		set {pre, pos} to {linefeed & RowId & "  " & linefeed & "> ", "  " & linefeed}
	else if RowStyles contains "Comment" then
		set {pre, pos} to {linefeed & "<!--", "-->" & linefeed}
	else if RowStyles contains "Ignore" then
		set {pre, pos} to {linefeed & "<!--", "-->" & linefeed}
	end if
	
	return {pre, pos}
end TranslateStylesToMarkdown

on TranslateRtfToMarkdown(lsText, lsFont, lsLinks, lsStyles)
	using terms from application "OmniOutliner"
		set outTxt to ""
		
		repeat with i from 1 to lsText's length
			
			set {thePrefix, theSufix} to {"", ""}
			set {aStr, aFont, aLink, aStyle} to {lsText's item i, lsFont's item i, lsLinks's item i, lsStyles's item i}
			
			if aFont contains "bold" or aFont contains "black" or aFont contains "Bd" then set {thePrefix, theSufix} to {"**", "**"}
			
			if aFont contains "italic" or aFont contains "It" then set {thePrefix, theSufix} to {thePrefix & "*", theSufix & "*"}
			
			
			if has local value of aLink then
				set theLink to value of aLink
				if theLink contains "omnioutliner" then
					set theLink to my replacetext(theLink, "omnioutliner:///open?row=", "#")
					set {thePrefix, theSufix} to {thePrefix & "[", "](" & theLink & ")" & theSufix}
				else
					set {thePrefix, theSufix} to {thePrefix & "[", "](" & theLink & ")" & theSufix}
				end if
			end if
			
			if aStyle contains "Index" then
				set {thePrefix, theSufix} to {"\\index{", "}"}
			end if
			
			set outTxt to outTxt & thePrefix & aStr & theSufix
			
		end repeat
		return outTxt
	end using terms from
end TranslateRtfToMarkdown

