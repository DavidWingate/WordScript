﻿-- WordScript for BibDesk

---

-- MIT License

-- Copyright (c) 2020 David Wingate

-- Permission is hereby granted, free of charge, to any person obtaining a copy
-- of this software and associated documentation files (the "Software"), to deal
-- in the Software without restriction, including without limitation the rights
-- to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
-- copies of the Software, and to permit persons to whom the Software is
-- furnished to do so, subject to the following conditions:

-- The above copyright notice and this permission notice shall be included in all
-- copies or substantial portions of the Software.

-- THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
-- IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
-- FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
-- AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
-- LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
-- OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
-- SOFTWARE.

---

-- Acknowledgements

-- The basic functionality of this software was inspired by three existing projects:

-- - BibFuse by Colin A. Smith (http://bibfuse.sourceforge.net)
-- - BibDeskToWord by Conan C. Albrecht
-- - NatBib Reference Sheet by Sébastien Merkel (http://merkel.texture.rocks/Latex/natbib.php)

-- Many thanks to the authors of these projects for sharing their hard work.

---

-- check that Word and BibDesk are ready
tell application "System Events"
	set {countWord, countBibDesk} to {0, 0}
	
	if (exists of process "Microsoft Word") is true then
		set countWord to (count of documents in application "Microsoft Word")
	end if
	
	if (exists of process "BibDesk") is true then
		set countBibDesk to (count of documents in application "BibDesk")
	end if
end tell

if {countWord, countBibDesk} contains 0 then
	activate
	
	try
		display dialog "Documents must be open in both Microsoft Word and BibDesk" buttons {"Cancel", "Open Now"} default button "Open Now"
		
		if button returned of result contains "Open Now" then
			tell application "BibDesk"
				launch
			end tell
			
			delay 3
		end if
		
	on error
		return
	end try
end if

-- check for selection in Word to trigger quick format
tell application "Microsoft Word"
	if content of text object of first paragraph of selection contains "ADDIN" then
		select text object of first paragraph of selection -- select recently edited reference
	end if
	
	set selectionStart to (selection start of selection)
	set selectionEnd to (selection end of selection)
end tell

if selectionStart is not selectionEnd then
	readBibliographySettings()
	
	set bibliographySettings to result
	
	if bibliographySettings is {} then
		defaultSettings()
		
		set bibliographySettings to result
		
		writeBibliographySettings(bibliographySettings)
		
		set bibliographySettings to result
		
		if bibliographySettings is {} then
			return
		end if
		
	end if
	formatReferences(bibliographySettings)
	
	return
end if

-- open WordScript Menu
activate

with timeout of 0 seconds
	choose from list {"Format References", "Unformat References", "Fill Empty References", "Tools", "Document Settings"} with title "WordScript" with prompt "You should save your work first!" default items "Format References" cancel button name {"Exit"} OK button name {"Choose"}
	
	set menuChoice to result
end timeout

if menuChoice is false then
	return
end if

if menuChoice contains "Format References" then
	readBibliographySettings()
	
	set bibliographySettings to result
	
	if bibliographySettings is {} then
		defaultSettings()
		
		set bibliographySettings to result
		
		writeBibliographySettings(bibliographySettings)
		
		set bibliographySettings to result
		
		if bibliographySettings is {} then
			return
		end if
	end if
	
	formatReferences(bibliographySettings)
	
else if menuChoice contains "Unformat References" then
	unformatReferences()
	
else if menuChoice contains "Fill Empty References" then
	fillEmptyReferences()
	
else if menuChoice contains "Document Settings" then
	readBibliographySettings()
	
	set bibliographySettings to result
	
	if bibliographySettings is {} then
		defaultSettings()
		
		set bibliographySettings to result
		
		writeBibliographySettings(bibliographySettings)
		
		set bibliographySettings to result
		
		if bibliographySettings is {} then
			return
		end if
	end if
	
	settingsMenu(bibliographySettings)
	
	if result is not bibliographySettings then
		set bibliographySettings to result
		
		writeBibliographySettings(bibliographySettings)
		
		set bibliographySettings to result
		
		if bibliographySettings is {} then
			return
		end if
		
		try
			tell application "Microsoft Word"
				activate
				
				display dialog "Settings have been changed, do you wish to format the references now?"
			end tell
			
			formatReferences(bibliographySettings)
			
		on error
			return
		end try
		
	else
		return
	end if
	
else if menuChoice contains "Tools" then
	try
		display dialog "References must be unformatted, do you wish to unformat now?"
		
		unformatReferences()
		
	on error
		return
	end try
	
	toolsMenu()
	
	try
		activate
		
		display dialog "Operation complete, do you wish to format the references now?"
		
		readBibliographySettings()
		
		set bibliographySettings to result
		
		formatReferences(bibliographySettings)
		
	on error
		return
	end try
end if

---

on defaultSettings()
	set bibDeskDocumentName to ""
	set bibliographyTemplatePath to ""
	set citepTemplatePath to ""
	set citetTemplatePath to ""
	set parenthesisOpen to "("
	set parenthesisClose to ")"
	set bibliographySortBy to "Cite Key" -- BibDesk field (leave blank for order of appearance)
	set bibliographyStyle to "Bibliography" -- leave blank to use formatting of template file
	set italicLatin to {"et al.", "ibid.", "cf.", "inter alia", "circa"}
	set italicPublicationTypes to {"film", "broadcast"}
	set superscriptReferences to "No"
	set numberedReferences to "No" -- overrides templates and sorting
	set numberSeperator to ","
	
	return {bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator}
end defaultSettings

---

on readBibliographySettings()
	-- global setup
	tell application "Microsoft Word"
		set wordDocument to active document
		set selectionRange to create range wordDocument start (selection start of selection) end (selection end of selection)
		set status bar to "Reading bibliography..."
	end tell
	
	-- look for a formatted bibliography
	set AppleScript's text item delimiters to ""
	
	tell application "Microsoft Word"
		set fieldCount to (count of fields in wordDocument)
		
		repeat until fieldCount is 0
			set currentField to (field fieldCount of wordDocument)
			set fieldCode to content of field code of currentField
			
			if field type of currentField is field addin and word 2 of fieldCode contains "bibliography" then
				set bibliographySettings to characters 21 thru -2 of fieldCode as string
				
				exit repeat
			end if
			
			set fieldCount to fieldCount - 1
		end repeat
		
		if fieldCount is 0 then -- look for an unformatted bibliography
			set findObject to find object of selection -- set up the find system in Word
			set match wildcards of findObject to true
			set wrap of findObject to find continue
			set content of findObject to "\\\\bibliography*\\}"
			
			select (create range wordDocument start 0 end 0)
			
			try -- count the number of instances found
				set findCount to 0
				
				repeat while (execute find findObject)
					set findCount to findCount + 1
				end repeat
				
			on error -- this can break if advanced find and replace has been used in Word
				display dialog "Something is wrong with the find system in Microsoft Word, please quit Microsoft Word then re-open" buttons {"OK"} cancel button {"OK"} default button {"OK"}
				
				return
			end try
			
			if the findCount is not 0 then -- extract raw bibiography settings
				set bibliographySettings to characters 15 thru -2 of (content of selection as string) as string
				
			else if findCount is 0 then -- no bibliography found
				return {}
			end if
		end if
	end tell
	
	try -- extract bibliography settings as a list
		set AppleScript's text item delimiters to {"<bibDeskDocumentName=", "><bibliographyTemplatePath=", "><citepTemplatePath=", "><citetTemplatePath=", "><parenthesisOpen=", "><parenthesisClose=", "><bibliographySortBy=", "><bibliographyStyle=", "><italicLatin=", "><italicPublicationTypes=", "><superscriptReferences=", "><numberedReferences=", "><numberSeperator=", ">"}
		set {bibliographyCommand, bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator} to text items of bibliographySettings
		
	on error -- if anything is amiss, recreate the bibliography field
		tell application "Microsoft Word"
			display dialog "There is something wrong with the bibliography data, please delete the bibliography and try again" buttons {"OK"} default button {"OK"}
		end tell
		
		return
	end try
	
	set AppleScript's text item delimiters to "," -- recover lists from settings
	set italicLatin to text items of italicLatin
	set italicPublicationTypes to text items of italicPublicationTypes
	set AppleScript's text item delimiters to ""
	
	try -- check to make sure that the template files exist
		tell application "Finder"
			get (exists of (POSIX file (bibliographyTemplatePath as string) as alias))
			get (exists of (POSIX file (citepTemplatePath as string) as alias))
			get (exists of (POSIX file (citetTemplatePath as string) as alias))
		end tell
		
	on error -- if not, recreate the bibliography field
		tell application "Microsoft Word"
			display dialog "Template files not found, please delete the bibliography and try again" buttons {"OK"} default button {"OK"}
		end tell
		
		return
	end try
	
	-- restore Word to its original position
	tell application "Microsoft Word"
		try
			select selectionRange
			
		on error -- produces error if cursor was at the end of the document
			select (create range wordDocument start (selection end of selection) end (selection end of selection))
		end try
		
		set status bar to "Finished reading bibliography"
	end tell
	
	return {bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator}
end readBibliographySettings

---

on writeBibliographySettings(bibliographySettings)
	-- global setup
	tell application "Microsoft Word"
		set wordDocument to active document
		set selectionRange to create range wordDocument start (selection start of selection) end (selection end of selection)
		set status bar to "Writing bibliography..."
	end tell
	
	tell application "BibDesk"
		set bibDeskDocument to front document
	end tell
	
	-- parse bibliography settings
	set AppleScript's text item delimiters to ""
	set {bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator} to bibliographySettings
	-- parse custom lists from the bibliography settings
	set AppleScript's text item delimiters to ","
	set italicLatin to text items of italicLatin
	set italicPublicationTypes to text items of italicPublicationTypes
	set AppleScript's text item delimiters to ""
	
	-- look for formatted bibliography
	tell application "Microsoft Word"
		set fieldCount to (count of fields in wordDocument)
		
		if fieldCount is not 0 then
			repeat until fieldCount is 0
				if content of field code of field fieldCount of wordDocument contains "ADDIN bibliography" then
					set bibliographyField to field fieldCount of wordDocument
					set show codes of bibliographyField to true
					
					exit repeat
				end if
				
				set fieldCount to fieldCount - 1
			end repeat
		end if
	end tell
	
	-- if none found, create the bibliography field
	if fieldCount is 0 then
		try
			tell application "Microsoft Word"
				display dialog "No bibliography found, do you wish to create one now?"
				set bibliographyField to make new field at end of text object of wordDocument with properties {field type:field quote, show codes:true, field text:"*"}
			end tell
			
		on error -- return the bibliography settings empty
			return {}
		end try
		
		if bibDeskDocumentName is "" then -- if default settings were used, set the BibDesk library
			tell application "BibDesk"
				set bibDeskDocumentName to name of bibDeskDocument
			end tell
		end if
		
		if bibliographyTemplatePath is "" then -- if default settings were used, set the template paths
			tell application "Microsoft Word"
				set bibliographyTemplatePath to POSIX path of (choose file with prompt "Please choose WordScript Template Bibliography" of type {"public.rtf", "com.microsoft.word.doc"})
			end tell
		end if
		
		set AppleScript's text item delimiters to "/" --  guess the other template paths
		set assumedTemplatePath to (text items 1 through -2 of bibliographyTemplatePath) as string
		set AppleScript's text item delimiters to ""
		
		tell application "Finder"
			try
				set citepTemplatePath to POSIX path of (POSIX file (assumedTemplatePath & "/" & "WordScript Template Author-Year (citep).txt" as string) as alias)
				set citetTemplatePath to POSIX path of (POSIX file (assumedTemplatePath & "/" & "WordScript Template Year Only (citet).txt" as string) as alias)
			end try
		end tell
		
		if citepTemplatePath is "" then --  otherwise choose manually
			tell application "Microsoft Word"
				set citepTemplatePath to POSIX path of (choose file with prompt "Please choose WordScript Template Author-Year (citep)" of type {"public.plain-text"})
			end tell
		end if
		
		if citetTemplatePath is "" then
			tell application "Microsoft Word"
				set citetTemplatePath to POSIX path of (choose file with prompt "Please choose WordScript Template Year Only (citet)" of type {"public.plain-text"})
			end tell
		end if
	end if
	
	-- write the bibiography field in Word
	tell application "Microsoft Word"
		set AppleScript's text item delimiters to ","
		set content of field code of bibliographyField to " ADDIN bibliography{" & "<bibDeskDocumentName=" & bibDeskDocumentName & "><bibliographyTemplatePath=" & bibliographyTemplatePath & "><citepTemplatePath=" & citepTemplatePath & "><citetTemplatePath=" & citetTemplatePath & "><parenthesisOpen=" & parenthesisOpen & "><parenthesisClose=" & parenthesisClose & "><bibliographySortBy=" & bibliographySortBy & "><bibliographyStyle=" & bibliographyStyle & "><italicLatin=" & italicLatin & "><italicPublicationTypes=" & italicPublicationTypes & "><superscriptReferences=" & superscriptReferences & "><numberedReferences=" & numberedReferences & "><numberSeperator=" & numberSeperator & ">}"
		set AppleScript's text item delimiters to ""
		
		try -- restore the cursor to its original position
			select selectionRange
			
		on error -- produces error if cursor was at the end of the document
			select (create range wordDocument start (selection end of selection) end (selection end of selection))
		end try
	end tell
	
	-- restore Word to its original position
	tell application "Microsoft Word"
		try
			select selectionRange
			
		on error -- produces error if cursor was at the end of the document
			select (create range wordDocument start (selection end of selection) end (selection end of selection))
		end try
		
		set status bar to "Finished writing bibliography"
	end tell
	
	return {bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator}
end writeBibliographySettings

---

on formatReferences(bibliographySettings)
	-- global setup
	tell application "Microsoft Word"
		set wordDocument to active document
		set selectionRange to create range wordDocument start (selection start of selection) end (selection end of selection)
		set status bar to "Formatting references..."
	end tell
	
	tell application "BibDesk"
		set bibDeskDocument to front document
	end tell
	
	-- parse bibliography settings
	set AppleScript's text item delimiters to ""
	set {bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator} to bibliographySettings
	
	-- parse custom lists from the bibliography settings
	set AppleScript's text item delimiters to ","
	set italicLatin to text items of italicLatin
	set italicPublicationTypes to text items of italicPublicationTypes
	set AppleScript's text item delimiters to ""
	
	-- read the plain text template data (should this move to 'format author-year references'?)
	set citepTemplate to read citepTemplatePath
	set citetTemplate to read citetTemplatePath
	
	-- intitialise document lists
	set citeKeyList to {}
	set publicationList to {}
	set ibidCiteKeyList to {}
	set italicReferences to {}
	
	-- begin formatting the references
	tell application "Microsoft Word"
		activate
		
		-- establish the format range
		if (selection start of selection) is equal to (selection end of selection) then
			if content of selection is "]" then -- just edited page number
				set formatRange to words of sentences of selection
				
			else
				set formatRange to wordDocument -- format entire document
				select (create range wordDocument start 0 end 0)
			end if
			
		else
			set formatRange to selectionRange -- format selection only
		end if
		
		-- set up the find system in Word
		set findObject to find object of selection
		set match wildcards of findObject to true
		set wrap of findObject to find continue
		
		-- convert unformatted cite commands to ADDIN fields
		repeat with citeCommand in {"citep", "citet", "citealp", "citealt", "bibliography"}
			set content of findObject to "\\\\" & citeCommand & "*\\}"
			
			select formatRange
			
			try
				repeat while (execute find findObject)
					set citeCommand to content of selection
					set content of selection to ""
					
					make new field at (text object of selection) with properties {field type:field quote, show codes:true, field text:"*"}
					
					set currentField to first field of (create range wordDocument start ((selection start of selection) - 1) end (selection end of selection))
					set content of field code of currentField to " ADDIN " & characters 2 thru -1 of citeCommand
				end repeat
				
			on error -- this can break if advanced find and replace has been used in Word
				display dialog "Something is wrong with the find system in Microsoft Word, please quit Microsoft Word then re-open" buttons {"OK"} default button {"OK"}
				
				return
			end try
		end repeat
		
		repeat with fieldNumber from 1 to count of fields of formatRange -- convert ADDIN fields to formatted references
			set currentField to field fieldNumber of formatRange
			set fieldCode to content of field code of currentField
			
			if field type of currentField is field addin and {"citep", "citet", "citealp", "citealt"} contains word 2 of fieldCode then -- extract cite key and prefix/suffix
				select currentField
				
				set AppleScript's text item delimiters to {"{", "}"} -- parse cite key
				set citeKey to text item 2 of fieldCode
				set AppleScript's text item delimiters to {"[", "]"} -- parse prefix/suffix
				set referencePrefix to ""
				set referenceSuffix to ""
				
				if (count of text items of fieldCode) is 5 then -- two pairs of square brackets
					set referencePrefix to text item 2 of fieldCode
					set referenceSuffix to text item 4 of fieldCode
					
				else if (count of text items of fieldCode) is 3 then -- one pair of square brackets
					set referenceSuffix to text item 2 of fieldCode
				end if
				
				tell application "BibDesk" -- match cite keys to publications
					set multiplePublications to {} -- intitialise the multiple reference lists
					set multipleCiteKeys to {}
					set AppleScript's text item delimiters to "," -- multiple cite keys must be seperated by commas in BibDesk
					
					repeat with eachInteger from 1 to count of text items of citeKey -- loop through every cite key and match to a publication
						set currentCiteKey to item eachInteger of text items of citeKey
						set currentPublication to (publications in bibDeskDocument whose cite key is currentCiteKey)
						
						if (count of currentPublication) is 0 then -- no match found
							tell application "Microsoft Word"
								set show codes of currentField to true
								
								select field code of currentField
								
								select (create range wordDocument start ((selection start of selection) - 1) end (selection end of selection) + 1)
								
								display dialog "No BibDesk reference matches this cite key: " & currentCiteKey buttons {"OK"} default button {"OK"}
								
								return
							end tell
						end if
						
						if ibidCiteKeyList is not {} and last item of ibidCiteKeyList is currentCiteKey then -- check for ibid. references
							set ibidReference to true
							
						else
							set ibidReference to false
						end if
						
						set ibidCiteKeyList to ibidCiteKeyList & currentCiteKey
						
						if citeKeyList does not contain currentCiteKey then -- list all cite keys
							set citeKeyList to citeKeyList & currentCiteKey
						end if
						
						repeat with eachInteger from 1 to the count of citeKeyList -- get number for numbered references
							if item eachInteger of citeKeyList is currentCiteKey then
								set citeKeyNumber to eachInteger
							end if
						end repeat
						
						set multiplePublications to multiplePublications & currentPublication -- list multiple publications in current field
						
						-- list of all publications
						if publicationList does not contain currentPublication then set publicationList to publicationList & currentPublication
						
						-- list references to be made italic such as films
						if italicPublicationTypes contains type of (publications in bibDeskDocument whose cite key is citeKey) then
							set italicReferenceTitle to title of (publications in bibDeskDocument whose cite key is citeKey)
							
							if italicReferences does not contain italicReferenceTitle then
								set italicReferences to italicReferences & italicReferenceTitle
							end if
						end if
						
						set multipleCiteKeys to multipleCiteKeys & citeKeyNumber
						set citeKeyNumberText to multipleCiteKeys
					end repeat
				end tell
				
				if (numberedReferences as boolean) is false then
					-- format author-year references
					if fieldCode contains "citep" then
						if ibidReference is true then
							set templatedText to "ibid."
							
						else
							tell application "BibDesk"
								set templatedText to (templated text document bibDeskDocumentName using text citepTemplate for multiplePublications)
							end tell
						end if
						
						set formattedReference to parenthesisOpen & referencePrefix & templatedText & referenceSuffix & parenthesisClose
						
					else if fieldCode contains "citet" then
						if ibidReference is true then
							set templatedText to "ibid."
							
						else
							tell application "BibDesk"
								set templatedText to (templated text document bibDeskDocumentName using text citetTemplate for multiplePublications)
							end tell
						end if
						
						set formattedReference to parenthesisOpen & referencePrefix & templatedText & referenceSuffix & parenthesisClose
						
					else if fieldCode contains "citealp" then
						if ibidReference is true then
							set templatedText to "ibid."
							
						else
							tell application "BibDesk"
								set templatedText to (templated text document bibDeskDocumentName using text citepTemplate for multiplePublications)
							end tell
						end if
						
						set formattedReference to referencePrefix & templatedText & referenceSuffix
						
					else if fieldCode contains "citealt" then
						if ibidReference is true then
							set templatedText to "ibid."
							
						else
							tell application "BibDesk"
								set templatedText to (templated text document bibDeskDocumentName using text citetTemplate for multiplePublications)
							end tell
						end if
						
						set formattedReference to referencePrefix & templatedText & referenceSuffix
					end if
					
					set content of result range of currentField to formattedReference
					
				else if (numberedReferences as boolean) is true then
					-- format numbered references
					set formattedReference to ""
					set AppleScript's text item delimiters to numberSeperator
					set content of result range of currentField to parenthesisOpen & citeKeyNumberText & parenthesisClose
				end if
				
				-- set superscript for current field
				if superscriptReferences is false then
					set superscript of font object of result range of currentField to false
					
				else if superscriptReferences is true then
					set superscript of font object of result range of currentField to true
				end if
				
				set show codes of currentField to false
				
				-- set italics for current field
				set italicCombined to italicLatin & italicReferences
				
				repeat with eachItalic in italicCombined
					if formattedReference contains eachItalic then
						set AppleScript's text item delimiters to ","
						
						repeat with italicFind in italicCombined
							set findObject to find object of selection
							set content of findObject to italicFind
							set content of replacement of findObject to italicFind
							set italic of font object of replacement of findObject to true
							execute find findObject replace replace all
						end repeat
						
						set AppleScript's text item delimiters to ""
					end if
				end repeat
			end if
		end repeat
		
		-- format bibliography
		set fieldCount to count of fields of formatRange
		
		repeat until fieldCount is 0
			set currentField to field fieldCount of formatRange
			set fieldCode to content of field code of currentField
			
			if field type of currentField is field addin and word 2 of fieldCode contains "bibliography" then
				set bibliographyField to currentField
				
				select bibliographyField
				
				-- compile the bibliography text and export to a temp file
				tell application "System Events"
					set tempDirectory to path to temporary items from user domain
					set bibliographyTemplateFile to POSIX file bibliographyTemplatePath as string
					set bibliographyTemplateFileName to name of file bibliographyTemplatePath
					set tempfile to (tempDirectory as string) & bibliographyTemplateFileName
				end tell
				
				tell application "BibDesk"
					if bibliographySortBy is not "" and (numberedReferences as boolean) is false then
						sort publicationList by bibliographySortBy
						set publicationList to result
					end if
					
					export bibDeskDocument using template file (bibliographyTemplateFile as string) to file tempfile for publicationList
				end tell
				
				-- insert exported text into the bibliography field
				set show codes of bibliographyField to false
				set content of result range of bibliographyField to "*"
				
				insert file at (create range wordDocument start (start of content of result range of bibliographyField) end (start of content of result range of bibliographyField)) file name tempfile
				
				tell application "System Events" to delete file tempfile
				
				-- set Word style of bibliography and force font object
				set style of result range of bibliographyField to bibliographyStyle
				set name of font object of result range of bibliographyField to name of font object of Word style bibliographyStyle of wordDocument
				
				-- remove superfluous characters
				delete (create range wordDocument start ((end of content of result range of bibliographyField) - 2) end (end of content of result range of bibliographyField))
				
				exit repeat
			end if
			
			set fieldCount to fieldCount - 1
		end repeat
	end tell
	
	-- restore Word to its original position
	tell application "Microsoft Word"
		try
			select selectionRange
			
		on error -- produces error if cursor was at the end of the document
			select (create range wordDocument start (selection end of selection) end (selection end of selection))
		end try
		
		set status bar to "Finished formatting references"
	end tell
	
	return bibliographySettings
end formatReferences

---

on unformatReferences()
	-- global setup
	tell application "Microsoft Word"
		set wordDocument to active document
		set selectionRange to create range wordDocument start (selection start of selection) end (selection end of selection)
		set status bar to "Unformatting references..."
	end tell
	
	set AppleScript's text item delimiters to ""
	
	-- unformat references
	tell application "Microsoft Word"
		activate
		
		if (selection start of selection) is equal to (selection end of selection) then
			set formatRange to wordDocument
			
			try
				display dialog "Are you sure? This will unformat the entire document!" buttons {"Cancel", "OK"} default button "OK"
				
			on error
				return
			end try
			
		else
			set formatRange to selectionRange
		end if
		
		set fieldList to fields of formatRange
		
		if (count of fieldList) is 0 then
			display dialog "There are no formatted references in this selection" buttons {"OK"} default button "OK"
			
			return
		end if
		
		repeat with currentField in reverse of fieldList
			set fieldCode to content of field code of currentField
			
			if field type of currentField is field addin and {"cite", "citep", "citet", "citealp", "citealt", "bibliography"} contains word 2 of fieldCode then
				if fieldCode contains "bibliography" then set style of result range of currentField to "Normal"
				
				set fieldCode to characters 8 thru -1 of fieldCode as string
				
				select currentField
				
				insert text "\\" & fieldCode at end of text object of selection
				
				delete currentField
			end if
		end repeat
	end tell
	
	-- restore Word to its original position
	tell application "Microsoft Word"
		try
			select selectionRange
			
		on error -- produces error if cursor was at the end of the document
			select (create range wordDocument start (selection end of selection) end (selection end of selection))
		end try
		
		set status bar to "Finished unformatting references"
	end tell
end unformatReferences

---

on fillEmptyReferences()
	-- find empty references contained by {}
	set AppleScript's text item delimiters to ""
	
	tell application "Microsoft Word"
		set wordDocument to active document
		set selectionRange to (create range wordDocument start (selection start of selection) end (selection end of selection))
		set status bar to "Filling empty references..."
		
		-- start from the beginning of the document
		select (create range wordDocument start 0 end 0)
		
		-- set up the find system in Word
		set findObject to find object of selection
		set match wildcards of findObject to true
		
		-- prevent wrapping to enable skip option
		set wrap of findObject to find stop
		
		-- find empty references and contents but not cite commands
		set content of findObject to "[!\\]]\\{*\\}"
		set theCounter to 0
		
		repeat while (execute find findObject)
			set status bar to "Searching for empty references... (" & theCounter & ")"
			set emptyReferenceText to content of (create range wordDocument start (selection start of selection) + 1 end (selection end of selection))
			set emptyReferenceContext to "\"" & characters 1 through -2 of (content of sentences of selection as string) & "\""
			
			try
				set {citepList, citetList, citealpList, citeSkip} to {"(Author, YEAR)", "(YEAR)", "(Custom...)", "Skip this instance"}
				
				choose from list {citepList, citetList, citealpList, citeSkip} with title "Fill Empty References {}" with prompt "Please select reference[s] in BibDesk for:

" & emptyReferenceContext & "
" default items citepList
				
				set theResult to result
				
				if theResult does not contain citeSkip then
					if theResult contains citepList then
						tell application "BibDesk" to set templatedText to templated text front document using text "<$publications><$#=1?>\\citep[][]{<$citeKey/><?$#?>,<$citeKey/></$#?></$publications>}" for selection of front document
						
					else if theResult contains citetList then
						tell application "BibDesk" to set templatedText to templated text front document using text "<$publications><$#=1?>\\citet[][]{<$citeKey/><?$#?>,<$citeKey/></$#?></$publications>}" for selection of front document
						
					else if theResult contains citealpList then
						set emptyReferenceContext to emptyReferenceContext & "
	
(Use 'citealp' for Author, YEAR, and 'citealt' for YEAR)"
						tell application "BibDesk" to set templatedText to templated text front document using text "<$publications><$#=1?>\\citealp[][]{<$citeKey/>}<?$#?>\\citealp[; ][]{<$citeKey/>}</$#?></$publications>" for selection of front document
						
					else if theResult is false then
						exit repeat
					end if
					
					display dialog "Please confirm the cite command for:

" & emptyReferenceContext with title "Fill Empty References {}" default answer templatedText default button "OK"
					
					set templatedText to text returned of result
					set content of (create range wordDocument start (selection start of selection) + 1 end (selection end of selection)) to (templatedText as string)
					set theCounter to theCounter + 1
				end if
				
			on error
				-- prevent "no empty references" dialog
				set theCounter to -1
				
				-- exit repeat instead of return to avoid breaking the find system in Word
				exit repeat
			end try
			
			-- continue search from end of filled reference onwards
			select (create range wordDocument start (selection end of selection) end (selection end of selection))
		end repeat
		
		try
			select selectionRange
		end try
	end tell
	
	if theCounter is 0 then
		tell application "Microsoft Word"
			display dialog "No empty references filled" buttons {"OK"} default button "OK"
		end tell
		
	else if theCounter is greater than 0 then
		try
			tell application "Microsoft Word"
				display dialog "Empty references filled, format references now?" buttons {"No", "Yes"} cancel button "No" default button "Yes"
				
				if button returned of result contains "Yes" then menuChoiceFormat()
			end tell
		end try
	end if
	
	tell application "Microsoft Word"
		set status bar to "Finished filling empty references"
	end tell
	
	return
end fillEmptyReferences

---

on toolsMenu()
	-- global setup
	tell application "Microsoft Word"
		set wordDocument to active document
		set selectionRange to create range active document start (selection start of selection) end (selection end of selection)
	end tell
	
	-- open tools menu
	repeat
		tell application "Microsoft Word"
			activate
			
			choose from list {"Convert from \\cite to \\citep", "Covert from \\citep to \\cite", "Convert p./pp. to :", "Convert : to p./pp.", "Add Leading Space", "Remove Leading Space"} with title "WordScript Tools" with prompt "You should save your work first!" default items "Convert from \\cite to \\citep" cancel button name {"Exit"} OK button name {"Choose"}
			
			set menuChoice to result
		end tell
		
		if menuChoice is false then
			exit repeat
		end if
		
		if menuChoice contains "Export to Static Group" then
			
			-- rewrite from scratch
			
		else if menuChoice contains "Convert from \\cite to \\citep" then
			set AppleScript's text item delimiters to ""
			
			tell application "Microsoft Word"
				activate
				
				set findObject to find object of text object of wordDocument
				set match wildcards of findObject to false
				
				repeat with citeCommand in {"\\cite{", "\\cite["}
					set content of findObject to citeCommand
					
					if citeCommand as string is "\\cite{" then
						set content of replacement of findObject to "\\citep{"
						execute find findObject replace replace all
						
					else if citeCommand as string is "\\cite[" then
						set content of replacement of findObject to "\\citep["
						
						execute find findObject replace replace all
					end if
				end repeat
			end tell
			
			return
			
		else if menuChoice contains "Covert from \\citep to \\cite" then
			set AppleScript's text item delimiters to ""
			
			tell application "Microsoft Word"
				activate
				
				set findObject to find object of text object of wordDocument
				set match wildcards of findObject to false
				
				repeat with citeCommand in {"\\citep{", "\\citep["}
					set content of findObject to citeCommand
					
					if citeCommand as string is "\\citep{" then
						set content of replacement of findObject to "\\cite{"
						
						execute find findObject replace replace all
						
					else if citeCommand as string is "\\citep[" then
						set content of replacement of findObject to "\\cite["
						
						execute find findObject replace replace all
					end if
				end repeat
			end tell
			
			return
			
		else if menuChoice contains "Convert p./pp. to :" then
			set AppleScript's text item delimiters to ""
			
			tell application "Microsoft Word"
				activate
				
				set findObject to find object of text object of wordDocument
				set match wildcards of findObject to false
				
				repeat with citeCommand in {"[, p.", "[, pp."}
					set content of findObject to citeCommand
					
					if citeCommand as string is "[, p." then
						set content of replacement of findObject to "[: " -- normal colon
						
						execute find findObject replace replace all
						
					else if citeCommand as string is "[, pp." then
						set content of replacement of findObject to "[꞉ " -- NBSp colon
						
						execute find findObject replace replace all
					end if
				end repeat
			end tell
			
			return
			
		else if menuChoice contains "Convert : to p./pp." then
			set AppleScript's text item delimiters to ""
			
			tell application "Microsoft Word"
				activate
				
				set findObject to find object of text object of wordDocument
				set match wildcards of findObject to false
				
				repeat with citeCommand in {"[: ", "[꞉ "} -- normal colon and modifier letter colon
					set content of findObject to citeCommand
					if citeCommand as string is "[: " then -- normal colon
						set content of replacement of findObject to "[, p."
						
						execute find findObject replace replace all
						
					else if citeCommand as string is "[꞉ " then -- modifier letter colon
						set content of replacement of findObject to "[, pp."
						
						execute find findObject replace replace all
					end if
				end repeat
			end tell
			
			return
			
		else if menuChoice contains "Add Leading Space" then
			set AppleScript's text item delimiters to ""
			
			tell application "Microsoft Word"
				activate
				
				set findObject to find object of text object of wordDocument
				set match wildcards of findObject to false
				
				repeat with citeCommand in {"\\citep", "\\citet", "\\citealp", "\\citealt"}
					set content of findObject to citeCommand
					
					if citeCommand as string is "\\citep" then
						set content of replacement of findObject to " \\citep"
						
						execute find findObject replace replace all
						
					else if citeCommand as string is "\\citet" then
						set content of replacement of findObject to " \\citet"
						
						execute find findObject replace replace all
						
					else if citeCommand as string is "\\citealp" then
						set content of replacement of findObject to " \\citealp"
						
						execute find findObject replace replace all
						
					else if citeCommand as string is "\\citealt" then
						set content of replacement of findObject to " \\citealt"
						
						execute find findObject replace replace all
					end if
				end repeat
			end tell
			
			return
			
		else if menuChoice contains "Remove Leading Space" then
			set AppleScript's text item delimiters to ""
			
			tell application "Microsoft Word"
				activate
				
				set findObject to find object of text object of wordDocument
				set match wildcards of findObject to false
				
				repeat with citeCommand in {" \\citep", " \\citet", " \\citealp", " \\citealt"}
					set content of findObject to citeCommand
					
					if citeCommand as string is " \\citep" then
						set content of replacement of findObject to "\\citep"
						
						execute find findObject replace replace all
						
					else if citeCommand as string is " \\citet" then
						set content of replacement of findObject to "\\citet"
						
						execute find findObject replace replace all
						
					else if citeCommand as string is " \\citealp" then
						set content of replacement of findObject to "\\citealp"
						
						execute find findObject replace replace all
						
					else if citeCommand as string is " \\citealt" then
						set content of replacement of findObject to "\\citealt"
						
						execute find findObject replace replace all
					end if
				end repeat
			end tell
			
			return
			
		end if
		
		-- restore Word to its original position
		tell application "Microsoft Word"
			try
				select selectionRange
				
			on error -- produces error if cursor was at the end of the document
				select (create range wordDocument start (selection end of selection) end (selection end of selection))
			end try
		end tell
	end repeat
	
	return
end toolsMenu

---

on settingsMenu(bibliographySettings)
	-- global setup
	tell application "BibDesk"
		set bibDeskDocument to front document
	end tell
	
	-- parse bibliography settings
	set AppleScript's text item delimiters to ""
	set {bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator} to bibliographySettings
	
	-- parse custom lists from the bibliography settings
	set AppleScript's text item delimiters to ","
	set italicLatin to text items of italicLatin
	set italicPublicationTypes to text items of italicPublicationTypes
	set AppleScript's text item delimiters to ""
	
	(*
	-- check to make sure that the template files exist
	try
		tell application "Finder"
			get exists of (POSIX file (bibliographyTemplatePath as string) as alias)
			get exists of (POSIX file (citepTemplatePath as string) as alias)
			get exists of (POSIX file (citetTemplatePath as string) as alias)
		end tell
	on error
		tell application "Microsoft Word"
			display dialog "Template files not found, please delete the bibliography and try again" buttons {"OK"} default button {"OK"}
		end tell
		return
	end try
	*)
	
	repeat
		-- parse menu choices
		set bibDeskDocumentNameChoice to "BibDesk Document: 		" & bibDeskDocumentName
		set AppleScript's text item delimiters to "/"
		set bibliographyTemplatePathChoice to "Bibliography Template: 	" & (text item -1 of bibliographyTemplatePath)
		set citepTemplatePathChoice to "Author-Year Template: 		" & (text item -1 of citepTemplatePath)
		set citetTemplatePathChoice to "Year Only Template: 		" & (text item -1 of citetTemplatePath)
		set AppleScript's text item delimiters to ""
		set parenthesisOpenChoice to "Parenthesis Open: 		" & parenthesisOpen
		set parenthesisCloseChoice to "Parenthesis Close: 		" & parenthesisClose
		set AppleScript's text item delimiters to ","
		set italicLatinChoice to "Italic Latin: 				" & italicLatin
		set italicPublicationTypesChoice to "Italic Publication Types: 	" & italicPublicationTypes
		set AppleScript's text item delimiters to ""
		set bibliographySortByChoice to "Bibliography Sort By: 		" & bibliographySortBy
		set bibliographyStyleChoice to "Bibliography Style: 		" & bibliographyStyle
		set superscriptReferencesChoice to "Superscript references: 	" & superscriptReferences
		set numberedReferencesChoice to "Numbered references: 		" & numberedReferences
		set numberSeperatorChoice to "Number Seperator: 		" & numberSeperator
		
		set builtInPreset1Choice to "Built-In Quick Settings: 	Harvard Referencing"
		set builtInPreset2Choice to "Built-In Quick Settings: 	Numbered Endnotes"
		
		tell application "Microsoft Word"
			choose from list {bibDeskDocumentNameChoice, bibliographyTemplatePathChoice, citepTemplatePathChoice, citetTemplatePathChoice, parenthesisOpenChoice, parenthesisCloseChoice, italicLatinChoice, italicPublicationTypesChoice, bibliographySortByChoice, bibliographyStyleChoice, superscriptReferencesChoice, numberedReferencesChoice, numberSeperatorChoice, " ", builtInPreset1Choice, builtInPreset2Choice} with title "WordScript Settings" with prompt "Please choose which setting to edit" cancel button name {"Exit"} OK button name {"Choose"}
			
			set menuChoice to result
		end tell
		
		-- save and exit
		if menuChoice is false then
			exit repeat
		end if
		
		-- Document Settings
		if menuChoice contains bibDeskDocumentNameChoice then
			try
				tell application "Microsoft Word"
					display dialog "Open desired library in BibDesk then click OK"
				end tell
				
				tell application "BibDesk"
					set bibDeskDocumentName to name of bibDeskDocument
				end tell
				
			on error -- do nothing
			end try
			
		else if menuChoice contains bibliographyTemplatePathChoice then
			tell application "Microsoft Word"
				set bibliographyTemplatePath to POSIX path of (choose file with prompt "Please choose bibliography template" of type {"public.rtf", "com.microsoft.word.doc"})
			end tell
			
		else if menuChoice contains citepTemplatePathChoice then
			tell application "Microsoft Word"
				set citepTemplatePath to POSIX path of (choose file with prompt "Please choose author-year template (citep)" of type {"public.plain-text"})
			end tell
			
		else if menuChoice contains citetTemplatePathChoice then
			tell application "Microsoft Word"
				set citetTemplatePath to POSIX path of (choose file with prompt "Please choose year only template (citet)" of type {"public.plain-text"})
			end tell
			
		else if menuChoice contains parenthesisOpenChoice then
			tell application "Microsoft Word"
				display dialog "Set Parenthesis Open
Example: ( or [ or <" buttons {"OK"} default button "OK" default answer parenthesisOpen

				set parenthesisOpen to text returned of result
			end tell
			
		else if menuChoice contains parenthesisCloseChoice then
			tell application "Microsoft Word"
				display dialog "Set Parenthesis Close
Example: ) or ] or =" buttons {"OK"} default button "OK" default answer parenthesisClose

				set parenthesisClose to text returned of result
			end tell
			
		else if menuChoice contains italicLatinChoice then
			set AppleScript's text item delimiters to ","
			
			tell application "Microsoft Word"
				display dialog "Set phrases to be italisised, separated by commas
Example: et al.,ibid.,cf.,inter alia,circa" buttons {"OK"} default button "OK" default answer italicLatin as string

				set italicLatin to text returned of result
				set italicLatin to (text items of italicLatin) as list
			end tell
			
			set AppleScript's text item delimiters to ""
			
		else if menuChoice contains italicPublicationTypesChoice then
			set AppleScript's text item delimiters to ","
			
			tell application "Microsoft Word"
				display dialog "Set reference types to be italisised, separated by commas
Example: film,broadcast" buttons {"OK"} default button "OK" default answer italicPublicationTypes as string

				set italicPublicationTypes to text returned of result
				set italicPublicationTypes to (text items of italicPublicationTypes) as list
			end tell
			
			set AppleScript's text item delimiters to ""
			
		else if menuChoice contains bibliographySortByChoice then
			tell application "Microsoft Word"
				display dialog "Set Sort Bibliography By BibDesk Field
Example: Author or Cite Key
(Leave blank to sort in order of appearance)" buttons {"OK"} default button "OK" default answer bibliographySortBy

				set bibliographySortBy to text returned of result
			end tell
			
		else if menuChoice contains bibliographyStyleChoice then
			tell application "Microsoft Word"
				display dialog "Set bibliography style in Microsoft Word
Example: Bibliography or List Number or Normal
(Leave blank to use formatting from template file)" buttons {"OK"} default button "OK" default answer bibliographyStyle
				set bibliographyStyle to text returned of result
			end tell
			
		else if menuChoice contains superscriptReferencesChoice then
			tell application "Microsoft Word"
				display dialog "Use superscript references?" buttons {"No", "Yes"} default button superscriptReferences
				
				set superscriptReferences to text items of button returned of result
			end tell
			
		else if menuChoice contains numberedReferencesChoice then
			tell application "Microsoft Word"
				display dialog "Use numbered references instead of Harvard style?
(Overrides reference templates and sorting)" buttons {"No", "Yes"} default button numberedReferences

				set numberedReferences to text items of button returned of result
			end tell
			
		else if menuChoice contains numberSeperatorChoice then
			tell application "Microsoft Word"
				display dialog "Set seperator for numbered references
(Only applies to multiple numbers in a single field)
Example: , or ; or space" buttons {"OK"} default button "OK" default answer numberSeperator

				set numberSeperator to text returned of result
			end tell
			
		else if menuChoice contains builtInPreset1Choice then
			set parenthesisOpen to "("
			set parenthesisClose to ")"
			set bibliographySortBy to "Cite Key"
			set bibliographyStyle to "Bibliography"
			set italicLatin to {"et al.", "ibid.", "cf.", "inter alia", "circa"}
			set italicPublicationTypes to {"film", "broadcast"}
			set superscriptReferences to "No"
			set numberedReferences to "No"
			set numberSeperator to ""
			
		else if menuChoice contains builtInPreset2Choice then
			set parenthesisOpen to ""
			set parenthesisClose to ""
			set bibliographySortBy to ""
			set bibliographyStyle to "List Number"
			set italicLatin to {}
			set italicPublicationTypes to {}
			set superscriptReferences to "Yes"
			set numberedReferences to "Yes"
			set numberSeperator to ","
		end if
	end repeat
	
	return {bibDeskDocumentName, bibliographyTemplatePath, citepTemplatePath, citetTemplatePath, parenthesisOpen, parenthesisClose, bibliographySortBy, bibliographyStyle, italicLatin, italicPublicationTypes, superscriptReferences, numberedReferences, numberSeperator}
end settingsMenu
