* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
WHAT IS ZOTERO TOOLS AND WHAT CAN IT DO FOR YOU?
	
	Zotero Tools involve a bunch of functions I'm missing in Zotero. I've found some hints and VBA scripts in the internet to solve my problems. They worked more or less reliable and I've not been satisfied in every case. So I've decided to bundle some of these ideas and complete them to one tool. And these are Zotero Tools!
	
	The functions:
		Adjust punctuation surrounding citation groups
			This function corrects spaces and punctuation before and after the citations in your text. You can switch between:
				This is a phenomenal, (1) and unbelivable sentence. (2) This is not.
				This is a phenomenal,(1) and unbelivable sentence.(2) This is not.
				This is a phenomenal(1), and unbelivable sentence(2). This is not.
			or
				This is a phenomenal, (Shakespeare 1565) and unbelivable sentence. (Raleigh 1554) This is not.
				This is a phenomenal,(Shakespeare 1565) and unbelivable sentence.(Raleigh 1554) This is not.
				This is a phenomenal (Shakespeare 1565), and unbelivable sentence (Raleigh 1554). This is not.

		Join citation groups
			This function joins citation groups, inside the whole document or in a selected range of the document:
				The first sentence.[1], [2], [3] The next sentence.
			becomes
				The first sentence.[1-3] The next sentence.

		Resolve unreachable citation groups
			This function resolves copy-pasted Zotero citations to clear readable text in parts of the document Zotero doesn't deal with, i.e.
				comments
				headers
				footers

		Set internal linking between citations and references
			This function sets internal links between citations and bibliography (i.e. hyperlinks inside the document). These links can be:
				unidirectional, i.e. the citation inside the text is hyperlinked to its reference in the bibliography
			or
				bidirectional, i.e. the reference in the bibliography is also hyperlinked to all its citations in the text. This is very helpful during the writing process of a publication.
			This function also involves an undo function.

		Set web links in references
			This function sets hyperlinks on URIs (uniform resource identifiers) in Zotero-generated bibliographies. It works for the following URI formats:
				http or https addresses (the scheme must be involved, e.g. 'http://www.example.com' but not 'www.example.com')
				doi addresses (the scheme must be involved, e.g. 'doi:10.1234/5678' or 'doi: 10.1234/5678')
				short doi addresses (the scheme must be involved, e.g. 'short-doi:abcde' or 'short-doi: abcde' or 'shortdoi:abcde' or 'shortdoi: abcde')

	The technique:
		Zotero Tools are programmed as Visual Basic for Applications (VBA) macro. They can be configured via an XML file. By this they can be addapted to many (perhaps any) numeric citation style. Some of them do also work for author-year styles.
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
DISCLAMER AND COPYRIGHT

	Zotero Tools.
		Procedures for editing Word documents with Zotero citations.
		Copyright © 2019, Olaf Ahrens (user oahrens at zotero.org and github.com). All rights reserved.

	This software is under Revised ('New') BSD license:
	Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
		*	Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
		*	Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
		*	Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
		
		THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !
ATTENTION

	PLEASE, ALWAYS KEEP IN MIND:
		All procedures will only run on Word for Windows, not Word for Mac.
		All procedures only run on documents with Zotero citations inserted as Word fields, not as bookmarks. Take a look at 'Document preferences' on the 'Zotero' tab in Word!
		Check whether the chosen procedure will run on your citation style (numeric or author-year). This information can be found in the procedure descriptions shown when running the macro.
		Adapt the settings in 'ZtConfig.xml' file before running any of the procedures.
		Check whether the procedure did what you expected.
		Check whether Zotero is still able to work with your document, i.e. inserting, deleting, or changing citations, and refreshing bibliography.
		Check whether actions of the 'Set linking between citations and references' procedure can be redone by 'Remove linking between citations and references'.
! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !


* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
INSTALLATION

	1.	Import all files exept ZtConfig.xml and ZtReadMe.txt into Normal.dotm:
		a.	Open VBA editor by pressing Alt-F11 in Word.
		b.	Open a Windows explorer window and navigate to the folder where you have saved the Zotero Tools files.
		c.	Pull all files exept ZtConfig.xml and ZtReadMe.txt into the VBA editor window and let them fall on 'Normal' in the project explorer lefthand side in the editor.
	2.	Save the modification of Normal.dotm by pressing Ctrl-S inside the editor.
	3.	Copy the following files into the Normal.dotm directory: ZtConfig.xml and this file (ZtReadMe.txt).
		a.	The directory of the Normal.dotm file should be:
				for Windows 7 and following OS versions: 	C:\Users\[username]\AppData\Roaming\Microsoft\Templates
				for Windows XP:								C:\Documents and Settings\[username]\Application Data\Microsoft\Templates
		b.	If you can't find the AppData directory it may be hidden. Have a look at https://www.wordfast.net/wiki/How_to_make_hidden_folders_visible_in_Windows.
	4.	Add a macro button to the Word ribbon (this must be done in Word's main window, not in the editor):
		a.	Click File > Options > Customize Ribbon.
		b.	Under 'Choose commands from', click 'Macros'.
		c.	Click on 'Normal.ZtStart.Start'.
		d.	Under 'Customize the ribbon', click 'New Tab'.
		e.	Then click 'Rename' and type a name for your tab, for example 'Zotero Tools'.
		f.	Rename the automatically inserted new group: select it, click 'Rename' and type a name for your group, for example also 'Zotero Tools'.
		g.	Click Add.
		h.	Select the just inserted macro and click 'Rename' to choose an image for the macro and type the name you want, for example 'Run'.
		i.	Click 'OK' twice.
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
MICROSOFT WORD SETTINGS

	The following libraries must be referenced (inside the VBA editor: Tools -> References):
		Visual Basic For Applications (set automatically)
		Microsoft Forms 2.0 Object Library (set automatically)
		Microsoft Office nn.n Object Library (nn.n = your Office version)
		Microsoft Word nn.n Object Library (nn.n = your Word version)
		Microsoft Scripting Runtime
		Microsoft VBScript Regular Expressions 5.5
		Microsoft XML, v6.0

	The following setting should be done (in VBA-editor: Tools -> Options -> tab General):
		switch 'Error Trapping' to 'Break in Class Module'
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
STARTING ZOTERO TOOLS

	You have to modify the settings of the macro in 'ZtConfig.xml' file acording to the properties of your document and citation style.
	After this start the macro and switch on debugging mode before running any procedure. In debugging mode the different procedures show how your settings are applied to the document and you get informations for correcting the settings where appropriate.
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
ZTCONFIG.XML: XML

	Every character of an XML element or attribute stands for itself. This includes spaces, tabs, and line breaks, also leading and trailing ones.

	But the following characters must or should be escaped:
		&		escaped by		&amp;		(must)
		<		escaped by		&lt;		(must)
		>		escaped by		&gt;		(should; must in some contexts)
		"		escaped by		&quot;		(should; must in some contexts)
		'		escaped by		&apos;		(should; must in some contexts)

	Unicode characters can be written in one of these two formats:
		&#n;	where n is the decimal number of the character, with or without leading zeros
		&#xh;	where h is the hexadecimal number of the character, with or without leading zeros

	Control characters:
		XML version 1.0 doesn't allow the following unicode characters (and XML version 1.1 isn't supported):
			\u0000 - \u0008, \u000B, \u000C, \u000E - \u001F, \u007F - \u0084, \u0086 - \u009F, \uD800 - \uDFFF
		The combination of VBA and MSXml doesn't read the following unicode character correct:
			\u000D

		Instead of these use one of the following for any element of type string:
			{Chr(n)}	where n is the decimal number of the character, with or without leading zeros
			{Chr(xh)}	where h is the hexadecimal number of the character, with or without leading zeros
			e.g.	instead of	<punctuationBreakField>.,:;?!&#10;&#13;&#19;&#21;</punctuationBreakField>
					use			<punctuationBreakField>.,:;?!&#10;{Chr(13)}{Chr(19)}{Chr(21)}</punctuationBreakField>
			{Chr(n)} and {Chr(xh)} can be escaped by a single backslash ('\'), the backslash itself can be escaped by a double backslash.

	You should use an unicode-enabled XML editor for changing the settings in ZtConfig.xml:
		Notepad++:			https://notepad-plus-plus.org
							also install the extension 'XML Tools'
		XML Notepad 2007:	https://www.microsoft.com/en-us/download/details.aspx?id=7973.

	Don't use Windows built-in editors Notepad or WordPad: depending on your Windows version they may not save the file with the correct coding (UTF-8). They are also missing any syntax highlighting and syntax verification.
	Never use Word for editing ZtConfig.xml.
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
ZTCONFIG.XML: REGEXES

	In 'ZtConfig.xml' all string elements named as '...RegString' must be written in regex syntax.
	Therefore here is a short and very, very basic regex introduction:

		Most of regexes used in this macro are case insensitive, but not all.

		Escaping characters with backslash:
			The following characters do have special meanings and have to be escaped with a backslash if they are used in their literal sence outside self defined character classes:
				\^$.|?*+()[
			The following characters do also have special meanings and may not be escaped outside self defined character classes mandatory, but for clarity:
				{}]
				i.e.	'a\[]' has the same meaning as 'a\[\]'
						'a]' is a well formed regex Pattern but 'a[' isn't
						'a{' and 'a\{' have the same meaning and both are well formed regex patterns
						'a}' and 'a\}' have the same meaning and both are well formed regex patterns
			The following characters must be escaped inside self defined character classes if they are used in their literal sence:
				\-]^
			The following characters may not be escaped inside self defined character classes mandatory, but for clarity:
				$.|?*+()[{}
			All other characters stand for themself, including ' ' (space), " (double quote), ' (single quote), < (less than), > (greater than); they must not be escaped when they're used in their literal sence, but some of them get a special meaning when being escaped.

		Special signs and some predefined character classes:
			\		1. Escape sign, it can be escaped by doubling, i.e. '\\' hits '\'.
					2. Definition sign for special characters, e.g. '\t' for tabulator.
					3. Definition sign for predefined character groups, e.g. '\d' for a numeric character.
			?		1. Quantifier, see below.
					2. Controls grouping mechanism in many ways, an example see below.
			^		Start of the tested string (doesn't capture anything).
			$		1. End of the tested string (doesn't capture anything).
					2. Reference to a group (see below) in '$n', where n is a number from 1 to 9.
			()		For building groups, see below.
			[]		For building character classes, see below.
			{}		For quantification, see below.
			+		For quantification, see below.
			*		For quantification, see below.
			|		Alternation sign, see below.
			-		Range inside a self defined character class, see below; outside groups it has its literal meaning.
			\t		Tabulator.
			\d		Any single numeric character, i.e. 0-9, but not hexadecimal numeric characters A-F or a-f.
			\unnnn	Unicode sign, where 'nnnn' stand for the hexadecimal number of the sign; we always have to use all 4 digits, e.g. '\u0020' does mean a space sign.
			.		Any single character including unicode characters, but without line break.

		Some predefined character classes we should not use, because VBScript regex doesn't support unicode in predefined character classes (see below):
			\w		Should hit any 'word character'; but in VBScript regex e.g. 'é' isn't hit!
			\W		Should hit any 'non-word character' (negation of '\w'); but in VBScript regex e.g. 'é' is hit too!
			\s		Should hit any white space; but in VBScript regex e.g. ChrW$(8203) = no-break space isn't hit!
			\S		Should hit any non-white space (negation of '\s'); but in VBScript regex e.g. ChrW$(8203) = no-break space is hit too!

		Self defined character classes:
			[]
			i.e.	'h[au]t' hits 'hat' and 'hut'.
					'h[^ac]t' hits 'het', but not 'hat' or 'hct'.
					'h[a-z]t' hits 'hat' and 'hut', but not 'ht' or 'hAt'.
					'h[a-zA-Z]t' hits 'hat', 'hut', and 'hAt', but not 'h2t'.
					'[0-9]' has the same meaning as '\d'.

		Grouping: any number of characters or groups can be grouped by
			()
			E.g.	'(a\d )' is a group of one 'a', one number sign and one space.
			If you're using a self defined group in 'ZtConfigUser' you have to avoid 'capturing' of the group; otherwise your group is counted by the regex and the logic of the macro will be interfered; you can avoid this by introducing the group by:
				?:
				e.g.	'(?:a\d )' does the same as '(a\d )' but the first group isn't count (captured) while the second group is.

		Alternations: characters, words, character classes, and groups can be written as a sequence of alternatives (usually in a group) by
			|
			e.g.	'This (fish|cat|dog|[XY]) is called Wanda'
						hits		'This fish is called Wanda', 'This dog is called Wanda', 'This X is called Wanda'
						hits not	'This horse is called Wanda', 'This XY is called Wanda'

		Repetitions/quantifiers: the quantity of any single character, character class, or character group can be defined by:
			*		Any count: zero, one, or several; i.e. 'phone \d*' 
						hits		'phone 0123' and 'phone 1' and 'phone ' in 'phone call'
						hits		nothing in 'phonetic'
			+		At least one; e.g. 'le[a-z]+st' 
						hits		'least' and 'lebest'
						hits not	'lest'.
			?		Zero or one; e.g. 'lea?st' 
						hits		'lest' and 'least'
						hits not	'leaast'
			{2,4}	2, 3, or 4; i.e. 'a{0,1}' does mean the same as 'a?'.
			{,4}	Maximum 4; i.e. 'a{,1}' does mean the same as 'a?' and 'a{0,1}'.
			{2,}	Minimum 2; i.e. 'a{0,}' does mean the same as 'a*', and 'a{1,}' does mean the same as 'a+'.
			{2}		Exact 2; i.e. 'a{2}' does mean the same as 'aa'.

	Excellent introduction for regex beginners and source for regex professionals: https://www.regular-expressions.info/tutorial.html.

	Some special hints for using VB-Script regexes (from https://www.regular-expressions.info/vbscript.html):

		VBScript implements Perl-style regular expressions. However, it lacks quite a number of advanced features available in Perl and other modern regular expression flavors:
			No \A or \Z anchors to match the start or end of the string. Use a caret or dollar instead.
			Lookbehind is not supported at all. Lookahead is fully supported.
			No atomic grouping or possessive quantifiers.
			No Unicode support, except for matching single characters with \uFFFF.
			No named capturing groups. Use numbered capturing groups instead.
			No mode modifiers to set matching Options within the regular expression.
			No conditionals.
			No regular expression comments. Describe your regular expression with VBScript apostrophe comments instead, outside the regular expression string.

	A simple regex tester specialized in VBScript regex):
		https://www.regular-expressions.info/vbscriptexample.html (IE only; press F12 and switch document mode to 9 before using
	Powerful regex testers not specialized in VBScript regex:
		http://regexstorm.net/tester
		http://regexhero.net/tester/ (IE only)
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
CODING OF SOME SPECIAL CHARACTERS

	Dashes:

		VBA (decimal unicode)		regex (hexadecimal unicode)		XML (decimal unicode)		XML (hexadecimal unicode)		description
		---------------------------------------------------------------------------------------------------------------------------------------------------------
		ChrW$(45)					\u002D							&#45;						&#x2D;							(normal, typewriter) hyphen-minus
		Chr$(150)					./.								./.							./.								en-Dash in Windows 1252 code page
		ChrW$(8208)					\u2010							&#8208;						&#x2010;						(true, typographic) hyphen
		ChrW$(8209)					\u2011							&#8209;						&#x2011;						non-breaking hyphen
		ChrW$(8210)					\u2012							&#8210;						&#x2012;						figure dash
		ChrW$(8211)					\u2013							&#8211;						&#x2013;						en-dash
		ChrW$(8212)					\u2014							&#8212;						&#x2014;						em-dash

	Spaces:

		VBA (decimal unicode)		regex (hexadecimal unicode)		XML (decimal unicode)		XML (hexadecimal unicode)		description
		---------------------------------------------------------------------------------------------------------------------------------------------------------
		ChrW$(32)					\u0020							&#32;						&#x20;							space
		ChrW$(160)					\u00A0							&#160;						&#xA0;							no-break space
		ChrW$(8203)					\u200B							&#8203;						&#x200B;						zero width space
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


