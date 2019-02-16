# Zotero-Tools
Tools for Zotero (citation and reference management software)

<p>Zotero Tools involve a bunch of functions I'm missing in Zotero. I've found some hints and VBA scripts in the internet to solve my problems. They worked more or less reliable and I've not been satisfied in every case. So I've desided to bundle some of these ideas and complete them to one tool. And these are Zotero Tools!</p>
<p><h3>The functions:</h3></p>
	<p><h4>Adjust punctuation surrounding citation groups</h4></p>
		<p>This function corrects spaces and punctuation before and after the citations in your text. You can switch between:</p>
			<ul><li>This is a phenomenal, <sup>(1)</sup> and unbelivable sentence. <sup>(2)</sup> This is not.</li>
	<li>This is a phenomenal,<sup>(1)</sup> and unbelivable sentence.<sup>(2)</sup> This is not.</li>
				<li>This is a phenomenal<sup>(1)</sup>, and unbelivable sentence<sup>(2)</sup>. This is not.</li></ul>
		<p>or</p>
			<ul><li>This is a phenomenal, (1) and unbelivable sentence. (2) This is not.</li>
                        	<li>This is a phenomenal (1), and unbelivable sentence (2). This is not.</li></ul>
		<p>or</p>
			<ul><li>This is a phenomenal, (Shakespeare 1565) and unbelivable sentence. (Raleigh 1554) This is not.</li>
				<li>This is a phenomenal,(Shakespeare 1565) and unbelivable sentence.(Raleigh 1554) This is not.</li>
				<li>This is a phenomenal (Shakespeare 1565), and unbelivable sentence (Raleigh 1554). This is not.</li></ul>
	<p><h4>Join citation groups</h4></p>
		<p>This function joins citation groups, inside the whole document or in a selected range of the document:<br>
			<ul><li>The first sentence.[1], [2], [3] The next sentence.</li></ul>
		<p>becomes</p>
			<ul><li>The first sentence.[1-3] The next sentence.</li></ul>
	<p><h4>Resolve unreachable citation groups</h4></p>
		<p>This function resolves copy-pasted Zotero citations to clear readable text in parts of the document Zotero doesn't deal with, i.e.</p>
			<ul><li>comments</li>
				<li>headers</li>
				<li>footers</li></ul>
	<p><h4>Set internal linking between citations and references</h4></p>
		<p>This function sets internal links between citations and bibliography (i.e. hyperlinks inside the document). These links can be:</p>
			<ul><li>unidirectional, i.e. the citation inside the text is hyperlinked to its reference in the bibliography</li></ul>
		<p>or</p>
			<ul><li>bidirectional, i.e. the reference in the bibliography is also hyperlinked to all its citations in the text. This is very helpful during the writing process of a publication.</li></ul>
		<p>This function also involves an undo function.</p>
	<p><h4>Set web links in references</h4></p>
		<p>This function sets hyperlinks on web addressed in Zotero-generated bibliographies. It works for following address formats:</p>
			<ul><li>http/https links</li>
				<li>doi links</li>
				<li>short doi links</li></ul>
<p><h3>The technique:</h3></p>
<p>Zotero Tools are programmed as Visual Basic for Applications (VBA) macro. They can be configured via an XML file. By this they can be addapted to many (perhaps any) numeric citation style. Some of them also work for author-year styles.</p>
<p><h3>Installation:</h3></p>
	<ol type="1"><li>Import all files exept <i>ZtConfig.xml</i> and <i>ZtReadMe.txt</i> into <i>Normal.dotm</i>:</li>
	<ol type="i"><li>Open VBA editor by pressing Alt-F11 in Word.</li>
		<li>Open a Windows explorer window and navigate to the folder where you have saved the Zotero Tools files.</li>
		<li>Pull all files exept <i>ZtConfig.xml</i> and <i>ZtReadMe.txt</i> into the VBA editor window and let them fall on 'Normal' in the project explorer lefthand side in the editor.</li></ol>
	<li>Save the modification of <i>Normal.dotm</i> by pressing Ctrl-S inside the editor.</li>
	<li>Add a macro button to the Word ribbon (this must be done in Word's main window, not in the editor):</li>
	<ol type="i"><li>Click File -&gt; Options -&gt; Customize Ribbon.</li>
		<li>Under 'Choose commands from', click 'Macros'.</li>
		<li>Click on 'Normal.ZtStart.Start'.</li>
		<li>Under 'Customize the ribbon', click 'New Tab'.</li>
		<li>Then click 'Rename' and type a name for your tab, for example 'Zotero Tools'.</li>
		<li>Rename the automatically inserted new group: select it, click 'Rename' and type a name for your group, for example also 'Zotero Tools'.</li>
		<li>Click Add.</li>
		<li>Select the just inserted macro and click 'Rename' to choose an image for the macro and type the name you want, for example 'Run'.</li>
		<li>Click 'OK' twice.</li></ol></ol>
<p><b>More informations you can find in the <i>ZtReadMe.txt</i> file.</b>
