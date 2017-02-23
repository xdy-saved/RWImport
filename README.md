# RWImport

This is a Powershell script to import spells into Realm Works from a CSV file.
The CSV file can be downloaded here: http://www.pathfindercommunity.net/home/databases/spells

I have sanitized my version by correcting errors and making certain elements consistent. I have also modified my Pathfinder game structure by adding tags from CSV.

I wanted to upload this with all the working parts, but I can't personally speak to the legality of that, so I'm uploading this as proof-of-concept code for those who asked to see how I was doing what I was doing.

This isn't polished. This isn't complete. My testing has been light and limited in scope. I offer no guarantees or warranties of anything. If you do use this code, the responsibility is yours and yours alone.

Make sure you backup your database.

Make sure that, prior to importing anything, you copy the realm that you will be importing into.

Before you start, you will have to edit the script. Preferably, edit it with the Powershell ISE or with Notepad++ or some other editor that highlights Powershell syntax.

Scroll down to near the bottom of the script where you will see a block of text as such:
<pre>
#############################################################################
#############   E D I T   Y O U R   F I L N A M E S   H E R E   #############
#############################################################################
####
&lt;##&gt; $Structure_File = "Pathfinder_Structure_Augmented.rwexport"
&lt;##&gt; $Spell_Spreadsheet = "spell_full - Updated 29Jan2017 - sanitized.csv"
&lt;##&gt; $New_XML_File = "Pathfinder_Spells.rwexport"
####
#############################################################################
</pre>
In this block of text, edit the filenames as you please.
You will have to export your own structure file, and download your own spreadsheet.

Known issues:
In the spreadsheet that you will download, Detect Evil has an error in the HTML code. Detect Evil is in row 136, and the error is in column S. Scroll through to the very end of the text in that cell, and you will see it end with the following: &lt;/td&gt;&lt;/tr&gt;&lt;/tfoot&gt;&lt;/table&gt;&gt;&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;&lt;/p&gt;

Change that to: &lt;/td&gt;&lt;/tr&gt;&lt;/tfoot&gt;&lt;/table&gt;&lt;/p&gt;

Even when you do this, there will be a set of table headers that are offset one cell to the right. You can edit this after the import.

Also, there is a bug with assigning tag values to snippets, so right now, I'm putting everything in the annotation field. Ignore the tags.

I have included code to assign tag values, but I've commented it out for now as it doesn't really do what I want it to do until LWD fixes the bug. And I can't really test it properly until then either.

I think that's it, though it's possible I've forgotten some things.
