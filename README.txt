'WinGrep version 1.2
'
'For all those UNIX people out there, "grep" should need little
'introduction.  It is simply a powerful tool for extracting lines
'from files.  This is the GUI version for Windows.
'
'To those of you who don't know UNIX, "grep" looks for lines matching a
'regular expression in a file, and if it finds a matching line, then
'it outputs it to a new file.  For an explanation of "regular
'expressions " please examine the VB help on the " Like " operator."
'
'For example if I have a file containing the lines
'
'Good Morning
'and how
'are
'you?
'Good Afternoon
'and Goodbye!
'
'And put the following regular expression into WinGrep and run it
'"Good"
'then the following lines will be added into the output file:
'
'Good Morning
'Good Afternoon
'and Goodbye!
'
'Up to 10 Regular expressions may be specified in WinGrep by putting
'semicolons between the Regular expressions.
'For example, if I was looking for either the words "Good" or "you" then
'my regular expression would be: "Good;you"
'and my output would be:
'
'Good Morning
'you?
'Good Afternoon
'and Goodbye!
'
'Characters in pattern Matches in string:
'? Any single character.
'* Zero or more characters.
'# Any single digit (0-9).
'[charlist] Any single character in charlist.
'[!charlist] Any single character not in charlist.
'
'Please note that the "Exact Matches only" option takes the regular
'expression and puts a * at the beginning and end of the regular
'expression.
'
'
'I know this help file is short - if you need any more information you
'can contact me at bigcalm@hotmail.com
'
'Bug fixes, Improvements, and Suggestions are most welcome too.
'
'
' New functionality for Version 1.1
' 1) Pre-Post line extraction
' 2) Counting facility
' 3) Find-Files is now finished.
' Numerous minor bug-fixes.

' New functionality for Version 1.2
' 1) UNIX Command Line style text box for those moving to WinGrep from Unix or Dos Grep
' 2) Load and Save searches
' 3) "HOT" (quick search) buttons for those regularly searches (registry keys)
' 4) Improved file finding dialog on main form.
' 5) DisplayFilenames option changed + Seperator included.
' 6) Can be run from the command line and explorer.  i.e. It processes it's arguments, and now runs using a Sub Main() procedure.
' 7) Post-extraction works correctly now.
' 8) I've put some comments into the Grep functions, 'cos they were getting horribly complex.
' 9) Silent modes for command line running

' Still outstanding stuff:
' Need to Shell Associate automatically. (Yuk, registry hacking?).
' I have supplied a .reg file which will need to be edited to work correctly.
