# SpellCheckLinter
Command line spell check linter designed specifically for .pas code files.

Can be launched from explorer, it will halt at the end by default to show the results.
Can be also launched from CMD where optional params can be set.
Text within single quotes will be checked against the dictionary file.
Words must be at least 3 characters in length in order to get checked.
PascalCase and camelCase text will get split into seperate words.
The first word in the camel case text must be at least 4 chars in length to get checked.

Startup parameters (optional):
1) Language file (%s)
2) Source path or full filename (.\)
3) File extension mask (*.pas)
4) Scan folders recursively (1)
5) Add to ignore prompt (1)
6) Path for the ignore files (.\)

Ignore files, must be in the same dir as exe, or you change the IgnoreFilePath classs property:
The IgnoreWords file will ignore the word if the word is in this file.
The IgnorePaths file will ignore files if the text in the file is contained in the path.
The IgnoreCode file will ignore the line if the text before the quote is in this file.
The IgnoreLines file will ignore lines if the text in the file is equal to the line.
The IgnoreContains file will ignore lines if the text in the file is contained in the the line.

Use Hunspell to inflate the .dic files

Link to zip containing exe, source and language files
https://drive.google.com/file/d/1PZOK4FGcbcqM5_-ZOOvg9Yytile9b0uP/view?usp=sharing
