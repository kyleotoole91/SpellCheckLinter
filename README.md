# SpellCheckLinter
 Command line spell check linter designed specifically for .pas code files.

Startup params:
LanguageFilename=1; //.dic file
SourceFilename=2; //source path or filename
ExtFilter=3; //eg *.pas
Resursive=4; //when 0, resusive scan is disabled
Halt=5; //when 0, halting at the end of run is disabled
cQuoteSym=6; //defaulted to a single quote ' for .pas files. Strings between these quote symbols are checked

Defaults to the working directory, scanning resursively for .pas files.
By default it looks for a en_US.dic file in the working directory.
The default source path is the working directory.
Resursive scan can be disabled by supplying 0 as a param.
Halting at the end of the run can be disable by supplying 0 as a param.

Hidden feature: submit r or R to rerun without needing to restart the application. Usefull when launching from Windows Explorer

Ingnore files may be placed in the working directory
IgnoreWords.dic (list of words to ignore, extends the language file)
IgnoreFiles.dic (list of filepaths to ignore)
IgnoreLines.dic (ignore lines of code, not by line number, but by value by pasting the entire line in here)

Link to zip containing exe, source and language files
https://drive.google.com/file/d/1rzD_JmpPTy9X6oh-4CjanN559ibzDP0n/view?usp=sharing
