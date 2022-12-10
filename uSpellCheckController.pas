unit uSpellCheckController;

interface

const
  cLanguageFilename=1; //.dic file
  cSourceFilename=2; //source path or filename
  cIgnoreFilepath=3; //defaulted to working dir
  cExtFilter=4; //eg *.pas
  cResursive=5; //when 0, resusive scan is disabled
  cHalt=6; //when 0 the program will stop when finished
  cProvideSuggestions=7; //toggle suggestions (these take time to generate)

type
  TErrorCodes = (ecExitCodeOk=0, ecExitCodeErrors=1, ecExitCodeException=2);

  function RunSpellCheck: integer;

implementation

uses
  System.SysUtils, DateUtils, uSpellCheckLinter, uConstants;

  function RunSpellCheck: integer;
  var
    input: string;
    spellCheck: TSpellCheckLinter;
    procedure NoErrorsMsg;
    begin
      if ParamStr(cHalt) <> '0' then begin
        Writeln('Press Enter to close or R to restart');
        ReadLn(input);
      end;
    end;
    procedure ShowManual;
    begin
      Writeln('Welcome to SpellCheckLinter, specifically designed for Delphi (.pas/.dfm) files. ');
      Writeln('Can be launched from explorer, it will halt at the end by default to show the results. ');
      Writeln('Can be launched from the command line where optional parameters can be set. ');
      Writeln('Only text within single quotes will be checked against the dictionary file.');
      Writeln('PascalCase and camelCase text will get split into separate words.');
      Writeln(Format('Only words %d characters in length or greater will be checked. %d for the first camelCase/PascalCase word.', [cMinCheckLength, cMinCheckLength+1]));
      Writeln('The program exit code will be non 0 when there are errors.');
      Writeln('');
      Writeln('Startup parameters (optional):');
      Writeln(Format('1) Dictionary file (%s)', [cDefaultlanguageName]));
      Writeln(Format('2) Source directory or file path (%s)', [cDefaultSourcePath]));
      Writeln(Format('3) Path for the ignore files (%s)', [spellCheck.IngoreFilePath]));
      Writeln(Format('4) File extension mask (%s)', [cDefaultExtFilter]));
      Writeln('5) Scan sub directories (1)');
      Writeln('6) Add to ignore prompt (1)');
      Writeln('7) Provide suggestions (1)');
      Writeln('');
      Writeln('Ignore files:');
      Writeln(Format('The %s file will ignore the word if the word is in this file. An extension of %s.', [cIgnoreWordsName, cDefaultlanguageName]));
      Writeln(Format('The %s file will ignore the line if the text before the quote is in this file.', [cIgnoreCodeName]));
      Writeln(Format('The %s file will ignore files if the text in the file is contained in the path.', [cIgnorePathsName]));
      Writeln(Format('The %s file will ignore lines if the text in the file is equal to the line.', [cIgnoreLinesName]));
      Writeln(Format('The %s file will ignore lines if the text in the file is contained in the the line..', [cIgnoreContainsName]));
      Writeln('');
      Writeln('Ignore file names:');
      Writeln(cIgnoreWordsName);
      Writeln(cIgnoreCodeName);
      Writeln(cIgnoreFilesName);
      Writeln(cIgnorePathsName);
      Writeln(cIgnoreLinesName);
    end;
    procedure WriteErrors;
    var
      a: integer;
    begin
      for a := 0 to spellCheck.Errors.Count-1 do
        Writeln(spellCheck.Errors.Strings[a]);
    end;
  begin
    input := '';
    result := NativeInt(ecExitCodeOk);
    spellCheck := TSpellCheckLinter.Create;
    try
      if (Trim(ParamStr(cLanguageFilename)) = 'help') or
         (Trim(ParamStr(cLanguageFilename)) = 'man') then
        ShowManual
      else begin
        if Trim(ParamStr(cLanguageFilename)) <> '' then
          spellCheck.LanguageFilename := ParamStr(cLanguageFilename);
        if Trim(ParamStr(cSourceFilename)) <> '' then
          spellCheck.SourcePath := ParamStr(cSourceFilename);
        if Trim(ParamStr(cExtFilter)) <> '' then
          spellCheck.FileExtFilter := ParamStr(cExtFilter);
        spellCheck.Recursive := ParamStr(cResursive) = '1';
        if Trim(ParamStr(cIgnoreFilepath)) <> '' then
          spellCheck.IngoreFilePath := ParamStr(cIgnoreFilepath);
        if Trim(ParamStr(cProvideSuggestions)) <> '' then
          spellCheck.ProvideSuggestions := ParamStr(cProvideSuggestions) = '1';
        Writeln('Spell checking files, please wait...');
        spellCheck.Run;
        if SecondsBetween(spellCheck.StartTime, spellCheck.EndTime) < 10 then
          Writeln(Format('Checked %d files in %d milliseconds', [spellCheck.FileCount,
                                                      MilliSecondsBetween(spellCheck.StartTime, spellCheck.EndTime)]))
        else
          Writeln(Format('Checked %d files in %d seconds', [spellCheck.FileCount,
                                                     SecondsBetween(spellCheck.StartTime, spellCheck.EndTime)]));
        Writeln(Format('Error count %d', [spellCheck.ErrorsWords.Count]));
        Writeln(' ');
        WriteErrors;
        if spellCheck.Errors.Count > 0 then begin
          result := NativeInt(ecExitCodeErrors);
          if ParamStr(cHalt) <> '0' then begin
            if (spellCheck.ErrorsWords.Count > 0) then begin
              Writeln('Would you like to add these words to '+cIgnoreWordsName+'? Y/N or R to run again');
            if ParamStr(cExtFilter) <> '0' then
              Readln(input);
            if (input = 'y') or (input = 'Y') then
              spellCheck.AddToIgnoreFile;
            end else
              NoErrorsMsg;
          end;
        end else
          NoErrorsMsg;
      end;
    finally
      spellCheck.Free;
      if (input ='R') or
         (input ='r') then
        result := RunSpellCheck;
    end;
  end;
end.
