program SpellCheckLinter;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  DateUtils,
  uConstants in 'uConstants.pas',
  uSpellChecker in 'uSpellChecker.pas';

const
    cLanguageFilename=1; //.dic file
    cSourceFilename=2; //source path or filename
    cExtFilter=3; //eg *.pas
    cResursive=4; //when 0, resusive scan is disabled
    cHalt=5; //when 0, halting at the end of run is disabled
    cQuoteSym=6; //defaulted to a single quote for .pas files. Strings between these quote symbols are checked

  procedure RunSpellChecker;
  var
    a: integer;
    input: string;
    spellChecker: TSpellChecker;
    procedure NoErrorsMsg;
    begin
      if ParamStr(cHalt) <> '0' then begin
        Writeln('Press Enter to close');
        ReadLn(input);
      end;
    end;
  begin
    input := '';
    spellChecker := TSpellChecker.Create;
    try
      if Trim(ParamStr(cLanguageFilename)) <> '' then
        spellChecker.LanguageFilename := ParamStr(cLanguageFilename);
      if Trim(ParamStr(cSourceFilename)) <> '' then
        spellChecker.SourcePath := ParamStr(cSourceFilename);
      if Trim(ParamStr(cExtFilter)) <> '' then
        spellChecker.FileExtFilter := ParamStr(cExtFilter);
      spellChecker.Recursive := ParamStr(cResursive) <> '0';
      if Trim(ParamStr(cQuoteSym)) <> '' then
        spellChecker.QuoteSym := ParamStr(cQuoteSym);
      Writeln('Linting files, please wait...');
      spellChecker.Run;
      for a := 0 to spellChecker.Errors.Count-1 do
        Writeln(spellChecker.Errors.Strings[a]);
      if SecondsBetween(spellChecker.StartTime, spellChecker.EndTime) = 0 then
        Writeln(Format('Checked %d files in %dms', [spellChecker.FileCount,
                                                    MilliSecondsBetween(spellChecker.StartTime, spellChecker.EndTime)]))
      else
        Writeln(Format('Checked %d files in %ds', [spellChecker.FileCount,
                                                   SecondsBetween(spellChecker.StartTime, spellChecker.EndTime)]));
      Writeln(Format('Unmatch count %d', [spellChecker.ErrorsWords.Count]));
      if spellChecker.Errors.Count > 0 then begin
        if ParamStr(cHalt) <> '0' then begin
          if (spellChecker.ErrorsWords.Count > 0) then begin
            Writeln('Would you like to add these words to '+cIgnoreWords+'? Y/N');
          if ParamStr(cExtFilter) <> '0' then
            Readln(input);
          if (input = 'y') or (input = 'Y') then
            spellChecker.AddToIgnoreFile;
          end else
            NoErrorsMsg;
        end;
      end else
        NoErrorsMsg;
    finally
      spellChecker.Free;
      if (input ='R') or
         (input ='r') then
        RunSpellChecker;
    end;
  end;

begin
  try
    RunSpellChecker;
  except
    on e: Exception do begin
      Writeln(e.ClassName, ': ', e.Message);
      if ParamStr(cHalt) <> '0' then
        Readln(input);
    end;
  end;
end.
