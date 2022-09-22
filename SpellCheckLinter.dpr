program SpellCheckLinter;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  DateUtils,
  uSpellChecker in '..\SpellChecker\uSpellChecker.pas',
  uConstants in '..\SpellChecker\uConstants.pas';

procedure RunSpellChecker;
  const
    cLanguageFilename=1; //.dic file
    cSourceFilename=2; //source path or filename
    cExtFilter=3; //eg *.pas
    cResursive=4; //when 0, resusive scan is disabled
    cHalt=5; //when 0, halting at the end of run is disabled
    cQuoteSym=6; //defaulted to a single quote ' for .pas files. Strings between these quote symbols are checked
  var
    a: integer;
    input: string;
    spellChecker: TSpellChecker;
  begin
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
      spellChecker.Run;
      Writeln(Format('Checked %d files (%s)', [spellChecker.FileCount, spellChecker.FileExtFilter]));
      Writeln(Format('Checked files in %d seconds', [SecondsBetween(spellChecker.StartTime, spellChecker.EndTime)]));
      if spellChecker.Errors.Count > 0 then
        Writeln(Format('Error count %d', [spellChecker.Errors.Count]))
      else
        Writeln('No errors found');
      for a := 0 to spellChecker.Errors.Count-1 do
        Writeln(spellChecker.Errors.Strings[a]);
      if ParamStr(cExtFilter) <> '0' then
        Readln(input);
    finally
      spellChecker.Free;
    end;
  end;

begin
  try
    RunSpellChecker;
  except
    on e: Exception do
      Writeln(e.ClassName, ': ', e.Message);
  end;
end.
