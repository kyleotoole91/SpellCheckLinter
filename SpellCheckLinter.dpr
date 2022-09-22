program SpellCheckLinter;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  uSpellChecker in '..\SpellChecker\uSpellChecker.pas',
  uConstants in '..\SpellChecker\uConstants.pas';

procedure RunSpellChecker;
  const
    cLanguageFilename=1;
    cSourceFilename=2;
    cExtFilter=3;
    cResursive=4;
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
      spellChecker.Run;
      Writeln(Format('Checked %d files', [spellChecker.FileCount]));
      Writeln(Format('Error count %d', [spellChecker.Errors.Count]));
      if spellChecker.Errors.Count > 0 then
        Writeln('Errors:');
      for a := 0 to spellChecker.Errors.Count-1 do
        Writeln(spellChecker.Errors.Strings[a]);
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
