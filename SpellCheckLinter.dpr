program SpellCheckLinter;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  uConstants in 'uConstants.pas',
  uSpellChecker in 'uSpellChecker.pas',
  uSpellCheckerController in 'uSpellCheckerController.pas';

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
