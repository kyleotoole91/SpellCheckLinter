program SpellCheckLinter;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  uConstants in 'uConstants.pas',
  uSpellCheck in 'uSpellCheck.pas',
  uSpellCheckController in 'uSpellCheckController.pas';

begin
  try
    RunSpellCheck;
  except
    on e: Exception do begin
      Writeln(e.ClassName, ': ', e.Message);
      if ParamStr(cHalt) <> '0' then
        Readln(input);
    end;
  end;
end.
