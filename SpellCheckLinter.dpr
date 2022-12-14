program SpellCheckLinter;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  uConstants in 'uConstants.pas',
  uSpellCheckLinter in 'uSpellCheckLinter.pas',
  uSpellCheckController in 'uSpellCheckController.pas',
  uSpellCheckFile in 'uSpellCheckFile.pas';

begin
  try
    ExitCode := RunSpellCheck;
  except
    on e: Exception do begin
      ExitCode := NativeInt(ecException);
      Writeln(e.ClassName, ': ', e.Message);
      if ParamStr(cHalt) <> '0' then
        Readln(input);
    end;
  end;
end.
