unit uConstants;

interface

const
  cGuidLen=36;
  cFileExtLen=4; //includes period eg .txt = 4
  cMinCheckLength=2; //there must be X number of chars in order for it to get checked
  cDefaultExtFilter='*.pas';
  cDefaultSourcePath='.\';
  cDefaultQuote='''';
  cIgnoreFiles='IgnoreFiles.dic';
  cIgnoreWords='IgnoreWords.dic';
  cIgnoreLines='IgnoreLines.dic';
  cDefaultlanguagePath='en_US.dic';
  //Regular Expressions
  cRegExCamelPascalCaseSpliter='([A-Z]+|[A-Z]?[a-z]+)(?=[A-Z]|\b)';
  cRegExKeepLettersAndQuotes='[^A-Z/''a-z/''.]';

implementation

end.
