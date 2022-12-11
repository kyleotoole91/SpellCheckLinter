unit uConstants;

interface

const
  cMaxThreads=10;
  cSleepTime=10;
  cMaxStringLengthDFM=64;
  cMaxSuggestions=100;
  cSuggestLengthOffset=3;
  cGuidLen=36;
  cFileExtLen=5; //includes period eg .jpeg = 5
  cMinCheckLength=3; //there must be X number of chars in order for it to get checked
  cIgnoreCamelPrefixLength=4; //there must be X number of chars in order for it to get checked
  cWordSeparator='/';
  cSkipLineEndString=';';
  cSkipLineEndStringDFM=';';
  cSpellCheckOn='//SPELL_CHECK_ON';
  cSpellCheckOff='//SPELL_CHECK_OFF';
  cSkippedLineCountLimit=220; //
  cDefaultExtFilter='*.pas|*.dfm';
  cDefaultSourcePath='.\';
  cDefaultIgnorePath= '.\';
  cDefaultQuote='''';
  cIgnoreContainsName='IgnoreContains.dic';
  cIgnoreCodeName='IgnoreCode.dic';
  cIgnorePathsName='IgnorePaths.dic';
  cIgnoreFilesName='IgnoreFiles.dic';
  cIgnoreWordsName='IgnoreWords.dic';
  cIgnoreLinesName='IgnoreLines.dic';
  cDefaultlanguageName='en_US.dic';
  //Regular Expressions
  cRegExCamelPascalCaseSpliter='([A-Z]+|[A-Z]?[a-z]+)(?=[A-Z]|\b)';
  cRegExKeepLettersAndQuotes='[^A-Z/0-9/''a-z/''.]';
  cRegExRemoveNumbers='[^A-Z/''a-z/''.]';

implementation

end.
