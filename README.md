# maven-excel

Parsing Microsoft Excel files thru **XSSF (poi and poi-ooxml maven dependencies)**

Two metods created:

  1. <code>getDataFull</code> -> returns *HashMap*, the whole row in format key => value **{TestCases=Purchase, col2=Data2.2, col3=Data2.3, col1=Data2.1}**
  2. <code>getData</code> -> returns *ArryList*, the row, where index 0 is a row name, all following indexes in array - data **[Purchase, Data2.1, Data2.2, Data2.3]**
