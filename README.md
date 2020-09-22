# sharesight-file-import-scripts
Scripts to import Saxo Markets trade CSV files into Sharesight


Python script that converts trades exported from [SaxoTrader](https://www.saxotrader.com) into [Sharesight](https://www.sharesight.com/).

Sample usage:
```%>./saxo-to-sharesight.py --file TradesExecuted.xlsx > sharesight.csv```

Note
1. Does not support corporate actions - you may need to amend those trades manually in Sharesight during import
2. Script mainly works for USD trades. GBP trades need special handling during Sharesight import. Sharesight may mark them as GBP (pounds) instead of GBp (pence) so you may need to manually amend during import. 
