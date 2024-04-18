# Useful-Excel-code
Handy tips for using Excel



To find a text string within a column - even if it doesn't match the entire entry in a cell, and return one value if it is there and another if it isn't: 
```
# This is looking up the text in cell J33, in sheet Results, within the colum D2-D31 in sheet Choi_Orthologues. If the text is found, it returns 'YES' and 'NA' if not.
=IF(OR(COUNTIF(Choi_Orthologues!$D$2:$D$31,"*"&Results!J33&"*")),"YES","NA")
```
(see https://www.youtube.com/watch?v=2u4BbyjL-fY for more)


To find a text string within a column - even if it doesn't match the entire entry in a cell, and return the value of the adjacent cell: 
```
# If you don't have FALSE, but rather have TRUE then it will return the first cell in the Choi_Orthologues list
# If you don't have the "*"& wildcard and & sign, then it will expect it to be an EXACT match and will only return the adjacent cell if it is so.
# This is saying to look up the value in J33, and then search for it within the range in the Choi Orthologues. It it has a hit (even if it doesn't match the entire cell) then it is to return the value in the adjacent cell, in column 2 of the lookup range.
=VLOOKUP("*"&J33&"*",Choi_Orthologues!$B$2:$C$31,2,FALSE)
```
(see https://www.youtube.com/watch?v=11nza4eq8dw for more)
