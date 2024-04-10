# Using-macros-on-a-Multivariate-Wine-Data-Set-in-Excel-VBA

*Block codes are posted below. Attached is an example of what the macro in Excel can do to automate our work and another example of how to apply the WRONG methodology on a data set.*

*To run this macro, you need to copy and paste this code into your module. By tapping alt + F11 you can open VBA.
Once this is done, save it, close it and connect each sub with each button. The reason I did it like this was because Jithub Syst doesn't like to put my macro there.
It's really simple. Trust me! and good lucj!*

*The data winequality-red.cvs was obtained from https://archive.ics.uci.edu/dataset/186/wine+quality*

*The data winequality-red.cvs has been renamed as wdata for convenience*

*Please see the files. Tx!*


      Sub FitAndColor()
      
      ' FitAndColor Macro
      
      ' This macro fits columns & colors rows
      '
      ' Keyboard Shortcut: Ctrl+n
      
          Range("A1:L1").Select
          
          Selection.Columns.AutoFit
          
          With Selection.Interior
          
              .Pattern = xlSolid
              
              .PatternColorIndex = xlAutomatic
              
              .ThemeColor = xlThemeColorAccent6
              
              .TintAndShade = 0.799981688894314
              
              .PatternTintAndShade = 0
              
          End With
          
          Selection.AutoFilter
          
          ActiveSheet.Range("$A$1:$L$1600").AutoFilter Field:=12, Criteria1:="3"
          
          Range("A461").Select
          
          Range(Selection, Selection.End(xlToRight)).Select
          
          Range(Selection, Selection.End(xlDown)).Select
          
          With Selection.Interior
          
              .Pattern = xlSolid
              
              .PatternColorIndex = xlAutomatic
              
              .Color = 6750207
              
              .TintAndShade = 0
              
              .PatternTintAndShade = 0
              
          End With
          
          ActiveSheet.Range("$A$1:$L$1600").AutoFilter Field:=12, Criteria1:="4"
          
          Range("A47").Select
          
          Selection.End(xlUp).Select
          
          Range("A20").Select
          
          Range(Selection, Selection.End(xlToRight)).Select
      
          Range(Selection, Selection.End(xlDown)).Select
          
          With Selection.Interior
          
              .Pattern = xlSolid
              
              .PatternColorIndex = xlAutomatic
              
              .Color = 16751001
              
              .TintAndShade = 0
              
              .PatternTintAndShade = 0
              
          End With
          
          ActiveSheet.Range("$A$1:$L$1600").AutoFilter Field:=12, Criteria1:="5"
          
          Selection.End(xlUp).Select
          
          Range("A2").Select
          
          Range(Selection, Selection.End(xlToRight)).Select
          
          Range(Selection, Selection.End(xlDown)).Select
          
          With Selection.Interior
          
              .Pattern = xlSolid
              
              .PatternColorIndex = xlAutomatic
              
              .Color = 39423
              
              .TintAndShade = 0
              
              .PatternTintAndShade = 0
              
          End With
          
          ActiveSheet.Range("$A$1:$L$1600").AutoFilter Field:=12, Criteria1:="6"
          Range("A5").Select
          Range(Selection, Selection.End(xlToRight)).Select
          Range(Selection, Selection.End(xlDown)).Select
          With Selection.Interior
              .Pattern = xlSolid
              .PatternColorIndex = xlAutomatic
              .Color = 13311
              .TintAndShade = 0
              .PatternTintAndShade = 0
              
          End With
          
          ActiveSheet.Range("$A$1:$L$1600").AutoFilter Field:=12, Criteria1:="7"
          
          Range("A9").Select
          
          Range(Selection, Selection.End(xlToRight)).Select
          
          Range(Selection, Selection.End(xlDown)).Select
          
          With Selection.Interior
          
              .Pattern = xlSolid
              
              .PatternColorIndex = xlAutomatic
              
              .Color = 3368601
              
              .TintAndShade = 0
              
              .PatternTintAndShade = 0
              
          End With
          
          ActiveSheet.Range("$A$1:$L$1600").AutoFilter Field:=12, Criteria1:="8"
          
          Range("A269").Select
          
          Range(Selection, Selection.End(xlToRight)).Select
          
          Range(Selection, Selection.End(xlDown)).Select
          
          With Selection.Interior
          
              .Pattern = xlSolid
              
              .PatternColorIndex = xlAutomatic
              
              .ThemeColor = xlThemeColorAccent6
              
              .TintAndShade = -0.249977111117893
              
              .PatternTintAndShade = 0
              
          End With
          
          Range("A1").Select
          
          ActiveSheet.ShowAllData
          
          Selection.AutoFilter
      End Sub
      
      Sub Undo()
      
          ' Undo Macro
          
          ' Undo everything : Deshace todo
          '
          ' Keyboard Shortcut: Ctrl+m
          '
              Range("A1").Select
              
              Range(Selection, Selection.End(xlToRight)).Select
              
              Range(Selection, Selection.End(xlDown)).Select
              
              With Selection.Interior
              
                  .Pattern = xlNone
                  
                  .TintAndShade = 0
                  
                  .PatternTintAndShade = 0
                  
              End With
              
          Selection.ColumnWidth = 8.43
          
          Range("U20").Select
      End Sub



