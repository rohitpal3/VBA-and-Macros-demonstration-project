# VBA-and-Macros-demonstration-project
1.In this project I used Macros and VBA to perform and set of task and with help of macros recorded  the pattern and applied the same pattern in other worksheet.</br>

<b> Tool Used: </b></br>
Microsoft Excel</br>

<b> Code </b></br>
Sub Macro1()</br>
'
' Macro1 Macro</br>
'
' Keyboard Shortcut: Ctrl+r</br>
'
    Range("A3").Select</br>
    Range(Selection, Selection.End(xlDown)).Select</br>
    Range(Selection, Selection.End(xlToRight)).Select</br>
    Range(Selection, Selection.End(xlToRight)).Select</br>
    Range(Selection, Selection.End(xlToLeft)).Select</br>
    Range(Selection, Selection.End(xlToRight)).Select</br>
    ActiveSheet.Shapes.AddChart2(286, xl3DColumnStacked).Select</br>
    ActiveChart.SetSourceData Source:=Range("Sheet5!$A$3:$D$7")</br></br>
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)</br>
End Sub</br>
</br>
<b>Output</b></br>
![image](https://user-images.githubusercontent.com/111983642/189544876-f6b5b7ce-cfdb-45a2-a7fa-bf90772c47ff.png)
