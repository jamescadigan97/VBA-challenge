{\rtf1\ansi\ansicpg1252\cocoartf2512
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Analysis2()\
Dim ws As Worksheet\
For Each ws In Worksheets\
ws.Activate\
 \
'Added Headers\
Range("I1").Value = "Ticker"\
Range("j1").Value = "Yearly Change"\
Range("k1").Value = "Precentage Change"\
Range("l1").Value = "Total Volume"\
Range("p1").Value = "Ticker"\
Range("q1").Value = "Value"\
Range("o2").Value = "Greatest % Increase"\
Range("o3").Value = "Greatest % Decrease"\
Range("o4").Value = "Largest volume"\
\
'Last row calculations for data\
Dim lastrow As Double\
lastrow = Cells(Rows.Count, 1).End(xlUp).Row\
\
' Add Each New ticker to the top\
Dim top_t As Double\
top_t = 1\
For i = 2 To lastrow\
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
        top_t = top_t + 1\
        Cells(top_t, 9).Value = Cells(i, 1).Value\
    End If\
Next i\
\
'Lastrow for calculations\
Dim lastr As Long\
lastr = Cells(Rows.Count, 9).End(xlUp).Row\
\
\
'Added volumes from similiar stocks\
 Dim top_v As Integer\
 top_v = 2\
For i = 2 To lastrow\
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
        top_v = top_v + 1\
    ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then\
        Cells(top_v, 12).Value = Cells(top_v, 12).Value + Cells(i, 7).Value\
    End If\
Next i\
\
'Found the difference\
Dim top_c As Integer\
top_c = 2\
Range("J2").Value = 0 - Range("C2").Value\
\
For i = 2 To lastrow - 1\
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
        Cells(top_c, 10).Value = Cells(top_c, 10).Value + Cells(i, 6).Value\
        top_c = top_c + 1\
        Cells(top_c, 10).Value = Cells(top_c, 10).Value - Cells(i + 1, 3).Value\
    End If\
Next i\
\
'Formatted colors\
For i = 2 To lastr\
Cells(i, 11).Value = Cells(i, 10).Value\
    If Cells(i, 10).Value > 0 Then\
        Cells(i, 10).Interior.ColorIndex = 4\
    Else\
        Cells(i, 10).Interior.ColorIndex = 3\
    End If\
Next i\
\
' Found precent change\
top_p = 1\
Range("K2").Value = Range("J2").Value / Range("C2").Value\
\
For i = 2 To lastrow\
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then\
    top_p = top_p + 1\
        If Cells(i + 1, 3).Value <> 0 Then\
            Cells(top_p, 11).Value = Cells(top_p, 10).Value / Cells(i + 1, 3).Value\
        Else\
            Cells(top_p, 11).Value = 0\
        End If\
End If\
Next i\
\
'Formatted precents\
For i = 2 To lastr\
    Cells(i, 11).NumberFormat = "0.00%"\
Next i\
\
'Found greatest increase and decrease\
For j = 2 To lastr\
    If Cells(j, 11).Value > Range("q2").Value Then\
        Range("q2").Value = Cells(j, 11).Value\
        Range("p2").Value = Cells(j, 9).Value\
    ElseIf Cells(j, 11).Value < Range("q3").Value Then\
        Range("q3").Value = Cells(j, 11).Value\
        Range("p3").Value = Cells(j, 9).Value\
    End If\
Next j\
\
'Formatted as precents\
Cells(2, 17).NumberFormat = "0.00%"\
Cells(3, 17).NumberFormat = "0.00%"\
\
'Found largest volume\
For j = 3 To lastr\
    If Cells(j, 12).Value > Range("q4").Value Then\
        Range("q4").Value = Cells(j, 12).Value\
        Range("p4").Value = Cells(j, 9).Value\
    End If\
Next j\
Next\
End Sub}