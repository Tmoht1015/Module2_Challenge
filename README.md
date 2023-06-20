# Module2_Challenge
Module Challenge #2
Sources for Data Help:

+' For Each ws In Worksheets & Next ws & wsname = ws.Name  '
Sources: https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
&Tutoring session regarding the subject on 06/20/23 w/ Marc Calache

+'  PercentChg = ((ws.Cells(x, 6).Value - ws.Cells(y, 3).Value) / ws.Cells(y, 3).Value) "
Source: https://stackoverflow.com/questions/64707941/i-keep-getting-an-error-that-says-compile-error-for-without-next

+' For the use of Formats like "Percent" & "Scientific"
Source: https://www.techonthenet.com/excel/formulas/format_string.php

+' LastRowTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row ' &  ' Dim LastRowTicker As Long '
Source: **See screenshot attached with a dialog in Slack w/ peer**

+'  ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(x, 7))) '
Source: https://www.automateexcel.com/vba/sum-function/
