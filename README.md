Program to track investments on Excel & calculating DRIP

How to use?
- Enable Developer tab in Microsoft Excel
- Open visual basic editor from Develop tab
- Replace the code with the following
- Sub SampleCall()
  mymodule = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
  RunPython "import portfolio_tracking; portfolio_tracking.main()"
  End Sub
-


Thank you, Sven Bo for sharing the starter project.
https://github.com/Sven-Bo/portfolio-tracking-excel-python

Thank you, Adrian for sharing his financial knowledge which inspired me to create this project.
https://www.youtube.com/watch?v=ouyXwaTOfhU
