Dim xlApp
Set xlApp = CreateObject("Excel.Application")

xlApp.Visible = True
xlApp.Workbooks.Add

xlApp.VBE.ActiveVBProject.VBComponents.Import "C:\Users\Nilay.Baranwal\Desktop\Module1.bas"
xlApp.VBE.ActiveVBProject.VBComponents.Import "C:\Users\Nilay.Baranwal\Desktop\Module2.bas"
xlApp.VBE.ActiveVBProject.VBComponents.Import "C:\Users\Nilay.Baranwal\Desktop\Module3.bas"

xlApp.Run "runMe"
xlApp.ActiveWorkbook.SaveAs Year(now) & Right("0" & Month(Now), 2) & Right("0" & Day(now), 2) & "_XCAL_LogList.xlsx"
xlApp.Quit