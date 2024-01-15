Dim xlApp
Set xlApp = CreateObject("Excel.Application")

xlApp.Visible = True
xlApp.Workbooks.Add

xlApp.VBE.ActiveVBProject.VBComponents.Import "C:\Users\Nilay.Baranwal\Desktop\Module1.bas"
xlApp.VBE.ActiveVBProject.VBComponents.Import "C:\Users\Nilay.Baranwal\Desktop\Module2.bas"
xlApp.VBE.ActiveVBProject.VBComponents.Import "C:\Users\Nilay.Baranwal\Desktop\Module3.bas"

xlApp.Run "runMe"