Attribute VB_Name = "Module1"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ThreadProc()
    Do
        VBMAN.Json.RootItems.Add VBMAN.ToolsStr.GetRandStr()
    Loop
End Sub
