Attribute VB_Name = "Mod_Main"
Option Explicit

Sub Main()
    Mod_HookSkin.Attach Frm_Main.hWnd
    Frm_Main.Show
End Sub
