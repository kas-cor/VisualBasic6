VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' We must implement this interface to be considered a GenWrapper plugin.
' This means code must be written for the three methods:
'  IRjlWinAmpGenPlugin_Configure
'  IRjlWinAmpGenPlugin_Initialize
'  IRjlWinAmpGenPlugin_Quit
' (Select IRjlWinAmpGenPlugin from the left hand dropdown list above.)
Implements IRjlWinAmpGenPlugin

' Instance variables
'  We remember the wrapper instance so we can get properties from WinAmp
'  such as its main window handle.
'  The hotkey value is also an instance variable. This is zero if we have
'  no hotkey or the ascii code - 64 of the alphabetic hotkey.
Private m_Wrapper As IRjlWinAmpGenWrapper

Private Sub IRjlWinAmpGenPlugin_Configure()
' This is called by GenWrapper when plugin should configure itself

' Display configuration form
MsgBox App.FileDescription

End Sub

Private Sub IRjlWinAmpGenPlugin_Initialize(ByVal Wrapper As RjlWinAmpGenLib.IRjlWinAmpGenWrapper)
' This is called by GenWrapper when plugin should initialize.
' The value of Wrapper should be remembered for future reference.

' Ensure there's only one of us
If Not This Is Nothing Then
    Err.Raise vbObjectError + 1, , "Already have a plugin instance"
    Exit Sub
End If
Set This = Me

' Remember our wrapper object - we need it for WinAmp properties
Set m_Wrapper = Wrapper

' Set our description
m_Wrapper.Description = "tcpctrl Plugin v." & App.Major & "." & App.Minor & "." & App.Revision & " (gen_tcpctrl.dll)"

main.ld
 
End Sub

Private Sub IRjlWinAmpGenPlugin_Quit()
    ' Called when plugin is about to be unloaded, so should clean up

    ' Clean up
    Unload frmHidden
    
End Sub

Public Sub Info()

' Display a messagebox showing information available from the wrapper
'MsgBox _
    "Plugin Description: " & vbCrLf & _
    m_Wrapper.Description & vbCrLf & vbCrLf & _
    "Wrapper Description: " & vbCrLf & _
    m_Wrapper.WrapperDescription & vbCrLf & vbCrLf & _
    "WinAmp Window Handle: 0x" & Hex(m_Wrapper.HwndParent) & vbCrLf & _
   "Wrapper DLL Handle: 0x" & Hex(m_Wrapper.WrapperDllInstance) & vbCrLf & _
    "Wrapper Version: " & Trim(CStr(m_Wrapper.WrapperVersion \ &H10000)) & _
        "." & Trim(CStr(m_Wrapper.WrapperVersion And &HFFFF&)), _
        vbInformation, "VbTRC Information"
        
 MsgBox "Plugin Description: " & vbCrLf & m_Wrapper.Description & vbCrLf & _
    "WinAmp Window Handle: 0x" & Hex(m_Wrapper.HwndParent) _
    , vbInformation, "tcpctrl Information"
    
End Sub
