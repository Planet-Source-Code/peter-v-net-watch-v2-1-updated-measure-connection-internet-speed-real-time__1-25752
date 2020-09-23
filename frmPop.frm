VERSION 5.00
Begin VB.Form frmPop 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   LinkTopic       =   "Form1"
   ScaleHeight     =   225
   ScaleWidth      =   705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------
' Created by Peter Verburgh.
' @ 20001 at Brugge, Belgium
' this application may NOT used  commercional ..
' before  asking me first..
' This is Freeware..
'------------------------------------------------------------------------------------------
' This form will be used if the user click on the icon is system tray..
' then it will show a popup window..
' but if your moues - cursor is out the popupbox ... it keep standing ...
' and the user wants after sertain time (0.5 to 1 sec ) that the popup windows
'automatic hides..
'So what i do is... in the frmMenu application , detects the mouse settings X.Y,
' and if the mouse is out the popup rectangular .. it detects & then it will call a timer
' after a time .. it calls this window .. but hidden and unload it directely..
' on that moment it the popup doesnt get the focus anymore and hide..
' on the hidden window here.
' i think it could be done by  using api calls to the handle of the menu..
' but why not using the easy way  !
'---------------------------------------------------------------------------------------------------
