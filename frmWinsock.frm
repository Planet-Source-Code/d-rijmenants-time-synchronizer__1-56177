VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWinsock 
   Caption         =   "frmWinsock"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1815
   Icon            =   "frmWinsock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   1815
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   450
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

