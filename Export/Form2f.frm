VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please wait..."
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Loading data..."
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
