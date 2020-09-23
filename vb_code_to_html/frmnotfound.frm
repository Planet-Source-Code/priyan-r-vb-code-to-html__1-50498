VERSION 5.00
Begin VB.Form frmnotfound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Missing Fies"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5565
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "The Following files are missing in the VB  project"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmnotfound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

