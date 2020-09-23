VERSION 5.00
Begin VB.Form frmWork 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicNumeros 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   9660
      Picture         =   "frmWork.frx":0000
      ScaleHeight     =   19
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox PicBakBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   421
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   673
      TabIndex        =   1
      Top             =   0
      Width           =   10155
   End
   Begin VB.PictureBox PicClean 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6375
      Left            =   720
      ScaleHeight     =   421
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   673
      TabIndex        =   0
      Top             =   360
      Width           =   10155
   End
End
Attribute VB_Name = "frmWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

