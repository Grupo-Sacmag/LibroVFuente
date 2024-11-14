VERSION 5.00
Begin VB.Form Directorio 
   Caption         =   "Establecer directorio de Trabajo"
   ClientHeight    =   3036
   ClientLeft      =   96
   ClientTop       =   480
   ClientWidth     =   4572
   Icon            =   "Directorio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3036
   ScaleWidth      =   4572
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Canc 
      Caption         =   "Cancelar y Salir"
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Acept 
      Caption         =   "Aceptar y Salir"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.DirListBox dir 
      Height          =   1665
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Directorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Acept_Click()
   Sub_dir = dir + "\"
   Unload Directorio
End Sub

Private Sub Canc_Click()
    Unload Directorio
End Sub

Private Sub dir_Change()
 On Error GoTo CAMBIO
    ChDir (dir)
    Directorio.Caption = dir
    Exit Sub
CAMBIO:
   ChDir "C:\"
End Sub

Private Sub dir_Click()

   dir_Change
   
End Sub

Private Sub dir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Directorio.Caption <> dir.List(dir.ListIndex) Then
        dir.Path = dir.List(dir.ListIndex)
        dir_Change
     End If
  End If
End Sub

Private Sub Drive1_Change()

    ChDrive (Drive1)
    dir.Path = Drive1.Drive
    Directorio.Caption = dir
    SCont.guarda = Trim(dir)
    Put 1, 1, SCont
    
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Drive1_Change
  End If
End Sub

Private Sub Form_Load()
   Directorio.Caption = dir
End Sub
