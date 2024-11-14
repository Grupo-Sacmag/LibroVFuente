VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Visor 
   Caption         =   "Visor subctas"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10290
   Begin MSFlexGridLib.MSFlexGrid Vsr 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11668
      _Version        =   393216
   End
End
Attribute VB_Name = "Visor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
          Vsr.Cols = 15
          Vsr.Rows = 2
          For i = 0 To 14
             Vsr.ColWidth(i) = 1300
          Next i
          For i = 0 To 14
              Vsr.Row = 0
              Vsr.Col = i
              Vsr.Text = Mm(i)
              Vsr.CellAlignment = 4
              Vsr.CellFontBold = True
         Next i
    Close 6, 7
         Open Sub_dir + "DEBES.GCS" For Random As 6 Len = Len(MvDebeS)
    EM = LOF(6) / Len(MvDebeS)
    Open Sub_dir + "HABERS.GCS" For Random As 7 Len = Len(MvHaberS)
    Fm = LOF(7) / Len(MvHaberS)
    If EM > Fm Then
            finalf = EM
            Else
            finalf = Fm
    End If
    Vsr.Rows = 1
    
    For w1 = 1 To finalf
        Get 6, w1, MvDebeS
        Get 7, w1, MvHaberS
        Vsr.AddItem w1 & Chr(9) & Format(MvDebeS.Inc, z1) & Chr(9) & _
                    Format(MvDebeS.Ene, z1) & Chr(9) & _
                    Format(MvDebeS.Feb, z1) & Chr(9) & _
                    Format(MvDebeS.Mar, z1) & Chr(9) & _
                    Format(MvDebeS.Abr, z1) & Chr(9) & _
                    Format(MvDebeS.May, z1) & Chr(9) & _
                    Format(MvDebeS.Jun, z1) & Chr(9) & _
                    Format(MvDebeS.Jul, z1) & Chr(9) & _
                    Format(MvDebeS.Ago, z1) & Chr(9) & _
                    Format(MvDebeS.Sep, z1) & Chr(9) & _
                    Format(MvDebeS.Oct, z1) & Chr(9) & _
                    Format(MvDebeS.Nov, z1) & Chr(9) & _
                    Format(MvDebeS.Dic, z1)
       Vsr.AddItem w1 & Chr(9) & Format(MvHaberS.Inc, z1) & Chr(9) & _
                    Format(MvHaberS.Ene, z1) & Chr(9) & _
                    Format(MvHaberS.Feb, z1) & Chr(9) & _
                    Format(MvHaberS.Mar, z1) & Chr(9) & _
                    Format(MvHaberS.Abr, z1) & Chr(9) & _
                    Format(MvHaberS.May, z1) & Chr(9) & _
                    Format(MvHaberS.Jun, z1) & Chr(9) & _
                    Format(MvHaberS.Jul, z1) & Chr(9) & _
                    Format(MvHaberS.Ago, z1) & Chr(9) & _
                    Format(MvHaberS.Sep, z1) & Chr(9) & _
                    Format(MvHaberS.Oct, z1) & Chr(9) & _
                    Format(MvHaberS.Nov, z1) & Chr(9) & _
                    Format(MvHaberS.Dic, z1)
                    
    Next w1
    
End Sub

Private Sub Form_Resize()
 If Visor.WindowState <> 1 Then
      Vsr.Height = ScaleHeight - 200
      Vsr.Width = ScaleWidth - 200
      Rem F_Aum = (Bcha1.Width - 400) / 9200
      Rem ColDfn
   End If

End Sub
