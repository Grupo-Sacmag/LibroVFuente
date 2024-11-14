VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Balancha 
   Caption         =   "Balanza con Movimientos de Subcuentas"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   9435
   Begin MSFlexGridLib.MSFlexGrid Bcha1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10398
      _Version        =   393216
      BackColorBkg    =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Balancha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_Aum As Double, L As Long, LL As Long, RR As Long
Dim MD As Currency, MH As Currency, Md_1 As Currency, Mh_1 As Currency
Dim Cla_ve As String

Private Sub Form_Activate()
    BALANZON = 4
    LIBROSV.Ci.Visible = False
End Sub

Sub DIBUJA(mes_lim)
    mes_lim = final
    mes_lim1 = final
    If mes_lim < 1 Or mes_lim > Arc_FinaL Then
        MsgBox "No es posible mostrar una balanza a esa fecha ", vbCritical
        Exit Sub
    End If
    Dim Sdo_ini As Currency, M_D As Currency, M_H As Currency
    Dim M_ue As Integer, z1 As String, Sdo_Fin As Currency
    z1 = "#,##0.00;(#,##0.00)"
    Close
    Open Sub_dir + "DATOS" For Random As 1 Len = Len(Datos)
    Balanza.Caption = "Balanza de comprobacion a " + Trim(Mm(mes_lim)) + " de " + Datos.a_o
    Close
    Bcha1.Clear
    Bcha1.Rows = 1
    Bcha1.Row = 0: Bcha1.Col = 0: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "Cuenta"
    Bcha1.Row = 0: Bcha1.Col = 1: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "Nombre"
    Bcha1.Row = 0: Bcha1.Col = 2: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "Saldo " + Mm(mes_lim - 1)
    Bcha1.Row = 0: Bcha1.Col = 3: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "Debe"
    Bcha1.Row = 0: Bcha1.Col = 4: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "Haber"
    Bcha1.Row = 0: Bcha1.Col = 5: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "Saldo " + Mm(mes_lim)
    Bcha1.Row = 0: Bcha1.Col = 6: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "CL."
    Bcha1.Row = 0: Bcha1.Col = 7: Bcha1.CellFontBold = True: Bcha1.CellAlignment = 4: Bcha1.Text = "REAL"
    Bcha1.FixedCols = 2
    Open Sub_dir + "CATMAY" For Random As 2 Len = Len(CATMAY)
    Open Sub_dir + "CATAUX" For Random As 3 Len = Len(CATAUX)
    Open Sub_dir + "DEBE.GCS" For Random As 4 Len = Len(MvDebe)
    Cm = LOF(4) / Len(MvDebe)
    Open Sub_dir + "HABER.GCS" For Random As 5 Len = Len(MvHaber)
    Dm = LOF(5) / Len(MvHaber)
    Open Sub_dir + "DEBES.GCS" For Random As 6 Len = Len(MvDebeS)
    Cm = LOF(6) / Len(MvDebeS)
    Open Sub_dir + "HABERS.GCS" For Random As 7 Len = Len(MvHaberS)
    Dm = LOF(7) / Len(MvHaberS)

    If Dm > Cm Then Cm = Dm
        Rem Bcha1.AddItem "Cuenta" & Chr(9) & "Nombre" & Chr(9) & "Saldo " + mm(Mes_Lim - 1) _
                                         & Chr(9) & "Debe" & Chr(9) & "Haber" & Chr(9) & "Saldo " + mm(Mes_Lim)
     
    For L = 1 To Cm
        M_ue = 0: Sdo_ini = 0: M_D = 0: M_H = 0: Sdo_Fin = 0
        Get 2, L, CATMAY
        Get 4, L, MvDebe: Get 5, L, MvHaber
        For r = 0 To mes_lim
            Select Case r
               Case 0
                Sdo_ini = Sdo_ini + MvDebe.Inc + MvHaber.Inc
               Case 1
                 Sdo_ini = Sdo_ini + MvDebe.Ene + MvHaber.Ene
                 M_D = MvDebe.Ene: M_H = MvHaber.Ene
               Case 2
                 Sdo_ini = Sdo_ini + MvDebe.Feb + MvHaber.Feb
                 M_D = MvDebe.Feb: M_H = MvHaber.Feb
               Case 3
                 Sdo_ini = Sdo_ini + MvDebe.Mar + MvHaber.Mar
                 M_D = MvDebe.Mar: M_H = MvHaber.Mar
               Case 4
                 Sdo_ini = Sdo_ini + MvDebe.Abr + MvHaber.Abr
                 M_D = MvDebe.Abr: M_H = MvHaber.Abr
               Case 5
                 Sdo_ini = Sdo_ini + MvDebe.May + MvHaber.May
                 M_D = MvDebe.May: M_H = MvHaber.May
               Case 6
                 Sdo_ini = Sdo_ini + MvDebe.Jun + MvHaber.Jun
                 M_D = MvDebe.Jun: M_H = MvHaber.Jun
               Case 7
                 Sdo_ini = Sdo_ini + MvDebe.Jul + MvHaber.Jul
                 M_D = MvDebe.Jul: M_H = MvHaber.Jul
               Case 8
                 Sdo_ini = Sdo_ini + MvDebe.Ago + MvHaber.Ago
                 M_D = MvDebe.Ago: M_H = MvHaber.Ago
               Case 9
                 Sdo_ini = Sdo_ini + MvDebe.Sep + MvHaber.Sep
                 M_D = MvDebe.Sep: M_H = MvHaber.Sep
               Case 10
                 Sdo_ini = Sdo_ini + MvDebe.Oct + MvHaber.Oct
                 M_D = MvDebe.Oct: M_H = MvHaber.Oct
               Case 11
                 Sdo_ini = Sdo_ini + MvDebe.Nov + MvHaber.Nov
                 M_D = MvDebe.Nov: M_H = MvHaber.Nov
               Case 12
                 Sdo_ini = Sdo_ini + MvDebe.Dic + MvHaber.Dic
                 M_D = MvDebe.Dic: M_H = MvHaber.Dic
            End Select
               
        Next r
        
        If Sdo_ini <> 0 Then M_ue = 1
        If M_D <> 0 Then M_ue = 1
        If M_H <> 0 Then M_ue = 1
        Sdo_Fin = Sdo_ini
        Sdo_ini = Sdo_ini - M_D - M_H
        If M_ue = 1 Then
          Bcha1.AddItem Format(Val(CATMAY.B1), "####0") & Chr(9) & CATMAY.B2 & Chr(9) & Format(Sdo_ini, z1) _
                        & Chr(9) & Format(M_D, z1) & Chr(9) & Format(M_H, z1) & Chr(9) & Format(Sdo_Fin, z1) _
                        & Chr(9) & "C"
          If IsNumeric(CATMAY.B4) And Val(CATMAY.B4) > 0 Then
                        
                        PrAux mes_lim, Val(CATMAY.B4), Val(CATMAY.B5)
                        Else
                        
                        Bcha1.AddItem Format(1, "####0") & Chr(9) & "VARIOS" & Chr(9) & Format(Sdo_ini, z1) _
                        & Chr(9) & Format(M_D, z1) & Chr(9) & Format(M_H, z1) & Chr(9) & Format(Sdo_Fin, z1) _
                        & Chr(9) & "A"
                        
          End If
        End If
        
    Next L
       Rem Verifica_Mvtos
       Suma_Mvtos
       FinParcial
       Rem Verifica_Mvtos
    End Sub
Sub FinParcial()
    Rem Aqui termina
    Verifica_Mvtos
End Sub
Sub PrAux(mes_lim, Aux_ini As Long, Aux_Fin As Long)
    Dim Sdo_ini1 As Currency, M_D1 As Currency, M_H1 As Currency
    Dim M_ue1 As Integer, z11 As String, Sdo_Fin1 As Currency
    z11 = "#,##0.00;(#,##0.00)"
    For LL = Aux_ini To Aux_Fin
        
        M_ue1 = 0: Sdo_ini1 = 0: M_D1 = 0: M_H1 = 0: Sdo_Fin1 = 0
        Get 3, LL, CATAUX
        
       If (IsNumeric(CATAUX.C1)) And (Val(CATAUX.C1) > 0) Then
        Get 6, LL, MvDebeS: Get 7, LL, MvHaberS
        
        For RR = 0 To mes_lim
            Select Case RR
               Case 0
               
                Sdo_ini1 = Sdo_ini1 + MvDebeS.Inc + MvHaberS.Inc
                
               Case 1
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Ene + MvHaberS.Ene
                 M_D1 = MvDebeS.Ene: M_H1 = MvHaberS.Ene
               Case 2
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Feb + MvHaberS.Feb
                 M_D1 = MvDebeS.Feb: M_H1 = MvHaberS.Feb
               Case 3
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Mar + MvHaberS.Mar
                 M_D1 = MvDebeS.Mar: M_H1 = MvHaberS.Mar
               Case 4
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Abr + MvHaberS.Abr
                 M_D1 = MvDebeS.Abr: M_H1 = MvHaberS.Abr
               Case 5
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.May + MvHaberS.May
                 M_D1 = MvDebeS.May: M_H1 = MvHaberS.May
               Case 6
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Jun + MvHaberS.Jun
                 M_D1 = MvDebeS.Jun: M_H1 = MvHaberS.Jun
               Case 7
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Jul + MvHaberS.Jul
                 M_D1 = MvDebeS.Jul: M_H1 = MvHaberS.Jul
               Case 8
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Ago + MvHaberS.Ago
                 M_D1 = MvDebeS.Ago: M_H1 = MvHaberS.Ago
               Case 9
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Sep + MvHaberS.Sep
                
                 M_D1 = MvDebeS.Sep: M_H1 = MvHaberS.Sep
               Case 10
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Oct + MvHaberS.Oct
                 
                 M_D1 = MvDebeS.Oct: M_H1 = MvHaberS.Oct
               Case 11
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Nov + MvHaberS.Nov
                 M_D1 = MvDebeS.Nov: M_H1 = MvHaberS.Nov
               Case 12
                 Sdo_ini1 = Sdo_ini1 + MvDebeS.Dic + MvHaberS.Dic
                 M_D1 = MvDebeS.Dic: M_H1 = MvHaberS.Dic
            End Select
               
        Next RR
        
        If Sdo_ini1 <> 0 Then M_ue1 = 1
        If M_D1 <> 0 Then M_ue1 = 1
        If M_H1 <> 0 Then M_ue1 = 1
        Sdo_Fin1 = Sdo_ini1
        Sdo_ini1 = Sdo_ini1 - M_D1 - M_H1
        
        If M_ue1 = 1 Then
          
          Bcha1.AddItem Format(Val(CATAUX.C1), "####0") & Chr(9) & CATAUX.C2 & Chr(9) & Format(Sdo_ini1, z11) _
                        & Chr(9) & Format(M_D1, z11) & Chr(9) & Format(M_H1, z11) & Chr(9) & Format(Sdo_Fin1, z11) _
                        & Chr(9) & "A" & Chr(9) & LL
                        
        End If
        End If
    Next LL
       Rem Suma_Mvtos

End Sub

Private Sub Form_Load()
    Rem DIBUJA 12
    Balancha.Refresh
    Bcha1.Cols = 8: Bcha1.Rows = 2: Bcha1.FixedRows = 1
    Bcha1.Font = "Arial"
    F_Aum = 1
    Balancha.WindowState = 2
    Form_Resize
    Rem Tot_Act
    DIBUJA final
    
End Sub
Private Sub Form_Resize()
   If Balancha.WindowState <> 1 Then
      Bcha1.Height = ScaleHeight - 200
      Bcha1.Width = ScaleWidth - 200
      F_Aum = (Bcha1.Width - 400) / 9200
      ColDfn
   End If
End Sub
Sub ColDfn()
    Bcha1.FontWidth = 3 * F_Aum
    Bcha1.ColWidth(0) = 800 * F_Aum
    Bcha1.ColWidth(1) = 2300 * F_Aum
    Bcha1.ColWidth(2) = 1500 * F_Aum
    Bcha1.ColWidth(3) = 1500 * F_Aum
    Bcha1.ColWidth(4) = 1500 * F_Aum
    Bcha1.ColWidth(5) = 1500 * F_Aum
    Bcha1.ColWidth(6) = 300 * F_Aum
    Bcha1.ColWidth(7) = 300 * F_Aum
End Sub

Sub Suma_Mvtos()
   Dim Sum_a(5) As Currency
   For r = 1 To Bcha1.Rows - 1
       For G = 2 To Bcha1.Cols - 3
            If IsNumeric(Bcha1.TextMatrix(r, G)) Then
                Sum_a(G - 1) = Sum_a(G - 1) + Bcha1.TextMatrix(r, G)
            End If
       Next G
   Next r
   Bcha1.AddItem "" & Chr(9) & "Sumas" & Chr(9) & Format(Sum_a(1), "#,##0.00;(#,##0.00)") & Chr(9) & Format(Sum_a(2), "#,##0.00;(#,##0.00)") _
               & Chr(9) & Format(Sum_a(3), "#,##0.00;(#,##0.00)") & Chr(9) & Format(Sum_a(4), "#,##0.00;(#,##0.00)")
End Sub


Private Sub Verifica_Mvtos()
Rem GoTo INUTILIZA:
     Dim InAx As Long
     For Ww = 1 To Bcha1.Rows - 1
          Cla_ve = Bcha1.TextMatrix(Ww, 6)
          
          If Bcha1.TextMatrix(Ww, 6) = "C" Then
              MD = Bcha1.TextMatrix(Ww, 3): MH = Bcha1.TextMatrix(Ww, 4)
              Md_1 = 0: Mh_1 = 0
          End If
          
          If Cla_ve = "A" Then
                 
                 InAx = Ww
                 Cla_ve = Bcha1.TextMatrix(InAx, 6)
                 
                 Do Until Cla_ve <> "A"
                    Md_1 = Md_1 + Bcha1.TextMatrix(InAx, 3)
                    Mh_1 = Mh_1 + Bcha1.TextMatrix(InAx, 4)
                    InAx = InAx + 1
                    Cla_ve = Bcha1.TextMatrix(InAx, 6)
                 Loop
                 If MD <> Md_1 Then MD = Md_1: Bcha1.TextMatrix(Ww - 1, 3) = Format(MD, z1): Md_1 = 0: MD = 0
                 If MH <> Mh_1 Then MH = Mh_1: Bcha1.TextMatrix(Ww - 1, 4) = Format(MH, z1): Mh_1 = 0: MH = 0
                 Ww = InAx - 1
          End If
     Next Ww
INUTILIZA:
End Sub
