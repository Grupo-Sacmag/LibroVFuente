VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Balanza 
   Caption         =   "Balanza de comprobacion"
   ClientHeight    =   6528
   ClientLeft      =   96
   ClientTop       =   480
   ClientWidth     =   9432
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6528
   ScaleWidth      =   9432
   Begin MSFlexGridLib.MSFlexGrid Blz 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16108
      _ExtentY        =   10393
      _Version        =   393216
      BackColorBkg    =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Balanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_Aum As Double
Dim A_ltura As Long, A_ncho As Long, Lim_largo As Long, P_One As Long
Dim Tam_largo As Long, Tam_ancho As Long, Lim_ancho As Long, TeR_mina As Long
Dim Aj_te As Currency

Private Sub Form_Activate()
    BALANZON = 1
    'LIBROSV.Meses.Visible = True
    'LIBROSV.MayArBal.Visible = True
    'LIBROSV.MayArDia.Visible = True
   LIBROSV.Ci.Visible = False
End Sub
Sub Suma_Mvtos()
   Dim Sum_a(5) As Currency
   For r = 1 To Blz.Rows - 1
       For G = 2 To Blz.Cols - 1
            If Blz.TextMatrix(r, G) <> "" Then
                    Sum_a(G - 1) = Sum_a(G - 1) + Blz.TextMatrix(r, G)
            End If
       Next G
   Next r
   Blz.AddItem "" & Chr(9) & "Sumas" & Chr(9) & Format(Sum_a(1), "#,##0.00;(#,##0.00)") & Chr(9) & Format(Sum_a(2), "#,##0.00;(#,##0.00)") _
               & Chr(9) & Format(Sum_a(3), "#,##0.00;(#,##0.00)") & Chr(9) & Format(Sum_a(4), "#,##0.00;(#,##0.00)")
End Sub
Sub impre_titulos()
    Printer.Font.Bold = True
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim(Datos.D1))) / 2
    Printer.Print Datos.D1
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim(Balanza.Caption))) / 2
    Printer.Print Balanza.Caption
    Printer.CurrentX = (Tam_ancho - A_ncho) / 2
    TeR_mina = Printer.CurrentX
    For r = 0 To Blz.Cols - 1
         TeR_mina = TeR_mina + (Blz.ColWidth(r) * Aj_te)
         P_One = Printer.CurrentX + ((Blz.ColWidth(r) * Aj_te) - TextWidth(Blz.TextMatrix(0, r))) / 2
         Fin_Cuad = Printer.CurrentY + TextHeight(Blz.TextMatrix(r, L))
         Printer.Line (Printer.CurrentX + 50, Printer.CurrentY)-(TeR_mina + 50, Fin_Cuad), , B
         Printer.CurrentY = Fin_Cuad - TextHeight(Blz.TextMatrix(r, L))
         Printer.CurrentX = P_One
         Printer.Print Blz.TextMatrix(0, r);
         Printer.CurrentX = TeR_mina
    Next r
      Printer.Print
      Rem Printer.Print
      TeR_mina = 0
End Sub
Sub impre_Balnza()
    Tam_largo = Printer.ScaleHeight
    Tam_ancho = Printer.ScaleWidth
    A_ltura = 0: A_ncho = 0
    Lim_largo = 0
    Printer.Font.Size = Blz.Font.Size
    For r = 0 To Blz.Cols - 1
        Blz.Row = 0: Blz.Col = r
        A_ncho = A_ncho + Blz.CellWidth
    Next r
    For r = 0 To Blz.Rows - 1
        Blz.Row = r: Blz.Col = 0
        A_ltura = A_ltura + Blz.CellHeight
    Next r
    Rem ****** Si el ancho es mayor que el de la hoja se ajustan los valores
    If A_ncho > Tam_ancho Then
          Aj_te = ((Tam_ancho - 400) / A_ncho)
          Lim_largo = Tam_largo - 400
          Else
          Aj_te = 1
          Lim_largo = Tam_largo - 400
    End If
    Rem ****** Si el largo es mayor que el de la hoja se ajustan los valores
    If Tam_largo > A_ltura Then
          A_ncho = A_ncho * Aj_te
          Printer.CurrentY = (Tam_largo - A_ltura) / 2
          Else
          Printer.CurrentY = 400
    End If
    
    impre_titulos
    Printer.FontBold = False
    For r = 1 To Blz.Rows - 1
        Rem Blz.Row = r
        Printer.CurrentX = (Tam_ancho - A_ncho) / 2
        TeR_mina = Printer.CurrentX
        For L = 0 To Blz.Cols - 1
            TeR_mina = TeR_mina + (Blz.ColWidth(L) * Aj_te)
            Rem Blz.Col = L
            Select Case L
                Case 1
                    Rem izquierda
                    P_One = Printer.CurrentX + 55
                Case 0
                    Rem Cento
                    P_One = Printer.CurrentX + ((Blz.ColWidth(L) * Aj_te) - TextWidth(Blz.TextMatrix(r, L))) / 2
                Case 2, 3, 4, 5
                    Rem Derecha
                    P_One = Printer.CurrentX + (Blz.ColWidth(L) * Aj_te) - TextWidth(Blz.TextMatrix(r, L))
                Case Else
                    P_One = Printer.CurrentX
            End Select
            
            Fin_Cuad = Printer.CurrentY + TextHeight(Blz.TextMatrix(r, L))
            Printer.Line (Printer.CurrentX + 50, Printer.CurrentY)-(TeR_mina + 50, Fin_Cuad), , B
            Printer.CurrentY = Fin_Cuad - TextHeight(Blz.TextMatrix(r, L))
            Printer.CurrentX = P_One
            Printer.Print Blz.TextMatrix(r, L);
            Printer.CurrentX = TeR_mina
        Next L
        Printer.Print
        
        If Printer.CurrentY > Lim_largo Then
                 Printer.NewPage
        End If
    Next r
    Printer.EndDoc
    Printer.Orientation = 1
End Sub

Private Sub Form_GotFocus()
     LIBROSV.Meses.Visible = False
     LIBROSV.MayArBal.Visible = False
     LIBROSV.MayArDia.Visible = False

End Sub

Private Sub Form_Load()
    BALANZON = 1
    Balanza.Refresh
    Blz.Cols = 6: Blz.Rows = 2: Blz.FixedRows = 1
    Blz.Font = "Arial"
    F_Aum = 1
    Balanza.WindowState = 2
    Form_Resize
    Tot_Act
    DIBUJA final
    
End Sub
Sub DIBUJA(mes_lim)
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
    Blz.Clear
    Blz.Rows = 1
    Blz.Row = 0: Blz.Col = 0: Blz.CellFontBold = True: Blz.CellAlignment = 4: Blz.Text = "Cuenta"
    Blz.Row = 0: Blz.Col = 1: Blz.CellFontBold = True: Blz.CellAlignment = 4: Blz.Text = "Nombre"
    Blz.Row = 0: Blz.Col = 2: Blz.CellFontBold = True: Blz.CellAlignment = 4: Blz.Text = "Saldo " + Mm(mes_lim - 1)
    Blz.Row = 0: Blz.Col = 3: Blz.CellFontBold = True: Blz.CellAlignment = 4: Blz.Text = "Debe"
    Blz.Row = 0: Blz.Col = 4: Blz.CellFontBold = True: Blz.CellAlignment = 4: Blz.Text = "Haber"
    Blz.Row = 0: Blz.Col = 5: Blz.CellFontBold = True: Blz.CellAlignment = 4: Blz.Text = "Saldo " + Mm(mes_lim)
    Blz.FixedCols = 2
    Open Sub_dir + "CATMAY" For Random As 2 Len = Len(CATMAY)
    Open Sub_dir + "DEBE.GCS" For Random As 4 Len = Len(MvDebe)
    cm = LOF(4) / Len(MvDebe)
    Open Sub_dir + "HABER.GCS" For Random As 5 Len = Len(MvHaber)
    Dm = LOF(5) / Len(MvHaber)
    If Dm > cm Then cm = Dm
        Rem Blz.AddItem "Cuenta" & Chr(9) & "Nombre" & Chr(9) & "Saldo " + mm(Mes_Lim - 1) _
                                         & Chr(9) & "Debe" & Chr(9) & "Haber" & Chr(9) & "Saldo " + mm(Mes_Lim)

    For L = 1 To cm
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
          Blz.AddItem Format(Val(CATMAY.B1), "####0") & Chr(9) & CATMAY.B2 & Chr(9) & Format(Sdo_ini, z1) _
                        & Chr(9) & Format(M_D, z1) & Chr(9) & Format(M_H, z1) & Chr(9) & Format(Sdo_Fin, z1)
        End If
        
    Next L
       Suma_Mvtos
    End Sub
Sub Tot_Act()
  Dim NomArch As String, MiArchivo As String
  Close
On Error Resume Next
  Open Sub_dir + "DATOS" For Random As 1 Len = Len(Datos)
  cm = LOF(1) / Len(Datos)
If cm > 0 Then
  Get 1, 1, Datos
  
  MiArchivo = dir(Sub_dir + "DEBE.GCS")
  If MiArchivo <> "" Then Kill Sub_dir + "DEBE.GCS"
  MiArchivo = dir(Sub_dir + "HABER.GCS")
  If MiArchivo <> "" Then Kill Sub_dir + "HABER.GCS"
  MiArchivo = dir(Sub_dir + "DEBES.GCS")
  If MiArchivo <> "" Then Kill Sub_dir + "DEBES.GCS"
  MiArchivo = dir(Sub_dir + "HABERS.GCS")
  If MiArchivo <> "" Then Kill Sub_dir + "HABERS.GCS"
   
  Open Sub_dir + "DEBE.GCS" For Random As 4 Len = Len(MvDebe)
  Open Sub_dir + "HABER.GCS" For Random As 5 Len = Len(MvHaber)
  Open Sub_dir + "DEBES.GCS" For Random As 6 Len = Len(MvDebeS)
  Open Sub_dir + "HABERS.GCS" For Random As 7 Len = Len(MvHaberS)
   For r = 0 To 12
   
     NomArch = Trim(Datos.No_arch)
     Select Case r
         Case 0
            NomArch = Sub_dir + NomArch + "13"
         Case 1 To 9
            NomArch = Sub_dir + NomArch + "0" + LTrim(Str(r))
         Case 10 To 12
            NomArch = Sub_dir + NomArch + LTrim(Str(r))
     End Select
     MiArchivo = dir(NomArch)
     If MiArchivo <> "" Then
          final = r: Arc_FinaL = r
          Close 3
          Open NomArch For Random As 3 Len = Len(oper)
          Dm = LOF(3) / Len(oper)
          For F = 1 To Dm: Get 3, F, oper
                Select Case oper.identi
                   Case "B"
                    If Val(oper.impte) > 0 Then
                          Get 4, Val(oper.real), MvDebe
                          Select Case r
                             Case 0
                             MvDebe.Inc = MvDebe.Inc + Val(oper.impte)
                             Case 1
                             MvDebe.Ene = MvDebe.Ene + Val(oper.impte)
                             Case 2
                             MvDebe.Feb = MvDebe.Feb + Val(oper.impte)
                             Case 3
                             MvDebe.Mar = MvDebe.Mar + Val(oper.impte)
                             Case 4
                             MvDebe.Abr = MvDebe.Abr + Val(oper.impte)
                             Case 5
                             MvDebe.May = MvDebe.May + Val(oper.impte)
                             Case 6
                             MvDebe.Jun = MvDebe.Jun + Val(oper.impte)
                             Case 7
                             MvDebe.Jul = MvDebe.Jul + Val(oper.impte)
                             Case 8
                             MvDebe.Ago = MvDebe.Ago + Val(oper.impte)
                             Case 9
                             MvDebe.Sep = MvDebe.Sep + Val(oper.impte)
                             Case 10
                             MvDebe.Oct = MvDebe.Oct + Val(oper.impte)
                             Case 11
                             MvDebe.Nov = MvDebe.Nov + Val(oper.impte)
                             Case 12
                             MvDebe.Dic = MvDebe.Dic + Val(oper.impte)
                          End Select
                          Put 4, Val(oper.real), MvDebe

                             
                    Else
                         Get 5, Val(oper.real), MvHaber
                         Select Case r
                             Case 0
                             
                             MvHaber.Inc = MvHaber.Inc + Val(oper.impte)
                             
                             Case 1
                             MvHaber.Ene = MvHaber.Ene + Val(oper.impte)
                             Case 2
                             MvHaber.Feb = MvHaber.Feb + Val(oper.impte)
                             Case 3
                             MvHaber.Mar = MvHaber.Mar + Val(oper.impte)
                             Case 4
                             MvHaber.Abr = MvHaber.Abr + Val(oper.impte)
                             Case 5
                             MvHaber.May = MvHaber.May + Val(oper.impte)
                             Case 6
                             MvHaber.Jun = MvHaber.Jun + Val(oper.impte)
                             Case 7
                             MvHaber.Jul = MvHaber.Jul + Val(oper.impte)
                             Case 8
                             MvHaber.Ago = MvHaber.Ago + Val(oper.impte)
                             Case 9
                             MvHaber.Sep = MvHaber.Sep + Val(oper.impte)
                             Case 10
                             MvHaber.Oct = MvHaber.Oct + Val(oper.impte)
                             Case 11
                             MvHaber.Nov = MvHaber.Nov + Val(oper.impte)
                             Case 12
                             MvHaber.Dic = MvHaber.Dic + Val(oper.impte)
                          End Select
                          Put 5, Val(oper.real), MvHaber

                    End If
                    Case "C"
                    If Val(oper.impte) > 0 Then
                          Get 6, Val(oper.CTA), MvDebeS
                          Select Case r
                             Case 0
                             
                             MvDebeS.Inc = MvDebeS.Inc + Val(oper.impte)
                             
                             Case 1
                             MvDebeS.Ene = MvDebeS.Ene + Val(oper.impte)
                             Case 2
                             MvDebeS.Feb = MvDebeS.Feb + Val(oper.impte)
                             Case 3
                             MvDebeS.Mar = MvDebeS.Mar + Val(oper.impte)
                             Case 4
                             MvDebeS.Abr = MvDebeS.Abr + Val(oper.impte)
                             Case 5
                             MvDebeS.May = MvDebeS.May + Val(oper.impte)
                             Case 6
                             MvDebeS.Jun = MvDebeS.Jun + Val(oper.impte)
                             Case 7
                             MvDebeS.Jul = MvDebeS.Jul + Val(oper.impte)
                             Case 8
                             MvDebeS.Ago = MvDebeS.Ago + Val(oper.impte)
                             Case 9
                             MvDebeS.Sep = MvDebeS.Sep + Val(oper.impte)
                             Case 10
                             MvDebeS.Oct = MvDebeS.Oct + Val(oper.impte)
                             Case 11
                             MvDebeS.Nov = MvDebeS.Nov + Val(oper.impte)
                             Case 12
                             MvDebeS.Dic = MvDebeS.Dic + Val(oper.impte)
                          End Select
                          Put 6, Val(oper.CTA), MvDebeS

                             
                    Else
                         Get 7, Val(oper.CTA), MvHaberS
                         Select Case r
                             Case 0
                             MvHaberS.Inc = MvHaberS.Inc + Val(oper.impte)
                             Case 1
                             MvHaberS.Ene = MvHaberS.Ene + Val(oper.impte)
                             Case 2
                             MvHaberS.Feb = MvHaberS.Feb + Val(oper.impte)
                             Case 3
                             MvHaberS.Mar = MvHaberS.Mar + Val(oper.impte)
                             Case 4
                             MvHaberS.Abr = MvHaberS.Abr + Val(oper.impte)
                             Case 5
                             MvHaberS.May = MvHaberS.May + Val(oper.impte)
                             Case 6
                             MvHaberS.Jun = MvHaberS.Jun + Val(oper.impte)
                             Case 7
                             MvHaberS.Jul = MvHaberS.Jul + Val(oper.impte)
                             Case 8
                             MvHaberS.Ago = MvHaberS.Ago + Val(oper.impte)
                             Case 9
                             MvHaberS.Sep = MvHaberS.Sep + Val(oper.impte)
                             Case 10
                             MvHaberS.Oct = MvHaberS.Oct + Val(oper.impte)
                             Case 11
                             MvHaberS.Nov = MvHaberS.Nov + Val(oper.impte)
                             Case 12
                             MvHaberS.Dic = MvHaberS.Dic + Val(oper.impte)
                          End Select
                          Put 7, Val(oper.CTA), MvHaberS

                    End If
                End Select
          
          Next F
       End If
          Close 3
     
     MiArchivo = ""
     NomArch = ""
   Next r
   Close
   Else
   Close

   MsgBox "En este subdirectorio no existen " & Chr(13) & _
            "datos para continuar elija la opcion Cambio de subdirectorio ", vbCritical
   End If
 
End Sub
Private Sub Form_Resize()
   If Balanza.WindowState <> 1 Then
      Blz.Height = ScaleHeight - 200
      Blz.Width = ScaleWidth - 200
      F_Aum = (Blz.Width - 400) / 9200
      ColDfn
   End If
End Sub
Sub ColDfn()
    Blz.FontWidth = 3 * F_Aum
    Rem Blz.RowHeight(0) = Blz.FontWidth
    Blz.ColWidth(0) = 800 * F_Aum
    Blz.ColWidth(1) = 2300 * F_Aum
    Blz.ColWidth(2) = 1500 * F_Aum
    Blz.ColWidth(3) = 1500 * F_Aum
    Blz.ColWidth(4) = 1500 * F_Aum
    Blz.ColWidth(5) = 1500 * F_Aum
    Rem Balanza.Refresh
    Rem Blz.FixedRows = 1
End Sub
