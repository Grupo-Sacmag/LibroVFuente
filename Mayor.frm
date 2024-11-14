VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Mayor 
   Caption         =   "LIBRO MAYOR"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9765
   Begin MSFlexGridLib.MSFlexGrid MAYOR1 
      Height          =   5892
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9132
      _ExtentX        =   16113
      _ExtentY        =   10398
      _Version        =   393216
   End
End
Attribute VB_Name = "Mayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_D As Currency, M_H As Currency, Sdo_ini As Currency, In_Gsa As Integer
Dim r As Long, Tamno As Currency, HIni As Long, Nom_Cta As String, F_olioM1 As Integer

Private Sub Form_Activate()
  BALANZON = 5
  LIBROSV.Ci.Visible = True
End Sub

Private Sub Form_Load()
   Dim Cue_tota As Integer
   BALANZON = 5
   Mayor.Caption = Mayor.Caption + " DE " + Datos.a_o
   Mayor.Refresh
   Mayor.WindowState = 2
   MAYOR1.Clear
   MAYOR1.Cols = 7: F_olioM1 = 0
   MAYOR1.Rows = 1
   MAYOR1.Row = 0: Tamno = 0
    MAYOR1.Row = 0: MAYOR1.Col = 0: MAYOR1.ColWidth(0) = 1200: Tamno = Tamno + 1200: MAYOR1.CellFontBold = True: MAYOR1.CellAlignment = 4: MAYOR1.Text = "MES"
    MAYOR1.Row = 0: MAYOR1.Col = 1: MAYOR1.ColWidth(1) = 1200: Tamno = Tamno + 1200: MAYOR1.CellFontBold = True: MAYOR1.CellAlignment = 4: MAYOR1.Text = "No."
    MAYOR1.Row = 0: MAYOR1.Col = 2: MAYOR1.ColWidth(2) = 2800: Tamno = Tamno + 2800: MAYOR1.CellFontBold = True: MAYOR1.CellAlignment = 4: MAYOR1.Text = "Cuenta"
    MAYOR1.Row = 0: MAYOR1.Col = 3: MAYOR1.ColWidth(3) = 3200: Tamno = Tamno + 3200: MAYOR1.CellFontBold = True: MAYOR1.CellAlignment = 4: MAYOR1.Text = "CONCEPTO"
    MAYOR1.Row = 0: MAYOR1.Col = 4: MAYOR1.ColWidth(4) = 1200: Tamno = Tamno + 1200: MAYOR1.CellFontBold = True: MAYOR1.CellAlignment = 4: MAYOR1.Text = "Debe"
    MAYOR1.Row = 0: MAYOR1.Col = 5: MAYOR1.ColWidth(5) = 1200: Tamno = Tamno + 1200: MAYOR1.CellFontBold = True: MAYOR1.CellAlignment = 4: MAYOR1.Text = "Haber"
    MAYOR1.Row = 0: MAYOR1.Col = 6: MAYOR1.ColWidth(6) = 1200: Tamno = Tamno + 1200: MAYOR1.CellFontBold = True: MAYOR1.CellAlignment = 4: MAYOR1.Text = "Saldo "
    Rem  Mm(mes_lim)
    MAYOR1.Width = Tamno + 400
    Close
    Open Sub_dir + "CATMAY" For Random As 2 Len = Len(CATMAY)
    w1 = LOF(2) / Len(CATMAY)
    Open Sub_dir + "DEBE.GCS" For Random As 4 Len = Len(MvDebe)
    Cm = LOF(4) / Len(MvDebe)
    Open Sub_dir + "HABER.GCS" For Random As 5 Len = Len(MvHaber)
    Dm = LOF(5) / Len(MvHaber)
   
    If Dm > Cm Then Cm = Dm
        Rem Blz.AddItem "Cuenta" & Chr(9) & "Nombre" & Chr(9) & "Saldo " + mm(Mes_Lim - 1) _
                                         & Chr(9) & "Debe" & Chr(9) & "Haber" & Chr(9) & "Saldo " + mm(Mes_Lim)
    In_Gsa = 0
    For L = 1 To w1
        If In_Gsa = 1 Then
            F_olioM1 = F_olioM1 + 1
            Mov_Cierre
        End If
        
        M_ue = 0: Sdo_ini = 0: M_D = 0: M_H = 0: Sdo_Fin = 0
        Get 2, L, CATMAY
        Get 4, L, MvDebe: Get 5, L, MvHaber
        
        If (Val(CATMAY.B1) > 0) Then
        
        
        For r = 0 To 12
        
            Select Case r
               Case 0
                Sdo_ini = Sdo_ini + MvDebe.Inc + MvHaber.Inc
                M_D = MvDebe.Inc: M_H = MvHaber.Inc
                If (M_D <> 0) Or (M_H <> 0) Then
                  Revela
                End If
               Case 1
                 Sdo_ini = Sdo_ini + MvDebe.Ene + MvHaber.Ene
                 M_D = MvDebe.Ene: M_H = MvHaber.Ene
                 If (M_D <> 0) Or (M_H <> 0) Then
                    Revela
                 End If
                 
               Case 2
                 Sdo_ini = Sdo_ini + MvDebe.Feb + MvHaber.Feb
                 M_D = MvDebe.Feb: M_H = MvHaber.Feb
                 If (M_D <> 0) Or (M_H <> 0) Then
                            Revela
                 End If
               Case 3
                 Sdo_ini = Sdo_ini + MvDebe.Mar + MvHaber.Mar
                 M_D = MvDebe.Mar: M_H = MvHaber.Mar
                 
                 If (M_D <> 0) Or (M_H <> 0) Then
                       Revela
                 End If

               Case 4
                 Sdo_ini = Sdo_ini + MvDebe.Abr + MvHaber.Abr
                 M_D = MvDebe.Abr: M_H = MvHaber.Abr
                 If (M_D <> 0) Or (M_H <> 0) Then
                    Revela
                 End If
               Case 5
                 Sdo_ini = Sdo_ini + MvDebe.May + MvHaber.May
                 M_D = MvDebe.May: M_H = MvHaber.May
                 If (M_D <> 0) Or (M_H <> 0) Then
                        Revela
                 End If

               Case 6
                 Sdo_ini = Sdo_ini + MvDebe.Jun + MvHaber.Jun
                 M_D = MvDebe.Jun: M_H = MvHaber.Jun
                 If (M_D <> 0) Or (M_H <> 0) Then
                            Revela
                 End If

               Case 7
                 Sdo_ini = Sdo_ini + MvDebe.Jul + MvHaber.Jul
                 M_D = MvDebe.Jul: M_H = MvHaber.Jul
                 If (M_D <> 0) Or (M_H <> 0) Then
                        Revela
                 End If

               Case 8
                 Sdo_ini = Sdo_ini + MvDebe.Ago + MvHaber.Ago
                 M_D = MvDebe.Ago: M_H = MvHaber.Ago
                 If (M_D <> 0) Or (M_H <> 0) Then
                       Revela
                 End If
               Case 9
                 Sdo_ini = Sdo_ini + MvDebe.Sep + MvHaber.Sep
                 M_D = MvDebe.Sep: M_H = MvHaber.Sep
                 If (M_D <> 0) Or (M_H <> 0) Then
                        Revela
                 End If

               Case 10
                 Sdo_ini = Sdo_ini + MvDebe.Oct + MvHaber.Oct
                 M_D = MvDebe.Oct: M_H = MvHaber.Oct
                 If (M_D <> 0) Or (M_H <> 0) Then
                        Revela
                 End If

               Case 11
                 Sdo_ini = Sdo_ini + MvDebe.Nov + MvHaber.Nov
                 M_D = MvDebe.Nov: M_H = MvHaber.Nov
                 If (M_D <> 0) Or (M_H <> 0) Then
                        Revela
                 End If
               Case 12
                 Sdo_ini = Sdo_ini + MvDebe.Dic + MvHaber.Dic
                 
                 M_D = MvDebe.Dic: M_H = MvHaber.Dic
                 If (M_D <> 0) Or (M_H <> 0) Then
                    Revela
                 End If

            End Select
               
        Next r
        
    End If
      If In_Gsa = 2 Then
           F_olioM1 = F_olioM1 + 1
           In_Gsa = 0
      End If
    Next L
    Rem RDOEJ.Show 1
End Sub
Sub Revela()
       MAYOR1.AddItem Mm(r) & Chr(9) & CATMAY.B1 & Chr(9) & CATMAY.B2 & Chr(9) & _
            "Movimientos del mes" & Chr(9) & Format(M_D, z1) & Chr(9) & _
            Format(M_H, z1) & Chr(9) & Format(Sdo_ini, z1)
       If CATMAY.B1 > 4000 Then
            In_Gsa = 1
            Else
            In_Gsa = 2
       End If
      
End Sub
Sub imp_sora()
    impre_titulos
    impre_diario
End Sub
Sub impre_titulos()
    Dim M_der As Long, M_Izq As Long, M_derb As Long
    Dim Anc_col As Long, Anc_txt As Long
    Printer.PaperSize = 5
    Printer.Orientation = 2
    Tam_largo = Printer.ScaleHeight
    Tam_ancho = Printer.ScaleWidth
    A_ltura = 0: A_ncho = 0
    Lim_largo = 0
    Rem Printer.Font.Size = Blz.Font.Size
    Printer.Font.Bold = True
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim(Datos.D1))) / 2
    Printer.Print Datos.D1
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim(Diario.Caption))) / 2
    Printer.Print Diario.Caption; " R.F.C. "; Datos.D2
    Printer.CurrentX = (Tam_ancho - A_ncho) / 2
    TeR_mina = Printer.CurrentX: Aj_te = 1
    F_olio = F_olio + 1
    Printer.CurrentX = (Tam_ancho - 2000)
    Printer.Print "FOLIO "; Format(F_olio, "#,##0")
    Printer.Print
    M_der = 0: M_Izq = 1200: M_derb = Printer.CurrentY
    For r = 0 To 13
    
         M_der = M_Izq + (Diario.Diario1.ColWidth(r) * Aj_te)
         'TeR_mina = TeR_mina + (Diario.Diario1.ColWidth(r) * Aj_te)
         
         Rem P_One = M_Izq + Int((Diario.Diario1.ColWidth(r) * Aj_te) - TextWidth(Left(Trim(Diario.Diario1.TextMatrix(0, r)), 11)) / 2)
         Anc_col = Diario.Diario1.ColWidth(r): Anc_txt = TextWidth(Left(Trim(Diario.Diario1.TextMatrix(0, r)), 11))
         Anc_col = (Anc_col - Anc_txt) / 2
         P_One = M_Izq + Anc_col
         
         Rem Punto final Bajo del cuadrado *****************************************************
         
         M_derb = Printer.CurrentY + TextHeight(Diario.Diario1.TextMatrix(0, r))
         Rem Es el cuadro ***********************************************************************
         Printer.CurrentX = M_Izq
         
         Printer.Line (Printer.CurrentX, Printer.CurrentY)-(M_der, M_derb), , B
         
         Rem recupera la altura inicial *********************************************************
         

         Printer.CurrentY = M_derb - TextHeight(Diario.Diario1.TextMatrix(0, r))
         
         
         Printer.CurrentX = P_One
         
         Printer.Print Left(Trim(Diario.Diario1.TextMatrix(0, r)), 11);
         
         M_Izq = M_der: Printer.CurrentX = M_Izq
         
    Next r
      Printer.Print
      Printer.Print
      TeR_mina = 0
      'Printer.PaperSize = 1
      ' Printer.Orientation = 1

End Sub

Sub impre_diario()
    Dim M_der1 As Long, M_Izq1 As Long, M_derb1 As Long, W2 As Long
    Dim Anc_col1 As Long, Anc_txt1 As Long, CuenTa1 As Integer
    CuenTa1 = 0
    For W2 = 1 To Diario.Diario1.Rows - 1
       Aj_te = 1
       M_der1 = 0: M_Izq1 = 1200: M_derb1 = Printer.CurrentY: CuenTa = 0
       
       For r = 0 To 13
         Diario.Diario1.Row = W2
         M_der1 = M_Izq1 + (Diario.Diario1.ColWidth(r) * Aj_te)
         Anc_col1 = Diario.Diario1.ColWidth(r): Anc_txt1 = TextWidth(Diario.Diario1.TextMatrix(W2, r))
         Anc_col1 = (Anc_col1 - Anc_txt1)
         Select Case r
            Case 0, 2
             P_One1 = M_Izq1 + 20
            Case 1
             P_One1 = M_Izq1 + ((Anc_col1) / 2)
            Case Else
                P_One1 = M_der1 - (Anc_txt1 + 40)
         End Select
         Rem Punto final Bajo del cuadrado *****************************************************
         
         M_derb1 = Printer.CurrentY + TextHeight(Diario.Diario1.TextMatrix(W2, r))
         Rem Es el cuadro ***********************************************************************
         Printer.CurrentX = M_Izq1
         
         Printer.Line (Printer.CurrentX, Printer.CurrentY)-(M_der1, M_derb1), , B
         
         Rem recupera la altura inicial *********************************************************
         

         Printer.CurrentY = M_derb1 - TextHeight(Diario.Diario1.TextMatrix(W2, r))
         
         
         Printer.CurrentX = P_One1
         
         Printer.Print Diario.Diario1.TextMatrix(W2, r);
         
         M_Izq1 = M_der1: Printer.CurrentX = M_Izq1
         
    Next r
      Printer.Print
      CuenTa1 = CuenTa1 + 1
       
         If (CuenTa1 >= 49) Then
              If W2 < (Diario.Diario1.Rows - 1) Then
                Printer.NewPage
                impre_titulos
                CuenTa1 = 0
              End If
         End If
    Next W2
      TeR_mina = 0
      Printer.EndDoc
      Printer.PaperSize = 1
    Printer.Orientation = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
     LIBROSV.Ci.Visible = False
     
End Sub
Sub tit_mayor()
    Dim M_der As Long, M_Izq As Long, M_derb As Long
    Dim Anc_col As Long, Anc_txt As Long
    Tam_largo = Printer.ScaleHeight
    Tam_ancho = Printer.ScaleWidth
    A_ltura = 0: A_ncho = 0
    Lim_largo = 0
    Rem Printer.Font.Size = Blz.Font.Size
    Printer.CurrentY = Tam_largo / 3
    Printer.Font.Bold = True
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim(Datos.D1))) / 2
    Printer.Print Datos.D1
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim(Mayor.Caption))) / 2
    Printer.Print Mayor.Caption
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim("R.F.C:" + Datos.D2))) / 2
    Printer.Print "R.F.C:"; Datos.D2;
    Printer.CurrentX = (Tam_ancho - A_ncho) / 2
    TeR_mina = Printer.CurrentX: Aj_te = 1
    F_olioM = F_olioM + 1
    Printer.CurrentX = (Tam_ancho - 2000)
    Printer.Print "FOLIO "; Format(F_olioM, "#,##0")
    Printer.Print
    Printer.CurrentX = (Printer.ScaleWidth - TextWidth(Trim(Nom_Cta))) / 2
    Printer.Print Nom_Cta
    Printer.Print
    Printer.Line (1800, Printer.CurrentY)-(Printer.ScaleWidth - 1000, Printer.CurrentY + 50), , BF
    M_der = 0: M_Izq = 1800: M_derb = Printer.CurrentY
    Printer.Print
    TeR_mina = 0
    For r = 0 To 6
         Select Case r
            Case 0, 3, 4, 5, 6
                M_der = M_Izq + (MAYOR1.ColWidth(r) * Aj_te)
                Anc_col = MAYOR1.ColWidth(r): Anc_txt = TextWidth(Left(Trim(MAYOR1.TextMatrix(0, r)), 11))
                Anc_col = (Anc_col - Anc_txt) / 2
                P_One = M_Izq + Anc_col
                M_derb = Printer.CurrentY + TextHeight(MAYOR1.TextMatrix(0, r))
                Printer.CurrentX = M_Izq
                Printer.Line (Printer.CurrentX, Printer.CurrentY)-(M_der, M_derb), , B
                Printer.CurrentY = M_derb - TextHeight(MAYOR1.TextMatrix(0, r))
                Printer.CurrentX = P_One
                Printer.Print Left(Trim(MAYOR1.TextMatrix(0, r)), 11);
                M_Izq = M_der: Printer.CurrentX = M_Izq
           Case 1, 2
              Rem NADA
        End Select
     Next r
     Printer.Print
     Rem Printer.Print

End Sub

Sub Impr_may()
    Dim M_der1 As Long, M_Izq1 As Long, M_derb1 As Long, W2 As Long
    Dim Anc_col1 As Long, Anc_txt1 As Long, CuenTa1 As Integer, Responde
    Dim Cam_bio As Integer
    Responde = MsgBox("Esta listo para imprimir el mayor, necesita hojas tamaño carta ", vbYesNo, "Impresion Libro Mayor")
    If Responde = vbYes Then
        Cam_bio = MAYOR1.TextMatrix(1, 1)
        Printer.FontBold = True
        Nom_Cta = Trim(MAYOR1.TextMatrix(1, 1)) + " " + Trim(MAYOR1.TextMatrix(1, 2))
        tit_mayor
        Printer.CurrentX = (Tam_ancho - 2000)
        M_der = 0: M_Izq = 1800: M_derb = Printer.CurrentY
        TeR_mina = 0
        For W2 = 1 To MAYOR1.Rows - 1
           If Cam_bio <> MAYOR1.TextMatrix(W2, 1) Then
             Printer.Print
             Printer.Line (1800, Printer.CurrentY)-(Printer.ScaleWidth - 1000, Printer.CurrentY + 50), , BF
             Printer.NewPage
             Cam_bio = MAYOR1.TextMatrix(W2, 1)
             Nom_Cta = Trim(MAYOR1.TextMatrix(W2, 1)) + " " + Trim(MAYOR1.TextMatrix(W2, 2))
             tit_mayor
             CuenTa1M = 0
        End If
        Aj_te = 1
        M_der1 = 0: M_Izq1 = 1800: M_derb1 = Printer.CurrentY: CuenTa = 0
       For r = 0 To 6
         MAYOR1.Row = W2
         M_der1 = M_Izq1 + (MAYOR1.ColWidth(r) * Aj_te)
         Anc_col1 = MAYOR1.ColWidth(r): Anc_txt1 = TextWidth(MAYOR1.TextMatrix(W2, r))
         Anc_col1 = (Anc_col1 - Anc_txt1)
         Select Case r
            Case 0, 3
             P_One1 = M_Izq1 + 20
            Case 1, 2
             P_One1 = M_Izq1 + ((Anc_col1) / 2)
            Case Else
                P_One1 = M_der1 - (Anc_txt1 + 40)
         End Select
         Rem Punto final Bajo del cuadrado *****************************************************
         Select Case r
             Case 0, 3, 4, 5, 6
                M_derb1 = Printer.CurrentY + TextHeight(MAYOR1.TextMatrix(W2, r))
                Rem Es el cuadro ***********************************************************************
                Printer.CurrentX = M_Izq1
                Printer.Line (Printer.CurrentX, Printer.CurrentY)-(M_der1, M_derb1), , B
                Rem recupera la altura inicial *********************************************************
                Printer.CurrentY = M_derb1 - TextHeight(MAYOR1.TextMatrix(W2, r))
                Printer.CurrentX = P_One1
                Printer.Print MAYOR1.TextMatrix(W2, r);
                M_Izq1 = M_der1: Printer.CurrentX = M_Izq1
            Case 1, 2
                Rem 'Nada ********************
        End Select
    Next r
      Printer.Print
      Rem CuenTa1M = CuenTa1M + 1
       
         Rem If CuenTa1M > 53 Then
                Rem Printer.NewPage
                
                'impre_titulos
                Rem CuenTa1M = 0
         Rem End If
         
    Next W2
      Printer.NewPage
      Nom_Cta = "Resultado del ejercicio "
      RDOEJ.Suma_res
      Rem RDOEJ.Show 1
      tit_mayor
      Printer.Print
      Rem Printer.CurrentX = (Tam_ancho - 2000)
      M_der = 0: M_Izq = 1800: M_derb = Printer.CurrentY
      TeR_mina = 0
      
      For W2 = 1 To RDOEJ.Res1.Rows - 1
        
        RDOEJ.Res1.Row = W2
        Aj_te = 1
        M_der1 = 0: M_Izq1 = 1800: M_derb1 = Printer.CurrentY: CuenTa = 0

        For r = 0 To 6
         M_der1 = M_Izq1 + (RDOEJ.Res1.ColWidth(r) * Aj_te)
         Anc_col1 = RDOEJ.Res1.ColWidth(r): Anc_txt1 = TextWidth(RDOEJ.Res1.TextMatrix(W2, r))
         Anc_col1 = (Anc_col1 - Anc_txt1)
         Anc_col1 = (Anc_col1 - Anc_txt1)
         Select Case r
            Case 0, 3
             P_One1 = M_Izq1 + 20
            Case 1, 2
             P_One1 = M_Izq1 + ((Anc_col1) / 2)
            Case Else
                P_One1 = M_der1 - (Anc_txt1 + 40)
         End Select
           Select Case r
             Case 0, 3, 4, 5, 6
                M_derb1 = Printer.CurrentY + TextHeight(RDOEJ.Res1.TextMatrix(W2, r))
                Rem Es el cuadro ***********************************************************************
                Printer.CurrentX = M_Izq1
                Printer.Line (Printer.CurrentX, Printer.CurrentY)-(M_der1, M_derb1), , B
                Rem recupera la altura inicial *********************************************************
                Printer.CurrentY = M_derb1 - TextHeight(RDOEJ.Res1.TextMatrix(W2, r))
                Printer.CurrentX = P_One1
                Printer.Print RDOEJ.Res1.TextMatrix(W2, r);
                M_Izq1 = M_der1: Printer.CurrentX = M_Izq1
            Case 1, 2
                Rem 'Nada ********************
        End Select
     Next r
       Printer.Print
      Next W2
      TeR_mina = 0
      Printer.EndDoc
      Printer.PaperSize = 1
      Printer.Orientation = 1
      Else
      MsgBox "Puedes imprimir mas Tarde, pasa el Diario por pantalla y luego, imprimes el Libro Mayor", vbCritical, "Impresion Libro Mayor pendiente"
End If
End Sub
Sub Mov_Cierre()
      Dim Sdo_Res As Currency, Ren_ex As Long
      Ren_ex = MAYOR1.Rows - 1
    'If In_Gsa = 1 Then
      Sdo_Res = MAYOR1.TextMatrix(Ren_ex, 6)
      If Sdo_Res > 0 Then
          MAYOR1.AddItem "DICIEMBRE" & Chr(9) & CATMAY.B1 & Chr(9) & CATMAY.B2 & Chr(9) & "Traspaso a Resultado del Ejercicio" _
                                     & Chr(9) & "" & Chr(9) & Format((Sdo_Res * -1), z1) & Chr(9) & Format(0, z1)
          RDOEJ.Res1.AddItem "DICIEMBRE" & Chr(9) & MAYOR1.TextMatrix(Ren_ex, 1) & Chr(9) & MAYOR1.TextMatrix(Ren_ex, 2) & Chr(9) & "Del Folio " + Str(F_olioM1) _
                                     & Chr(9) & Format(Sdo_Res, z1) & Chr(9) & "" & Chr(9) & Format(0, z1)
                           
          In_Gsa = 0
          Else
          MAYOR1.AddItem "DICIEMBRE" & Chr(9) & CATMAY.B1 & Chr(9) & CATMAY.B2 & Chr(9) & "Traspaso a Resultado del Ejercicio" _
                                     & Chr(9) & Format((Sdo_Res * -1), z1) & Chr(9) & "" & Chr(9) & Format(0, z1)
          RDOEJ.Res1.AddItem "DICIEMBRE" & Chr(9) & MAYOR1.TextMatrix(Ren_ex, 1) & Chr(9) & MAYOR1.TextMatrix(Ren_ex, 2) & Chr(9) & "Del Folio " + Str(F_olioM1) _
                                     & Chr(9) & "" & Chr(9) & Format(Sdo_Res, z1) & Chr(9) & Format(0, z1)
                           
          In_Gsa = 0
      End If
   'End If
End Sub
