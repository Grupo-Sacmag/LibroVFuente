VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Diario 
   Caption         =   "Actualización Libros"
   ClientHeight    =   4875
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   9180
   Icon            =   "LIBROS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   9180
   Begin MSFlexGridLib.MSFlexGrid Diario1 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   9120
      TabIndex        =   1
      Top             =   0
      Width           =   9180
   End
End
Attribute VB_Name = "Diario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nom_bre As String
Dim Q1 As Long, HOJA As Single, Hj_comp As Single
Sub Suma_Deb()
    Dim sumad As Currency
    Diario1.Rows = Diario1.Rows + 2
    Diario1.Row = Diario1.Rows - 2: Diario1.Col = 2: Diario1.CellFontBold = True: Diario1.CellAlignment = 8
    Diario1.TextMatrix(Diario1.Rows - 2, 2) = "Suma Debe :":
    Diario1.Row = Diario1.Rows - 1: Diario1.Col = 2: Diario1.CellFontBold = True: Diario1.CellAlignment = 8
    Diario1.TextMatrix(Diario1.Rows - 1, 2) = "Suma Haber :"
    sumad = 0: sumah = 0
  For I = 3 To (Diario1.Cols - 1)
    For r = 1 To Diario1.Rows - 3
        If IsNumeric(Diario1.TextMatrix(r, I)) Then
                If Diario1.TextMatrix(r, I) >= 0 Then
                   sumad = sumad + Diario1.TextMatrix(r, I)
                End If
                If Diario1.TextMatrix(r, I) < 0 Then
                   sumah = sumah + Diario1.TextMatrix(r, I)
                End If
                
        End If
    Next r
    Diario1.TextMatrix(Diario1.Rows - 2, I) = Format(sumad, "###,###,##0.00;(###,###,##0.00)"): sumad = 0
    Diario1.TextMatrix(Diario1.Rows - 1, I) = Format(sumah, "#,##0.00;(###,###,##0.00)"): sumah = 0
  Next I
End Sub



Private Sub Form_Activate()
    BALANZON = 2
    LIBROSV.Ci.Visible = False
End Sub

Private Sub Form_Load()
  Diario.Refresh
  BALANZON = 2
  Diario.WindowState = 2
  final = 12
  Ape_Operaciones
If Cm > 1 Then
  Ide_Mayor
  Suma_Deb
  Diario1.FixedCols = 3
  Diario.Caption = "Movimientos Libro Diario del mes de " + Trim(Mm(final)) + " de " + Datos.a_o
  Else
  Rem nada
End If
End Sub


Private Sub Form_Resize()
   Diario1.Height = ScaleHeight
   Diario1.Width = ScaleWidth
End Sub
Sub Local_iza(Nom_bre1)
  
  For I = 0 To final
     Select Case I
      Case 0
         Nom_bre = Trim(Nom_bre1) + "13"
         MiArchivo = dir(Sub_dir + Nom_bre)
         If MiArchivo = "" Then
             MsgBox "No existe archivo de incorporacion no es posible continuar" & Chr(13) & _
                     "Posiblemente no es el subdirectorio adecuado", vbCritical
                     
             Exit Sub
             final = -1
         End If
    
      Case 1 To 9
         Nom_bre = Trim(Nom_bre1) + "0" + Trim(Str(I))
         MiArchivo = dir(Sub_dir + Nom_bre)
         
         If MiArchivo = "" Then
             I = I - 1
             Nom_bre = Nom_bre1 + "0" + Trim(Str(I))
             final = I
             Exit Sub
         End If
      Case 10 To 12
         Nom_bre = Trim(Nom_bre1) + Trim(Str(I))
         MiArchivo = dir(Sub_dir + Nom_bre)
         If MiArchivo = "" Then
            If (I - 1) < 10 Then
             I = I - 1
             Nom_bre = Trim(Nom_bre1) + "0" + Trim(Str(I))
             final = I
             Exit Sub
             Else
             I = I - 1
             Nom_bre = Trim(Nom_bre1) + Trim(Str(I))
             final = I
             Exit Sub
            End If
         End If
     End Select
  Next I
 
End Sub
Sub Ape_Operaciones()
  On Error Resume Next
    Close
    Open Sub_dir + "DATOS" For Random As 1 Len = Len(Datos)
    Get 1, 1, Datos
    Cm = LOF(1) / Len(Datos)
 If Cm > 0 Then
    Open Sub_dir + "CATMAY" For Random As 2 Len = Len(CATMAY)
    Dm = LOF(2) / Len(CATMAY)
    Local_iza Trim(Datos.No_arch)
    Open Sub_dir + Nom_bre For Random As 3 Len = Len(oper)
    Cm = LOF(3) / Len(oper)
    
    Diario1.Rows = 4
    
    V_ctas
    
    For r = 1 To Cm: Get 3, r, oper
       If oper.identi = "B" Then
         entro = 0
         For I = 0 To Diario1.Rows - 1
                If Val(Diario1.TextMatrix(I, 0)) = Val(oper.Cta) Then
                   entro = 1
                   Exit For
                End If
                
         Next I
         If entro = 0 Then
            Get 2, Val(oper.real), CATMAY
            Diario1.AddItem Format(Val(oper.Cta), "####0") + " " + Trim(CATMAY.B2)
         End If
      End If
   Next r
   Diario1.Col = 0
   Diario1.Row = 0
   Diario1.Sort = 1:
   
   Diario1.Cols = Diario1.Rows - 1
   Diario1.Col = 1: Diario1.Row = 1
   For r = 1 To (Diario1.Rows - 1)
        Diario1.Row = 0: Diario1.Col = r - 1
        Diario1.CellFontBold = True: Diario1.CellAlignment = 1
        Diario1.Text = Diario1.TextMatrix(r, 0)
        Diario1.TextMatrix(r, 0) = ""
   Next r
   Diario1.Rows = 1
   Else
   Rem nada
   End If
End Sub
Sub Ide_Mayor()
 Diario1.ColWidth(0) = 800
 Diario1.ColWidth(1) = 800: Diario1.RowHeight(0) = 800
 Diario1.ColWidth(2) = 3500: For I = 3 To Diario1.Cols - 1: Diario1.ColWidth(I) = 1300: Next I
 Dim Valor_cel As Currency
 
     Diario1.TextMatrix(0, 0) = "Fecha": Rem  Diario1.TextMatrix(1, 0) = "Fecha"
     Diario1.TextMatrix(0, 1) = "Poliza": Rem Diario1.TextMatrix(1, 1) = "Poliza"
     Diario1.TextMatrix(0, 2) = "Descripcion": Rem Diario1.TextMatrix(1, 2) = "Descripcion"
     For r = 1 To Cm: Get 3, r, oper
       
       Select Case oper.identi
         Case "A"
            Select Case final
             Case 0
                fechita = "01/01/" + Trim(Datos.a_o)
             Case 1 To 9
                fechita = oper.fe + "/0" + Trim(Str(final)) + "/" + Trim(Datos.a_o)
             Case 10 To 12
                fechita = oper.fe + "/" + Trim(Str(final)) + "/" + Trim(Datos.a_o)
            End Select
             Diario1.AddItem fechita & Chr(9) & Format(Val(oper.Cta), "####0") & Chr(9) & " " + oper.descr
         Case "B"
             Valor_cel = 0
             For I = 3 To Diario1.Cols - 1
                  If Val(Diario1.TextMatrix(0, I)) = Val(oper.Cta) Then
                     If Diario1.TextMatrix(Diario1.Rows - 1, I) <> "" Then
                             Diario1.Rows = Diario1.Rows + 1
                             Rem Valor_cel = Diario1.TextMatrix(Diario1.Rows - 1, i)
                     End If
                     Rem Valor_cel = Valor_cel + Val(oper.impte)
                     Diario1.TextMatrix(Diario1.Rows - 1, I) = Format(Val(oper.impte), "#,##0.00;(#,##0.00)")
                     Rem If Diario1.Rows > 3 Then Diario1.MergeCells = 3: Diario1.MergeRow(Diario1.Rows - 1) = False
                     Rem Diario1.TextMatrix(Diario1.Rows - 1, i) = Format(Valor_cel, " #,##0.00")
                  End If
             Next I
      End Select
      
     Next r
End Sub
Sub V_ctas()
    Dim W_1 As Long, C_T_1 As Integer, Ve_z As Integer, W_2 As Integer, In_I As Integer
    Dim Dent As Integer
    VCtas.VCtas1.Cols = 2
    VCtas.VCtas1.Rows = 1
    VCtas.VCtas1.Row = 0: Tamno = 0
    VCtas.VCtas1.Row = 0: VCtas.VCtas1.Col = 0: VCtas.VCtas1.ColWidth(0) = 800:  VCtas.VCtas1.CellFontBold = True: VCtas.VCtas1.CellAlignment = 4: VCtas.VCtas1.Text = "Veces"
    VCtas.VCtas1.Row = 0: VCtas.VCtas1.Col = 1: VCtas.VCtas1.ColWidth(1) = 800:  VCtas.VCtas1.CellFontBold = True: VCtas.VCtas1.CellAlignment = 4: VCtas.VCtas1.Text = "Cta."
    Ve_z = 0
    In_I = 1: Dent = 0
    For W_1 = 1 To Cm: Get 3, W_1, oper
         If oper.identi = "B" Then
              C_T_1 = Val(oper.Cta)
              
              If VCtas.VCtas1.Rows > 1 Then
               
                  For W_2 = 1 To VCtas.VCtas1.Rows - 1
                         If C_T_1 = VCtas.VCtas1.TextMatrix(W_2, 1) Then
                               
                               VCtas.VCtas1.TextMatrix(W_2, 0) = VCtas.VCtas1.TextMatrix(W_2, 0) + 1
                               
                               Dent = 1
                               Exit For
                               Else
                               Dent = 0
                         End If
                  Next W_2
                  Else
                   VCtas.VCtas1.AddItem In_I & Chr(9) & C_T_1
                   
                   Dent = 1
              End If
              If Dent = 0 Then
                VCtas.VCtas1.AddItem In_I & Chr(9) & C_T_1
                
              End If
         End If
    Next W_1
    VCtas.VCtas1.Row = 1: VCtas.VCtas1.Col = 0
    VCtas.VCtas1.Sort = 2
    Rem VCtas.Show
End Sub
Sub Diario_2()
    Close
    Open Sub_dir + "CATMAY" For Random As 2 Len = Len(CATMAY)
    Dm = LOF(2) / Len(CATMAY)
    Open Sub_dir + "DATOS" For Random As 1 Len = Len(Datos)
    Get 1, 1, Datos
    For Q1 = 0 To 12:
            Rem Local_iza Trim(Datos.No_arch)
            Select Case Q1
                   Case 0
                   Nom_bre = Trim(Datos.No_arch) + "13"
                   Case 1 To 9
                       Nom_bre = Trim(Datos.No_arch) + "0" + Trim(Str(Q1))
                   Case 10 To 12
                       Nom_bre = Trim(Datos.No_arch) + Trim(Str(Q1))
            End Select
            
            Close 3
            Open Sub_dir + Nom_bre For Random As 3 Len = Len(oper)
            Cm = LOF(3) / Len(oper)
            Ide_Mayor1
            'Debug.Print Nom_bre
    Next Q1
End Sub
Sub Ide_Mayor1()
Dim Lim_Ctas As Long, No_apli As Integer
Dim Respuesta
 Diario1.Clear
 Diario1.Cols = 14: Diario1.Rows = 1
 Diario1.Row = 0
 Rem
 Diario1.Col = 0: Diario1.ColWidth(0) = 800: Diario1.CellAlignment = 4: Diario1.CellBackColor = vbYellow: Diario1.RowHeight(0) = 300:: Diario1.CellFontBold = True
 Diario1.Col = 1: Diario1.ColWidth(1) = 800:  Diario1.CellAlignment = 4: Diario1.CellBackColor = vbYellow
 Diario1.Col = 2: Diario1.ColWidth(2) = 3100: Diario1.CellAlignment = 4: Diario1.CellBackColor = vbYellow
 For I = 3 To 12: Diario1.ColWidth(I) = 1300: Diario1.Col = I: Diario1.CellAlignment = 0: Diario1.CellBackColor = vbYellow: Next I
 Diario1.Col = 13: Diario1.ColWidth(13) = 800: Diario1.CellAlignment = 4: Diario1.CellBackColor = vbYellow
 Dim Valor_cel As Currency
  V_ctas
  If VCtas.VCtas1.Rows - 1 <= 9 Then
        Lim_Ctas = VCtas.VCtas1.Rows - 1
        Else
        Lim_Ctas = 9
  End If
     
     Diario1.TextMatrix(0, 0) = "Fecha": Rem  Diario1.TextMatrix(1, 0) = "Fecha"
     Diario1.TextMatrix(0, 1) = "Poliza": Rem Diario1.TextMatrix(1, 1) = "Poliza"
     Diario1.TextMatrix(0, 2) = "Descripcion": Rem Diario1.TextMatrix(1, 2) = "Descripcion"
     For r = 1 To Lim_Ctas:
                Diario1.TextMatrix(0, r + 2) = VCtas.VCtas1.TextMatrix(r, 1)
                For I = 1 To Dm: Get 2, I, CATMAY
                   If Val(CATMAY.B1) = Val(VCtas.VCtas1.TextMatrix(r, 1)) Then
                      Diario1.TextMatrix(0, r + 2) = Trim(Diario1.TextMatrix(0, r + 2) + " " + Left(Trim(CATMAY.B2), 13))
                      Exit For
                   End If
                Next I
                
     Next r
     Diario1.TextMatrix(0, 12) = "VARIAS CTAS."
     Diario1.TextMatrix(0, 13) = "CTA."
     Unload VCtas
     
     For r = 1 To Cm: Get 3, r, oper
    
       Select Case oper.identi
         Case "A"
            Select Case Q1
             Case 0
                fechita = "01/01/" + Trim(Datos.a_o)
             Case 1 To 9
                fechita = oper.fe + "/0" + Trim(Str(Q1)) + "/" + Trim(Datos.a_o)
             Case 10 To 12
                fechita = oper.fe + "/" + Trim(Str(Q1)) + "/" + Trim(Datos.a_o)
            End Select
             Diario1.AddItem fechita & Chr(9) & Format(Val(oper.Cta), "####0") & Chr(9) & " " + oper.descr
         Case "B"
             
             Valor_cel = 0: No_apli = 0
             For I = 3 To 12
                  If Val(Diario1.TextMatrix(0, I)) = Val(oper.Cta) Then
                     If Diario1.TextMatrix(Diario1.Rows - 1, I) <> "" Then
                             Diario1.Rows = Diario1.Rows + 1
                             Rem Valor_cel = Diario1.TextMatrix(Diario1.Rows - 1, i)
                     End If
                     Rem Valor_cel = Valor_cel + Val(oper.impte)
                     Diario1.TextMatrix(Diario1.Rows - 1, I) = Format(Val(oper.impte), "#,##0.00;(#,##0.00)")
                     Rem If Diario1.Rows > 3 Then Diario1.MergeCells = 3: Diario1.MergeRow(Diario1.Rows - 1) = False
                     Rem Diario1.TextMatrix(Diario1.Rows - 1, i) = Format(Valor_cel, " #,##0.00")
                     No_apli = 1
                     Exit For
                     
                  End If
             Next I
             If No_apli = 0 Then
                     If Diario1.TextMatrix(Diario1.Rows - 1, 12) <> "" Then
                             Diario1.Rows = Diario1.Rows + 1
                             Rem Valor_cel = Diario1.TextMatrix(Diario1.Rows - 1, i)
                     End If
                     Diario1.TextMatrix(Diario1.Rows - 1, 12) = Format(Val(oper.impte), "#,##0.00;(#,##0.00)")
                     Diario1.TextMatrix(Diario1.Rows - 1, 13) = oper.Cta
                     No_apli = 0
            End If
      End Select
      
     Next r
     Suma_Deb
     
     Diario1.TextMatrix(Diario1.Rows - 1, 13) = "": Diario1.TextMatrix(Diario1.Rows - 2, 13) = ""
     Diario.SetFocus: Rem Diario1.TextMatrix(Diario1.Rows - 1, 13) = "": Diario1.TextMatrix(Diario1.Rows - 2, 13) = ""
     Diario.Caption = "Movimientos Libro Diario del mes de " + Trim(Mm(Q1)) + " de " + Datos.a_o
     Respuesta = MsgBox("Desea Imprimir, Necesita un monton de papel oficio y luego carta ", vbYesNo, "Cierre de Libros")
     If Respuesta = vbYes Then ' El usuario eligió el botón Sí.
            Mayor.imp_sora
            For I = 1 To Mayor.MAYOR1.Rows - 1
                  If Mm(Q1) = Mayor.MAYOR1.TextMatrix(I, 0) Then
                         Mayor.MAYOR1.TextMatrix(I, 3) = "SALDO DEL FOLIO DE DIARIO " & Format(F_olio, "#,##0")
                  End If
            Next I
            Else   ' El usuario eligió el botón No.
            Rem MiCadena = "No"   ' Ejecuta una acción.
            HOJA = (Diario1.Rows / 49): Hj_comp = HOJA - Int(HOJA)
            
            If Hj_comp > 0.05 Then
                
                HOJA = Int(HOJA) + 1
                Else
                HOJA = Int(HOJA)
            End If
            
            F_olio = F_olio + Int(HOJA)
            For I = 1 To Mayor.MAYOR1.Rows - 1
                  If Mm(Q1) = Mayor.MAYOR1.TextMatrix(I, 0) Then
                         Mayor.MAYOR1.TextMatrix(I, 3) = "SALDO DEL FOLIO DE DIARIO " & Format(F_olio, "#,##0")
                  End If
            Next I
     End If
     Mayor.SetFocus
End Sub
