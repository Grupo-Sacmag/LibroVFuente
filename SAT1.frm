VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SAT1 
   Caption         =   "CATALOGO SAT"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6228
   ScaleMode       =   0  'User
   ScaleWidth      =   9432
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5172
   End
   Begin MSFlexGridLib.MSFlexGrid HcSat 
      Height          =   5652
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9132
      _ExtentX        =   16113
      _ExtentY        =   9975
      _Version        =   393216
   End
End
Attribute VB_Name = "SAT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_Aum1 As Long, T_cta As Integer, T_ubi As Long, T_Naux As Integer, T_Inicio As Long
Dim T_si As Currency, T_Md As Currency, T_Mh As Currency, T_Sf As Currency
Dim w As Long, w1 As Long, r As Long, EncHc As Double, T_Cta1 As Single, BscHc As Integer
Dim MomA As Currency, Sa_ini As Currency, MomA1 As Currency

Private Sub Form_Activate()
  BALANZON = 6
    LIBROSV.Ci.Visible = False
End Sub

Private Sub Form_Load()
       BALANZON = 6
       HcSat.Clear
       HcSat.Cols = 13: HcSat.Rows = 20
      F_Aum1 = 1
      Form_Resize
      SAT1.Refresh
    HcSat.Cols = 13: HcSat.Rows = 2: HcSat.FixedRows = 1
    HcSat.Font = "Arial"
    F_Aum1 = 1
    Balancha.WindowState = 2
    HcSat.Clear
    HcSat.Rows = 1
    HcSat.Clear
    HcSat.Rows = 1
    HcSat.Row = 0: HcSat.Col = 0: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Cod.Agrp"
    HcSat.Row = 0: HcSat.Col = 1: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Cta.Cont."
    HcSat.Row = 0: HcSat.Col = 2: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "NombreCont"
    HcSat.Row = 0: HcSat.Col = 3: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Cta/Sbcta"
    HcSat.Row = 0: HcSat.Col = 4: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Natur."
    HcSat.Row = 0: HcSat.Col = 5: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Saldo " + Mm(mes_lim1 - 1)
    HcSat.Row = 0: HcSat.Col = 6: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Debe"
    HcSat.Row = 0: HcSat.Col = 7: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Haber"
    HcSat.Row = 0: HcSat.Col = 8: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Saldo " + Mm(mes_lim1)
    HcSat.Row = 0: HcSat.Col = 9: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Tipo"
    HcSat.Row = 0: HcSat.Col = 10: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Clave"
    HcSat.Row = 0: HcSat.Col = 11: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Rgtro"
    HcSat.Row = 0: HcSat.Col = 12: HcSat.CellFontBold = True: HcSat.CellAlignment = 4: HcSat.Text = "Aux"
    HcSat.FixedCols = 2
    Rem SatRecorre
    Alternativa
    Suma_Mvtos1
End Sub
Sub SatRecorre()
    Close: Open "CATSAT1.CG" For Random As 1 Len = Len(CatSat)
    Cm = LOF(1) / Len(CatSat)
    If Cm < 1 Then
       MsgBox "NO EXISTE EL ARCHIVO CATSAT1.CG O ESTA VACIO", vbCritical
       Exit Sub
    End If
    For r = 1 To Cm: Get 1, r, CatSat
        Sa_ini = 0
        
        HcSat.AddItem CatSat.Csta & Chr(9) & _
                      CatSat.Numero & Chr(9) & _
                      CatSat.Nombre & Chr(9) & _
                      Format(Sa_ini, "#,##0.00") & Chr(9) & _
                      Format(Sa_ini, "#,##0.00") & Chr(9) & _
                      Format(Sa_ini, "#,##0.00") & Chr(9) & _
                      Format(Sa_ini, "#,##0.00") & Chr(9) & _
                      CatSat.Clave & Chr(9) & _
                      r
                      
                       
         Select Case CatSat.Clave
             Rem T_cta = 0: T_ubi = 0
             Case 1
                For I = 1 To Balancha.Bcha1.Rows - 1
                If (IsNumeric(Balancha.Bcha1.TextMatrix(I, 0))) And (Balancha.Bcha1.TextMatrix(I, 6) = "C") Then
                    If Val(CatSat.Tipo) = Balancha.Bcha1.TextMatrix(I, 0) Then
                       HcSat.TextMatrix(HcSat.Rows - 1, 3) = Balancha.Bcha1.TextMatrix(I, 2)
                       HcSat.TextMatrix(HcSat.Rows - 1, 4) = Balancha.Bcha1.TextMatrix(I, 3)
                       HcSat.TextMatrix(HcSat.Rows - 1, 5) = Balancha.Bcha1.TextMatrix(I, 4)
                       HcSat.TextMatrix(HcSat.Rows - 1, 6) = Balancha.Bcha1.TextMatrix(I, 5)
                       T_cta = Balancha.Bcha1.TextMatrix(I, 0): T_ubi = I
                       Exit For
                    End If
                    Else
                    T_ubi = 0
                End If
                Next I
             Case 2
                Rem ****** AUXILIAR ESPECIFICO **********
                T_Inicio = T_ubi + 1
                HcSat.RemoveItem HcSat.Rows - 1
                t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
                Do Until t_cl <> "A"
                  Rem *****  DETERMINA EL RANGO DEL AUXILIAR *******
                     T_Naux = Balancha.Bcha1.TextMatrix(T_Inicio, 0)
                     r = r + T_Naux - 1: Get 1, r, CatSat
                     If r > T_max Then
                                T_max = r
                     End If
                     HcSat.AddItem CatSat.Csta & Chr(9) & _
                                   CatSat.Numero & Chr(9) & _
                                   CatSat.Nombre & Chr(9) & _
                                   Balancha.Bcha1.TextMatrix(I, 2) & Chr(9) & _
                                   Balancha.Bcha1.TextMatrix(I, 3) & Chr(9) & _
                                   Balancha.Bcha1.TextMatrix(I, 4) & Chr(9) & _
                                   Balancha.Bcha1.TextMatrix(I, 5) & Chr(9) & _
                                   CatSat.Clave
                     T_Inicio = T_Inicio + 1
                     t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
                Loop
                
                If r < T_max Then
                        r = T_max
                        T_max = 0
                        Else
                        T_max = 0
                End If
                T_ubi = 0
             Case 3
                T_ubi = 0
             Case 4
                T_ubi = 0
             Case 5
               Rem Suma Auxiliares simple
             If T_ubi > 0 Then
               T_Inicio = T_ubi + 1:
               t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
               Do Until t_cl <> "A"
                   T_si = T_si + Balancha.Bcha1.TextMatrix(T_Inicio, 2)
                   T_Md = T_Md + Balancha.Bcha1.TextMatrix(T_Inicio, 3)
                   T_Mh = T_Mh + Balancha.Bcha1.TextMatrix(T_Inicio, 4)
                   T_Sf = T_Sf + Balancha.Bcha1.TextMatrix(T_Inicio, 5)
                   T_Inicio = T_Inicio + 1:
                   t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
               Loop
                   HcSat.TextMatrix(HcSat.Rows - 1, 3) = Format(T_si, z1)
                   HcSat.TextMatrix(HcSat.Rows - 1, 4) = Format(T_Md, z1)
                   HcSat.TextMatrix(HcSat.Rows - 1, 5) = Format(T_Mh, z1)
                   HcSat.TextMatrix(HcSat.Rows - 1, 6) = Format(T_Sf, z1)
                   T_cta = 0: T_ubi = 0
                   T_si = 0: T_Md = 0: T_Mh = 0: T_Sf = 0
                   T_ubi = 0
              End If
              Case 6
                  Rem Varias Cuentas Y Auxiliares
                  
                  
                  
                  
                  
                  
                  T_ubi = 0
              Case 7
                  T_ubi = 0
              Case 9
                 T_ubi = 0
              Case 10
                T_ubi = 0
             Case Else
                 T_ubi = 0
                Rem  Nada
         End Select
        
    Next r
    
End Sub
Sub Col22()
    HcSat.FontWidth = 3 * F_Aum1
    HcSat.ColWidth(0) = 800 * F_Aum1
    HcSat.ColWidth(1) = 800 * F_Aum1
    HcSat.ColWidth(2) = 2300 * F_Aum1
    HcSat.ColWidth(3) = 200 * F_Aum1
    HcSat.ColWidth(4) = 200 * F_Aum1
    HcSat.ColWidth(5) = 1500 * F_Aum1
    HcSat.ColWidth(6) = 1500 * F_Aum1
    HcSat.ColWidth(7) = 1500 * F_Aum1
    HcSat.ColWidth(8) = 1500 * F_Aum1
    HcSat.ColWidth(9) = 800 * F_Aum1
    HcSat.ColWidth(9) = 800 * F_Aum1
    HcSat.ColWidth(9) = 800 * F_Aum1
End Sub

Private Sub Form_Resize()
  If SAT1.WindowState <> 1 Then
      HcSat.Height = ScaleHeight - 400
      HcSat.Width = ScaleWidth - 400
      F_Aum1 = (HcSat.Width - 400) / 9200
      Col22
   End If

End Sub
Sub Alternativa()
    Dim SATClave As String, SATCta As Integer, SATClave1 As Integer
    Dim M_sat As Long, BscHc As Single
    Close
    Open "CATSAT1.CG" For Random As 1 Len = Len(CatSat)
    Cm = LOF(1) / Len(CatSat)
    Open "CruceCg.Cg" For Random As 2 Len = Len(CruSat)
    Dm = LOF(2) / Len(CruSat)
    If Cm < 1 Then MsgBox "No existe archivo CATSAT1.CG o esta vacio"
    If Dm < 1 Then MsgBox "No existe archivo CruceCg.Cg o esta vacio"
    Rem ***** Carga Catalogo Sat
       Sa_ini = 0
       For r = 1 To Cm: Get 1, r, CatSat
       
       If CatSat.Clave = 0 Then
                    CatSat.Csta = "0"
                    Ctita = 0: NOMBRECITO = "": compa = Trim(CatSat.Numero)
                    If compa = "100" Then Ctita = 1000: NOMBRECITO = "ACTIVO": CatSat.Tipo = "D"
                    If compa = "100.01" Then Ctita = 1100: NOMBRECITO = "CIRCULANTE": CatSat.Tipo = "D"
                    If compa = "100.02" Then Ctita = 1200: NOMBRECITO = "FIJO": CatSat.Tipo = "D"
                    If compa = "200" Then Ctita = 2000: NOMBRECITO = "PASIVO": CatSat.Tipo = "A"
                    If compa = "200.01" Then Ctita = 2100: NOMBRECITO = "CIRCULANTE": CatSat.Tipo = "A"
                    If compa = "200.02" Then Ctita = 2200: NOMBRECITO = "FIJO": CatSat.Tipo = "A"
                    If compa = "300" Then Ctita = 3000: NOMBRECITO = "HABER SOCIAL": CatSat.Tipo = "A"
                    If compa = "400" Then Ctita = 5000: NOMBRECITO = "INGRESOS": CatSat.Tipo = "A"
                    If compa = "500" Then Ctita = 4000: NOMBRECITO = "OTROS GASTOS": CatSat.Tipo = "D"
                    If compa = "600" Then Ctita = 4100: NOMBRECITO = "GASTOS": CatSat.Tipo = "D"
                    If compa = "700" Then Ctita = 4200: NOMBRECITO = "CARGOS FINANCIEROS": CatSat.Tipo = "D"
                    If compa = "702" Then Ctita = 5100: NOMBRECITO = "GANANCIA FINANCIERA": CatSat.Tipo = "D"
                    If compa = "704" Then Ctita = 5200: NOMBRECITO = "OTROS INGRESOS": CatSat.Tipo = "D"
                    If compa = "800" Then Ctita = 6000: NOMBRECITO = "CUENTAS DE ORDEN": CatSat.Tipo = "D"
                    If compa = "100.02" Then Ctita = 1200: NOMBRECITO = "FIJO": CatSat.Tipo = "D"
                    If NOMBRECITO = "" Then NOMBRECITO = Trim(CatSat.Nombre)
                    HcSat.AddItem Trim(CatSat.Numero) & Chr(9) & _
                                Ctita & Chr(9) & _
                                NOMBRECITO & Chr(9) & _
                                "1" & Chr(9) & _
                                CatSat.Tipo & Chr(9) & _
                                "0" & Chr(9) & _
                                "0" & Chr(9) & _
                                "0" & Chr(9) & _
                                "0" & Chr(9) & _
                                CatSat.Tipo & Chr(9) & _
                                CatSat.Clave & Chr(9) & _
                                r & Chr(9) & _
                                ""
                    Else
                    w = 0: Ctita = CDec(CatSat.Numero)
                    NOMBRECITO = CatSat.Nombre
                    For w = 1 To Dm: Get 2, w, CruSat
                        If (Val(CatSat.Numero) = Val(CruSat.SatNu)) Then
                            If IsNumeric(CruSat.Cta) Then
                                TrCtita = ""
                                Ctita = Val(CruSat.Cta): NOMBRECITO = Trim(CruSat.Nomb)
                                TrCtita = CruSat.Cta
                                Exit For
                                Else
                                Exit For
                            End If
                        End If
             
                    Next w
                    If Val(CatSat.Csta) = 2 Then
                        
                         axsat = axsat + 1
                        If TrCtita <> "" Then
                        
                         If axsat < 10 Then
                             Ctita = Trim(TrCtita) + ".0" + Trim(Str(axsat))
                            
                             Else
                             Ctita = Trim(TrCtita) + "." + Trim(Str(axsat))
                         End If
                          CatSat.Numero = Format(CatSat.Numero, "#0.00")
                        End If
                         Else
                         axsat = 0
                    End If
                    HcSat.AddItem Trim(CatSat.Numero) & Chr(9) & _
                                           Ctita & Chr(9) & _
                                           Trim(NOMBRECITO) & Chr(9) & _
                                           Val(CatSat.Csta) & Chr(9) & _
                                           CruSat.Nat & Chr(9) & _
                                           Format(Sa_ini, z1) & Chr(9) & _
                                           Format(Sa_ini, z1) & Chr(9) & _
                                           Format(Sa_ini, z1) & Chr(9) & _
                                           Format(Sa_ini, z1) & Chr(9) & _
                                           CatSat.Tipo & Chr(9) & _
                                           CatSat.Clave & Chr(9) & _
                                           r & Chr(9) & _
                                           ""
        End If
    
     Next r
     
    Rem ******* Recorrer Nuestra balanza
    For r = 1 To Balancha.Bcha1.Rows - 1
         SATClave = Balancha.Bcha1.TextMatrix(r, 6)
         If SATClave = "C" Then
             SATCta = Balancha.Bcha1.TextMatrix(r, 0)
             
             For w = 1 To Dm: Get 2, w, CruSat: Rem INICIA RECORRIDO ARCHIVO BUSCANDO CUENTA DEL SAT
                 
                 If IsNumeric(CruSat.Cta) Then
                 
                    If SATCta = CruSat.Cta Then
                        
                         BscHc = Val(CruSat.SatNu): Rem CUENTA DEL SAT COINCIDE Y CONSIGUE CUENTA DEL SAT
                            Rem HcSat.TextMatrix(r, 1) = Balancha.Bcha1.TextMatrix(r, 0)
                            Rem HcSat.TextMatrix(r, 2) = Balancha.Bcha1.TextMatrix(r, 1)
                         For w1 = 1 To HcSat.Rows - 1
                         
                            Rem BUSCA CUENTA DEL SAT EN LA HOJA DE CALCULO *************************
                            
                            EncHc = Val(HcSat.TextMatrix(w1, 0))
                            If BscHc = EncHc Then
                               If CruSat.SatNu >= 800 Then
                                 Debug.Print CruSat.SatNu; (Balancha.Bcha1.TextMatrix(r, 2)); MomA; (Balancha.Bcha1.TextMatrix(r, 5))
                                 
                               End If
                               Rem HcSat.TextMatrix(w1, 1) = Balancha.Bcha1.TextMatrix(r, 0)
                               Rem HcSat.TextMatrix(w1, 2) = Balancha.Bcha1.TextMatrix(r, 1)
                              MomA = 0
                               If IsNumeric(HcSat.TextMatrix(w1, 5)) Then
                                    
                                    MomA = (HcSat.TextMatrix(w1, 5))
                                    MomA = MomA + (Balancha.Bcha1.TextMatrix(r, 2))
                                    HcSat.TextMatrix(w1, 5) = Format(MomA, z1)
                                    
                                    Else
                                    MomA = Balancha.Bcha1.TextMatrix(r, 2)
                                    HcSat.TextMatrix(w1, 5) = Format(MomA, z1)
                               End If
                               If IsNumeric(HcSat.TextMatrix(w1, 6)) Then
                                    MomA = (HcSat.TextMatrix(w1, 6))
                                    MomA = MomA + (Balancha.Bcha1.TextMatrix(r, 3))
                                    HcSat.TextMatrix(w1, 6) = Format(MomA, z1)
                                    Else
                                    MomA = Balancha.Bcha1.TextMatrix(r, 3)
                                    HcSat.TextMatrix(w1, 6) = Format(MomA, z1)
                               End If
                               If IsNumeric(HcSat.TextMatrix(w1, 7)) Then
                                    MomA = (HcSat.TextMatrix(w1, 7))
                                    MomA = MomA + (Balancha.Bcha1.TextMatrix(r, 4))
                                    HcSat.TextMatrix(w1, 7) = Format(MomA, z1)
                                    Else
                                    MomA = Balancha.Bcha1.TextMatrix(r, 4)
                                    HcSat.TextMatrix(w1, 7) = Format(MomA, z1)
                               End If
                               If IsNumeric(HcSat.TextMatrix(w1, 8)) Then
                                    MomA = (HcSat.TextMatrix(w1, 8))
                                    MomA = MomA + (Balancha.Bcha1.TextMatrix(r, 5))
                                    HcSat.TextMatrix(w1, 8) = Format(MomA, z1)
                                    
                                    Else
                                    MomA = Balancha.Bcha1.TextMatrix(r, 5)
                                    HcSat.TextMatrix(w1, 8) = Format(MomA, z1)
                               End If
                          AxSolucion
                          
                         Exit For
                       End If
                    Next w1
                   End If
                 End If
             Next w
         
      End If
    Next r
    
End Sub
Sub AxSolucion()
Dim ClaHc As Integer, Cl_cl1 As Integer
 Rem GoTo BRINCAA:
 If w1 < HcSat.Rows - 1 Then
    ClaHc = HcSat.TextMatrix(w1 + 1, 10)
    Select Case ClaHc
        Case 5
           Rem Suma auxiliares completo a una sola subcuenta del sat  ******
           Rem If T_ubi > 0 Then
           
               CALLITO = Balancha.Bcha1.TextMatrix(T_Inicio, 0)
               T_Inicio = r + 1
               t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
               
               Do Until t_cl <> "A"
               
                   T_si = T_si + Balancha.Bcha1.TextMatrix(T_Inicio, 2)
                   T_Md = T_Md + Balancha.Bcha1.TextMatrix(T_Inicio, 3)
                   T_Mh = T_Mh + Balancha.Bcha1.TextMatrix(T_Inicio, 4)
                   T_Sf = T_Sf + Balancha.Bcha1.TextMatrix(T_Inicio, 5)
                   T_Inicio = T_Inicio + 1:
                   t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
               Loop
               
               Rem r = T_Inicio - 1
               If IsNumeric(HcSat.TextMatrix(w1 + 1, 5)) Then
                           
                           MomA = HcSat.TextMatrix(w1 + 1, 5)
                           MomA = MomA + T_si
                           HcSat.TextMatrix(w1 + 1, 5) = Format(MomA, z1)
                           Else
                           MomA = T_si
                           HcSat.TextMatrix(w1 + 1, 5) = Format(MomA, z1)
               End If
               If IsNumeric(HcSat.TextMatrix(w1 + 1, 6)) Then
                           MomA = HcSat.TextMatrix(w1 + 1, 6)
                           MomA = MomA + T_Md
                           HcSat.TextMatrix(w1 + 1, 6) = Format(MomA, z1)
                           Else
                           MomA = T_Md
                           HcSat.TextMatrix(w1 + 1, 6) = Format(MomA, z1)
               End If
               If IsNumeric(HcSat.TextMatrix(w1 + 1, 7)) Then
                           MomA = HcSat.TextMatrix(w1 + 1, 7)
                           MomA = MomA + T_Mh
                           HcSat.TextMatrix(w1 + 1, 7) = Format(MomA, z1)
                           Else
                           MomA = T_Mh
                           HcSat.TextMatrix(w1 + 1, 7) = Format(MomA, z1)
               End If
               If IsNumeric(HcSat.TextMatrix(w1 + 1, 8)) Then
                           MomA = HcSat.TextMatrix(w1 + 1, 8)
                           MomA = MomA + T_Sf
                           HcSat.TextMatrix(w1 + 1, 8) = Format(MomA, z1)
                           Else
                           MomA = T_Sf
                           HcSat.TextMatrix(w1 + 1, 8) = Format(MomA, z1)
               End If

                   ' HcSat.TextMatrix(w1 + 1, 4) = Format(T_Md, z1)
                   ' HcSat.TextMatrix(w1 + 1, 5) = Format(T_Mh, z1)
                   ' HcSat.TextMatrix(w1 + 1, 6) = Format(T_Sf, z1)
                   T_cta = 0: T_ubi = 0
                   T_si = 0: T_Md = 0: T_Mh = 0: T_Sf = 0
                   T_ubi = 0: MomA = 0
                   
           Rem  End If
           Case 2
               Rem buscar en el archivo el rango entre una cuenta y otra ******************
               
             For w4 = 1 To Dm: Get 2, w4, CruSat
                    If Val(CruSat.Cta) = Balancha.Bcha1.TextMatrix(r, 0) Then
                        Ini_ax = w4
                        Exit For
                    End If
             Next w4
             For w4 = Ini_ax + 1 To Dm: Get 2, w4, CruSat
                 If Val(CruSat.Cta) > 0 Then
                        Fin_ax = w4 - 1
                        Exit For
                 End If
             Next w4
             
             Rem queda establecido el rango de busqueda en el archivo ***************
             Rem Vas a buscar los auxiliares con un loop y vas a buscar el numero de auxiliar y asociarlo en la hcsat aunque a partir de la w1
             T_Inicio = r + 1
             t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
             T_cta = Balancha.Bcha1.TextMatrix(T_Inicio, 0)
               Do Until t_cl <> "A"
                  For w4 = Ini_ax To Fin_ax: Get 2, w4, CruSat
                    
                       If T_cta = Val(CruSat.Fin) Then
                            
                           For w5 = w1 + 1 To HcSat.Rows - 2
                                
                               If Val(CruSat.SatNu) = Val(HcSat.TextMatrix(w5, 0)) Then
                                   
                                    For w6 = 2 To 5
                                        MomA = HcSat.TextMatrix(w5, w6 + 3)
                                        MomA1 = Balancha.Bcha1.TextMatrix(T_Inicio, w6)
                                        MomA = MomA + MomA1
                                        HcSat.TextMatrix(w5, w6 + 3) = Format(MomA, z1)
                                        
                                    Next w6
                                    
                                    MomA = 0: MomA1 = 0
                               End If
                           Next w5
                        
                       Exit For
                       End If
                  Next w4
                   T_Inicio = T_Inicio + 1
                   t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
                   T_cta = Val(Balancha.Bcha1.TextMatrix(T_Inicio, 0))
               Loop
               Rem r = T_Inicio
               MomA = 0: MomA1 = 0
        Case 7 To 11
            
            Dim M_Na(4) As Currency, M_Ex(4) As Currency, M_PeB(4) As Currency, M_NaP(4) As Currency, M_EP(4) As Currency
            Rem VAMOS A UBICAR EL RANGO DE BUSQUEDA EN EL SAT **************************************
             Ini_ax = w1 + 1
             For w4 = Ini_ax To HcSat.Rows - 1:
                    If Val(HcSat.TextMatrix(w4, 3)) <> 2 Then
                        Fin_ax = w4 - 1
                        Exit For
                    End If
             Next w4
            
            
            T_Inicio = r + 1
             t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
             T_cta = Balancha.Bcha1.TextMatrix(T_Inicio, 0)
               Do Until t_cl <> "A"
                  Cl_cl = Left(Balancha.Bcha1.TextMatrix(T_Inicio, 1), 3)
                  Select Case Cl_cl
                      Case "EX "
                         Rem SUMA EXTRANJEROS CLAVE 7 ********************************************
                         For w3 = 2 To 5
                              MomA = Balancha.Bcha1.TextMatrix(T_Inicio, w3)
                              M_Ex(w3 - 1) = M_Ex(w3 - 1) + MomA
                         Next w3
                      Case "NP "
                          Rem SUMA NACIONAL PARTES RELACIONADAS  CLAVE 10 **************************
                         For w3 = 2 To 5
                              MomA = Balancha.Bcha1.TextMatrix(T_Inicio, w3)
                              M_NaP(w3 - 1) = M_NaP(w3 - 1) + MomA
                         Next w3
                      Case "EP "
                          Rem SUMA EXTRANJEROS PARTES RELACIONADAS CLAVE 8 ************************
                         For w3 = 2 To 5
                              MomA = Balancha.Bcha1.TextMatrix(T_Inicio, w3)
                              M_EP(w3 - 1) = M_EP(w3 - 1) + MomA
                         Next w3
                      Case "PB "
                          Rem SUMA PRESTAMOS BANCARIOS CALVE 11 ************************************
                          For w3 = 2 To 5
                              MomA = Balancha.Bcha1.TextMatrix(T_Inicio, w3)
                              M_PeB(w3 - 1) = M_PeB(w3 - 1) + MomA
                          Next w3
                      Case "NA "
                          Rem SUMA NACIONALES CLAVE 9 *********************************************
                          For w3 = 2 To 5
                              MomA = Balancha.Bcha1.TextMatrix(T_Inicio, w3)
                              M_Na(w3 - 1) = M_Na(w3 - 1) + MomA
                          Next w3
                          
                      Case Else
                          Rem SUMA NACIONALES  CLAVE  9 *********************************************
                          For w3 = 2 To 5
                              MomA = Balancha.Bcha1.TextMatrix(T_Inicio, w3)
                              M_Na(w3 - 1) = M_Na(w3 - 1) + MomA
                          Next w3

                  End Select
                   T_Inicio = T_Inicio + 1
                   t_cl = Balancha.Bcha1.TextMatrix(T_Inicio, 6)
                   T_cta = Val(Balancha.Bcha1.TextMatrix(T_Inicio, 0))
               Loop
               
               
               
               Rem CORREGIR CLAVE AHORA BUSCAR EL RANGO DE
             For w4 = Ini_ax To Fin_ax
             
                 Cl_cl1 = HcSat.TextMatrix(w4, 10)
                 
                 Select Case Cl_cl1
                       Case 7
                         For w5 = 3 To 6
                              MomA = HcSat.TextMatrix(w4, w5 + 2)
                              MomA = MomA + M_Ex(w5 - 2)
                              HcSat.TextMatrix(w4, w5 + 2) = Format(MomA, z1)
                         Next w5
                          
                       Case 8
                         For w5 = 3 To 6
                              MomA = HcSat.TextMatrix(w4, w5 + 2)
                              MomA = MomA + M_EP(w5 - 2)
                              HcSat.TextMatrix(w4, w5 + 2) = Format(MomA, z1)
                         Next w5
                       Case 9
                          For w5 = 3 To 6
                              MomA = HcSat.TextMatrix(w4, w5 + 2)
                              MomA = MomA + M_Na(w5 - 2)
                              HcSat.TextMatrix(w4, w5 + 2) = Format(MomA, z1)
                         Next w5
                            
                       Case 10
                          For w5 = 3 To 6
                              MomA = HcSat.TextMatrix(w4, w5 + 2)
                              MomA = MomA + M_NaP(w5 - 2)
                              HcSat.TextMatrix(w4, w5 + 2) = Format(MomA, z1)
                          Next w5
                        Case 11
                           For w5 = 3 To 6
                              MomA = HcSat.TextMatrix(w4, w5 + 2)
                              MomA = MomA + M_PeB(w5 - 2)
                              HcSat.TextMatrix(w4, w5 + 2) = Format(MomA, z1)
                          Next w5
                 End Select
             Next w4
    End Select
    End If
BRINCAA:
End Sub

Private Sub HcSat_EnterCell()
    Text1.Text = HcSat.Text
End Sub
Sub Suma_Mvtos1()
   Dim Sum_a(5) As Currency
   For r = 1 To HcSat.Rows - 1
       For G = 5 To 8
            If IsNumeric(HcSat.TextMatrix(r, G)) Then
                Sum_a(G - 4) = Sum_a(G - 4) + HcSat.TextMatrix(r, G)
            End If
       Next G
   Next r
   HcSat.AddItem "" & Chr(9) & "" & Chr(9) & "Sumas" & Chr(9) & _
                 "" & Chr(9) & _
                 "" & Chr(9) & _
                Format(Sum_a(1), "#,##0.00;(#,##0.00)") & Chr(9) & Format(Sum_a(2), "#,##0.00;(#,##0.00)") _
               & Chr(9) & Format(Sum_a(3), "#,##0.00;(#,##0.00)") & Chr(9) & Format(Sum_a(4), "#,##0.00;(#,##0.00)")
End Sub

