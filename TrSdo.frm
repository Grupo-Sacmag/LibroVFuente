VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form TrSdo 
   Caption         =   "Saldos del Ejercicio"
   ClientHeight    =   6252
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11028
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6252
   ScaleWidth      =   11028
   Begin VB.CheckBox Check1 
      Caption         =   "Mostrar cuentas sin movimientos"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid TrS1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10455
      _ExtentX        =   18436
      _ExtentY        =   8911
      _Version        =   393216
      Cols            =   50
   End
End
Attribute VB_Name = "TrSdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mi_Linea
Dim F_Aum As Double

Private Sub Check1_Click()
  Form_Activate
  Form_Load
End Sub

Private Sub Form_Activate()
  BALANZON = 3
  LIBROSV.Ci.Visible = False
End Sub

Private Sub Form_GotFocus()
   
   
     LIBROSV.Meses.Visible = False
     LIBROSV.MayArBal.Visible = False
     LIBROSV.MayArDia.Visible = False
   
End Sub

Private Sub Form_Load()
     'LIBROSV.Meses.Visible = False
     'LIBROSV.MayArBal.Visible = False
     'LIBROSV.MayArDia.Visible = False
     TrS1.ColWidth(0) = 800: TrS1.ColWidth(1) = 2300
     z1 = "#,##0.00;(#,##0.00)"
      For I = 2 To 49: TrS1.ColWidth(I) = 1300: Next I
      Diario.Refresh
      BALANZON = 3
      Diario.WindowState = 2
      Rem final=12
      Sdotr final
End Sub
Sub Sdotr(final)
Rem On Error GoTo SALEFIN
If cm > 1 Then
  TrSdo.Caption = "Saldos de mayor " + Trim(Mm(12)) + " de " + Datos.a_o
  Else
  Rem nada
End If
    TrS1.Clear
    TrS1.Rows = 1
    TrS1.Row = 0: TrS1.Col = 0: TrS1.CellFontBold = True: TrS1.CellAlignment = 4: TrS1.Text = "Cuenta"
    TrS1.Row = 0: TrS1.Col = 1: TrS1.CellFontBold = True: TrS1.CellAlignment = 4: TrS1.Text = "Nombre"
    For I = 2 To 15
      TrS1.Row = 0: TrS1.Col = I: TrS1.CellFontBold = True: TrS1.CellAlignment = 4: TrS1.Text = Trim(Mm(I - 2))
    Next I
    TrS1.FixedCols = 2
    Close
    Open Sub_dir + "CATMAY" For Random As 2 Len = Len(CATMAY)
    EM = LOF(2) / Len(CATMAY)
    Open Sub_dir + "DEBE.GCS" For Random As 4 Len = Len(MvDebe)
    cm = LOF(4) / Len(MvDebe)
    Open Sub_dir + "HABER.GCS" For Random As 5 Len = Len(MvHaber)
    Dm = LOF(5) / Len(MvHaber)
    For L = 1 To EM
        M_ue = 0: Sdo_ini = 0: M_D = 0: M_H = 0: Sdo_Fin = 0
        Get 2, L, CATMAY
        Get 4, L, MvDebe: Get 5, L, MvHaber
        Mi_Linea = ""
        If (IsNumeric(CATMAY.B1)) And (Val(CATMAY.B1) > 0) Then
            Mi_Linea = Format(CATMAY.B1, "####0") & Chr(9) & Trim(CATMAY.B2) & Chr(9)
            
            For I = 0 To final
              Select Case I
               Case 0
                Sdo_ini = Sdo_ini + MvDebe.Inc + MvHaber.Inc
                Mi_Linea = Mi_Linea + Format(Sdo_ini, z1)
               Case 1
                 Sdo_ini = Sdo_ini + MvDebe.Ene + MvHaber.Ene
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
               Case 2
                 Sdo_ini = Sdo_ini + MvDebe.Feb + MvHaber.Feb
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
               Case 3
                 Sdo_ini = Sdo_ini + MvDebe.Mar + MvHaber.Mar
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
               Case 4
                 Sdo_ini = Sdo_ini + MvDebe.Abr + MvHaber.Abr
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
               Case 5
                 Sdo_ini = Sdo_ini + MvDebe.May + MvHaber.May
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
               Case 6
                 Sdo_ini = Sdo_ini + MvDebe.Jun + MvHaber.Jun
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
               Case 7
                 Sdo_ini = Sdo_ini + MvDebe.Jul + MvHaber.Jul
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
                 
               Case 8
                 Sdo_ini = Sdo_ini + MvDebe.Ago + MvHaber.Ago
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
                 
               Case 9
                 Sdo_ini = Sdo_ini + MvDebe.Sep + MvHaber.Sep
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
                 
               Case 10
                 Sdo_ini = Sdo_ini + MvDebe.Oct + MvHaber.Oct
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
                 
               Case 11
                 Sdo_ini = Sdo_ini + MvDebe.Nov + MvHaber.Nov
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
                 
               Case 12
                 Sdo_ini = Sdo_ini + MvDebe.Dic + MvHaber.Dic
                 Mi_Linea = Mi_Linea + Chr(9) & Format(Sdo_ini, z1)
                 
            End Select
               
        Next I
        Rem Mi_Linea = Mi_Linea + Chr(10) & Chr(13)
        TrS1.AddItem Mi_Linea
        
       End If
       Next L
      If Check1.Value = 0 Then
            Ana_lss
      End If

End Sub
Sub Ana_lss()
Dim Bien As Integer, ValorReg As Currency
   I = 0
   Do Until I >= (TrS1.Rows - 2)
       I = I + 1
       Bien = 0
       For F = 2 To (final + 2)
        If IsNumeric(TrS1.TextMatrix(I, F)) Then
                  ValorReg = TrS1.TextMatrix(I, F)
                  If ValorReg <> 0 Then
                     Bien = 1
                     Exit For
                  End If
           
         End If
       Next F
       
       If Bien = 0 Then
                
                TrS1.RemoveItem (I)
                Rem TrS1.Rows = TrS1.Rows - 1
                I = I - 1
       End If
   Loop
End Sub

Private Sub Form_Resize()
   If TrSdo.WindowState <> 1 Then
      TrS1.Height = ScaleHeight - 600
      TrS1.Width = ScaleWidth - 200
      F_Aum = (TrS1.Width - 400) / 9200
      Rem ColDfn
   End If
End Sub
Sub ColDfn()
    TrS1.FontWidth = 3 * F_Aum
    Rem TRS1.RowHeight(0) = TRS1.FontWidth
    TrS1.ColWidth(0) = 300 * F_Aum
    TrS1.ColWidth(1) = 2300 * F_Aum
    For G = 2 To 15
        TrS1.ColWidth(I) = 1500 * F_Aum
    Next G
End Sub
