VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RDOEJ 
   Caption         =   "RESULTADO DEL EJERCICIO"
   ClientHeight    =   4548
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   12504
   LinkTopic       =   "Form1"
   ScaleHeight     =   4548
   ScaleWidth      =   12504
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Res1 
      Height          =   3252
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   5652
      _ExtentX        =   9970
      _ExtentY        =   5736
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "RDOEJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Def_Col()
   
    Res1.Cols = 7
   Res1.Rows = 1
   Res1.Row = 0: Tamno = 0
    Res1.Row = 0: Res1.Col = 0: Res1.ColWidth(0) = 1200: Tamno = Tamno + 1200: Res1.CellFontBold = True: Res1.CellAlignment = 4: Res1.Text = "MES"
    Res1.Row = 0: Res1.Col = 1: Res1.ColWidth(1) = 1200: Tamno = Tamno + 1200: Res1.CellFontBold = True: Res1.CellAlignment = 4: Res1.Text = "No."
    Res1.Row = 0: Res1.Col = 2: Res1.ColWidth(2) = 2800: Tamno = Tamno + 2800: Res1.CellFontBold = True: Res1.CellAlignment = 4: Res1.Text = "Cuenta"
    Res1.Row = 0: Res1.Col = 3: Res1.ColWidth(3) = 3200: Tamno = Tamno + 3200: Res1.CellFontBold = True: Res1.CellAlignment = 4: Res1.Text = "CONCEPTO"
    Res1.Row = 0: Res1.Col = 4: Res1.ColWidth(4) = 1200: Tamno = Tamno + 1200: Res1.CellFontBold = True: Res1.CellAlignment = 4: Res1.Text = "Debe"
    Res1.Row = 0: Res1.Col = 5: Res1.ColWidth(5) = 1200: Tamno = Tamno + 1200: Res1.CellFontBold = True: Res1.CellAlignment = 4: Res1.Text = "Haber"
    Res1.Row = 0: Res1.Col = 6: Res1.ColWidth(6) = 1200: Tamno = Tamno + 1200: Res1.CellFontBold = True: Res1.CellAlignment = 4: Res1.Text = "Saldo "
    Rem  Mm(mes_lim)
    Res1.Width = Tamno + 400
End Sub

Private Sub Form_Load()
   Def_Col
End Sub
Sub Suma_res()
   Dim Sdo_Res1 As Currency, Sma_mov As Currency
   For r = 1 To Res1.Rows - 1
       If IsNumeric(Res1.TextMatrix(r, 4)) Then
           Sma_mov = (Res1.TextMatrix(r, 4))
           Else
           Sma_mov = (Res1.TextMatrix(r, 5))
       End If
       Sdo_Res1 = Sdo_Res1 + Sma_mov
       Res1.TextMatrix(r, 6) = Format(Sdo_Res1, z1)
   Next r
End Sub
