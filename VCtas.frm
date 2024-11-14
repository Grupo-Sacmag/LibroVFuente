VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VCtas 
   Caption         =   "VCtas"
   ClientHeight    =   2484
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   ScaleHeight     =   2484
   ScaleWidth      =   2700
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid VCtas1 
      Height          =   2172
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2172
      _ExtentX        =   3831
      _ExtentY        =   3831
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "VCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ColumNas
End Sub
Sub ColumNas()
   
    Rem VCtas1.Row = 0: VCtas1.Col = 2: VCtas1.ColWidth(2) = 2200: Tamno = Tamno + 2200: VCtas1.CellFontBold = True: VCtas1.CellAlignment = 4: VCtas1.Text = "CONCEPTO"
    Rem VCtas1.Row = 0: VCtas1.Col = 3: VCtas1.ColWidth(3) = 1200: Tamno = Tamno + 1200: VCtas1.CellFontBold = True: VCtas1.CellAlignment = 4: VCtas1.Text = "Debe"
    Rem VCtas1.Row = 0: VCtas1.Col = 4: VCtas1.ColWidth(4) = 1200: Tamno = Tamno + 1200: VCtas1.CellFontBold = True: VCtas1.CellAlignment = 4: VCtas1.Text = "Haber"
    Rem VCtas1.Row = 0: VCtas1.Col = 5: VCtas1.ColWidth(5) = 1200: Tamno = Tamno + 1200: VCtas1.CellFontBold = True: VCtas1.CellAlignment = 4: VCtas1.Text = "Saldo "
    Rem  Mm(mes_lim)
    End Sub
