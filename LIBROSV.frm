VERSION 5.00
Begin VB.MDIForm LIBROSV 
   BackColor       =   &H8000000C&
   Caption         =   "Actualizacion libros"
   ClientHeight    =   7380
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14745
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MayAr 
      Caption         =   "&Archivo"
      Begin VB.Menu MayArDir 
         Caption         =   "&Establecer Directorio de Trabajo"
      End
      Begin VB.Menu MayArSep4 
         Caption         =   "-"
      End
      Begin VB.Menu MayArDia 
         Caption         =   "&Diario"
      End
      Begin VB.Menu MaySep1 
         Caption         =   "-"
      End
      Begin VB.Menu MayArBal 
         Caption         =   "&Balanza de Comprobacion"
      End
      Begin VB.Menu MayArSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MayArImp 
         Caption         =   "&Impresion"
      End
      Begin VB.Menu MayArSep3 
         Caption         =   "-"
      End
      Begin VB.Menu ArSat 
         Caption         =   "&Catalogo para el SAT"
      End
      Begin VB.Menu MayArSep5 
         Caption         =   "-"
      End
      Begin VB.Menu MayArSal 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "&Edicion"
      Begin VB.Menu EdiCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdiSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdiSelT 
         Caption         =   "&Seleccionar Todo"
      End
      Begin VB.Menu EdiSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdiTit 
         Caption         =   "&Eliminar Titulos"
      End
   End
   Begin VB.Menu Ci 
      Caption         =   "C&ierre Anual"
      Begin VB.Menu Ci_e 
         Caption         =   "&Cierre"
      End
      Begin VB.Menu CiSep1 
         Caption         =   "-"
      End
      Begin VB.Menu CiSale 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu Meses 
      Caption         =   "&Mes"
      Begin VB.Menu MesE 
         Caption         =   "&Incorporacion"
         Index           =   0
      End
      Begin VB.Menu MesE 
         Caption         =   "&Enero"
         Index           =   1
      End
      Begin VB.Menu MesE 
         Caption         =   "&Febrero"
         Index           =   2
      End
      Begin VB.Menu MesE 
         Caption         =   "&Marzo"
         Index           =   3
      End
      Begin VB.Menu MesE 
         Caption         =   "&Abril"
         Index           =   4
      End
      Begin VB.Menu MesE 
         Caption         =   "M&ayo"
         Index           =   5
      End
      Begin VB.Menu MesE 
         Caption         =   "&Junio"
         Index           =   6
      End
      Begin VB.Menu MesE 
         Caption         =   "J&ulio"
         Index           =   7
      End
      Begin VB.Menu MesE 
         Caption         =   "A&gosto"
         Index           =   8
      End
      Begin VB.Menu MesE 
         Caption         =   "&Septiembre"
         Index           =   9
      End
      Begin VB.Menu MesE 
         Caption         =   "&Octubre"
         Index           =   10
      End
      Begin VB.Menu MesE 
         Caption         =   "&Noviembre"
         Index           =   11
      End
      Begin VB.Menu MesE 
         Caption         =   "&Diciembre"
         Index           =   12
      End
   End
   Begin VB.Menu Ven1 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu VenAi 
         Caption         =   "Arreglar Iconos"
      End
   End
End
Attribute VB_Name = "LIBROSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ArSat_Click()
   If BALANZON <> 4 Then
       MsgBox "Necesita operar con la balanza con Movimientos de Subcuentas", vbCritical
       Else
       
       SAT1.Show
    
   End If
   
End Sub

Private Sub Ci_e_Click()
     If BALANZON = 5 Then
         'LIBROSV.Ci.Visible = True
         Diario.Diario_2
         Printer.PaperSize = 1
         Printer.Orientation = 1
         Rem Mayor.tit_mayor
         F_olio = 0
         Mayor.Impr_may
         F_olioM = 0
     End If
End Sub

Private Sub CiSale_Click()
   Unload LIBROSV
   Close
   End
   
End Sub

Private Sub EdiCop_Click()
 Dim Temporal1
  Select Case BALANZON
   Case 1
   Clipboard.Clear
   difer = Balanza.Blz.RowSel - Balanza.Blz.Row
   Temporal1 = LIBROSV.Caption & Chr(13) & Chr(10)
   For I = Balanza.Blz.Row To Balanza.Blz.RowSel
      Rem For f = 0 To balanza.blz.ColSel
      For F = Balanza.Blz.Col To Balanza.Blz.ColSel
            Temporal1 = Temporal1 + Balanza.Blz.TextMatrix(I, F)
            If F < Balanza.Blz.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next F
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
   Next I
   Clipboard.SetText Temporal1
   Balanza.Blz.FixedCols = 2
   Balanza.Blz.FixedRows = 1
   Case 2
        Clipboard.Clear
        Temporal1 = LIBROSV.Caption & Chr(13) & Chr(10)
        difer = Diario.Diario1.RowSel - Diario.Diario1.Row
        For I = Diario.Diario1.Row To Diario.Diario1.RowSel
        
            For F = 0 To Diario.Diario1.ColSel
                Temporal1 = Temporal1 + Diario.Diario1.TextMatrix(I, F)
                If F < Diario.Diario1.ColSel Then
                    Temporal1 = Temporal1 & Chr(9)
                End If
            Next F
         Temporal1 = Temporal1 & Chr(13) & Chr(10)
        Next I
        Clipboard.SetText Temporal1
        Diario.Diario1.FixedCols = 3
        Diario.Diario1.FixedRows = 1
  Case 3
   Clipboard.Clear
   Rem  Temporal1 = LIBROSV.Caption & Chr(13) & Chr(10)
   difer = TrSdo.TrS1.RowSel - TrSdo.TrS1.Row
   For I = TrSdo.TrS1.Row To TrSdo.TrS1.RowSel
      
      For F = 0 To TrSdo.TrS1.ColSel
            Temporal1 = Temporal1 + TrSdo.TrS1.TextMatrix(I, F)
            If F < TrSdo.TrS1.ColSel Then
                Temporal1 = Temporal1 + Chr(9)
             End If
      Next F
      
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next I
   
    Clipboard.SetText Temporal1
  Case 4
   Clipboard.Clear
   Rem  Temporal1 = LIBROSV.Caption & Chr(13) & Chr(10)
   difer = Balancha.Bcha1.RowSel - Balancha.Bcha1.Row
   For I = Balancha.Bcha1.Row To Balancha.Bcha1.RowSel
      
      For F = 0 To Balancha.Bcha1.ColSel
            Temporal1 = Temporal1 + Balancha.Bcha1.TextMatrix(I, F)
            If F < Balancha.Bcha1.ColSel Then
                Temporal1 = Temporal1 + Chr(9)
             End If
      Next F
      
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next I
   
    Clipboard.SetText Temporal1
    Case 5
    Clipboard.Clear
   Rem  Temporal1 = LIBROSV.Caption & Chr(13) & Chr(10)
   difer = Mayor.MAYOR1.RowSel - Mayor.MAYOR1.Row
   For I = Mayor.MAYOR1.Row To Mayor.MAYOR1.RowSel
      
      For F = Mayor.MAYOR1.Col To Mayor.MAYOR1.ColSel
            Temporal1 = Temporal1 + Mayor.MAYOR1.TextMatrix(I, F)
            If F < Mayor.MAYOR1.ColSel Then
                Temporal1 = Temporal1 + Chr(9)
             End If
      Next F
      
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next I
   
    Clipboard.SetText Temporal1

   Case 6
    Clipboard.Clear
   Rem  Temporal1 = LIBROSV.Caption & Chr(13) & Chr(10)
   difer = SAT1.HcSat.RowSel - SAT1.HcSat.Row
   For I = SAT1.HcSat.Row To SAT1.HcSat.RowSel
      
      For F = SAT1.HcSat.Col To SAT1.HcSat.ColSel
            Temporal1 = Temporal1 + SAT1.HcSat.TextMatrix(I, F)
            If F < SAT1.HcSat.ColSel Then
                Temporal1 = Temporal1 + Chr(9)
             End If
      Next F
      
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next I
   
    Clipboard.SetText Temporal1
   
   
End Select
sale1:
End Sub

Private Sub EdiSelT_Click()
   Select Case BALANZON
   Case 1
    Balanza.Blz.Row = 0: Balanza.Blz.Col = 0
    Balanza.Blz.RowSel = Balanza.Blz.Rows - 1
    Balanza.Blz.ColSel = Balanza.Blz.Cols - 1
   Case 2
    Diario.Diario1.Row = 0: Diario.Diario1.Col = 0
    Diario.Diario1.RowSel = Diario.Diario1.Rows - 1
    Diario.Diario1.ColSel = Diario.Diario1.Cols - 1
   Case 3
    TrSdo.TrS1.Row = 0: TrSdo.TrS1.Col = 0
    TrSdo.TrS1.RowSel = TrSdo.TrS1.Rows - 1
    TrSdo.TrS1.ColSel = TrSdo.TrS1.Cols - 1
   Case 4
    Balancha.Bcha1.Row = 0: Balancha.Bcha1.Col = 0
    Balancha.Bcha1.RowSel = Balancha.Bcha1.Rows - 1
    Balancha.Bcha1.ColSel = Balancha.Bcha1.Cols - 1
    Case 5
    Mayor.MAYOR1.Row = 0: Mayor.MAYOR1.Col = 0
    Mayor.MAYOR1.RowSel = Mayor.MAYOR1.Rows - 1
    Mayor.MAYOR1.ColSel = Mayor.MAYOR1.Cols - 1
    Case 6
    SAT1.HcSat.Row = 0: SAT1.HcSat.Col = 0
    SAT1.HcSat.RowSel = SAT1.HcSat.Rows - 1
    SAT1.HcSat.ColSel = SAT1.HcSat.Cols - 1
   End Select
End Sub

Private Sub EdiTit_Click()
  Select Case BALANZON
   Case 1
    Balanza.Blz.FixedCols = 0
    Balanza.Blz.FixedRows = 0
   Case 2
    Diario.Diario1.FixedCols = 0: Diario.Diario1.FixedRows = 0
   Case 3
    TrSdo.TrS1.FixedCols = 0: TrSdo.TrS1.FixedRows = 0
   Case 4
    Balancha.Bcha1.FixedCols = 0: Balancha.Bcha1.FixedRows = 0
   Case 5
    Mayor.MAYOR1.FixedCols = 0: Mayor.MAYOR1.FixedRows = 0
   Case 6
   SAT1.HcSat.FixedCols = 0: SAT1.HcSat.FixedRows = 0
  End Select
End Sub

Private Sub MayArBal_Click()
  Balanza.Show
End Sub

Private Sub MayArDia_Click()
   Diario.Show
End Sub

Private Sub MayArDir_Click()
   Directorio.Show 1
   
   Close 3
   Open "C:\GconTA\Gcont.Arr" For Random As 3 Len = Len(SCont)
   Get 3, 1, SCont
   SCont.guarda = Trim(Sub_dir)
   
   ChDir SCont.guarda
   Put 3, 1, SCont
   Close 3
   Close 1
   Unload Directorio
   Open Sub_dir + "DATOS" For Random As 1 Len = Len(Datos)
   Get 1, 1, Datos
   Unload Balanza
   Unload Diario
   Unload TrSdo
   LIBROSV.Caption = Trim(Datos.D1) + " Actualizacion libros "
   Unload Balancha
   Unload Mayor
   MDIForm_Load
   'LIBROSV.Caption = Trim(Datos.D1) + " Actualizacion libros "
   'Balanza.Tot_Act
   'Balanza.DIBUJA final
   'LIBROS.Ape_Operaciones
   'LIBROS.Ide_Mayor
   'LIBROS.Suma_Deb
   'If DIARIO.DIARIO1.Cols > 3 Then DIARIO.DIARIO1.FixedCols = 3
   'LIBROS.Caption = "Movimientos de libro diario del mes de " + Trim(Mm(final)) + " de " + Datos.a_o
   'Close 1
End Sub

Private Sub MayArImp_Click()
    Select Case BALANZON
        Case 1
          Balanza.impre_Balnza
        Case 5
          
          Mayor.imp_sora
    End Select
    
End Sub

Private Sub MayArSal_Click()
   Unload LIBROSV
   Close
   End
End Sub

Private Sub MDIForm_Load()
   Mm(0) = "Incorporacion": Mm(1) = "Enero": Mm(2) = "Febrero": Mm(3) = "Marzo": Mm(4) = "Abril"
   Mm(5) = "Mayo": Mm(6) = "Junio": Mm(7) = "Julio": Mm(8) = "Agosto"
   Mm(9) = "Septiembre": Mm(10) = "Octubre": Mm(11) = "Noviembre": Mm(12) = "Diciembre"
   Ide_ti
   Ci.Visible = False
   Balanza.Show
   Diario.Show
   TrSdo.Show
   LIBROSV.Caption = Trim(Datos.D1) + " Actualizacion libros " + Datos.a_o
   Balancha.Show
   Mayor.Show
   
End Sub
Sub Ide_ti()
   On Error Resume Next
    Open "C:\GconTA\Gcont.Arr" For Random As 3 Len = Len(SCont)
    Get 3, 1, SCont
    If Trim(SCont.guarda) <= "" Then
        Directorio.Show 1
        Unload Directorio
        
        Exit Sub
        Else
        Sub_dir = Trim(SCont.guarda)
        MsgBox "El directorio es : " + Sub_dir, vbCritical
         Aprt
         Open Sub_dir + "DATOS" For Random As 1 Len = Len(Datos)
         Get 1, 1, Datos
         
         Exit Sub
     End If
salta:
    
    Close 3
End Sub
Sub Aprt()
   On Error GoTo Pelas
        ChDir Sub_dir
   
   Exit Sub
Pelas:
   ChDir "C:"
   MsgBox "Esa Ruta de archivo ya no esta Disponible", vbCritical
   Directorio.Show 1
End Sub

Private Sub MDIForm_Terminate()
    Unload LIBROSV
    Close
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Close
    End
End Sub

Private Sub MesE_Click(INDEX As Integer)
      
      Select Case BALANZON
        Case 1
        final = INDEX
        Balanza.DIBUJA final
        Case 2
        final = INDEX
        Diario.Ape_Operaciones
        Diario.Ide_Mayor
        Diario.Suma_Deb
        Diario.Diario1.FixedCols = 3
        Diario.Caption = "Movimientos de Mayor del mes de " + Trim(Mm(final)) + " de " + Datos.a_o
        Case 3
        final = INDEX
        TrSdo.Caption = "Saldos de Mayor al mes de " + Trim(Mm(final)) + " de " + Datos.a_o
        TrSdo.Sdotr final
        Rem  no se si se vea
        Case 4
        final = INDEX
        Balancha.Caption = "Movimientos de Mayor y auxiliares del mes de " + Trim(Mm(final)) + " de " + Datos.a_o
        Balancha.DIBUJA final
        mes_lim1 = final
        Rem Balancha.Ide_Mayor
        Rem Balancha.Suma_Deb
        Rem Balancha.Diario1.FixedCols = 3

      End Select
    
End Sub

