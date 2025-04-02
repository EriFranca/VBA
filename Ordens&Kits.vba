'Autor: Fabio Mitsueda
'Contato: mitsueda.fabio@gmail.com
'Data Criação: 12/08/2018
Option Explicit
Sub Executar()
    'Variáveis do SAP
    Dim DLin As Long
    Dim i As Long, j As Long, col_ctn As Long, lastRow As Long
    Dim colOrd_arr As Variant, SAPTable_arr As Variant
    Dim SAPTable_cell As String
    Dim branchValue As String
    Dim inputDate As String, sapUser As String
    Dim currentPage As Long, totalPages As Long
    Dim actualDate As Date
    
    
    'session.findById("wnd[0]").iconify
    session.findById("wnd[0]").MAXIMIZE
    session.findById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
    session.findById("wnd[0]").sendVKey 0
    'CABEÇALHOS ORDENS
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOH000"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    'LAYOUT
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&LOAD"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell


    lastRow = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount
    col_ctn = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").ColumnCount
    ReDim colOrd_arr(0 To col_ctn - 1)
    
    For j = 0 To col_ctn - 1
        colOrd_arr(j) = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").columnOrder.Item(j)
    Next

    ReDim SAPTable_arr(0 To lastRow - 1, 0 To col_ctn - 1)
    
    For j = 0 To col_ctn - 1
        For i = 0 To lastRow - 1
            SAPTable_cell = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").getCellValue(i, colOrd_arr(j))
            SAPTable_arr(i, j) = SAPTable_cell
        Next
    Next
    
    ' Paginar para garantir que todos os dados sejam carregados
totalPages = lastRow \ 20 ' Número total de páginas completas
If lastRow Mod 20 > 0 Then ' Se houver linhas extras que não preenchem uma página completa
    totalPages = totalPages + 1 ' Adicionar uma página extra para as linhas restantes
End If

For currentPage = 0 To totalPages - 1 ' -1 porque o índice começa em 0
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellRow = currentPage * 20
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = currentPage * 20
    'session.findById("wnd[0]").sendVKey 2 ' Pagina para baixo
Next currentPage
    
    ' Copiar todos os dados carregados
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectAll
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItemByPosition "0"
    
    Sheets("cabeçalho").Select
    Range("A2:H1800").Select
    Selection.ClearContents
    
    ' Obter e lançar o usuário SAP na célula R1
'    sapUser = session.Info.User
'    ThisWorkbook.Sheets("CABEÇALHO").Range("I1").Value = sapUser
        
     ' Colar dados
    Call ColarDados1
        
        
    Application.Wait Now + TimeValue("00:00:2")
        
    'COMPONENTES
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    'LAYOUT
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&LOAD"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell


    lastRow = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount
    col_ctn = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").ColumnCount
    ReDim colOrd_arr(0 To col_ctn - 1)
    
    For j = 0 To col_ctn - 1
        colOrd_arr(j) = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").columnOrder.Item(j)
    Next
    
    ReDim SAPTable_arr(0 To lastRow - 1, 0 To col_ctn - 1)
    
    For j = 0 To col_ctn - 1
        For i = 0 To lastRow - 1
            SAPTable_cell = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").getCellValue(i, colOrd_arr(j))
            SAPTable_arr(i, j) = SAPTable_cell
        Next
    Next
    
    ' Paginar para garantir que todos os dados sejam carregados
totalPages = lastRow \ 20 ' Número total de páginas completas
If lastRow Mod 20 > 0 Then ' Se houver linhas extras que não preenchem uma página completa
    totalPages = totalPages + 1 ' Adicionar uma página extra para as linhas restantes
End If

For currentPage = 0 To totalPages - 1 ' -1 porque o índice começa em 0
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellRow = currentPage * 20
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = currentPage * 20

Next currentPage
    
    ' Copiar todos os dados carregados
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").SelectAll
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItemByPosition "0"
    
    Sheets("componentes").Select
    Range("A2:H1800").Select
    Selection.ClearContents
    
    ' Obter e lançar o usuário SAP na célula R1
'    sapUser = session.Info.User
'    ThisWorkbook.Sheets("COMPONENTES").Range("I1").Value = sapUser
    
    session.findById("wnd[0]").Close ' Fechar tela de relatório
    
    ' Colar dados
    Call ColarDados2
    
   Exit Sub
   
InvalidDate:
    MsgBox "Data inválida fornecida. Operação cancelada.", vbExclamation
    Exit Sub
    
    
End Sub