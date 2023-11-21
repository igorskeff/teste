
Option Explicit

Option Explicit

Option Explicit

Option Explicit

Option Explicit

Private Sub exemplo1()

End Sub

Private Sub exemplo2()
    MsgBox "teste2"
End Sub

Private Sub buttonCadastrar_Click()
    '1. criar as variaveis
    Dim codigo As Long, descricao As String, categoria As String
    Dim valor As Currency, qtdEstoque As Integer, linha As Integer
    Dim valorTotal As Currency
    
    '2. inspecionar os preenchimentos
    If Not IsNumeric(textCodigo.Text) Then
        MsgBox "Favor preencher corretamente o campo código"
        Exit Sub
    End If
    
    If textDescricao.Text = "" Then
        MsgBox "Favor preencher o campo Descrição"
        Exit Sub
    End If
    
    If comboCategoria.Text = "" Then
        MsgBox "Favor preencher o campo Categoria"
        Exit Sub
    End If
    
     If Not IsNumeric(textValor.Text) Then
        MsgBox "Favor preencher corretamente o campo valor"
        Exit Sub
    End If
    
     If Not IsNumeric(textQtdEstoque.Text) Then
        MsgBox "Favor preencher corretamente o campo quantidade em estoque"
        Exit Sub
     End If
     
     '3. passar os dados do formulario para as variaves
     codigo = textCodigo.Text
     descricao = textDescricao.Text
     categoria = comboCategoria.Text
     valor = textValor.Text
     qtdEstoque = textQtdEstoque.Text
     
     '4. calcular o valor total
     valorTotal = valor * qtdEstoque
     
     '5. pegar a linha da planilha de controle
     linha = PlanControle.Range("A2").Value
     
     '6. passar os dados das variáveis para a planilha de produtos
     PlanProdutos.Cells(linha, 1).Value = codigo
     PlanProdutos.Cells(linha, 2).Value = descricao
     PlanProdutos.Cells(linha, 3).Value = categoria
     PlanProdutos.Cells(linha, 4).Value = valor
     PlanProdutos.Cells(linha, 5).Value = qtdEstoque
     PlanProdutos.Cells(linha, 6).Value = valorTotal
     
     '7. mudar a numeracao da linha
     linha = linha + 1
     PlanControle.Range("A2").Value = linha
     
     '8. Limpar os dados do formulario
     textCodigo.Text = ""
     textDescricao.Text = ""
     comboCategoria.Text = ""
     textQtdEstoque.Text = ""
     textValor.Text = ""
     
     '9. colocar o foco no primeiro controle
     textCodigo.SetFocus
     
     MsgBox "Produto cadastrado com sucesso", vbInformation, "Sucesso"
        
    
End Sub

Private Sub buttonSair_Click()
    'fechar o formulario
    Unload Me
End Sub

Private Sub UserForm_Activate()
    'armazem de secos e molhados
    'duas opções secos e molhados
    comboCategoria.AddItem "Secos"
    comboCategoria.AddItem "Molhados"
End Sub
Option Explicit

Sub AbrirFormulario()
    FormProdutos.Show
End Sub


Sub FecharPrograma()
    Application.Quit
End Sub
Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub
 
Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
 
    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub
 
Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.codeModule.Lines(1, item.codeModule.CountOfLines)
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub
 
Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub
 
Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean
 
    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next
 
    Debug.Print "Exported at " & pathToExport
    
End Sub
 
Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub
 
CreateLogFile_Error:
 
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"
 
End Sub
