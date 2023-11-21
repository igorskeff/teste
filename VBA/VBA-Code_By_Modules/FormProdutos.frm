VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProdutos 
   Caption         =   "Formulário de Produtos"
   ClientHeight    =   4320
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6960
   OleObjectBlob   =   "FormProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
