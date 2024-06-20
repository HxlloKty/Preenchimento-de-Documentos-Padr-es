Sub PreencherMinutaMapfre()
    'variáveis principais
    Dim objWord As Object
    Dim wdDoc As Document
    Dim caminhoTemplate As String
    Dim caminhoDoc As String
    Dim caminhoPDF As String
    
    'atribuir planilha
    Set objPlanilha = ActiveWorkbook
    
    'caminhos docs
    caminhoTemplate = ThisWorkbook.Path & "\Modelos\Modelo Minuta Mapfre.docx"
    caminhoDoc = Dir("ThisWorkbook.Path" & "\Minutas Prontas\Minuta - " & "*.docx", vbNormal)
    caminhoPDF = ThisWorkbook.Path & "\Minutas Prontas\Minuta - " & Cells(10, 1).Value & " - " & Cells(2, 1) & Replace(docxFile, ".docx", ".pdf")
    
    'atribuir Word
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    Set wdDoc = objWord.Documents.Open(caminhoTemplate)
    
    'variáveis terceiro
    Dim strTerceiro As String
    Dim strCpfTerceiro As String
    Dim strEnderecoTerceiro As String
    Dim strCidadeTerceiro As String
    Dim strCepTerceiro As String
    
    strTerceiro = objPlanilha.Sheets(1).Range("A2").Value
    strCpfTerceiro = objPlanilha.Sheets(1).Range("B2").Value
    strEnderecoTerceiro = objPlanilha.Sheets(1).Range("C2").Value
    strCidadeTerceiro = objPlanilha.Sheets(1).Range("D2").Value
    strCepTerceiro = objPlanilha.Sheets(1).Range("E2").Value
    
    'variáveis segurado
    Dim strSegurado As String
    Dim strCpfSegurado As String
    Dim strEnderecoSegurado As String
    Dim strCidadeSegurado As String
    Dim strCepSegurado As String
    
    strSegurado = objPlanilha.Sheets(1).Range("A4").Value
    strCpfSegurado = objPlanilha.Sheets(1).Range("B4").Value
    strEnderecoSegurado = objPlanilha.Sheets(1).Range("C4").Value
    strCidadeSegurado = objPlanilha.Sheets(1).Range("D4").Value
    strCepSegurado = objPlanilha.Sheets(1).Range("E4").Value
    
    'variáveis acidente
    Dim strData As String
    Dim strEnderecoAcidente As String
    Dim strCidadeAcidente As String
    Dim strValorMinuta As String
    Dim strValorExtenso As String
    Dim strSinistro As String

    strData = objPlanilha.Sheets(1).Range("A6").Value
    strEnderecoAcidente = objPlanilha.Sheets(1).Range("B6").Value
    strCidadeAcidente = objPlanilha.Sheets(1).Range("C6").Value
    strValorMinuta = objPlanilha.Sheets(1).Range("D6").Value
    strValorExtenso = objPlanilha.Sheets(1).Range("E6").Value
    strSinistro = objPlanilha.Sheets(1).Range("A10").Value
    
    'substituição
    With objWord.ActiveDocument.Content.Find
    
    'terceiro
    .Text = "TERCEIRO"
    .Replacement.Text = strTerceiro
    .Execute Replace:=2
    
    .Text = "000.000.000-00"
    .Replacement.Text = strCpfTerceiro
    .Execute Replace:=2
    
    .Text = "Endereço 1 (Rua, Número, Bairro)"
    .Replacement.Text = strEnderecoTerceiro
    .Execute Replace:=2
    
    .Text = "Cidade 1/ES"
    .Replacement.Text = strCidadeTerceiro
    .Execute Replace:=2
    
    .Text = "CEP 1"
    .Replacement.Text = strCepTerceiro
    .Execute Replace:=2
    
    'segurado
    .Text = "SEGURADO"
    .Replacement.Text = strSegurado
    .Execute Replace:=2
    
    .Text = "111.111.111-11"
    .Replacement.Text = strCpfSegurado
    .Execute Replace:=2
    
    .Text = "Endereço 2 (Rua, Número, Bairro)"
    .Replacement.Text = strEnderecoSegurado
    .Execute Replace:=2
    
    .Text = "Cidade 2/ES"
    .Replacement.Text = strCidadeSegurado
    .Execute Replace:=2
    
    .Text = "CEP 2"
    .Replacement.Text = strCepSegurado
    .Execute Replace:=2
    
    'acidente
    .Text = "Data do Acidente"
    .Replacement.Text = strData
    .Execute Replace:=2
    
    .Text = "Endereço Acidente (Rua, Número)"
    .Replacement.Text = strEnderecoAcidente
    .Execute Replace:=2
    
    .Text = "Cidade Acidente/ES"
    .Replacement.Text = strCidadeAcidente
    .Execute Replace:=2
    
    .Text = "Valor Minuta (R$)"
    .Replacement.Text = strValorMinuta
    .Execute Replace:=2
    
    .Text = "Valor Minuta (por extenso)"
    .Replacement.Text = strValorExtenso
    .Execute Replace:=2
    
    .Text = "N° DO SINISTRO"
    .Replacement.Text = strSinistro
    .Execute Replace:=2
    
    End With
     
    'salvar em pdf e docx
    wdDoc.SaveAs2 caminhoPDF, FileFormat:=17
        
    wdDoc.ExportAsFixedFormat OutputFileName:=caminhoPDF, ExportFormat:=17
    
    MsgBox "Documento salvo!"

    wdDoc.SaveAs2 ThisWorkbook.Path & "\Minutas Prontas\Minuta - " & Cells(10, 1).Value & " - " & Cells(2, 1) & ".docx"
        
    wdDoc.Close False
    objWord.Quit
    
    Set wdDoc = Nothing
    Set objWord = Nothing
    Set objPlanilha = Nothing
    
End Sub
