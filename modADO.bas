Attribute VB_Name = "modADO"
''
' Módulo com funções para manipulação de conexões e recordsets do Microsoft ADO.
'
' Escrito por Christian Haagensen Gontijo, abril de 2005.
'

Option Explicit

''
' Abre uma conexão ao banco de dados.
'
' @param sStrConexao String de conexão ao banco de dados.
'
' @return Objeto "Connection" do ADO com a referência à conexão feita.
'
Public Function AbreCn(ByVal sStrConexao As String) As ADODB.Connection

    Dim cn As ADODB.Connection: Set cn = New ADODB.Connection

    If sStrConexao = "" Then Exit Function
    cn.ConnectionString = sStrConexao

'    '<DEBUG>
'    On Error Resume Next
'    '</DEBUG>

    cn.Open

'    '<DEBUG>
'    If Err.Number <> 0 Then
'        MsgBox "Erro em modADO::AbreCn: " & Error$
'        Stop
'    End If
'    '</DEBUG>

    
    Set AbreCn = cn

End Function

''
' Cria recordset a partir de dada frase SQL.
'
' @param sFraseSQL Frase com sentença SQL que populará o recordset.
' @param sStrConexao String de conexão ao banco.
' @param CursorLocation Local do cursor. Por padrão, usa um cursor do lado do cliente.
' @param CursorType Tipo do cursor. Por padrão, usa um cursor estático.
' @param Options Opções para o recordset.
'
' @return Objeto Recordset populado.
'
' @remarks Se o cursor especificado for cliente, o recordset é desconectado da fonte de dados assim que populado, de
'   forma a minimizar a carga no servidor. Se for necessário o uso de um recordset conectado, utilize a função
'   "AbreRsConectado".
'
Public Function AbreRs(ByVal sFraseSQL As String, _
                       ByVal sStrConexao As String, _
                       Optional ByVal CursorLocation As CursorLocationEnum = adUseClient, _
                       Optional ByVal CursorType As CursorTypeEnum = adOpenStatic, _
                       Optional ByVal Options As CommandTypeEnum = adCmdText) As ADODB.Recordset

    ' Sem "On Error Goto..." aqui. Erros serão gerados na rotina que chamou esta função.

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset

    ' Abre conexão ao banco
    Set cn = AbreCn(sStrConexao)

    ' Gera recordset com o cursor especificado
    Set rs = New ADODB.Recordset
    rs.CursorLocation = CursorLocation

    ' Se o cursor está do lado do cliente, desconecta-o e fecha conexão ao banco
    If CursorLocation = adUseClient Then
        rs.Open sFraseSQL, cn, CursorType, adLockBatchOptimistic, Options
        Set rs.ActiveConnection = Nothing
        cn.Close: Set cn = Nothing
    Else
        ' cursor do lado do servidor
        rs.Open sFraseSQL, cn, CursorType, adLockOptimistic, Options
    End If

    Set AbreRs = rs

End Function

''
' Cria recordset "firehose" a partir de dada frase SQL.
'
' @param sFraseSQL Frase com sentença SQL que populará o recordset.
' @param sStrConexao String de conexão ao banco.
' @param Options Opções para o recordset.
'
' @return Objeto Recordset populado.
'
' @remarks O recordset "firehose" obtém dados com a maior rapidez possível no ADO. São cursores do lado do
'   servidor, apenas para frente, e somente leitura.
'
Public Function AbreFirehoseRs(ByVal sFraseSQL As String, _
                               ByVal sStrConexao As String, _
                               Optional ByVal Options As CommandTypeEnum = adCmdText) As ADODB.Recordset

    ' Sem "On Error Goto..." aqui. Erros serão gerados na rotina que chamou esta função.

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset

    ' Abre conexão ao banco
    Set cn = AbreCn(sStrConexao)

    ' Gera recordset em modo "firehose"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sFraseSQL, cn, adOpenForwardOnly, adLockReadOnly, Options

    Set AbreFirehoseRs = rs

End Function

''
' Abre um recordset, mantendo a conexão com o banco.
'
' @param sFraseSQL Frase SQL a ser executada, ou tabela a abrir, no banco de dados.
' @param cn Objeto "connection" do ADO, já contendo a conexão ao banco, a ser usado pelo recordset.
'
' @return Objeto "Recordset" criado.
'
Public Function AbreRsConectado(ByVal sFraseSQL As String, _
                                ByVal cn As ADODB.Connection) As ADODB.Recordset

    Dim rs As ADODB.Recordset

    If Not cn Is Nothing Then

        Set rs = New ADODB.Recordset
        rs.Open sFraseSQL, cn, adOpenKeyset, adLockOptimistic
    
        Set AbreRsConectado = rs

    End If

End Function

''
' Executa, via ADO, a frase SQL informada.
'
' @param sFraseSQL Frase SQL a ser executada.
' @param sStrConexao String de conexão ao banco.
'
' @remarks A função cria, temporariamente, uma conexão ao banco para execução do pedido.
'
Public Sub ExecutaSQL(ByVal sFraseSQL As String, ByVal sStrConexao As String)

    ' Erros não devem ser tratados aqui, mas na função chamadora.

    Dim cn As ADODB.Connection
    
    Set cn = AbreCn(sStrConexao)
    cn.Execute sFraseSQL
    cn.Close
    Set cn = Nothing

End Sub

''
' Retorna a descrição dos erros encontrados numa operação nos dados através do ADO.
'
' @param errCol Coleção "Errors" do ADO, contendo objetos de erro de onde as descrições serão obtidas.
'
' @return Texto com a descrição dos erros.
'
' @remarks O erro NÃO é "limpo" na saída do procedimento.
'
Public Function AdoError(ByVal errCol As ADODB.Errors) As String

    Dim hc As Long, hf As String, num As Long, descr As String, src As String
    Dim sMsg As String
    Dim errObj As ADODB.Error

    ' guarda antes que se perca :)
    descr = Err.Description
    hc = Err.HelpContext
    hf = Err.HelpFile
    num = Err.Number
    src = Err.Source

    On Error Resume Next

    For Each errObj In errCol
        sMsg = sMsg & vbNewLine & errObj.NativeError & ": " & errObj.Description
    Next
    If Len(sMsg) > Len(vbNewLine) Then sMsg = Mid$(sMsg, Len(vbNewLine) + 1)

    AdoError = sMsg

    ' Retorna erro original
    Err.Description = descr
    Err.HelpContext = hc
    Err.HelpFile = hf
    Err.Number = num
    Err.Source = src

End Function

''
' Obtém o nome de cada campo da tabela que faça parte de determinado índice.
'
' @param sNomeTabela Nome da tabela a verificar.
' @param sNomeIndice Nome do índice desejado.
' @param sStrConexao String de conexão ao banco.
'
' @return Vetor contendo, em cada elemento, o nome de  um dos campos que forma o índice em questão.
'
Public Function ObtemNomeColunasIndice(ByVal sNomeTabela As String, ByVal sNomeIndice As String, ByVal sStrConexao As String) As Variant

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim vsNomes() As String

    On Error GoTo Erro

    ReDim vsNomes(0 To 100)

    Set cn = AbreCn(sStrConexao)
    Set rs = cn.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, sNomeTabela))
    Do While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            If rs.Fields(i).Name = "INDEX_NAME" Then
                If rs.Fields(i).Value = sNomeIndice Then
                    vsNomes(j) = rs.Fields("COLUMN_NAME").Value
                    j = j + 1
                End If
            End If
        Next
        rs.MoveNext
    Loop

    If j > 0 Then j = j - 1
    ReDim Preserve vsNomes(0 To j)

    ObtemNomeColunasIndice = vsNomes

    Exit Function

Erro:
Resume Fecha
Fecha:
Exit Function

End Function

''
' Obtém os valores atuais dos campos de um recordset.
'
' @param rs Recordset a ser verificado. Deve estar aberto e posicionado no registro desejado.
'
' @param Vetor contendo, em cada elemento, o valor atual de cada campo do recordset em questão.
'
Public Function CamposObtemRegistroAtual(ByVal rs As ADODB.Recordset) As Variant

    Dim i As Long
    ReDim vetCampos(rs.Fields.Count) As String

    On Error GoTo Erro

    For i = 0 To rs.Fields.Count - 1
        vetCampos(i) = IIf(IsNull(rs.Fields(i).Value), "", rs.Fields(i).Value)
    Next

    CamposObtemRegistroAtual = vetCampos

    Exit Function

Erro:
Resume Fecha
Fecha:
Exit Function

End Function

''
' Compara os valores dos campos de um recordset com os de um vetor, retornando uma string informando as diferenças.
'
' @param rs Recordset a ser verificado. Deve estar aberto e posicionado no registro desejado.
' @param CamposBD Vetor contendo, em cada elemento, um determinado valor que será comparado com o do recorset.
'
' @remarks Tipicamente, o vetor passado como parâmetro será populado através da função "CamposObtemRegistroAtual".
'   <p>O retorno da função pode ser usado para propósitos de "Log".
'
Public Function CamposObtemAlteracoes(ByVal rs As ADODB.Recordset, _
                                      CamposBD() As String) As String

    Dim i As Long
    Dim sCampo As String

    On Error GoTo Fecha

    For i = 0 To rs.Fields.Count - 1
        If (rs.Fields(i).Value & "") <> CamposBD(i) Then
            sCampo = sCampo & _
                     rs.Fields(i).Name & ": de [" & CamposBD(i) & _
                     "] para [" & _
                     (rs.Fields(i).Value & "") & "]" & _
                     vbNewLine
        End If
    Next

    CamposObtemAlteracoes = sCampo

    Exit Function

Erro:
Resume Fecha
Fecha:
Exit Function

End Function

Private Sub ajustaDefaults(rs As ADODB.Recordset)

    Dim bAchouIdentity As Boolean
    Dim i As Long
    Dim j As Long

    For i = 0 To rs.Fields.Count - 1

        '
        ' Salta campos identity, porque estes campos não podem receber valores via aplicativo.
        '
        If Not bAchouIdentity Then ' só há um por tabela
            For j = 0 To rs.Fields(i).Properties.Count - 1
                If rs.Fields(i).Properties(j).Name = "ISAUTOINCREMENT" And _
                   rs.Fields(i).Properties(j).Value = True Then
                    bAchouIdentity = True
                    GoTo proximo
                End If
            Next
        End If
    
        Select Case rs.Fields(i).Type
            Case adDate, adDBDate, adDBTime, adDBTimeStamp, adFileTime
                rs.Fields(i).Value = CDate("1980-01-01")
            Case adVarChar, adChar, adVarWChar, adWChar
                rs.Fields(i).Value = ""
            Case adNumeric, adTinyInt, adInteger, adBigInt, _
                 adCurrency, adDecimal, adDouble, adSingle, _
                 adSmallInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                rs.Fields(i).Value = 0
            Case adBoolean
                rs.Fields(i).Value = False
            Case Else
                MsgBox "Erro em AjustaDefaults: tipo nº " & rs.Fields(i).Type & " não reconhecido.", vbCritical, "Erro"
        End Select
proximo:
    Next

End Sub
