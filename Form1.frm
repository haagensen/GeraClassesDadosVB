VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gera classes Base e DAO"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "SQL Server"
      Height          =   195
      Index           =   1
      Left            =   4455
      TabIndex        =   2
      Top             =   180
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Oracle"
      Height          =   195
      Index           =   0
      Left            =   3195
      TabIndex        =   1
      Top             =   180
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton cmdDAO 
      Caption         =   "Gerar DAO"
      Height          =   465
      Left            =   6210
      TabIndex        =   11
      Top             =   3105
      Width           =   1275
   End
   Begin VB.TextBox txtStrCon 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   510
      Width           =   7365
   End
   Begin VB.CommandButton cmdColecao 
      Caption         =   "Gerar Coleção"
      Height          =   465
      Left            =   6210
      TabIndex        =   9
      Top             =   2475
      Width           =   1275
   End
   Begin VB.TextBox txtTabela 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1215
      Width           =   4545
   End
   Begin VB.CommandButton cmdBase 
      Caption         =   "Gerar Base"
      Height          =   465
      Left            =   6210
      TabIndex        =   7
      Top             =   1815
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gerar um arquivo de classe (xxx.CLS), contendo os métodos para acesso a dados (DAO) da tabela."
      Height          =   390
      Index           =   2
      Left            =   135
      TabIndex        =   10
      Top             =   3105
      Width           =   5865
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "String de conexão ao banco de dados:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gerar um arquivo de classe (xxx.CLS) para um arquivo de coleção para a classe-base criada anteriormente."
      Height          =   390
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   2475
      Width           =   5760
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   990
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gerar um arquivo de classe do VB6 (xxx.CLS), contendo as variáveis de instância, getters e setters para determinada tabela."
      Height          =   390
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   1830
      Width           =   5955
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Provider=OraOLEDB.Oracle;Data Source=master;User Id=TESTESOFFICEORACLE;Password=DISYS

Option Explicit

Private msNomeChavePrimaria() As String

Private Type TCampos
    Nome As String
    Tipo As String
    TipoADO As String
    Instancia As String
    Tamanho As Long
    ChavePrimaria As Boolean
End Type

Dim gsStrConexao As String

Private Sub ObtemChavePrimaria(ByVal sNomeTabela As String)

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo Erro

    Set cn = AbreCn(gsStrConexao)
    ' o UCASE abaixo é necessário no Oracle
    Set rs = cn.OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, UCase$(sNomeTabela)))
    
    Erase msNomeChavePrimaria
    
    Do While Not rs.EOF
        ReDim Preserve msNomeChavePrimaria(i)
        msNomeChavePrimaria(i) = rs("COLUMN_NAME")
        rs.MoveNext
        i = i + 1
    Loop

    GoTo Sai

Erro:
    MsgBox "Erro " & Err.Number & " em Form1::ObtemChavePrimaria: " & Err.Description, vbExclamation, "Erro"
GoTo Sai
Sai:
    On Error Resume Next
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing
Exit Sub

End Sub

Private Function ChavePrimariaComposta(tb() As TCampos) As Boolean

    On Error GoTo Erro

    Dim i As Long
    Dim lNumPKs As Long

    For i = 0 To UBound(tb)
        If tb(i).ChavePrimaria Then
            lNumPKs = lNumPKs + 1
        End If
    Next
    If lNumPKs > 1 Then ChavePrimariaComposta = True

    On Error GoTo 0
    Exit Function

Erro:
    MsgBox "Erro " & Err.Number & " em Form1::ChavePrimariaComposta: " & Err.Description, vbExclamation, "Erro"
GoTo Sai
Sai:
Exit Function

End Function

Private Function ObtemIndiceChavePrimaria(tb() As TCampos) As Long

    On Error GoTo Erro

    Dim i As Long

    ' Não funciona para chaves compostas!
    For i = 0 To UBound(tb)
        If tb(i).ChavePrimaria Then
            ObtemIndiceChavePrimaria = i
            Exit For
        End If
    Next

    On Error GoTo 0
    Exit Function

Erro:
    MsgBox "Erro " & Err.Number & " em Form1::ObtemIndiceChavePrimaria: " & Err.Description, vbExclamation, "Erro"
GoTo Sai
Sai:
Exit Function

End Function

Private Function ObtemDadosTabela(tb() As TCampos) As Boolean

    Dim i As Long
    Dim j As Long
    Dim sPesq As String
    Dim ub As Long
    Dim rs As ADODB.Recordset

    On Error GoTo Erro

    Screen.MousePointer = vbHourglass
    gsStrConexao = txtStrCon.Text
    
    If InStr(1, txtStrCon.Text, ".ORA", vbTextCompare) > 0 Then
        Set rs = AbreFirehoseRs("SELECT * FROM " & Trim$(txtTabela.Text), gsStrConexao)
    Else
        Set rs = AbreFirehoseRs("SELECT TOP(1) * FROM dbo." & Trim$(txtTabela.Text), gsStrConexao)
    End If
    ReDim tb(0 To rs.Fields.Count - 1) As TCampos
    
    Call ObtemChavePrimaria(Trim$(txtTabela.Text))

    On Error Resume Next
        ub = -1
        ub = UBound(msNomeChavePrimaria)
    On Error GoTo Erro

    For i = 0 To rs.Fields.Count - 1

        tb(i).Nome = rs.Fields(i).Name
        
        ' Chave primária?
        For j = 0 To ub
            If rs.Fields(i).Name = msNomeChavePrimaria(j) Then
                tb(i).ChavePrimaria = True
                Exit For
            End If
        Next

        Select Case rs.Fields(i).Type

            Case ADODB.adNumeric
                Select Case rs.Fields(i).Precision
                    Case Is < 5
                        tb(i).Tipo = "Long"
                        tb(i).Instancia = "ml" & rs.Fields(i).Name
                    Case Else
                        tb(i).Tipo = "Double"
                        tb(i).Instancia = "mdbl" & rs.Fields(i).Name
                End Select
                tb(i).TipoADO = "adNumeric" 'requer PRECISION e NUMERIC SCALE!!!
            Case ADODB.adDecimal
                tb(i).Tipo = "Double"
                tb(i).Instancia = "mdbl" & rs.Fields(i).Name
                tb(i).TipoADO = "adDecimal" 'requer PRECISION e NUMERIC SCALE!!!
            Case ADODB.adTinyInt
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adTinyInt"
            Case ADODB.adInteger
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adInteger"
            Case ADODB.adSmallInt
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adSmallInt"
            Case ADODB.adUnsignedInt
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adUnsignedInt"
            Case ADODB.adUnsignedSmallInt
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adUnsignedSmallInt"
            Case ADODB.adUnsignedTinyInt
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adUnsignedTinyInt"
            Case ADODB.adBigInt
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adBigInt"
            Case ADODB.adUnsignedBigInt
                tb(i).Tipo = "Long"
                tb(i).Instancia = "ml" & rs.Fields(i).Name
                tb(i).TipoADO = "adUnsignedBigInt"
            Case ADODB.adSingle
                tb(i).Tipo = "Double"
                tb(i).Instancia = "mdbl" & rs.Fields(i).Name
                tb(i).TipoADO = "adSingle"
            Case ADODB.adDouble
                tb(i).Tipo = "Double"
                tb(i).Instancia = "mdbl" & rs.Fields(i).Name
                tb(i).TipoADO = "adDouble"
            Case ADODB.adCurrency
                tb(i).Tipo = "Currency"
                tb(i).Instancia = "mc" & rs.Fields(i).Name
                tb(i).TipoADO = "adCurrency"
            
            Case ADODB.adDate
                tb(i).Tipo = "Date"
                tb(i).Instancia = "mdt" & rs.Fields(i).Name
                tb(i).TipoADO = "adDate"
            Case ADODB.adDBDate
                tb(i).Tipo = "Date"
                tb(i).Instancia = "mdt" & rs.Fields(i).Name
                tb(i).TipoADO = "adDBDate"
            Case ADODB.adDBTime
                tb(i).Tipo = "Date"
                tb(i).Instancia = "mdt" & rs.Fields(i).Name
                tb(i).TipoADO = "adDBTime"
            Case ADODB.adDBTimeStamp
                tb(i).Tipo = "Date"
                tb(i).Instancia = "mdt" & rs.Fields(i).Name
                tb(i).TipoADO = "adDBTimeStamp"
            
            Case ADODB.adVarChar, ADODB.adVarWChar, ADODB.adLongVarChar
                tb(i).Tipo = "String"
                tb(i).Instancia = "ms" & rs.Fields(i).Name
                tb(i).TipoADO = "adVarChar"
                tb(i).Tamanho = rs.Fields(i).DefinedSize
            Case ADODB.adChar, ADODB.adWChar
                tb(i).Tipo = "String"
                tb(i).Instancia = "ms" & rs.Fields(i).Name
                tb(i).TipoADO = "adChar"
                tb(i).Tamanho = rs.Fields(i).DefinedSize
            
            Case ADODB.adBoolean
                tb(i).Tipo = "Boolean"
                tb(i).Instancia = "mb" & rs.Fields(i).Name
                tb(i).TipoADO = "adBoolean"

            Case ADODB.adLongVarBinary 'blob,image
                tb(i).Tipo = "Variant"
                tb(i).Instancia = "mv" & rs.Fields(i).Name
                tb(i).TipoADO = "adLongVarBinary"
            
            Case Else
                tb(i).Tipo = "Variant"
                tb(i).Instancia = "mv" & rs.Fields(i).Name
                tb(i).TipoADO = rs.Fields(i).Type

        End Select

    Next
    rs.Close: Set rs = Nothing

    ObtemDadosTabela = True

Exit Function

Erro:
    MsgBox "Erro: " & Error$, vbCritical, "Erro"
Resume Fecha
Fecha:
    On Error Resume Next
    Screen.MousePointer = vbDefault
    rs.Close: Set rs = Nothing
Exit Function

End Function

Private Function MeuStrConv(ByVal sNome As String) As String

    Dim pos As Integer

    sNome = LCase$(sNome)

    Mid$(sNome, 1, 1) = UCase$(Mid$(sNome, 1, 1))

    If Left$(sNome, 1) = "C" Then
        If Len(sNome) > 1 Then
            Mid$(sNome, 2, 1) = UCase$(Mid$(sNome, 2, 1))
        End If
    End If

    Do While True
        pos = InStr(pos + 1, sNome, "_")
        If pos > 0 Then
            If pos < Len(sNome) Then
                Mid$(sNome, pos + 1, 1) = UCase$(Mid$(sNome, pos + 1, 1))
            End If
        Else
            Exit Do
        End If
    Loop

    MeuStrConv = sNome

End Function

Private Sub Cabecalho(ByVal f As Integer, ByVal sNome As String)

    Print #f, "VERSION 1.0 CLASS"
    Print #f, "BEGIN"
    Print #f, "  MultiUse = -1  'True"
    #If VB6 Then
        Print #f, "  Persistable = 0  'NotPersistable"
        Print #f, "  DataBindingBehavior = 0  'vbNone"
        Print #f, "  DataSourceBehavior = 0   'vbNone"
        Print #f, "  MTSTransactionMode = 0   'NotAnMTSObject"
    #End If
    Print #f, "End"
    Print #f, "Attribute VB_Name = """ & sNome & """"
    Print #f, "Attribute VB_GlobalNameSpace = False"
    Print #f, "Attribute VB_Creatable = True"
    Print #f, "Attribute VB_PredeclaredId = False"
    Print #f, "Attribute VB_Exposed = True"

    Print #f, "Option Explicit"
    Print #f, ""

End Sub

Private Sub GeraBase(tb() As TCampos)

    Dim i As Long
    Dim f As Integer
    Dim arq As String
    Dim ub As Long

    On Error Resume Next
        ub = UBound(tb)
        If Err.Number <> 0 Then
            MsgBox "Erro: não há dados.", vbCritical, "Erro"
            Exit Sub
        End If
    On Error GoTo 0

    Screen.MousePointer = vbHourglass

    f = FreeFile
    arq = Environ$("TEMP") & "\C" & MeuStrConv(txtTabela.Text) & ".cls"
    Open arq For Output As #f

    Cabecalho f, "C" & MeuStrConv(txtTabela.Text)

    ' Variáveis de instância
    For i = 0 To UBound(tb)
        Print #f, "Private " & tb(i).Instancia & " As " & tb(i).Tipo & IIf(tb(i).ChavePrimaria, "  ' PK", "")
    Next
    Print #f, ""
    Print #f, ""
    
    ' Getters & Setters
    For i = 0 To UBound(tb)
        Print #f, "Public Property Get " & tb(i).Nome & "() As " & tb(i).Tipo
        Print #f, Space$(4) & tb(i).Nome & " = " & tb(i).Instancia
        Print #f, "End Property"
        Print #f, "Public Property Let " & tb(i).Nome & "(ByVal v As " & tb(i).Tipo & ")"
        Print #f, Space$(4) & tb(i).Instancia & " = v"
        Print #f, "End Property"
        Print #f, ""
    Next



    ' Vou gerar um ToString() também
    Print #f, "Public Function ToString() As String"
    Print #f, ""
    Print #f, "    Dim sDesc As String"
    Print #f, ""
    Print #f, "    sDesc = """""
    Print #f, "    If " & tb(0).Instancia & " <> " & IIf(tb(0).Tipo = "String", """""", "0") & " Then ";
    Print #f, "sDesc = sDesc & """ & tb(0).Nome & ": "" & " & tb(0).Instancia
    For i = 1 To UBound(tb)
        Print #f, "    If " & tb(i).Instancia & " <> " & IIf(tb(i).Tipo = "String", """""", "0") & " Then ";
        Print #f, "sDesc = sDesc & "", " & tb(i).Nome & ": "" & " & tb(i).Instancia
    Next
    Print #f, ""
    Print #f, "    ToString = sDesc"
    Print #f, ""
    Print #f, "End Function"

    Close #f

    Screen.MousePointer = vbDefault
    MsgBox "OK, gerado arquivo " & arq, vbInformation, "Atenção"

End Sub

Private Sub GeraColecao()

    Dim ub As Long
    Dim f As Integer
    Dim arq As String
    Dim sNomeClasse As String

    Screen.MousePointer = vbHourglass

    sNomeClasse = "C" & MeuStrConv(txtTabela.Text)

    f = FreeFile
    arq = Environ$("TEMP") & "\" & sNomeClasse & "s.cls"
    Open arq For Output As #f

    Print #f, "VERSION 1.0 CLASS"
    Print #f, "BEGIN"
    Print #f, "  MultiUse = -1  'True"
    #If VB6 Then
        Print #f, "  Persistable = 0  'NotPersistable"
        Print #f, "  DataBindingBehavior = 0  'vbNone"
        Print #f, "  DataSourceBehavior = 0   'vbNone"
        Print #f, "  MTSTransactionMode = 0   'NotAnMTSObject"
    #End If
    Print #f, "End"
    Print #f, "Attribute VB_Name = """ & sNomeClasse & "s"""
    Print #f, "Attribute VB_GlobalNameSpace = False"
    Print #f, "Attribute VB_Creatable = True"
    Print #f, "Attribute VB_PredeclaredId = False"
    Print #f, "Attribute VB_Exposed = True"
    Print #f, "Attribute VB_Ext_KEY = ""SavedWithClassBuilder"" ,""Yes""" ' mentira, mas enfim
    Print #f, "Attribute VB_Ext_KEY = ""Collection"" ,""" & sNomeClasse & """"
    Print #f, "Attribute VB_Ext_KEY = ""Member0"" ,""" & sNomeClasse & """"
    Print #f, "Attribute VB_Ext_KEY = ""Top_Level"" ,""Yes"""

    Print #f, "Option Explicit"
    Print #f, ""
    Print #f, "Private mCol As Collection"
    Print #f, ""
    Print #f, "Friend Sub Add(ByVal Item As " & sNomeClasse & ", Optional ByVal Key)"
    Print #f, "    If IsMissing(Key) Then"
    Print #f, "        mCol.Add Item"
    Print #f, "    Else"
    Print #f, "        mCol.Add Item, Key"
    Print #f, "    End If"
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Public Property Get Item(vntIndexKey As Variant) As " & sNomeClasse
    Print #f, "Attribute Item.VB_UserMemId = 0"
    Print #f, "    Set Item = mCol(vntIndexKey)"
    Print #f, "End Property"
    Print #f, ""
    Print #f, "Public Property Get Count() As Long"
    Print #f, "    Count = mCol.Count"
    Print #f, "End Property"
    Print #f, ""
    Print #f, "Friend Sub Remove(vntIndexKey As Variant)"
    Print #f, "    mCol.Remove vntIndexKey"
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Public Property Get NewEnum() As IUnknown"
    Print #f, "Attribute NewEnum.VB_UserMemId = -4"
    Print #f, "Attribute NewEnum.VB_MemberFlags = ""40"""
    Print #f, "    Set NewEnum = mCol.[_NewEnum]"
    Print #f, "End Property"
    Print #f, ""
    Print #f, "Private Sub Class_Initialize()"
    Print #f, "    Set mCol = New Collection"
    Print #f, "End Sub"
    Print #f, ""
    Print #f, "Private Sub Class_Terminate()"
    Print #f, "    Set mCol = Nothing"
    Print #f, "End Sub"

    Close #f

    Screen.MousePointer = vbDefault
    MsgBox "OK, gerado arquivo " & arq, vbInformation, "Atenção"

End Sub

Private Sub GeraDAO(tb() As TCampos)

    Dim i As Long
    Dim f As Integer
    Dim sNomeDoProjeto As String, sNomeDaClasseDAO As String, sNomeDoProjetoEClasseDAO As String
    Dim arq As String
    Dim ub As Long
    Dim lIndicePK As Long
    Dim sPK As String
    Dim bUsouToDate As Boolean
    Dim bUsouToBool As Boolean
    Dim bUsouToCur As Boolean
    Dim bTodosOsCamposFazemParteDaPK As Boolean
    Dim bFlag As Boolean

    On Error Resume Next
        ub = UBound(tb)
        If Err.Number <> 0 Then
            MsgBox "Erro: não há dados.", vbCritical, "Erro"
            Exit Sub
        End If
    On Error GoTo 0

    If ChavePrimariaComposta(tb) = False Then
        lIndicePK = ObtemIndiceChavePrimaria(tb)
        sPK = tb(lIndicePK).Nome 'um atalho...
    End If

    bTodosOsCamposFazemParteDaPK = True
    For i = 0 To UBound(tb)
        If tb(i).ChavePrimaria = False Then
            bTodosOsCamposFazemParteDaPK = False
            Exit For
        End If
    Next

    Screen.MousePointer = vbHourglass

    sNomeDoProjeto = InputBox("Informe o nome do projeto onde a classe DAO será colocada", "Nome do projeto", "MeuProjetoDAO")
    sNomeDaClasseDAO = "C" & MeuStrConv(txtTabela.Text) & "DAO"
    If sNomeDoProjeto <> "" Then
        sNomeDoProjetoEClasseDAO = sNomeDoProjeto & "." & sNomeDaClasseDAO
    Else
        sNomeDoProjetoEClasseDAO = sNomeDaClasseDAO
    End If

    f = FreeFile
    arq = Environ$("TEMP") & "\C" & MeuStrConv(txtTabela.Text) & "DAO.cls"
    Open arq For Output As #f

    Cabecalho f, sNomeDaClasseDAO

    Print #f, "Private Const TABELA = """ & UCase$(Trim$(txtTabela.Text)) & """"
    Print #f, "Private mConn As ADODB.Connection"
    Print #f, "Private mbConexaoDoCliente as Boolean"
    Print #f, "Private msStrCon As String"
    Print #f, "Private msLogFile As String"
    Print #f, "Private msOrdenacao As String"
    Print #f, ""
    Print #f, "Event DebugLog(ByVal sTexto As String)"
    Print #f, "Event LogAcesso(ByVal sAcao As String, ByVal sTipo As EAcoesLog, ByVal sDescricao As String, ByVal sCodDocumentoRelacionado As String)"
    Print #f, ""

    Print #f, "Public Property Get Conexao() As ADODB.Connection"
    Print #f, "    Set Conexao = mConn"
    Print #f, "End Property"
    Print #f, "Public Property Set Conexao(ByVal con As ADODB.Connection)"
    Print #f, "    Set mConn = con"
    Print #f, "    mbConexaoDoCliente = True"
    Print #f, "End Property"
    Print #f, ""

    Print #f, "Public Property Get StringConexao() As String"
    Print #f, "    StringConexao = msStrCon"
    Print #f, "End Property"
    Print #f, "Public Property Let StringConexao(ByVal sStrCon As String)"
    Print #f, "    msStrCon = sStrCon"
    Print #f, "End Property"
    Print #f, ""

    Print #f, "Public Property Get LogFile() As String"
    Print #f, "    LogFile = msLogFile"
    Print #f, "End Property"
    Print #f, "Public Property Let LogFile(ByVal sLogFile As String)"
    Print #f, "    msLogFile = sLogFile"
    Print #f, "End Property"
    Print #f, ""

    Print #f, "' Possivelmente há formas mais elegantes de se fazer isso, mas KISS"
    Print #f, "Public Property Get OrderBy() As String"
    Print #f, "    OrderBy = msOrdenacao"
    Print #f, "End Property"
    Print #f, "Public Property Let OrderBy(ByVal sOrdenacao As String)"
    Print #f, "    msOrdenacao = sOrdenacao"
    Print #f, "End Property"

    Print #f, "Private Function ConPF() As ADODB.Connection"
    Print #f, "    Const NOMEROTINA As String = """ & sNomeDoProjetoEClasseDAO & "::ConPF"""
    Print #f, "    If mConn Is Nothing Then"
    Print #f, "        If msStrCon = """" Then"
    Print #f, "            Err.Raise vbObjectError + 1, NOMEROTINA, ""Uma das propriedades """"Conexao"""" ou """"StringConexao"""" precisam ser informadas."""
    Print #f, "        End If"
    Print #f, "        Set ConPF = AbreCn(msStrCon)"
    Print #f, "    Else"
    Print #f, "        If mConn.State <> adStateOpen Then"
    Print #f, "            If msStrCon = """" Then"
    Print #f, "                Err.Raise vbObjectError + 1, NOMEROTINA, ""Uma das propriedades """"Conexao"""" ou """"StringConexao"""" precisam ser informadas."""
    Print #f, "            End If"
    Print #f, "            Set mConn = AbreCn(msStrCon)"
    Print #f, "        End If"
    Print #f, "        Set ConPF = mConn"
    Print #f, "    End If"
    Print #f, "End Function"
    Print #f, ""

    Print #f, "Private Sub DebugLogPS(ByVal sTexto As String)"
    Print #f, "    Static oLog As FnLog"
    Print #f, "    If oLog Is Nothing Then Set oLog = New FnLog"
    Print #f, "    oLog.LogFile = msLogFile"
    Print #f, "    oLog.Ativar = (msLogFile <> """")"
    Print #f, "    oLog.DebugLog sTexto"
    Print #f, "    ' Como alternativa..."
    Print #f, "    RaiseEvent DebugLog(sTexto)"
    Print #f, "End Sub"
    Print #f, ""

    Print #f, "Public Function Adiciona(obj As C" & MeuStrConv(txtTabela.Text) & ") As Boolean"
    If ChavePrimariaComposta(tb) = False Then
        Print #f, ""
        Print #f, "    Dim lID As Long"
        Print #f, ""
        Print #f, "    lID = GeraSequenceGF(ConPF(), ""SEQ_"" & TABELA)"
        Print #f, "    If lID < 1 Then Exit Function"
        Print #f, "    obj." & sPK & " = lID"
        Print #f, ""
    End If
    Print #f, "    Adiciona = AdicionaAtualizaPF(True, obj)"
    Print #f, "End Function"
    Print #f, ""

    Print #f, "Public Function Atualiza(obj As C" & MeuStrConv(txtTabela.Text) & ") As Boolean"
    If bTodosOsCamposFazemParteDaPK Then
        Print #f, "    Const NOMEROTINA As String = """ & sNomeDoProjetoEClasseDAO & "::Atualiza"""
        Print #f, "    Err.Raise vbObjectError + 1, NOMEROTINA, ""Não há como atualizar, todas as colunas fazem parte da chave primária."""
    Else
        Print #f, "    Atualiza = AdicionaAtualizaPF(False, obj)"
    End If
    Print #f, "End Function"
    Print #f, ""

    Print #f, "Public Function Exclui(obj As C" & MeuStrConv(txtTabela.Text) & ") As Boolean"
    Print #f, ""
    Print #f, "    Const NOMEROTINA As String = """ & sNomeDoProjetoEClasseDAO & "::Exclui"""
    Print #f, "    Dim MyErr As CErro"
    Print #f, "    Dim cmd As ADODB.Command"
    Print #f, ""
    Print #f, "    On Error GoTo Erro"
    Print #f, ""
    Print #f, "    Set cmd = New ADODB.Command"
    Print #f, "    cmd.ActiveConnection = ConPF()"
    Print #f, "    cmd.CommandType = adCmdText"
    Print #f, "    cmd.CommandText = ""DELETE FROM "" & TABELA & "" WHERE ";
    Dim sAux As String, j As Long
    For i = 0 To UBound(tb)
        If tb(i).ChavePrimaria Then
            sAux = sAux & tb(i).Nome & " = ? AND "
        End If
    Next
    If Right$(sAux, 5) = " AND " Then sAux = Mid$(sAux, 1, Len(sAux) - 5) & """"
    Print #f, sAux
    For i = 0 To UBound(tb)
        If tb(i).ChavePrimaria Then
            If tb(i).Tamanho > 0 Then
                'string
                Print #f, "    cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Size:=" & tb(i).Tamanho & ", Value:=Left$(Trim$(obj." & tb(i).Nome & "), " & tb(i).Tamanho & "))"
            Else
                Print #f, "    cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Value:=obj." & tb(i).Nome & ")"
            End If
        End If
    Next
    Print #f, "    cmd.Execute Options:=adExecuteNoRecords"
    Print #f, ""
    Print #f, "    RaiseEvent LogAcesso(""Excluir"", Acao_Informacao, _"
    If sPK <> "" Then
        ' PK simples
        Print #f, "        ""Excluído registro código "" & obj." & sPK & " , _"
        Print #f, "        obj." & sPK & ")"
    Else
        Print #f, "        ""Excluído registro: ";
        ' PK composta
        sAux = ""
        For i = 0 To UBound(tb)
            If tb(i).ChavePrimaria Then
                sAux = sAux & tb(i).Nome & " = "" & obj." & tb(i).Nome & " & "", "
            End If
        Next
        If Right$(sAux, 5) = "& "", " Then sAux = Mid$(sAux, 1, Len(sAux) - 5)
        Print #f, sAux & ", """")"
    End If
    Print #f, ""
    Print #f, "    Exclui = True"
    Print #f, ""
    Print #f, "    GoTo Sai"
    Print #f, ""
    Print #f, "Erro:"
    Print #f, "    Set MyErr = New CErro"
    Print #f, "    Set MyErr.Erro = Err"
    Print #f, "Resume Sai"
    Print #f, "Sai:"
    Print #f, "    On Error Resume Next"
    Print #f, "    If Not cmd Is Nothing Then"
    Print #f, "        Set cmd.ActiveConnection = Nothing"
    Print #f, "        Set cmd = Nothing"
    Print #f, "    End If"
    Print #f, "    If Not MyErr Is Nothing Then"
    Print #f, "        If MyErr.Source = App.Title Then MyErr.Source = NOMEROTINA"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": ----> saindo com o erro "" & MyErr.Number & "" em "" & MyErr.Source & "": "" & MyErr.Description)"
    Print #f, "        On Error GoTo 0"
    Print #f, "        Err.Raise MyErr.Number, MyErr.Source, ""Erro ao excluir dados: "" & MyErr.Description"
    Print #f, "    End If"
    Print #f, "Exit Function"
    Print #f, ""
    Print #f, "End Function"
    Print #f, ""

    Print #f, "Public Function Localiza(obj As C" & MeuStrConv(txtTabela.Text) & ") As C" & MeuStrConv(txtTabela.Text) & "s"
    Print #f, ""
    Print #f, "    Const NOMEROTINA As String = """ & sNomeDoProjetoEClasseDAO & "::Localiza"""
    Print #f, "    Dim MyErr As CErro"
    Print #f, "    Dim cmd as ADODB.Command"
    Print #f, "    Dim sPesq As String"
    Print #f, ""
    Print #f, "    On Error Goto Erro"
    Print #f, ""
    Print #f, "    Set cmd = New ADODB.Command"
    Print #f, "    cmd.ActiveConnection = ConPF()"
    Print #f, "    cmd.CommandType = adCmdText"
    Print #f, ""
    Print #f, "    sPesq = ""SELECT * FROM "" & TABELA & "" WHERE (1=1)"""
    Print #f, ""
    For i = 0 To UBound(tb)
        If tb(i).Tipo = "String" Then
            Print #f, "    If obj." & tb(i).Nome & " <> """" Then"
        Else
            Print #f, "    If obj." & tb(i).Nome & " > 0 Then"
        End If
        Print #f, "        sPesq = sPesq & "" AND " & tb(i).Nome & " = ? """
        If tb(i).Tamanho > 0 Then
            ' string
            Print #f, "        cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Size:=" & tb(i).Tamanho & ", Value:=TLT(obj." & tb(i).Nome & ", " & tb(i).Tamanho & "))"
        Else
            Print #f, "        cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Value:=obj." & tb(i).Nome & ")"
        End If
        Print #f, "    End If"
        Print #f, ""
    Next
    Print #f, "    cmd.CommandText = sPesq"
    Print #f, ""
    Print #f, "    Set Localiza = LocalizaPF(cmd)"
    Print #f, ""
    Print #f, "    GoTo Sai"
    Print #f, ""
    Print #f, "Erro:"
    Print #f, "    Set MyErr = New CErro"
    Print #f, "    Set MyErr.Erro = Err"
    Print #f, "Resume Sai"
    Print #f, "Sai:"
    Print #f, "    On Error Resume Next"
    Print #f, "    If Not cmd Is Nothing Then"
    Print #f, "        Set cmd.ActiveConnection = Nothing"
    Print #f, "        Set cmd = Nothing"
    Print #f, "    End If"
    Print #f, "    If Not MyErr Is Nothing Then"
    Print #f, "        If MyErr.Source = App.Title Then MyErr.Source = NOMEROTINA"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": ----> saindo com o erro "" & MyErr.Number & "" em "" & MyErr.Source & "": "" & MyErr.Description)"
    Print #f, "        On Error GoTo 0"
    Print #f, "        Err.Raise MyErr.Number, MyErr.Source, MyErr.Description"
    Print #f, "    End If"
    Print #f, "Exit Function"
    Print #f, ""
    Print #f, "End Function"
    Print #f, ""
    
    Print #f, "Private Function LocalizaPF(ByVal cmd as ADODB.Command) As C" & MeuStrConv(txtTabela.Text) & "s"
    Print #f, ""
    Print #f, "    Const NOMEROTINA As String = """ & sNomeDoProjetoEClasseDAO & "::LocalizaPF"""
    Print #f, "    Dim MyErr as CErro"
    Print #f, "    Dim rs As ADODB.Recordset"
    Print #f, "    Dim obj As C" & MeuStrConv(txtTabela.Text)
    Print #f, "    Dim sParam As String"
    Print #f, ""
    Print #f, "    On Error GoTo Erro"
    Print #f, ""
    Print #f, "    sParam = CommandParametersToString(cmd)"
    Print #f, ""
    Print #f, "    Set rs = cmd.Execute()"
    Print #f, "    If rs Is Nothing Then"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": frase executada, mas objeto está vazio!"")"
    Print #f, "        GoTo Sai"
    Print #f, "    ElseIf rs.EOF Then"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": frase executada, mas não há dados"")"
    Print #f, "    End If"
    Print #f, ""
    Print #f, "    Set LocalizaPF = New C" & MeuStrConv(txtTabela.Text) & "s"
    Print #f, ""
    Print #f, "    Do While Not rs.EOF"
    Print #f, "        Set obj = New C" & MeuStrConv(txtTabela.Text)
    Print #f, ""
    For i = 0 To UBound(tb)
        Select Case tb(i).Tipo
            Case "String"
                Print #f, "        obj." & tb(i).Nome & " = rs(""" & tb(i).Nome & """) & """""
            Case "Long"
                Print #f, "        obj." & tb(i).Nome & " = Val(rs(""" & tb(i).Nome & """) & """")"
            Case "Double", "Currency"
                Print #f, "        obj." & tb(i).Nome & " = ToCurPF(rs(""" & tb(i).Nome & """))"
                bUsouToCur = True
            Case "Date"
                Print #f, "        obj." & tb(i).Nome & " = ToDatePF(rs(""" & tb(i).Nome & """))"
                bUsouToDate = True
            Case "Boolean"
                Print #f, "        obj." & tb(i).Nome & " = ToBoolPF(rs(""" & tb(i).Nome & """))"
                bUsouToBool = True
            Case Else
                Print #f, "        obj." & tb(i).Nome & " = rs(""" & tb(i).Nome & """)"
        End Select
    Next
    Print #f, ""
    Print #f, "        rs.MoveNext"
    Print #f, "        LocalizaPF.Add obj"
    Print #f, ""
    Print #f, "    Loop"
    Print #f, ""
    Print #f, "    Call DebugLogPS(NOMEROTINA & "": total de "" & LocalizaPF.Count & "" objetos na coleção."")"
    Print #f, ""
    Print #f, "    GoTo Sai"
    Print #f, ""
    Print #f, "Erro:"
    Print #f, "    Set MyErr = New CErro"
    Print #f, "    Set MyErr.Erro = Err"
    Print #f, "Resume Sai"
    Print #f, "Sai:"
    Print #f, "    On Error Resume Next"
    Print #f, "    If Not rs Is Nothing Then"
    Print #f, "        If rs.State <> adStateClosed Then rs.Close"
    Print #f, "        Set rs = Nothing"
    Print #f, "    End If"
    Print #f, "    If Not MyErr Is Nothing Then"
    Print #f, "        If MyErr.Source = App.Title Then MyErr.Source = NOMEROTINA"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": ----> saindo com o erro "" & MyErr.Number & "" em "" & MyErr.Source & "": "" & MyErr.Description)"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": string de conexão: "" & modADO.ObtemStringConexaoSegura(cmd.ActiveConnection.ConnectionString))"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": frase SQL a executar é ["" & cmd.CommandText & ""]"")"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": parâmetros: "" & sParam)"
    Print #f, "        On Error GoTo 0"
    Print #f, "        Err.Raise MyErr.Number, MyErr.Source, MyErr.Description"
    Print #f, "    End If"
    Print #f, "Exit Function"
    Print #f, ""
    Print #f, "End Function"
    Print #f, ""

    If bUsouToDate Then
        Print #f, "Private Function ToDatePF(ByVal vData As Variant) As Date"
        Print #f, ""
        Print #f, "    On Error Resume Next"
        Print #f, "    If IsDate(vData) Then ToDatePF = CDate(vData) Else ToDatePF = CDate(0)"
        Print #f, ""
        Print #f, "End Function"
        Print #f, ""
    End If

    If bUsouToBool Then
        Print #f, "Private Function ToBoolPF(ByVal pvValor As Variant) As Boolean"
        Print #f, ""
        Print #f, "    On Error Resume Next"
        Print #f, ""
        Print #f, "    If IsNull(pvValor) Then"
        Print #f, "        ToBoolPF = False"
        Print #f, "    Else"
        Print #f, "        ToBoolPF = CBool(pvValor)"
        Print #f, "    End If"
        Print #f, ""
        Print #f, "End Function"
        Print #f, ""
    End If

    If bUsouToCur Then
        Print #f, "Private Function ToCurPF(ByVal pvValor As Variant) As Currency"
        Print #f, ""
        Print #f, "    On Error Resume Next"
        Print #f, ""
        Print #f, "    If IsNull(pvValor) Then"
        Print #f, "        ToCurPF = 0"
        Print #f, "    Else"
        Print #f, "        ' transforma valores entre parênteses para negativo"
        Print #f, "        If Left$(pvValor, 1) = ""("" And Right$(pvValor, 1) = "")"" Then"
        Print #f, "            pvValor = ""-"" & Mid$(pvValor, 2)"
        Print #f, "        End If"
        Print #f, "        ' retira sinais e moedas"
        Print #f, "        pvValor = ReplacePF(ReplacePF(ReplacePF(ReplacePF(ReplacePF(ReplacePF(pvValor, _"
        Print #f, "                  ""%"", """"), _"
        Print #f, "                  "")"", """"), _"
        Print #f, "                  ""("", """"), _"
        Print #f, "                  ""US$"", """"), _"
        Print #f, "                  ""R$"", """"), _"
        Print #f, "                  ""$"", """")"
        Print #f, "        ToCurPF = RoundPF(CCur(pvValor), 2)"
        Print #f, "    End If"
        Print #f, ""
        Print #f, "End Function"
        Print #f, ""
        Print #f, "Private Function ReplacePF(ByVal sIn As String, ByVal sFind As String, ByVal sReplace As String, _"
        Print #f, "                        Optional nStart As Long = 1, Optional nCount As Long = -1, _"
        Print #f, "                        Optional bCompare As VbCompareMethod = vbBinaryCompare) As String"
        Print #f, ""
        Print #f, "    Dim nC As Long, nPos As Long"
        Print #f, "    Dim nFindLen As Long, nReplaceLen As Long"
        Print #f, ""
        Print #f, "    nFindLen = Len(sFind)"
        Print #f, "    nReplaceLen = Len(sReplace)"
        Print #f, ""
        Print #f, "    If (sFind <> """") And (sFind <> sReplace) Then"
        Print #f, "        nPos = InStr(nStart, sIn, sFind, bCompare)"
        Print #f, "        Do While nPos"
        Print #f, "            nC = nC + 1"
        Print #f, "            sIn = Left(sIn, nPos - 1) & sReplace & Mid(sIn, nPos + nFindLen)"
        Print #f, "            If nCount <> -1 And nC >= nCount Then Exit Do"
        Print #f, "            nPos = InStr(nPos + nReplaceLen, sIn, sFind, bCompare)"
        Print #f, "        Loop"
        Print #f, "    End If"
        Print #f, ""
        Print #f, "    ReplacePF = sIn"
        Print #f, ""
        Print #f, "End Function"
        Print #f, ""
        Print #f, "' Tirei de http://www.xbeat.net/vbspeed/c_Round.htm"
        Print #f, "Public Function RoundPF(ByVal v As Double, Optional ByVal lngDecimals As Long = 0) As Double"
        Print #f, ""
        Print #f, "  Dim xint As Double, yint As Double, xrest As Double"
        Print #f, "  Static PreviousValue    As Double"
        Print #f, "  Static PreviousDecimals As Long"
        Print #f, "  Static PreviousOutput   As Double"
        Print #f, "  Static M                As Double"
        Print #f, ""
        Print #f, "  If PreviousValue = v And PreviousDecimals = lngDecimals Then RoundPF = PreviousOutput: Exit Function"
        Print #f, "  If v = 0 Then Exit Function"
        Print #f, ""
        Print #f, "  If PreviousDecimals = lngDecimals Then"
        Print #f, "      If M = 0 Then M = 1  ' Initialization - M is never 0 (it is always 10 ^ n)"
        Print #f, "      Else"
        Print #f, "        PreviousDecimals = lngDecimals"
        Print #f, "        M = 10 ^ lngDecimals"
        Print #f, "      End If"
        Print #f, ""
        Print #f, "  If M = 1 Then xint = v Else xint = v * CDec(M)"
        Print #f, "  RoundPF = Fix(xint)"
        Print #f, ""
        Print #f, "  If Abs(Fix(10 * (xint - RoundPF))) > 4 Then"
        Print #f, "    If xint < 0 Then"
        Print #f, "      RoundPF = RoundPF - 1"
        Print #f, "    Else"
        Print #f, "      RoundPF = RoundPF + 1"
        Print #f, "    End If"
        Print #f, "  End If"
        Print #f, ""
        Print #f, "  If M = 1 Then Else RoundPF = RoundPF / M"
        Print #f, ""
        Print #f, "  PreviousOutput = RoundPF"
        Print #f, "  PreviousValue = v"
        Print #f, ""
        Print #f, "End Function"
        Print #f, ""
    End If

    Print #f, "Private Function AdicionaAtualizaPF(ByVal bAdicionar As Boolean, obj As C" & MeuStrConv(txtTabela.Text) & ") As Boolean"
    Print #f, ""
    Print #f, "    Const NOMEROTINA As String = """ & sNomeDoProjetoEClasseDAO & "::AdicionaAtualizaPF"""
    Print #f, "    Dim cmd As ADODB.Command"
    Print #f, "    Dim sSQL As String"
    Print #f, ""
    Print #f, "    On Error GoTo Erro"
    Print #f, ""
    Print #f, "    Set cmd = New ADODB.Command"
    Print #f, "    cmd.ActiveConnection = ConPF()"
    Print #f, "    cmd.CommandType = adCmdText"
    Print #f, ""

    Print #f, "    If bAdicionar Then"
    Print #f, "        sSQL = """""
    Print #f, "        sSQL = sSQL & ""INSERT INTO "" & TABELA"
    Print #f, "        sSQL = sSQL & "" (";
    For i = 0 To UBound(tb) - 1
        Print #f, tb(i).Nome & ", ";
    Next
    Print #f, tb(i).Nome & ")"""
    Print #f, "        sSQL = sSQL & "" VALUES (";
    For i = 0 To UBound(tb) - 1
        Print #f, "?, ";
    Next
    Print #f, "?)"""
    For i = 0 To UBound(tb)
        If tb(i).ChavePrimaria Then
            If tb(i).Tamanho > 0 Then
                'string
                Print #f, "        cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Size:=" & tb(i).Tamanho & ", Value:=TLT(obj." & tb(i).Nome & ", " & tb(i).Tamanho & "))"
            Else
                Print #f, "        cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Value:=obj." & tb(i).Nome & ")"
            End If
        End If
    Next

    Print #f, "    Else"
    If bTodosOsCamposFazemParteDaPK Then
        Print #f, "        Err.Raise vbObjectError + 1, NOMEROTINA, ""Atualização não é possível: todos os campos fazem parte da chave primária."""
    Else
        Print #f, "        sSQL = """""
        Print #f, "        sSQL = sSQL & ""UPDATE "" & TABELA"
        For i = 0 To UBound(tb)
            If Not tb(i).ChavePrimaria Then
                If Not bFlag Then
                    Print #f, "        sSQL = sSQL & ""   SET " & tb(i).Nome & " = ?"""
                    bFlag = True
                Else
                    Print #f, "        sSQL = sSQL & ""     , " & tb(i).Nome & " = ?"""
                End If
            End If
        Next

        Print #f, "        sSQL = sSQL & "" WHERE ";
        sAux = ""
        For i = 0 To UBound(tb)
            If tb(i).ChavePrimaria Then
                sAux = sAux & tb(i).Nome & " = ? AND "
            End If
        Next
        If Right$(sAux, 5) = " AND " Then sAux = Mid$(sAux, 1, Len(sAux) - 5)
        Print #f, sAux & """"
    End If
    Print #f, "    End If"

    Print #f, ""
    Print #f, "    cmd.CommandText = sSQL"
    Print #f, ""
    For i = 0 To UBound(tb)
        If Not tb(i).ChavePrimaria Then
            If tb(i).Tamanho > 0 Then
                ' char/varchar
                Print #f, "    cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Size:=" & tb(i).Tamanho & ", Value:=TLT(obj." & tb(i).Nome & ", " & tb(i).Tamanho & "))"
            Else
                Print #f, "    cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Value:=obj." & tb(i).Nome & ")"
            End If
        End If
    Next
    If Not bTodosOsCamposFazemParteDaPK Then
        Print #f, ""
        Print #f, "    If Not bAdicionar Then"
        For i = 0 To UBound(tb)
            If tb(i).ChavePrimaria Then
                If tb(lIndicePK).Tamanho > 0 Then
                    'string
                    Print #f, "         cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Size:=" & tb(i).Tamanho & ", Value:=TLT(obj." & tb(i).Nome & ", " & tb(i).Tamanho & "))"
                Else
                    Print #f, "         cmd.Parameters.Append cmd.CreateParameter(""" & tb(i).Nome & """, " & tb(i).TipoADO & ", Value:=obj." & tb(i).Nome & ")"
                End If
            End If
        Next
        Print #f, "    End If"
        Print #f, ""
    End If
    Print #f, "    cmd.Execute Options:=adExecuteNoRecords"
    Print #f, ""

    For i = 0 To UBound(tb)
        If tb(i).ChavePrimaria Then
            Print #f, "    obj." & tb(i).Nome & " = cmd.Parameters(""" & tb(i).Nome & """).Value"
        End If
    Next
    Print #f, ""
    Print #f, "    Set cmd.ActiveConnection = Nothing"
    Print #f, "    Set cmd = Nothing"
    Print #f, ""
    Print #f, "    RaiseEvent LogAcesso(""Salvar"", Acao_Informacao, _"
    Print #f, "        IIf(bAdicionar, ""Gravou"", ""Atualizou"") & "" registro: "" & obj.ToString(), _"
    Print #f, "        """")"
    Print #f, ""
    Print #f, "    AdicionaAtualizaPF = True"
    Print #f, ""
    Print #f, "    GoTo Sai"
    Print #f, ""
    Print #f, "Erro:"
    Print #f, "    Set MyErr = New CErro"
    Print #f, "    Set MyErr.Erro = Err"
    Print #f, "Resume Sai"
    Print #f, "Sai:"
    Print #f, "    On Error Resume Next"
    Print #f, "    If Not cmd Is Nothing Then"
    Print #f, "        Set cmd.ActiveConnection = Nothing"
    Print #f, "        Set cmd = Nothing"
    Print #f, "    End If"
    Print #f, "    If Not MyErr Is Nothing Then"
    Print #f, "        If MyErr.Source = App.Title Then MyErr.Source = NOMEROTINA"
    Print #f, "        Call DebugLogPS(NOMEROTINA & "": ----> saindo com o erro "" & MyErr.Number & "" em "" & MyErr.Source & "": "" & MyErr.Description)"
    Print #f, "        On Error GoTo 0"
    Print #f, "        Err.Raise MyErr.Number, MyErr.Source, ""Erro ao salvar dados: "" & MyErr.Description"
    Print #f, "    End If"
    Print #f, "Exit Function"
    Print #f, ""
    Print #f, "End Function"
    Print #f, ""

'
' Você pode optar por usar ou não o código a seguir...
'
'
'    Print #f, "Public Function ObtemPrimeiro(ByVal sCampo As String) As C" & MeuStrConv(txtTabela.Text)
'    Print #f, ""
'    Print #f, "    Dim sPesq As String"
'    Print #f, "    Dim cmd as ADODB.Command"
'    Print #f, "    Dim col As C" & MeuStrConv(txtTabela.Text) & "s"
'    Print #f, ""
'    Print #f, "    On Error GoTo Erro"
'    Print #f, ""
'    Print #f, "    sPesq = ""SELECT * FROM "" & TABELA & _"
'    Print #f, "            "" WHERE "" & sCampo & "" = (SELECT MIN("" & sCampo & "")"" & _"
'    Print #f, "                                      "" FROM "" & TABELA"
'    Print #f, "    sPesq = sPesq & "")"""
'    Print #f, ""
'    Print #f, "    Set cmd = New ADODB.Command"
'    Print #f, "    cmd.ActiveConnection = ConPF()"
'    Print #f, "    cmd.CommandType = adCmdText"
'    Print #f, "    cmd.CommandText = sPesq"
'    Print #f, ""
'    Print #f, "    Set col = LocalizaPF(cmd)"
'    Print #f, "    If col.Count > 0 Then"
'    Print #f, "        Set ObtemPrimeiro = col(1)"
'    Print #f, "    Else"
'    Print #f, "        Set ObtemPrimeiro = Nothing"
'    Print #f, "    End If"
'    Print #f, ""
'    Print #f, "    Exit Function"
'    Print #f, ""
'    Print #f, "Erro:"
'    Print #f, "    ErroX ""Erro ao localizar registro: "" & Error$, MENAME & ""::ObtemPrimeiro"""
'    Print #f, "Resume Fecha"
'    Print #f, "Fecha:"
'    Print #f, "    On Error Resume Next"
'    Print #f, "        Set col = Nothing"
'    Print #f, "        Set ObtemPrimeiro = Nothing"
'    Print #f, "    On Error GoTo 0"
'    Print #f, "Exit Function"
'    Print #f, ""
'    Print #f, "End Function"
'    Print #f, ""
'
'    Print #f, "Public Function ObtemUltimo(ByVal sCampo As String) As C" & MeuStrConv(txtTabela.Text)
'    Print #f, ""
'    Print #f, "    Dim sPesq As String"
'    Print #f, "    Dim cmd as ADODB.Command"
'    Print #f, "    Dim col As C" & MeuStrConv(txtTabela.Text) & "s"
'    Print #f, ""
'    Print #f, "    On Error GoTo Erro"
'    Print #f, ""
'    Print #f, "    sPesq = ""SELECT * FROM "" & TABELA & _"
'    Print #f, "            "" WHERE "" & sCampo & "" = (SELECT MAX("" & sCampo & "")"" & _"
'    Print #f, "                                      "" FROM "" & TABELA"
'    Print #f, "    sPesq = sPesq & "")"""
'    Print #f, ""
'    Print #f, "    Set cmd = New ADODB.Command"
'    Print #f, "    cmd.ActiveConnection = ConPF()"
'    Print #f, "    cmd.CommandType = adCmdText"
'    Print #f, "    cmd.CommandText = sPesq"
'    Print #f, ""
'    Print #f, "    Set col = LocalizaPF(cmd)"
'    Print #f, "    If col.Count > 0 Then"
'    Print #f, "        Set ObtemUltimo = col(1)"
'    Print #f, "    Else"
'    Print #f, "        Set ObtemUltimo = Nothing"
'    Print #f, "    End If"
'    Print #f, ""
'    Print #f, "    Exit Function"
'    Print #f, ""
'    Print #f, "Erro:"
'    Print #f, "    ErroX ""Erro ao localizar registro: "" & Error$, MENAME & ""::ObtemUltimo"""
'    Print #f, "Resume Fecha"
'    Print #f, "Fecha:"
'    Print #f, "    On Error Resume Next"
'    Print #f, "        Set col = Nothing"
'    Print #f, "        Set ObtemUltimo = Nothing"
'    Print #f, "    On Error GoTo 0"
'    Print #f, "Exit Function"
'    Print #f, ""
'    Print #f, "End Function"
'    Print #f, ""
'
'    Print #f, "Public Function ObtemAnterior(ByVal sCampo As String, ByVal sValor As String) As C" & MeuStrConv(txtTabela.Text)
'    Print #f, ""
'    Print #f, "    Dim sPesq As String"
'    Print #f, "    Dim cmd as ADODB.Command"
'    Print #f, "    Dim col As C" & MeuStrConv(txtTabela.Text) & "s"
'    Print #f, ""
'    Print #f, "    On Error GoTo Erro"
'    Print #f, ""
'    Print #f, "    sPesq = ""SELECT * FROM "" & TABELA & _"
'    Print #f, "            "" WHERE "" & sCampo & "" = (SELECT MAX("" & sCampo & "")"" & _"
'    Print #f, "                                     ""  FROM "" & TABELA & _"
'    Print #f, "                                     "" WHERE "" & sCampo & "" < '"" & sValor & ""'"""
'    Print #f, "    sPesq = sPesq & "")"""
'    Print #f, ""
'    Print #f, "    Set cmd = New ADODB.Command"
'    Print #f, "    cmd.ActiveConnection = ConPF()"
'    Print #f, "    cmd.CommandType = adCmdText"
'    Print #f, "    cmd.CommandText = sPesq"
'    Print #f, ""
'    Print #f, "    Set col = LocalizaPF(cmd)"
'    Print #f, "    If col.Count > 0 Then"
'    Print #f, "        Set ObtemAnterior = col(1)"
'    Print #f, "    Else"
'    Print #f, "        Set ObtemAnterior = ObtemPrimeiro(sCampo)"
'    Print #f, "    End If"
'    Print #f, ""
'    Print #f, "    Exit Function"
'    Print #f, ""
'    Print #f, "Erro:"
'    Print #f, "    ErroX ""Erro ao localizar registro: "" & Error$, MENAME & ""::ObtemAnterior"""
'    Print #f, "Resume Fecha"
'    Print #f, "Fecha:"
'    Print #f, "    On Error Resume Next"
'    Print #f, "        Set col = Nothing"
'    Print #f, "        Set ObtemAnterior = Nothing"
'    Print #f, "    On Error GoTo 0"
'    Print #f, "Exit Function"
'    Print #f, ""
'    Print #f, "End Function"
'    Print #f, ""
'
'    Print #f, "Public Function ObtemProximo(ByVal sCampo As String, ByVal sValor As String) As C" & MeuStrConv(txtTabela.Text)
'    Print #f, ""
'    Print #f, "    Dim sPesq As String"
'    Print #f, "    Dim cmd as ADODB.Command"
'    Print #f, "    Dim col As C" & MeuStrConv(txtTabela.Text) & "s"
'    Print #f, ""
'    Print #f, "    On Error GoTo Erro"
'    Print #f, ""
'    Print #f, "    sPesq = ""SELECT * FROM "" & TABELA & _"
'    Print #f, "            "" WHERE "" & sCampo & "" = (SELECT MIN("" & sCampo & "")"" & _"
'    Print #f, "                                     ""  FROM "" & TABELA & _"
'    Print #f, "                                     "" WHERE "" & sCampo & "" > '"" & sValor & ""'"""
'    Print #f, "    sPesq = sPesq & "")"""
'    Print #f, ""
'    Print #f, "    Set cmd = New ADODB.Command"
'    Print #f, "    cmd.ActiveConnection = ConPF()"
'    Print #f, "    cmd.CommandType = adCmdText"
'    Print #f, "    cmd.CommandText = sPesq"
'    Print #f, ""
'    Print #f, "    Set col = LocalizaPF(cmd)"
'    Print #f, "    If col.Count > 0 Then"
'    Print #f, "        Set ObtemProximo = col(1)"
'    Print #f, "    Else"
'    Print #f, "        Set ObtemProximo = ObtemUltimo(sCampo)"
'    Print #f, "    End If"
'    Print #f, ""
'    Print #f, "    Exit Function"
'    Print #f, ""
'    Print #f, "Erro:"
'    Print #f, "    ErroX ""Erro ao localizar registro: "" & Error$, MENAME & ""::ObtemProximo"""
'    Print #f, "Resume Fecha"
'    Print #f, "Fecha:"
'    Print #f, "    On Error Resume Next"
'    Print #f, "        Set col = Nothing"
'    Print #f, "        Set ObtemProximo = Nothing"
'    Print #f, "    On Error GoTo 0"
'    Print #f, "Exit Function"
'    Print #f, ""
'    Print #f, "End Function"
'    Print #f, ""
'
'    Print #f, "' Active Data para o relatório"
'    Print #f, "Public Function GeraDataSourceCrystal() As ADODB.Recordset"
'    Print #f, ""
'    Print #f, "    On Error GoTo Erro"
'    Print #f, ""
'    Print #f, "    Dim sb As SQLBuilder"
'    Print #f, ""
'    Print #f, "    Set sb = New SQLBuilder"
'    Print #f, "    sb.Banco = glSGBD"
'    Print #f, "    sb.SQLSelect = """ & tb(0).Nome & """"
'    Print #f, "    sb.From = """ & UCase$(txtTabela.Text) & """"
'    Print #f, "    sb.OrderBy = """ & tb(0).Nome & """"
'    Print #f, ""
'    Print #f, "    Set GeraDataSourceCrystal = AbreFirehoseRs(sb.FraseSQL, """", cn:=ConPF())"
'    Print #f, ""
'    Print #f, "    On Error GoTo 0"
'    Print #f, "    Exit Function"
'    Print #f, ""
'    Print #f, "Erro:"
'    Print #f, "    ErroX ""Erro ao gerar dados para o relatório: "" & Error$, MENAME & ""::GeraDataSourceCrystal"""
'    Print #f, "GoTo Sai"
'    Print #f, "Sai:"
'    Print #f, "    On Error Resume Next"
'    Print #f, "    Set sb = Nothing"
'    Print #f, "Exit Function"
'    Print #f, ""
'    Print #f, "End Function"
'    Print #f, ""

    Print #f, "Private Sub Class_Terminate()"
    Print #f, "    If mbConexaoDoCliente = False Then"
    Print #f, "        If Not mConn Is Nothing Then"
    Print #f, "            If mConn.State <> adStateClosed Then mConn.Close"
    Print #f, "            Set mConn = Nothing"
    Print #f, "        End If"
    Print #f, "    End If"
    Print #f, "End Sub"

    Close #f

    Screen.MousePointer = vbDefault
    MsgBox "OK, gerado arquivo " & arq, vbInformation, "Atenção"

End Sub

Private Sub cmdBase_Click()

    Dim tb() As TCampos

    If Not ObtemDadosTabela(tb) Then Exit Sub
    Call GeraBase(tb)

End Sub

Private Sub cmdColecao_Click()
    Call GeraColecao
End Sub

Private Sub cmdDAO_Click()

    Dim tb() As TCampos

    If Not ObtemDadosTabela(tb) Then Exit Sub
    Call GeraDAO(tb)

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtTabela.SetFocus
End Sub

Private Sub Form_Load()
    Call Option1_Click(0)
End Sub

Private Sub Option1_Click(Index As Integer)

    If Index = 0 Then
        txtStrCon.Text = "Provider=OraOLEDB.Oracle;Data Source=master;User Id=TESTESOFFICEORACLE;Password=DISYS"
    Else
        txtStrCon.Text = "Provider=SQLOLEDB.1;Initial Catalog=TESTESSSDESENV02;Persist Security Info=True;User ID=TESTESSSDESENV02;Password=DISYS;Data Source=desenv02\SQLEXPRESS"
    End If

End Sub
