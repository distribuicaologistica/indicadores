Attribute VB_Name = "EnviarIndicadores"
Option Explicit

' ============================================================
'  Envio de lançamentos para o Supabase
'
'  Como usar:
'  1. No Excel: Alt+F11 -> Arquivo -> Importar Arquivo -> este .bas
'  2. Monte a planilha assim, a partir da linha 2:
'
'       A            B          C       D
'       Data         Motorista  Valor   Observação (opcional)
'       01/08/2026   1042       2,40
'       01/08/2026   1077       1,90    chuva
'       02/08/2026   1042       3,15
'
'     A coluna B é o código do motorista no Promax.
'
'  3. Rode a macro EnviarLancamentos (Alt+F8).
'
'  A senha é pedida na hora e não fica gravada no arquivo.
' ============================================================

Private Const SUPABASE_URL As String = "https://nquoksknvkahnfvqvcoo.supabase.co"
Private Const SUPABASE_KEY As String = "sb_publishable_u0e1CVU5F7PTGdQ0U2Mx8g_wPRXUHhh"
Private Const DOMINIO_INTERNO As String = "taruma.com.br"

' Quantas linhas vão por requisição.
Private Const LOTE As Long = 400

Private mToken As String


' ============================================================
'  MACRO PRINCIPAL
' ============================================================
Public Sub EnviarLancamentos()
    Dim ws As Worksheet
    Dim indicador As String
    Dim ultimaLinha As Long, i As Long
    Dim enviados As Long, ignorados As Long
    Dim json As String, itens As Long
    Dim dataTexto As String, valorTexto As String, obs As String, motorista As String
    Dim codigos As Object

    Set ws = ActiveSheet

    indicador = LCase(Trim(InputBox( _
        "Código do indicador:" & vbCrLf & vbCrLf & _
        "devolucao, dispersao, refugo ou rating", _
        "Enviar lançamentos")))
    If indicador = "" Then Exit Sub

    If Not Autenticar() Then Exit Sub

    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "Não encontrei dados. A partir da linha 2: data na coluna A, valor na coluna B.", vbExclamation
        Exit Sub
    End If

    ' Cadastra os motoristas que ainda não existem: a chave estrangeira
    ' recusa lançamento de código desconhecido.
    Set codigos = CreateObject("Scripting.Dictionary")
    For i = 2 To ultimaLinha
        motorista = Trim(CStr(ws.Cells(i, 2).Value))
        If motorista <> "" Then
            If Not codigos.Exists(motorista) Then codigos.Add motorista, 1
        End If
    Next i
    If Not GarantirMotoristas(codigos) Then Exit Sub

    Application.StatusBar = "Enviando lançamentos..."
    json = "["
    itens = 0

    For i = 2 To ultimaLinha
        dataTexto = ParaDataISO(ws.Cells(i, 1).Value)
        motorista = Trim(CStr(ws.Cells(i, 2).Value))
        valorTexto = ParaNumero(ws.Cells(i, 3).Value)
        obs = Trim(CStr(ws.Cells(i, 4).Value))

        If dataTexto = "" Or valorTexto = "" Or motorista = "" Then
            ignorados = ignorados + 1
        Else
            If itens > 0 Then json = json & ","
            json = json & "{""indicador"":""" & indicador & """," & _
                          """data"":""" & dataTexto & """," & _
                          """motorista"":""" & EscaparJson(motorista) & """," & _
                          """valor"":" & valorTexto
            If obs <> "" Then json = json & ",""observacao"":""" & EscaparJson(obs) & """"
            json = json & "}"
            itens = itens + 1
            enviados = enviados + 1
        End If

        ' Fecha o lote e envia.
        If itens >= LOTE Or (i = ultimaLinha And itens > 0) Then
            json = json & "]"
            If Not Gravar(json) Then
                Application.StatusBar = False
                Exit Sub
            End If
            json = "["
            itens = 0
            Application.StatusBar = "Enviados " & enviados & " lançamentos..."
        End If
    Next i

    Application.StatusBar = False

    MsgBox enviados & " lançamento(s) enviado(s)." & _
           IIf(ignorados > 0, vbCrLf & ignorados & " linha(s) ignorada(s) por data, motorista ou valor inválido.", ""), _
           vbInformation, "Concluído"
End Sub


' ============================================================
'  LOGIN — pega o token com matrícula e senha
' ============================================================
Private Function Autenticar() As Boolean
    Dim http As Object
    Dim matricula As String, senha As String
    Dim corpo As String, resposta As String

    If mToken <> "" Then
        Autenticar = True
        Exit Function
    End If

    matricula = Trim(InputBox("Sua matrícula:", "Entrar"))
    If matricula = "" Then Exit Function

    senha = InputBox("Sua senha:", "Entrar")
    If senha = "" Then Exit Function

    corpo = "{""email"":""" & matricula & "@" & DOMINIO_INTERNO & """," & _
            """password"":""" & EscaparJson(senha) & """}"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", SUPABASE_URL & "/auth/v1/token?grant_type=password", False
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Content-Type", "application/json"
    http.send corpo

    resposta = http.responseText

    If http.Status <> 200 Then
        MsgBox "Não deu para entrar." & vbCrLf & vbCrLf & _
               "Confira matrícula e senha. Se você ainda usa a senha do CPF, " & _
               "entre uma vez no painel e crie sua senha antes.", vbCritical
        Exit Function
    End If

    mToken = ExtrairValor(resposta, "access_token")
    Autenticar = (mToken <> "")

    If Not Autenticar Then
        MsgBox "Login aceito, mas não consegui ler o token da resposta.", vbCritical
    End If
End Function


' ============================================================
'  GRAVAÇÃO — upsert em lancamentos
' ============================================================
Private Function Gravar(ByVal json As String) As Boolean
    Dim http As Object
    Dim url As String

    ' Nunca envia corpo vazio: o PostgREST responde PGRST102.
    If Len(Replace(Replace(json, "[", ""), "]", "")) = 0 Then
        Gravar = True
        Exit Function
    End If

    url = SUPABASE_URL & "/rest/v1/lancamentos?on_conflict=indicador,data,motorista"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Authorization", "Bearer " & mToken
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Prefer", "resolution=merge-duplicates,return=minimal"
    http.send json

    If http.Status >= 200 And http.Status < 300 Then
        Gravar = True
    Else
        Gravar = False
        Select Case http.Status
            Case 401
                mToken = ""
                MsgBox "Sua sessão expirou. Rode a macro de novo.", vbExclamation
            Case 403
                MsgBox "Sem permissão para gravar." & vbCrLf & vbCrLf & _
                       "Só quem tem papel de admin pode lançar dados.", vbCritical
            Case Else
                MsgBox "Erro " & http.Status & " ao gravar:" & vbCrLf & vbCrLf & _
                       Left(http.responseText, 400), vbCritical
        End Select
    End If
End Function


' ============================================================
'  MOTORISTAS — cria os códigos que ainda não estão cadastrados
' ============================================================
Private Function GarantirMotoristas(ByVal codigos As Object) As Boolean
    Dim http As Object
    Dim json As String
    Dim chave As Variant
    Dim primeiro As Boolean

    If codigos.Count = 0 Then
        GarantirMotoristas = True
        Exit Function
    End If

    json = "{" & Chr(34) & "codigos" & Chr(34) & ":["
    primeiro = True
    For Each chave In codigos.Keys
        If Not primeiro Then json = json & ","
        json = json & Chr(34) & EscaparJson(CStr(chave)) & Chr(34)
        primeiro = False
    Next chave
    json = json & "]}"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", SUPABASE_URL & "/rest/v1/rpc/garantir_motoristas", False
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Authorization", "Bearer " & mToken
    http.setRequestHeader "Content-Type", "application/json"
    http.send json

    If http.Status >= 200 And http.Status < 300 Then
        GarantirMotoristas = True
    Else
        GarantirMotoristas = False
        MsgBox "Erro " & http.Status & " ao cadastrar os motoristas:" & vbCrLf & vbCrLf & _
               Left(http.responseText, 400), vbCritical
    End If
End Function


' ============================================================
'  CONVERSÕES
' ============================================================
Private Function ParaDataISO(ByVal v As Variant) As String
    On Error GoTo Falhou
    If IsEmpty(v) Or Trim(CStr(v)) = "" Then Exit Function
    If IsDate(v) Then
        ParaDataISO = Format$(CDate(v), "yyyy-mm-dd")
    End If
    Exit Function
Falhou:
    ParaDataISO = ""
End Function

' O JSON exige ponto decimal, independente da configuração do Windows.
' Trata os três casos: célula numérica de verdade, texto com vírgula
' decimal (7,89) e texto com ponto decimal (7.89).
Private Function ParaNumero(ByVal v As Variant) As String
    Dim t As String
    Dim n As Double
    On Error GoTo Falhou

    If IsEmpty(v) Then Exit Function

    ' Célula numérica: o Excel entrega Double, sem ambiguidade de separador.
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal
            ParaNumero = NumeroParaJson(CDbl(v))
            Exit Function
    End Select

    t = Trim(CStr(v))
    If t = "" Then Exit Function

    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then
        t = Replace(t, ".", "")          ' 1.234,56 -> ponto é milhar
        t = Replace(t, ",", ".")
    ElseIf InStr(t, ",") > 0 Then
        t = Replace(t, ",", ".")         ' 7,89
    End If

    If Not TextoNumerico(t) Then Exit Function

    ' Val sempre usa ponto como decimal, ignorando a configuração regional.
    n = Val(t)
    ParaNumero = NumeroParaJson(n)
    Exit Function
Falhou:
    ParaNumero = ""
End Function

' Str$ sempre devolve ponto decimal e nunca deixa separador sobrando,
' ao contrário de Format$, que segue a configuração regional do Windows.
Private Function NumeroParaJson(ByVal n As Double) As String
    Dim t As String

    t = Trim$(Str$(n))
    If Left$(t, 1) = "." Then t = "0" & t
    If Left$(t, 2) = "-." Then t = "-0" & Mid$(t, 2)
    If Right$(t, 1) = "." Then t = Left$(t, Len(t) - 1)

    NumeroParaJson = t
End Function

Private Function TextoNumerico(ByVal t As String) As Boolean
    Dim i As Long, c As String, pontos As Long

    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c = "." Then
            pontos = pontos + 1
            If pontos > 1 Then Exit Function
        ElseIf c = "-" Then
            If i > 1 Then Exit Function
        ElseIf c < "0" Or c > "9" Then
            Exit Function
        End If
    Next i

    TextoNumerico = (Len(t) > 0)
End Function

Private Function EscaparJson(ByVal t As String) As String
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    EscaparJson = t
End Function

' Leitura simples de um campo de texto do JSON, sem biblioteca externa.
Private Function ExtrairValor(ByVal json As String, ByVal campo As String) As String
    Dim marca As String
    Dim inicio As Long, fim As Long

    marca = """" & campo & """:"""
    inicio = InStr(1, json, marca, vbTextCompare)
    If inicio = 0 Then Exit Function

    inicio = inicio + Len(marca)
    fim = InStr(inicio, json, """")
    If fim = 0 Then Exit Function

    ExtrairValor = Mid$(json, inicio, fim - inicio)
End Function


' ============================================================
'  Teste rápido de conexão e permissão
' ============================================================
Public Sub TestarConexao()
    Dim http As Object

    If Not Autenticar() Then Exit Sub

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", SUPABASE_URL & "/rest/v1/indicadores?select=codigo,nome", False
    http.setRequestHeader "apikey", SUPABASE_KEY
    http.setRequestHeader "Authorization", "Bearer " & mToken
    http.send

    If http.Status = 200 Then
        MsgBox "Conexão certa. Indicadores disponíveis:" & vbCrLf & vbCrLf & http.responseText, vbInformation
    Else
        MsgBox "Erro " & http.Status & ":" & vbCrLf & Left(http.responseText, 400), vbCritical
    End If
End Sub


' ============================================================
'  DIAGNÓSTICO — mostra o JSON das 3 primeiras linhas
'  Use quando o servidor recusar o corpo, para ver o que saiu.
' ============================================================
Public Sub VerJsonDeAmostra()
    Dim ws As Worksheet
    Dim i As Long, ultima As Long
    Dim json As String
    Dim d As String, m As String, v As String

    Set ws = ActiveSheet
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima > 4 Then ultima = 4

    json = "["
    For i = 2 To ultima
        d = ParaDataISO(ws.Cells(i, 1).Value)
        m = Trim(CStr(ws.Cells(i, 2).Value))
        v = ParaNumero(ws.Cells(i, 3).Value)

        If i > 2 Then json = json & "," & vbCrLf
        json = json & "{""indicador"":""EXEMPLO""," & _
                      """data"":""" & d & """," & _
                      """motorista"":""" & m & """," & _
                      """valor"":" & v & "}"
    Next i
    json = json & "]"

    MsgBox json, vbInformation, "JSON gerado"
End Sub
