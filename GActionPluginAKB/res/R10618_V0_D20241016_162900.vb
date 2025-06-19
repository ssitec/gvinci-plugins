Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R10618

    ' Tipo de dado do AKB
    Private Enum AKBTypeConst
        cAKBTypeNone = -1
        cAKBTypeNumeric = 1
        cAKBTypeDate = 2
        cAKBTypeString = 3
        cAKBTypeLogical = 4

        ' atv 13.453-107.368
        cAKBTypeUndefined = 0
        cAKBTypeQuery = 5
    End Enum

    Public Const AKB_DecimalPoint = ","
    Public Const DecimalSymbol = "."
    Public Const DateIdentifier = "'"
    Public Const StringIdentifier = "'"
    Public Shared Sub Main ()

    End Sub

    Public Shared Function R10618(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P34144 As Double, ByVal P34145 As Double) As DataTable
    ' 
    ' 10618 - Gravar Observação de Produtos Inativos no Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R10618
        Dim sCurrComponent as String

        Dim C191255DataType() As Short
        Dim C191256 As Object
        Dim C191256DataType As Short
        Dim QueryC191260 As New Object
        Dim RsC191260 As New Object
        Dim C191260DataType() As Short
        Dim RsC191260_EOF As Boolean
        Dim C191261 As Object
        Dim C191261DataType As Short
        Dim C191262 As Object
        Dim C191262DataType As Short
        Dim C191263 As Boolean
        Dim C191263DataType As Short
        Dim C191264 As Object
        Dim C191264DataType As Short
        Dim C191266 As Object
        Dim C191266DataType As Short
        Dim C191267 As Object
        Dim C191267DataType As Short
        Dim C191268 As Object
        Dim C191268DataType As Short
        Dim C191269 As Boolean
        Dim C191269DataType As Short
        Dim C191271 As Boolean
        Dim C191271DataType As Short
        Dim C191272 As Object
        Dim C191272DataType As Short
        Dim QueryC191273 As New Object
        Dim C191273 As Integer
        Dim C191273DataType As Short
        Dim C191274 As Object
        Dim C191274DataType As Short
        Dim C191275 As Object
        Dim C191275DataType As Short
        Dim C191276 As Object
        Dim C191276DataType As Short
        Dim C191277 As Object
        Dim C191277DataType As Short
        Dim C191278 As Object
        Dim C191278DataType As Short
        Dim C191279 As Object
        Dim C191279DataType As Short
        Dim QueryC191280 As New Object
        Dim C191280 As Integer
        Dim C191280DataType As Short
        Dim QueryC191281 As New Object
        Dim C191281 As Integer
        Dim C191281DataType As Short
        Dim QueryC191473 As New Object
        Dim RsC191473 As New Object
        Dim C191473DataType() As Short
        Dim RsC191473_EOF As Boolean
        Dim C191474 As Object
        Dim C191474DataType As Short
        Dim C191475 As Boolean
        Dim C191475DataType As Short
        Dim C191476 As Object
        Dim C191476DataType As Short
        Dim C191477 As Short
        Dim C191477DataType As Short
        Dim C191477ReturnDataType() As Short

        ReDim ReturnDatatype(0)

        GoTo Comp_C191256

    Comp_C191255:

        ' Ret
        sCurrComponent = "Ret"
        Dim tb_C191255 As DataTable = New DataTable()
        tb_C191255.Columns.Add("vRet" & "")
        Dim row_C191255 As DataRow = tb_C191255.NewRow()
        row_C191255(0) = C191256
        tb_C191255.Rows.Add(row_C191255)
        R10618 = tb_C191255
        ReDim C191255DataType(1)
        C191255DataType(1) = C191256DataType
        ReturnDataType = C191255DataType

        GoTo Exit_R10618

    Comp_C191256:

        ' vRet
        sCurrComponent = "vRet"
        C191256 = 1
        C191256DataType = 4
        GoTo Comp_C191274

    Comp_C191260:

        ' Sel_ProdInativo
        sCurrComponent = "Sel_ProdInativo"
        QueryC191260 = con.CreateCommand()
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "Select WF_PEDIDO_ITENS.Sigla_Prod, "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "              NVL(WF_PEDIDO_ITENS.Acesso,0),"
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "              WF_ESTAB_PARAM_SIGLA.Departamento "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & ""
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "From AKBUSER01.WF_PEDIDO, AKBUSER01.WF_PEDIDO_ITENS, AKBUSER01.WF_ACESSOS,  AKBUSER01.WF_ESTAB_PARAM_SIGLA"
        QueryC191260.CommandText = QueryC191260.CommandText & " " & ""
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "Where WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_PEDIDO_ITENS.Sigla_Prod = WF_ACESSOS.Sigla_Prod "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_PEDIDO_ITENS.Acesso = WF_ACESSOS.Acesso"
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_PEDIDO.Cod_Pessoa_Estab_Juridico  = WF_ESTAB_PARAM_SIGLA.Cod_Pessoa_Estab_Juridico "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_PEDIDO_ITENS.Sigla_Prod = WF_ESTAB_PARAM_SIGLA.Sigla_Prod"
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_ACESSOS.Produto_Inativo = 1"
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_PEDIDO.Tp_Produto  = " & _ 
ConvertValue(P34144, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And      WF_PEDIDO.Pedido = " & _ 
ConvertValue(P34145, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191260.CommandText = QueryC191260.CommandText & " " & ""
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "And       WF_ESTAB_PARAM_SIGLA.Departamento Is Not Null"
        QueryC191260.CommandText = QueryC191260.CommandText & " " & "Order by  WF_PEDIDO_ITENS.Sigla_Prod, WF_PEDIDO_ITENS.Acesso"
        QueryC191260.Transaction = txn
        RsC191260 = QueryC191260.ExecuteReader()
        Dim iC191260 As Short
        ReDim C191260DataType(RsC191260.FieldCount)
        For iC191260 = 0 to RsC191260.FieldCount - 1
            Select Case RsC191260.GetDataTypeName(iC191260).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C191260DataType(iC191260 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C191260DataType(iC191260 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C191260DataType(iC191260 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC191260
        RsC191260_EOF = Not RsC191260.Read()

        GoTo Comp_C191261

    Comp_C191261:

        ' First_ProdInat
        sCurrComponent = "First_ProdInat"
        C191261DataType = 4

        GoTo Comp_C191262

    Comp_C191262:

        ' Eof_ProdInat
        sCurrComponent = "Eof_ProdInat"
        C191262DataType = 4
        C191262 = RsC191260_EOF
        GoTo Comp_C191263

    Comp_C191263:

        ' É Fim Prod Inativo?
        sCurrComponent = "É Fim Prod Inativo?"
        C191263 = (fn_ConvertValueCompiled(C191262, C191262DataType, AKB_DecimalPoint, False) = 1)
        C191263DataType = AKBTypeConst.cAKBTypeLogical
        If C191263 Then
            GoTo Comp_C191473
        Else
            GoTo Comp_C191266
        End If

    Comp_C191264:

        ' Acesso_Prod
        sCurrComponent = "Acesso_Prod"
        C191264DataType = 0
        C191264DataType = C191260Datatype(2)
        If C191264DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC191260(1)) Then
          C191264 = Strings.RTrim(RsC191260(1))
        Else 
          C191264 = RsC191260(1)
        End If 

        GoTo Comp_C191278

    Comp_C191266:

        ' Sigla_Prod
        sCurrComponent = "Sigla_Prod"
        C191266DataType = 0
        C191266 = RsC191260(0)
        C191266DataType = C191260Datatype(1)
        If C191266DataType = AKBTypeConst.cAKBTypeString Then
          C191266 = IIF(IsDBNull(C191266), C191266, Strings.RTrim(C191266))
        End If 

        GoTo Comp_C191276

    Comp_C191267:

        ' Next_ProdInat
        sCurrComponent = "Next_ProdInat"
        C191267DataType = 4
        RsC191260_EOF = Not RsC191260.Read()
        C191267 = True

        GoTo Comp_C191262

    Comp_C191268:

        ' vSigla
        sCurrComponent = "vSigla"
        C191268 = "0"
        C191268DataType = 3
        GoTo Comp_C191260

    Comp_C191269:

        ' vSigla=0?
        sCurrComponent = "vSigla=0?"
        C191269 = (fn_ConvertValueCompiled(C191268, C191268DataType, AKB_DecimalPoint, False) = 0)
        C191269DataType = AKBTypeConst.cAKBTypeLogical
        If C191269 Then
            GoTo Comp_C191281
        Else
            GoTo Comp_C191264
        End If

    Comp_C191271:

        ' Não inseriu a sigla?
        sCurrComponent = "Não inseriu a sigla?"
        C191271 = ((fn_ConvertValueCompiled(C191268, C191268DataType, AKB_DecimalPoint, False) = 0) OR (fn_ConvertValueCompiled(C191268, C191268DataType, AKB_DecimalPoint, False) <> fn_ConvertValueCompiled(C191266, C191266DataType, AKB_DecimalPoint, False) ))
        C191271DataType = AKBTypeConst.cAKBTypeLogical
        If C191271 Then
            GoTo Comp_C191272
        Else
            GoTo Comp_C191280
        End If

    Comp_C191272:

        ' Seq_Obs
        sCurrComponent = "Seq_Obs"
        C191272DataType = 1
        C191272 = ObjTable_NewIdentity (con, "WF_PED_OBSERV")
        GoTo Comp_C191273

    Comp_C191273:

        ' Ins_Obs
        sCurrComponent = "Ins_Obs"
        QueryC191273 = con.CreateCommand()
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_OBSERV ( WF_PED_OBSERV.SEQUENCIA, "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "WF_PED_OBSERV.Tp_Produto, "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "WF_PED_OBSERV.Pedido,"
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "WF_PED_OBSERV.Departamento, "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "WF_PED_OBSERV.Observacao, "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "WF_PED_OBSERV.Data_Geracao, "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "WF_PED_OBSERV.Usuario_Geracao, "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "WF_PED_OBSERV.Destinatario ) "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & ""
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "VALUES( " & _ 
ConvertValue(C191272, C191272DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "" & _ 
ConvertValue(P34144, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "" & _ 
ConvertValue(P34145, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "" & _ 
ConvertValue(C191278, C191278DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & ""
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "'Produtos na sigla ' || " & _ 
ConvertValue(C191266, C191266DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " || ' com cadastros Inativos: ' ||CHR(13)||CHR(10)||'Acesso ' || " & _ 
ConvertValue(C191264, C191264DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & ""
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "" & _ 
ConvertValue(C191277, C191277DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "" & _ 
ConvertValue(C191276, C191276DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC191273.CommandText = QueryC191273.CommandText & " " & "'Sistema' )"
        QueryC191273.Transaction = txn
        C191273 = QueryC191273.ExecuteNonQuery()
        C191273DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C191275

    Comp_C191274:

        ' vSeq_Observ
        sCurrComponent = "vSeq_Observ"
        C191274 = 0
        C191274DataType = 1
        GoTo Comp_C191268

    Comp_C191275:

        ' Att_SeqObs
        sCurrComponent = "Att_SeqObs"
        C191275DataType = 4
        C191274 = fn_ConvertValueCompiled(C191272 , 1, AKB_DecimalPoint)
        C191275 = True
        GoTo Comp_C191279

    Comp_C191276:

        ' User
        sCurrComponent = "User"
        C191276DataType = 3
        C191276 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C191277

    Comp_C191277:

        ' SysDate
        sCurrComponent = "SysDate"
        C191277DataType = 2
        C191277 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C191269

    Comp_C191278:

        ' Depto
        sCurrComponent = "Depto"
        C191278DataType = 0
        C191278DataType = C191260Datatype(3)
        If C191278DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC191260(2)) Then
          C191278 = Strings.RTrim(RsC191260(2))
        Else 
          C191278 = RsC191260(2)
        End If 

        GoTo Comp_C191271

    Comp_C191279:

        ' Att_Sigla2
        sCurrComponent = "Att_Sigla2"
        C191279DataType = 4
        C191268 = fn_ConvertValueCompiled(C191266 , 3, AKB_DecimalPoint)
        C191279 = True
        GoTo Comp_C191267

    Comp_C191280:

        ' Up_Obs
        sCurrComponent = "Up_Obs"
        QueryC191280 = con.CreateCommand()
        QueryC191280.CommandText = QueryC191280.CommandText & " " & "UPDATE AKBUSER01.WF_PED_OBSERV"
        QueryC191280.CommandText = QueryC191280.CommandText & " " & "SET WF_PED_OBSERV.Observacao= ( WF_PED_OBSERV.Observacao || ',' ||"
        QueryC191280.CommandText = QueryC191280.CommandText & " " & "          ' Acesso ' || " & _ 
ConvertValue(C191264, C191264DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC191280.CommandText = QueryC191280.CommandText & " " & "WHERE WF_PED_OBSERV.Tp_Produto = " & _ 
ConvertValue(P34144, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191280.CommandText = QueryC191280.CommandText & " " & "And          WF_PED_OBSERV.Pedido = " & _ 
ConvertValue(P34145, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191280.CommandText = QueryC191280.CommandText & " " & "And          WF_PED_OBSERV.SEQUENCIA = " & _ 
ConvertValue(C191274, C191274DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191280.Transaction = txn
        C191280 = QueryC191280.ExecuteNonQuery()
        C191280DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C191267

    Comp_C191281:

        ' Del_ObsSistema
        sCurrComponent = "Del_ObsSistema"
        QueryC191281 = con.CreateCommand()
        QueryC191281.CommandText = QueryC191281.CommandText & " " & "DELETE FROM AKBUSER01.WF_PED_OBSERV "
        QueryC191281.CommandText = QueryC191281.CommandText & " " & "WHERE WF_PED_OBSERV.Tp_Produto = " & _ 
ConvertValue(P34144, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191281.CommandText = QueryC191281.CommandText & " " & "AND  WF_PED_OBSERV.Pedido = " & _ 
ConvertValue(P34145, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191281.CommandText = QueryC191281.CommandText & " " & "AND  WF_PED_OBSERV.Destinatario = 'Sistema' "
        QueryC191281.Transaction = txn
        C191281 = QueryC191281.ExecuteNonQuery()
        C191281DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C191264

    Comp_C191473:

        ' Sel_VerifDepto
        sCurrComponent = "Sel_VerifDepto"
        QueryC191473 = con.CreateCommand()
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "Select Distinct WF_PEDIDO_ITENS.Sigla_Prod"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & ""
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "From AKBUSER01.WF_PEDIDO, AKBUSER01.WF_PEDIDO_ITENS, AKBUSER01.WF_ACESSOS"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & ""
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "Where WF_PEDIDO.Tp_Produto = WF_PEDIDO_ITENS.Tp_Produto "
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "And      WF_PEDIDO.Pedido = WF_PEDIDO_ITENS.Pedido "
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "And      WF_PEDIDO_ITENS.Sigla_Prod = WF_ACESSOS.Sigla_Prod "
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "And      WF_PEDIDO_ITENS.Acesso = WF_ACESSOS.Acesso"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "And      WF_ACESSOS.Produto_Inativo = 1"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "And      WF_PEDIDO.Tp_Produto  = " & _ 
ConvertValue(P34144, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "And      WF_PEDIDO.Pedido = " & _ 
ConvertValue(P34145, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC191473.CommandText = QueryC191473.CommandText & " " & ""
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "And     Not Exists ( Select WF_ESTAB_PARAM_SIGLA.Sigla_Prod"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "                                    From AKBUSER01.WF_ESTAB_PARAM_SIGLA"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "                                    Where WF_ESTAB_PARAM_SIGLA.Cod_Pessoa_Estab_Juridico = WF_PEDIDO.Cod_Pessoa_Estab_Juridico"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "                                    And      WF_ESTAB_PARAM_SIGLA.Sigla_Prod = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "                                    And      WF_ESTAB_PARAM_SIGLA.Departamento Is Not Null"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "                                  )"
        QueryC191473.CommandText = QueryC191473.CommandText & " " & ""
        QueryC191473.CommandText = QueryC191473.CommandText & " " & "Order by  WF_PEDIDO_ITENS.Sigla_Prod"
        QueryC191473.Transaction = txn
        RsC191473 = QueryC191473.ExecuteReader()
        Dim iC191473 As Short
        ReDim C191473DataType(RsC191473.FieldCount)
        For iC191473 = 0 to RsC191473.FieldCount - 1
            Select Case RsC191473.GetDataTypeName(iC191473).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C191473DataType(iC191473 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C191473DataType(iC191473 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C191473DataType(iC191473 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC191473
        RsC191473_EOF = Not RsC191473.Read()

        GoTo Comp_C191474

    Comp_C191474:

        ' EOF_VerifDepto
        sCurrComponent = "EOF_VerifDepto"
        C191474DataType = 4
        C191474 = RsC191473_EOF
        GoTo Comp_C191475

    Comp_C191475:

        ' EOF_VerifDepto é fim?
        sCurrComponent = "EOF_VerifDepto é fim?"
        C191475 = (fn_ConvertValueCompiled(C191474, C191474DataType, AKB_DecimalPoint, False) = 1)
        C191475DataType = AKBTypeConst.cAKBTypeLogical
        If C191475 Then
            GoTo Comp_C191255
        Else
            GoTo Comp_C191476
        End If

    Comp_C191476:

        ' ML_SiglaSemDepto
        sCurrComponent = "ML_SiglaSemDepto"
        C191476DataType = 5
        C191476 = ""
        Do Until RsC191473(0).EOF
            If RTrim(C191476) <> "" Then
                C191476 = C191476 & ", "
            End If
            C191476 = C191476 & ConvertValue(RsC191473(0), 0, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss")
            RsC191473(0).MoveNext
        Loop

        GoTo Comp_C191477

    Comp_C191477:

        ' Msg_SiglaSemDepto
        sCurrComponent = "Msg_SiglaSemDepto"
        C191477 = 1
        C191477DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C191255

    Exit_R10618:

        Exit Function

    Err_R10618:
        If Not IsNothing(txn) Then
            txn.Rollback()
        End If
    End Function
    '
    ' Converte um valor String para o tipo apropriado
    '
    Public Shared Function fn_ConvertValueCompiled(ByVal vValue As Object, ByVal lDatatype As Integer, ByVal sDecimalPoint As String, _
                                            Optional bExpression As Boolean = false) As Object

        If Not IsDBNull(vValue) Then

            ' Verifica o tipo do dado
            Select Case lDatatype
                Case AKBTypeConst.cAKBTypeNumeric
                    If CStr(vValue) <> "" Then
                        fn_ConvertValueCompiled = CDbl(vValue)
                    End If
                Case AKBTypeConst.cAKBTypeDate
                    fn_ConvertValueCompiled = CDate(vValue)
                Case AKBTypeConst.cAKBTypeString, AKBTypeConst.cAKBTypeQuery
                    fn_ConvertValueCompiled = vValue
                Case AKBTypeConst.cAKBTypeLogical
                    fn_ConvertValueCompiled = IIf(CBool(vValue) = 1 Or CBool(vValue) = True, 1, 0)
                Case Else
                    If CStr(vValue) = "True" Or CStr(vValue) = "False" Then
                        fn_ConvertValueCompiled = IIf(vValue = 1 Or vValue = True, 1, 0)
                    ElseIf IsNumeric(vValue) Then
                        fn_ConvertValueCompiled = CDbl(vValue)
                    ElseIf IsDate(vValue) Then
                        fn_ConvertValueCompiled = CDate(vValue)
                    Else
                        fn_ConvertValueCompiled = vValue
                    End If
            End Select
        Else
            If Not bExpression Then
                fn_ConvertValueCompiled = System.DBNull.Value
            Else

                ' Verifica o tipo do dado
                Select Case lDatatype
                    Case AKBTypeConst.cAKBTypeNumeric
                        fn_ConvertValueCompiled = 0
                    Case AKBTypeConst.cAKBTypeDate
                        fn_ConvertValueCompiled = ""
                    Case AKBTypeConst.cAKBTypeString, AKBTypeConst.cAKBTypeQuery
                        fn_ConvertValueCompiled = ""
                    Case AKBTypeConst.cAKBTypeLogical
                        fn_ConvertValueCompiled = 0
                    Case Else
                        fn_ConvertValueCompiled = ""
                End Select
            End If
        End If
    End Function
    '
    ' Retorna o inteiro maior ou igual ao número
    '
    Public Shared Function fn_Ceiling(ByVal dParam As Double) As Integer

        ' Obtém valor
        If dParam <= 0 Then
            fn_Ceiling = Fix(dParam)
        Else
            If Int(dParam) = dParam Then
                fn_Ceiling = dParam
            Else
                fn_Ceiling = Fix(dParam) + 1
            End If
        End If
    End Function
    '
    ' Retorna o inteiro menor ou igual ao número
    '
    Public Shared Function fn_Floor(ByVal dParam As Double) As Integer

        ' Obtém valor
        If dParam >= 0 Then
            fn_Floor = Fix(dParam)
        Else
            If Int(dParam) = dParam Then
                fn_Floor = dParam
            Else
                fn_Floor = Fix(dParam) - 1
            End If
        End If
    End Function
    Public Shared ReturnDatatype() As Short
    Private Shared iColumns As Short
    Private Shared iCol As Short
    Public Shared Function GetReturnDatatype() As Short()
        GetReturnDatatype = ReturnDatatype
    End Function

    '
    'Retorna o valor por extenso
    '
    Public Shared Function Extenso(valor)

       Dim g(10, 1) As String
       Dim CEN(9) As String
       Dim UNI(9) As String
       Dim DEZ(19) As String
       Dim PARTE(9) As String
       Dim vlr As Double
       Dim Numero$
       Dim p
       Dim i
       Dim TEXTO
       Dim c, D, U
       Dim DE
       Dim VALOR_ANT
       Dim VALOR_POS
       Dim X1, X2, X3
       Dim E1, E2, E3
       Dim GRANDEZA

       If Not IsNumeric(valor) Then
           Extenso = "Valor Não Numerico"
           Exit Function
       End If

       g(1, 0) = "CENTAVO"
       g(2, 0) = ""
       g(3, 0) = "MIL"
       g(4, 0) = "MILHÃO"
       g(5, 0) = "BILHÃO"
       g(6, 0) = "TRILHÃO"

       g(1, 1) = "CENTAVOS"
       g(2, 1) = ""
       g(3, 1) = "MIL"
       g(4, 1) = "MILHÕES"
       g(5, 1) = "BILHÕES"
       g(6, 1) = "TRILHÕES"

       CEN(0) = ""
       CEN(1) = "CENTO"
       CEN(2) = "DUZENTOS"
       CEN(3) = "TREZENTOS"
       CEN(4) = "QUATROCENTOS"
       CEN(5) = "QUINHENTOS"
       CEN(6) = "SEISCENTOS"
       CEN(7) = "SETECENTOS"
       CEN(8) = "OITOCENTOS"
       CEN(9) = "NOVECENTOS"

       DEZ(0) = ""
       DEZ(1) = "DEZ"
       DEZ(2) = "VINTE"
       DEZ(3) = "TRINTA"
       DEZ(4) = "QUARENTA"
       DEZ(5) = "CINQUENTA"
       DEZ(6) = "SESSENTA"
       DEZ(7) = "SETENTA"
       DEZ(8) = "OITENTA"
       DEZ(9) = "NOVENTA"
       DEZ(11) = "ONZE"
       DEZ(12) = "DOZE"
       DEZ(13) = "TREZE"
       DEZ(14) = "QUATORZE"
       DEZ(15) = "QUINZE"
       DEZ(16) = "DEZESSEIS"
       DEZ(17) = "DEZESSETE"
       DEZ(18) = "DEZOITO"
       DEZ(19) = "DEZENOVE"

       UNI(0) = ""
       UNI(1) = "UM"
       UNI(2) = "DOIS"
       UNI(3) = "TRES"
       UNI(4) = "QUATRO"
       UNI(5) = "CINCO"
       UNI(6) = "SEIS"
       UNI(7) = "SETE"
       UNI(8) = "OITO"
       UNI(9) = "NOVE"

       vlr = CDbl(valor)
       Numero$ = Format$(valor, "000000000000000.00")
       p = 1
       For i = 6 To 2 Step -1
          PARTE$(i) = Mid$(Numero$, p, 3)
          p = p + 3
       Next
       PARTE$(1) = 0 & Right$(Numero$, 2)

       TEXTO = ""

       For i = 6 To 1 Step -1
           If Val(PARTE$(i)) <> 0 Then
              c = Val(Left$(PARTE$(i), 1))
              D = Val(Mid$(PARTE$(i), 2, 1))
              U = Val(Mid$(PARTE$(i), 3, 1))
              DE = Val(Mid$(PARTE$(i), 2, 2))
              VALOR_ANT = "0"
              VALOR_POS = "0"
              If i < 6 Then
                 VALOR_ANT = Left$(Numero$, (6 - i) * 3)
              End If
              If i > 2 Then
                 VALOR_POS = Mid$(Numero$, (7 - i) * 3 + 1, (i - 2) * 3)
              End If
              X1 = ""
              X2 = ""
              X3 = ""
              If c = 1 And DE = 0 Then
                   X1 = "CEM"
              Else
                   X1 = CEN(c)
              End If
              If DE < 11 Or DE > 19 Then
                   X2 = DEZ(D)
                   X3 = UNI(U)
              Else
                   X2 = DEZ(DE)
              End If
              'DETERMINA E's
              E1 = " "
              E2 = ""
              E3 = ""
              'Entre as centenas, dezenas e unidades
              If c > 0 And DE > 0 Then
                   E2 = " E "
              End If
              If D > 1 And U > 0 Then
                   E3 = " E "
              End If

              'Entre grandezas diferentes
              If i > 1 And c = 0 And DE <> 0 And Val(VALOR_ANT) <> 0 And Val(VALOR_POS) = 0 Then
                   E2 = " E "
              End If
              If i > 1 And c > 0 And DE = 0 And Val(VALOR_ANT) <> 0 And Val(VALOR_POS) = 0 Then
                   E1 = " E "
              End If

              'Entre a parte inteira e os centavos
              If i = 1 And DE <> 0 And Int(valor) <> 0 Then
                   E1 = " E "
              End If
              If Val(PARTE$(i)) > 1 Then
                   GRANDEZA = g(i, 1)
              Else
                   GRANDEZA = g(i, 0)
              End If
              TEXTO = TEXTO & E1 & X1 & E2 & X2 & E3 & X3 & " " & GRANDEZA
           End If
           If i = 2 Then
               If Int(valor) = 1 Then
                   TEXTO = TEXTO & " REAL"
               Else
                   If PARTE$(3) = "000" And PARTE$(2) = "000" Then
                       TEXTO = TEXTO & " DE REAIS"
                   Else
                       TEXTO = TEXTO & " REAIS"
                   End If
               End If
           End If

       Next
       Extenso = Trim(TEXTO)
    End Function
    '
    ' Validar CGC
    '
    Public Shared Function CalculaCGC(Numero As String) As String

    Dim i As Short
    Dim prod As Short
    Dim mult As Short
    Dim digito As Short

    If Not IsNumeric(Numero) Then
      CalculaCGC = """"
      Exit Function
    End If

    mult = 2
    For i = Len(Numero) To 1 Step -1
     prod = prod + Val(Mid(Numero, i, 1)) * mult
     mult = IIf(mult = 9, 2, mult + 1)
    Next

    digito = 11 - Int(prod Mod 11)
    digito = IIf(digito = 10 Or digito = 11, 0, digito)

    CalculaCGC = Trim(Str(digito))

    End Function
    '
    'Validar CGC
    '
    Public Shared Function ValidaCGC(CGC As String) As Boolean

       'Retira os caracteres de formatação
       CGC = Replace(CGC, ".", "")
       CGC = Replace(CGC, "/", "")
       CGC = Replace(CGC, "-", "")

    If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
      ValidaCGC = False
      Exit Function
    End If

    If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
      ValidaCGC = False
      Exit Function
    End If

    ValidaCGC = True

    End Function

    '
    'Validar CPF
    '
    Public Shared Function calculacpf(CPF As String) As Boolean

       'Retira os caracteres de formatação
       CPF = Replace(CPF, ".", "")
       CPF = Replace(CPF, "/", "")
       CPF = Replace(CPF, "-", "")

       'Esta rotina foi adaptada da revista Fórum Access
       On Error GoTo Err_CPF
       Dim i As Short        'utilizada nos FOR... NEXT
       Dim strcampo As String  'armazena do CPF que será utilizada para o cálculo
       Dim strCaracter As String   'armazena os dígitos do CPF da direita para a esquerda
       Dim intNumero As Short    'armazena o digito separado para cálculo (uma a um)
       Dim intMais As Short  'armazena o digito específico multiplicado pela sua base
       Dim lngSoma As Integer     'armazena a soma dos dígitos multiplicados pela sua base(intmais)
       Dim dblDivisao As Double    'armazena a divisão dos dígitos * base por 11
       Dim lngInteiro As Integer  'armazena inteiro da divisão
       Dim intResto As Short     'armazena o resto
       Dim intDig1 As Short  'armazena o 1º digito verificador
       Dim intDig2 As Short  'armazena o 2º digito verificador
       Dim strConf As String   'armazena o digito verificador

       lngSoma = 0
       intNumero = 0
       intMais = 0
       strcampo = Left(CPF, 9)

       'Inicia cálculos do 1º dígito
       For i = 2 To 10
           strCaracter = Right(strcampo, i - 1)
           intNumero = Left(strCaracter, 1)
           intMais = intNumero * i
           lngSoma = lngSoma + intMais
       Next i
       dblDivisao = lngSoma / 11

       lngInteiro = Int(dblDivisao) * 11
       intResto = lngSoma - lngInteiro
       If intResto = 0 Or intResto = 1 Then
           intDig1 = 0
       Else
           intDig1 = 11 - intResto
       End If

       strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
       lngSoma = 0
       intNumero = 0
       intMais = 0
       'Inicia cálculos do 2º dígito
       For i = 2 To 11
           strCaracter = Right(strcampo, i - 1)
           intNumero = Left(strCaracter, 1)
           intMais = intNumero * i
           lngSoma = lngSoma + intMais
       Next i
       dblDivisao = lngSoma / 11
       lngInteiro = Int(dblDivisao) * 11
       intResto = lngSoma - lngInteiro
       If intResto = 0 Or intResto = 1 Then
           intDig2 = 0
       Else
           intDig2 = 11 - intResto
       End If
       strConf = intDig1 & intDig2
       'Caso o CPF esteja errado dispara a mensagem
       If strConf <> Right(CPF, 2) Then
           calculacpf = False
       Else
           calculacpf = True
       End If
       Exit Function

    Exit_CPF:
       Exit Function
    Err_CPF:
       MsgBox(Err.Description)
       Resume Exit_CPF
    End Function
    '
    'Converte um valor para uma String no formato esperado pelo Banco
    '
    Public Shared Function ConvertValue(ByVal Value As Object, ByVal ValueType As Integer, ByVal DecimalSymbol As String, ByVal DateIdentifier As String, ByVal StringIdentifier As String, ByVal DateFormat as String) As String

       On Error GoTo ConvertValue_Error

       Dim sTemp As String

       If Not IsDBNull(Value) Then

           'Verifica o tipo do dado
           Select Case ValueType

               Case AKBTypeConst.cAKBTypeNumeric

                   'Substitui o ponto pelo separador do banco
                   sTemp = Trim(Str(Value))
                   ConvertValue = Replace(sTemp, ".", DecimalSymbol)

               Case AKBTypeConst.cAKBTypeDate
                   ConvertValue = "TO_DATE(" & DateIdentifier & Format(CDate(Value), DateFormat) & DateIdentifier & ",'dd-mm-yyyy hh24:mi:ss')"

               Case AKBTypeConst.cAKBTypeString
                   ConvertValue = StringIdentifier & Replace(Value, StringIdentifier, StrDup(2, StringIdentifier)) & StringIdentifier

               Case AKBTypeConst.cAKBTypeLogical
                   ConvertValue = IIf(Value, 1, 0)

               Case Else
                   ConvertValue = Value

           End Select
       Else

           'Se for nulo retorna a String "Null"
           ConvertValue = "Null"
       End If

       Exit Function

    ConvertValue_Error:
       MsgBox("Erro ao executar a versão compilada GVINCI da função ConvertValue ")
    End Function
    '
    ' Retorna a data atual do banco de dados
    '
    Public Shared Function GetDate(ByRef con As OracleConnection, ByRef txn As OracleTransaction) As Object
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select sysdate from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Query.Transaction = txn
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            GetDate = Rs(0)
        End If
        Rs.Close()
    End Function
    '
    ' Retorna a data atual do banco de dados
    '
    Public Shared Function ObjTable_NewIdentity(ByRef con As OracleConnection, ByVal TableName as String) As Long
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select SEQ_" & TableName & ".nextval from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            ObjTable_NewIdentity = Rs(0)
        End If
        Rs.Close()
    End Function

End Class
