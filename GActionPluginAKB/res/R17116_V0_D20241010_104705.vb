Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17116

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

    Public Shared Function R17116(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P61519 As Double, ByVal P62736 As Object) As DataTable
    ' 
    ' 17116 - Verifica Cadastro de Pessoas - Versão: 0
    ' 
        'On Error GoTo Err_R17116
        Dim sCurrComponent as String

        Dim QueryC374418 As New Object
        Dim RsC374418 As New Object
        Dim C374418DataType() As Short
        Dim RsC374418_EOF As Boolean
        Dim C374419 As Object
        Dim C374419DataType As Short
        Dim C374420 As Object
        Dim C374420DataType As Short
        Dim C374421 As Object
        Dim C374421DataType As Short
        Dim C374422 As Object
        Dim C374422DataType As Short
        Dim C374423 As Object
        Dim C374423DataType As Short
        Dim C374424 As Object
        Dim C374424DataType As Short
        Dim C374425 As Object
        Dim C374425DataType As Short
        Dim C374426 As Object
        Dim C374426DataType As Short
        Dim C374427 As Object
        Dim C374427DataType As Short
        Dim C374428 As Object
        Dim C374428DataType As Short
        Dim C374429 As Boolean
        Dim C374429DataType As Short
        Dim C374430 As Boolean
        Dim C374430DataType As Short
        Dim C374431 As Boolean
        Dim C374431DataType As Short
        Dim C374432 As Boolean
        Dim C374432DataType As Short
        Dim C374433 As Boolean
        Dim C374433DataType As Short
        Dim C374435 As Boolean
        Dim C374435DataType As Short
        Dim C374436 As Boolean
        Dim C374436DataType As Short
        Dim C374437 As Boolean
        Dim C374437DataType As Short
        Dim C374438 As Object
        Dim C374438DataType As Short
        Dim C374439 As Object
        Dim C374439DataType As Short
        Dim C374440 As Object
        Dim C374440DataType As Short
        Dim C374442 As Object
        Dim C374442DataType As Short
        Dim C374443 As Object
        Dim C374443DataType As Short
        Dim C374444 As Object
        Dim C374444DataType As Short
        Dim C374445 As Object
        Dim C374445DataType As Short
        Dim C374446 As Object
        Dim C374446DataType As Short
        Dim C374447 As Object
        Dim C374447DataType As Short
        Dim C374448 As Object
        Dim C374448DataType As Short
        Dim C374449 As Object
        Dim C374449DataType As Short
        Dim C374450 As Object
        Dim C374450DataType As Short
        Dim C374451 As Object
        Dim C374451DataType As Short
        Dim C374452 As Object
        Dim C374452DataType As Short
        Dim C374453 As Object
        Dim C374453DataType As Short
        Dim C374454 As Object
        Dim C374454DataType As Short
        Dim C374455 As Short
        Dim C374455DataType As Short
        Dim C374455ReturnDataType() As Short

        Dim C374456 As Short
        Dim C374456DataType As Short
        Dim C374456ReturnDataType() As Short

        Dim C374457 As Object
        Dim C374457DataType As Short
        Dim C374458DataType() As Short
        Dim C374459 As Boolean
        Dim C374459DataType As Short
        Dim C374461 As Boolean
        Dim C374461DataType As Short
        Dim C374462 As Object
        Dim C374462DataType As Short
        Dim C374671 As Object
        Dim C374671DataType As Short
        Dim C374672 As Object
        Dim C374672DataType As Short
        Dim C374682 As Boolean
        Dim C374682DataType As Short
        Dim C374685 As Object
        Dim C374685DataType As Short
        Dim C374686 As Object
        Dim C374686DataType As Short
        Dim C374687 As Boolean
        Dim C374687DataType As Short
        Dim C374688 As Object
        Dim C374688DataType As Short
        Dim C374689 As Object
        Dim C374689DataType As Short
        Dim QueryC374690 As New Object
        Dim RsC374690 As New Object
        Dim C374690DataType() As Short
        Dim RsC374690_EOF As Boolean
        Dim C376373 As Object
        Dim C376373DataType As Short
        Dim C376590 As Object
        Dim C376590DataType As Short
        Dim C376591 As Boolean
        Dim C376591DataType As Short
        Dim C376592 As Object
        Dim C376592DataType As Short
        Dim C376593 As Short
        Dim C376593DataType As Short
        Dim C376593ReturnDataType() As Short

        Dim QueryC377336 As New Object
        Dim RsC377336 As New Object
        Dim C377336DataType() As Short
        Dim RsC377336_EOF As Boolean
        Dim C381423 As Object
        Dim C381423DataType As Short
        Dim C381424 As Boolean
        Dim C381424DataType As Short
        Dim C381425 As Object
        Dim C381425DataType As Short
        Dim C381426 As Object
        Dim C381426DataType As Short
        Dim C447651 As DataTable
        Dim C447651CurrentRow As Integer
        Dim C447651DataType() As Short
        Dim C447652 As Object
        Dim C447652DataType As Short
        Dim C465574 As Object
        Dim C465574DataType As Short
        Dim C465575 As Boolean
        Dim C465575DataType As Short
        P62736 = IIf(IsDBNull(P62736), "S", P62736)

        ReDim ReturnDatatype(0)

        GoTo Comp_C374457

    Comp_C374418:

        ' selDadosPessoa
        sCurrComponent = "selDadosPessoa"
        QueryC374418 = con.CreateCommand()
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "SELECT NVL(WF_PESSOAS.Endereco ,'NULL'),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL(WF_PESSOAS.Ender_Compl , 'NULL'),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL(WF_PESSOAS.Bairro ,'NULL'),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL(WF_PESSOAS.Codigo_Cidade ,0),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL(WF_PESSOAS.CEP ,'NULL'),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL(WF_PESSOAS.Numero_endereco , 0),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL(WF_PES_JURIDICA.CGC ,'NULL'),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL(WF_PES_JURIDICA.Inscr_Estadual, 'NULL'),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "Decode(NVL(WF_PES_JURIDICA.CGC_INVALIDO ,0),1,'CNPJ Inválido', 'CNPJ Válido'),"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL((SELECT WF_COMUNIC.Num_Comunicacao "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "FROM AKBUSER01.WF_COMUNIC "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "WHERE WF_COMUNIC.Cod_Pessoa = WF_PESSOAS.Cod_Pessoa "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND WF_COMUNIC.Cod_Comunic = 1"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND WF_COMUNIC.Cod_Pessoa  = WF_PESSOAS.Cod_Pessoa),'NULL') Fone,"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL((SELECT WF_COMUNIC.Num_Comunicacao "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "FROM AKBUSER01.WF_COMUNIC "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "WHERE WF_COMUNIC.Cod_Pessoa = WF_PESSOAS.Cod_Pessoa "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND WF_COMUNIC.Cod_Comunic = 2"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND WF_COMUNIC.Cod_Pessoa  = WF_PESSOAS.Cod_Pessoa),'NULL') EMail,"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "WF_PES_JURIDICA.Nao_Contribuinte_ICMS ,"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "WF_PESSOAS.Nome_Pessoa ,"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "NVL((SELECT WF_COMUNIC.Num_Comunicacao "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "FROM AKBUSER01.WF_COMUNIC , AKBUSER01.WF_ESTAB_JURIDICO_PARAM"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "WHERE WF_COMUNIC.Cod_Pessoa = WF_PESSOAS.Cod_Pessoa "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND WF_COMUNIC.Cod_Comunic = WF_ESTAB_JURIDICO_PARAM.Cod_Comun_Email_DANFE"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND WF_COMUNIC.Cod_Pessoa = WF_PESSOAS.Cod_Pessoa"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(C447652, C447652DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND ROWNUM < 2 ),'NULL') EMail_Danfe,"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "WF_PES_JURIDICA.codigo_ident_destinatario"
        QueryC374418.CommandText = QueryC374418.CommandText & " " & ""
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "FROM AKBUSER01.WF_PESSOAS , AKBUSER01.WF_PES_JURIDICA "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & ""
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "WHERE WF_PES_JURIDICA.Cod_Pessoa = WF_PESSOAS.Cod_Pessoa "
        QueryC374418.CommandText = QueryC374418.CommandText & " " & "AND  WF_PESSOAS.Cod_Pessoa = " & _ 
ConvertValue(P61519, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC374418.Transaction = txn
        RsC374418 = QueryC374418.ExecuteReader()
        Dim iC374418 As Short
        ReDim C374418DataType(RsC374418.FieldCount)
        For iC374418 = 0 to RsC374418.FieldCount - 1
            Select Case RsC374418.GetDataTypeName(iC374418).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C374418DataType(iC374418 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C374418DataType(iC374418 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C374418DataType(iC374418 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC374418
        RsC374418_EOF = Not RsC374418.Read()

        GoTo Comp_C377336

    Comp_C374419:

        ' varDadosInexistentes
        sCurrComponent = "varDadosInexistentes"
        C374419 = System.DBNull.Value
        C374419DataType = 3
        GoTo Comp_C447651

    Comp_C374420:

        ' vEndereço
        sCurrComponent = "vEndereço"
        C374420DataType = 0
        C374420 = RsC374418(0)
        C374420DataType = C374418Datatype(1)
        If C374420DataType = AKBTypeConst.cAKBTypeString Then
          C374420 = IIF(IsDBNull(C374420), C374420, Strings.RTrim(C374420))
        End If 

        GoTo Comp_C374422

    Comp_C374421:

        ' vIE
        sCurrComponent = "vIE"
        C374421DataType = 0
        C374421DataType = C374418Datatype(8)
        If C374421DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(7)) Then
          C374421 = Strings.RTrim(RsC374418(7))
        Else 
          C374421 = RsC374418(7)
        End If 

        GoTo Comp_C374428

    Comp_C374422:

        ' vEndCompl
        sCurrComponent = "vEndCompl"
        C374422DataType = 0
        C374422DataType = C374418Datatype(2)
        If C374422DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(1)) Then
          C374422 = Strings.RTrim(RsC374418(1))
        Else 
          C374422 = RsC374418(1)
        End If 

        GoTo Comp_C374423

    Comp_C374423:

        ' vBairro
        sCurrComponent = "vBairro"
        C374423DataType = 0
        C374423DataType = C374418Datatype(3)
        If C374423DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(2)) Then
          C374423 = Strings.RTrim(RsC374418(2))
        Else 
          C374423 = RsC374418(2)
        End If 

        GoTo Comp_C374424

    Comp_C374424:

        ' vCidade
        sCurrComponent = "vCidade"
        C374424DataType = 0
        C374424DataType = C374418Datatype(4)
        If C374424DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(3)) Then
          C374424 = Strings.RTrim(RsC374418(3))
        Else 
          C374424 = RsC374418(3)
        End If 

        GoTo Comp_C374425

    Comp_C374425:

        ' vCEP
        sCurrComponent = "vCEP"
        C374425DataType = 0
        C374425DataType = C374418Datatype(5)
        If C374425DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(4)) Then
          C374425 = Strings.RTrim(RsC374418(4))
        Else 
          C374425 = RsC374418(4)
        End If 

        GoTo Comp_C374426

    Comp_C374426:

        ' vNumEndereço
        sCurrComponent = "vNumEndereço"
        C374426DataType = 0
        C374426DataType = C374418Datatype(6)
        If C374426DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(5)) Then
          C374426 = Strings.RTrim(RsC374418(5))
        Else 
          C374426 = RsC374418(5)
        End If 

        GoTo Comp_C374427

    Comp_C374427:

        ' vCNPJ
        sCurrComponent = "vCNPJ"
        C374427DataType = 0
        C374427DataType = C374418Datatype(7)
        If C374427DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(6)) Then
          C374427 = Strings.RTrim(RsC374418(6))
        Else 
          C374427 = RsC374418(6)
        End If 

        GoTo Comp_C374421

    Comp_C374428:

        ' vCNPJInval
        sCurrComponent = "vCNPJInval"
        C374428DataType = 0
        C374428DataType = C374418Datatype(9)
        If C374428DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(8)) Then
          C374428 = Strings.RTrim(RsC374418(8))
        Else 
          C374428 = RsC374418(8)
        End If 

        GoTo Comp_C374671

    Comp_C374429:

        ' Sem Endereço?
        sCurrComponent = "Sem Endereço?"
        C374429 = (fn_ConvertValueCompiled(C374420, C374420DataType, AKB_DecimalPoint, False) = "NULL")
        C374429DataType = AKBTypeConst.cAKBTypeLogical
        If C374429 Then
            GoTo Comp_C374444
        Else
            GoTo Comp_C374437
        End If

    Comp_C374430:

        ' Sem Cidade?
        sCurrComponent = "Sem Cidade?"
        C374430 = (fn_ConvertValueCompiled(C374424, C374424DataType, AKB_DecimalPoint, False) = 0)
        C374430DataType = AKBTypeConst.cAKBTypeLogical
        If C374430 Then
            GoTo Comp_C374447
        Else
            GoTo Comp_C381424
        End If

    Comp_C374431:

        ' Sem Número?
        sCurrComponent = "Sem Número?"
        C374431 = (fn_ConvertValueCompiled(C374426, C374426DataType, AKB_DecimalPoint, False) = 0)
        C374431DataType = AKBTypeConst.cAKBTypeLogical
        If C374431 Then
            GoTo Comp_C374449
        Else
            GoTo Comp_C374432
        End If

    Comp_C374432:

        ' Sem CNPJ?
        sCurrComponent = "Sem CNPJ?"
        C374432 = (fn_ConvertValueCompiled(C374427, C374427DataType, AKB_DecimalPoint, False) = "NULL" AND fn_ConvertValueCompiled(RsC377336(0), C377336DataType(1), AKB_DecimalPoint, False)  = "BRASIL")
        C374432DataType = AKBTypeConst.cAKBTypeLogical
        If C374432 Then
            GoTo Comp_C374451
        Else
            GoTo Comp_C374433
        End If

    Comp_C374433:

        ' Sem IE?
        sCurrComponent = "Sem IE?"
        C374433 = (fn_ConvertValueCompiled(C374421, C374421DataType, AKB_DecimalPoint, False) = "NULL"  AND fn_ConvertValueCompiled(C376373, C376373DataType, AKB_DecimalPoint, False) = 0  AND fn_ConvertValueCompiled(RsC377336(0), C377336DataType(1), AKB_DecimalPoint, False)  = "BRASIL")
        C374433DataType = AKBTypeConst.cAKBTypeLogical
        If C374433 Then
            GoTo Comp_C465575
        Else
            GoTo Comp_C374682
        End If

    Comp_C374435:

        ' Sem Bairro?
        sCurrComponent = "Sem Bairro?"
        C374435 = (fn_ConvertValueCompiled(C374423, C374423DataType, AKB_DecimalPoint, False) = "NULL")
        C374435DataType = AKBTypeConst.cAKBTypeLogical
        If C374435 Then
            GoTo Comp_C374446
        Else
            GoTo Comp_C374430
        End If

    Comp_C374436:

        ' Sem CEP?
        sCurrComponent = "Sem CEP?"
        C374436 = (fn_ConvertValueCompiled(C374425, C374425DataType, AKB_DecimalPoint, False) = "NULL" AND fn_ConvertValueCompiled(RsC377336(0), C377336DataType(1), AKB_DecimalPoint, False)  = "BRASIL")
        C374436DataType = AKBTypeConst.cAKBTypeLogical
        If C374436 Then
            GoTo Comp_C374448
        Else
            GoTo Comp_C374431
        End If

    Comp_C374437:

        ' Sem EndCompl
        sCurrComponent = "Sem EndCompl"
        C374437 = (1=2)
        C374437DataType = AKBTypeConst.cAKBTypeLogical
        If C374437 Then
            GoTo Comp_C374445
        Else
            GoTo Comp_C374435
        End If

    Comp_C374438:

        ' atvEndereço
        sCurrComponent = "atvEndereço"
        C374438DataType = 4
        C374419 = fn_ConvertValueCompiled(C374444 , 3, AKB_DecimalPoint)
        C374438 = True
        GoTo Comp_C374437

    Comp_C374439:

        ' atvEndCompl
        sCurrComponent = "atvEndCompl"
        C374439DataType = 4
        C374419 = fn_ConvertValueCompiled(C374445 , 3, AKB_DecimalPoint)
        C374439 = True
        GoTo Comp_C374435

    Comp_C374440:

        ' atvBairro
        sCurrComponent = "atvBairro"
        C374440DataType = 4
        C374419 = fn_ConvertValueCompiled(C374446 , 3, AKB_DecimalPoint)
        C374440 = True
        GoTo Comp_C374430

    Comp_C374442:

        ' atvCidade
        sCurrComponent = "atvCidade"
        C374442DataType = 4
        C374419 = fn_ConvertValueCompiled(C374447 , 3, AKB_DecimalPoint)
        C374442 = True
        GoTo Comp_C381424

    Comp_C374443:

        ' atvCEP
        sCurrComponent = "atvCEP"
        C374443DataType = 4
        C374419 = fn_ConvertValueCompiled(C374448 , 3, AKB_DecimalPoint)
        C374443 = True
        GoTo Comp_C374431

    Comp_C374444:

        ' Concat1
        sCurrComponent = "Concat1"
        C374444DataType = 3
        C374444 = C374419  & "Endereço /"
        GoTo Comp_C374438

    Comp_C374445:

        ' Concat2
        sCurrComponent = "Concat2"
        C374445DataType = 3
        C374445 = C374419  & "Endereço Complemento /"
        GoTo Comp_C374439

    Comp_C374446:

        ' Concat3
        sCurrComponent = "Concat3"
        C374446DataType = 3
        C374446 = C374419  & "Bairro /"
        GoTo Comp_C374440

    Comp_C374447:

        ' Concat4
        sCurrComponent = "Concat4"
        C374447DataType = 3
        C374447 = C374419  & "Cidade /"
        GoTo Comp_C374442

    Comp_C374448:

        ' Concat6
        sCurrComponent = "Concat6"
        C374448DataType = 3
        C374448 = C374419  & "CEP /"
        GoTo Comp_C374443

    Comp_C374449:

        ' Concat7
        sCurrComponent = "Concat7"
        C374449DataType = 3
        C374449 = C374419  & "Número Endereço /"
        GoTo Comp_C374450

    Comp_C374450:

        ' atvNumeroEnd
        sCurrComponent = "atvNumeroEnd"
        C374450DataType = 4
        C374419 = fn_ConvertValueCompiled(C374449 , 3, AKB_DecimalPoint)
        C374450 = True
        GoTo Comp_C374432

    Comp_C374451:

        ' Concat8
        sCurrComponent = "Concat8"
        C374451DataType = 3
        C374451 = C374419  & "CNPJ /"
        GoTo Comp_C374452

    Comp_C374452:

        ' atvCNPJ
        sCurrComponent = "atvCNPJ"
        C374452DataType = 4
        C374419 = fn_ConvertValueCompiled(C374451 , 3, AKB_DecimalPoint)
        C374452 = True
        GoTo Comp_C374433

    Comp_C374453:

        ' Concat9
        sCurrComponent = "Concat9"
        C374453DataType = 3
        C374453 = C374419  & "Inscrição Estadual /"
        GoTo Comp_C374454

    Comp_C374454:

        ' atvIE
        sCurrComponent = "atvIE"
        C374454DataType = 4
        C374419 = fn_ConvertValueCompiled(C374453 , 3, AKB_DecimalPoint)
        C374454 = True
        GoTo Comp_C374682

    Comp_C374455:

        ' msgComCNPJInvalido
        sCurrComponent = "msgComCNPJInvalido"
        C374455 = 1
        C374455DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C374458

    Comp_C374456:

        ' msgmSemCNPJInvalido
        sCurrComponent = "msgmSemCNPJInvalido"
        C374456 = 1
        C374456DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C374458

    Comp_C374457:

        ' vRet
        sCurrComponent = "vRet"
        C374457 = 0
        C374457DataType = 4
        GoTo Comp_C374419

    Comp_C374458:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C374458 As DataTable = New DataTable()
        tb_C374458.Columns.Add("vRet" & "")
        Dim row_C374458 As DataRow = tb_C374458.NewRow()
        row_C374458(0) = C374457
        tb_C374458.Rows.Add(row_C374458)
        R17116 = tb_C374458
        ReDim C374458DataType(1)
        C374458DataType(1) = C374457DataType
        ReturnDataType = C374458DataType

        GoTo Exit_R17116

    Comp_C374459:

        ' Tem dados inexistentes?
        sCurrComponent = "Tem dados inexistentes?"
        C374459 = (fn_ConvertValueCompiled(RsC374690(0), C374690DataType(1), AKB_DecimalPoint, False)  <> "NULL"  And fn_ConvertValueCompiled(C374428, C374428DataType, AKB_DecimalPoint, False) = "CNPJ Inválido")
        C374459DataType = AKBTypeConst.cAKBTypeLogical
        If C374459 Then
            GoTo Comp_C374455
        Else
            GoTo Comp_C374461
        End If

    Comp_C374461:

        ' Tem Dados Inexistente e CNPJ Válido?
        sCurrComponent = "Tem Dados Inexistente e CNPJ Válido?"
        C374461 = (fn_ConvertValueCompiled(RsC374690(0), C374690DataType(1), AKB_DecimalPoint, False) <> "NULL"  And fn_ConvertValueCompiled(C374428, C374428DataType, AKB_DecimalPoint, False) = "CNPJ Válido")
        C374461DataType = AKBTypeConst.cAKBTypeLogical
        If C374461 Then
            GoTo Comp_C374456
        Else
            GoTo Comp_C374462
        End If

    Comp_C374462:

        ' atvTrue
        sCurrComponent = "atvTrue"
        C374462DataType = 4
        C374457 = fn_ConvertValueCompiled(1, 4, AKB_DecimalPoint)
        C374462 = True
        GoTo Comp_C374458

    Comp_C374671:

        ' vFone
        sCurrComponent = "vFone"
        C374671DataType = 0
        C374671DataType = C374418Datatype(10)
        If C374671DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(9)) Then
          C374671 = Strings.RTrim(RsC374418(9))
        Else 
          C374671 = RsC374418(9)
        End If 

        GoTo Comp_C374672

    Comp_C374672:

        ' vEmail
        sCurrComponent = "vEmail"
        C374672DataType = 0
        C374672DataType = C374418Datatype(11)
        If C374672DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(10)) Then
          C374672 = Strings.RTrim(RsC374418(10))
        Else 
          C374672 = RsC374418(10)
        End If 

        GoTo Comp_C376373

    Comp_C374682:

        ' Sem Fone?
        sCurrComponent = "Sem Fone?"
        C374682 = (fn_ConvertValueCompiled(C374671, C374671DataType, AKB_DecimalPoint, False) = "NULL")
        C374682DataType = AKBTypeConst.cAKBTypeLogical
        If C374682 Then
            GoTo Comp_C374685
        Else
            GoTo Comp_C374687
        End If

    Comp_C374685:

        ' Concat10
        sCurrComponent = "Concat10"
        C374685DataType = 3
        C374685 = C374419  & "Telefone /"
        GoTo Comp_C374686

    Comp_C374686:

        ' atvFone
        sCurrComponent = "atvFone"
        C374686DataType = 4
        C374419 = fn_ConvertValueCompiled(C374685 , 3, AKB_DecimalPoint)
        C374686 = True
        GoTo Comp_C374687

    Comp_C374687:

        ' Sem E-Mail?
        sCurrComponent = "Sem E-Mail?"
        C374687 = (1=2)
        C374687DataType = AKBTypeConst.cAKBTypeLogical
        If C374687 Then
            GoTo Comp_C374688
        Else
            GoTo Comp_C376592
        End If

    Comp_C374688:

        ' Concat11
        sCurrComponent = "Concat11"
        C374688DataType = 3
        C374688 = C374419  & "E-mail /"
        GoTo Comp_C374689

    Comp_C374689:

        ' atvEMail
        sCurrComponent = "atvEMail"
        C374689DataType = 4
        C374419 = fn_ConvertValueCompiled(C374688 , 3, AKB_DecimalPoint)
        C374689 = True
        GoTo Comp_C376592

    Comp_C374690:

        ' selDados
        sCurrComponent = "selDados"
        QueryC374690 = con.CreateCommand()
        QueryC374690.CommandText = QueryC374690.CommandText & " " & "SELECT nvl(" & _ 
ConvertValue(C374419, C374419DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",'NULL')  FROM DUAL"
        QueryC374690.Transaction = txn
        RsC374690 = QueryC374690.ExecuteReader()
        Dim iC374690 As Short
        ReDim C374690DataType(RsC374690.FieldCount)
        For iC374690 = 0 to RsC374690.FieldCount - 1
            Select Case RsC374690.GetDataTypeName(iC374690).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C374690DataType(iC374690 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C374690DataType(iC374690 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C374690DataType(iC374690 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC374690
        RsC374690_EOF = Not RsC374690.Read()

        GoTo Comp_C374459

    Comp_C376373:

        ' vNãoContribuinteICMS
        sCurrComponent = "vNãoContribuinteICMS"
        C376373DataType = 0
        C376373DataType = C374418Datatype(12)
        If C376373DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(11)) Then
          C376373 = Strings.RTrim(RsC374418(11))
        Else 
          C376373 = RsC374418(11)
        End If 

        GoTo Comp_C376590

    Comp_C376590:

        ' vNome
        sCurrComponent = "vNome"
        C376590DataType = 0
        C376590DataType = C374418Datatype(13)
        If C376590DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(12)) Then
          C376590 = Strings.RTrim(RsC374418(12))
        Else 
          C376590 = RsC374418(12)
        End If 

        GoTo Comp_C381423

    Comp_C376591:

        ' Nome Maior que 60?
        sCurrComponent = "Nome Maior que 60?"
        C376591 = (fn_ConvertValueCompiled(C376592, C376592DataType, AKB_DecimalPoint, False) > 60)
        C376591DataType = AKBTypeConst.cAKBTypeLogical
        If C376591 Then
            GoTo Comp_C376593
        Else
            GoTo Comp_C374690
        End If

    Comp_C376592:

        ' LengthNome
        sCurrComponent = "LengthNome"
        C376592DataType = 1
        C376592 = Len(C376590 & "")
        GoTo Comp_C376591

    Comp_C376593:

        ' MSG1
        sCurrComponent = "MSG1"
        C376593 = 1
        C376593DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C374690

    Comp_C377336:

        ' SelPais
        sCurrComponent = "SelPais"
        QueryC377336 = con.CreateCommand()
        QueryC377336.CommandText = QueryC377336.CommandText & " " & "SELECT TRIM(wf_paises.Descr_Pais)"
        QueryC377336.CommandText = QueryC377336.CommandText & " " & ""
        QueryC377336.CommandText = QueryC377336.CommandText & " " & "FROM wf_pessoas, wf_cidades, wf_estados, wf_paises"
        QueryC377336.CommandText = QueryC377336.CommandText & " " & ""
        QueryC377336.CommandText = QueryC377336.CommandText & " " & "WHERE wf_pessoas.cod_pessoa = " & _ 
ConvertValue(P61519, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC377336.CommandText = QueryC377336.CommandText & " " & "AND wf_pessoas.codigo_cidade = wf_cidades.codigo_cidade"
        QueryC377336.CommandText = QueryC377336.CommandText & " " & "AND wf_cidades.Sigla_estado = wf_estados.Sigla_estado"
        QueryC377336.CommandText = QueryC377336.CommandText & " " & "AND wf_estados.codigo_pais = wf_paises.codigo_pais "
        QueryC377336.Transaction = txn
        RsC377336 = QueryC377336.ExecuteReader()
        Dim iC377336 As Short
        ReDim C377336DataType(RsC377336.FieldCount)
        For iC377336 = 0 to RsC377336.FieldCount - 1
            Select Case RsC377336.GetDataTypeName(iC377336).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C377336DataType(iC377336 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C377336DataType(iC377336 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C377336DataType(iC377336 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC377336
        RsC377336_EOF = Not RsC377336.Read()

        GoTo Comp_C374420

    Comp_C381423:

        ' vE-Mail_Danfe
        sCurrComponent = "vE-Mail_Danfe"
        C381423DataType = 0
        C381423DataType = C374418Datatype(14)
        If C381423DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(13)) Then
          C381423 = Strings.RTrim(RsC374418(13))
        Else 
          C381423 = RsC374418(13)
        End If 

        GoTo Comp_C465574

    Comp_C381424:

        ' Tem Email Danfe
        sCurrComponent = "Tem Email Danfe"
        C381424 = (fn_ConvertValueCompiled(C381423, C381423DataType, AKB_DecimalPoint, False) = "NULL" AND fn_ConvertValueCompiled(P62736, 3, AKB_DecimalPoint, False)  = "S")
        C381424DataType = AKBTypeConst.cAKBTypeLogical
        If C381424 Then
            GoTo Comp_C381425
        Else
            GoTo Comp_C374436
        End If

    Comp_C381425:

        ' Concat5
        sCurrComponent = "Concat5"
        C381425DataType = 3
        C381425 = C374419  & "E-Mail Danfe /"
        GoTo Comp_C381426

    Comp_C381426:

        ' atvE-MailDanfe
        sCurrComponent = "atvE-MailDanfe"
        C381426DataType = 4
        C374419 = fn_ConvertValueCompiled(C381425 , 3, AKB_DecimalPoint)
        C381426 = True
        GoTo Comp_C374436

    Comp_C447651:

        ' EstabDefault
        sCurrComponent = "EstabDefault"
        C447651 = clsRuleDynamicallyCompiled_R1770.R1770(con, txn)
        C447651CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C447651) Then
          iColumns = C447651.Columns.Count
        End If
        ReDim C447651DataType(iColumns)
        For iCol = 1 To iColumns
            C447651DataType(iCol) = clsRuleDynamicallyCompiled_R1770.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C447652

    Comp_C447652:

        ' vCodEstab
        sCurrComponent = "vCodEstab"
        C447652DataType = 0
        C447652 = C447651.rows(C447651CurrentRow - 1)(1)
        C447652DataType = C447651DataType(2)
        GoTo Comp_C374418

    Comp_C465574:

        ' vCodIdentDesti
        sCurrComponent = "vCodIdentDesti"
        C465574DataType = 0
        C465574DataType = C374418Datatype(15)
        If C465574DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC374418(14)) Then
          C465574 = Strings.RTrim(RsC374418(14))
        Else 
          C465574 = RsC374418(14)
        End If 

        GoTo Comp_C374429

    Comp_C465575:

        ' vCodIdent=2
        sCurrComponent = "vCodIdent=2"
        C465575 = (fn_ConvertValueCompiled(C465574, C465574DataType, AKB_DecimalPoint, False) = 2)
        C465575DataType = AKBTypeConst.cAKBTypeLogical
        If C465575 Then
            GoTo Comp_C374682
        Else
            GoTo Comp_C374453
        End If

    Exit_R17116:

        Exit Function

    Err_R17116:
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
                    fn_ConvertValueCompiled = IIf(vValue = 1 Or vValue = True, 1, 0)
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
    Public Shared Function GetDate(ByRef con As OracleConnection) As Object
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select sysdate from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
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
