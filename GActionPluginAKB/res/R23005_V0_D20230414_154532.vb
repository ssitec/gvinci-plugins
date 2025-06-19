Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R23005

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

    Public Shared Function R23005(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P89115 As Double, ByVal P89116 As Object, ByVal P89117 As Double, ByVal P89146 As Double, ByVal P89147 As Object, ByVal P89148 As Double, ByVal P89149 As Double, ByVal P89215 As Object, ByVal P89283 As Double, ByVal P89285 As Object, ByVal P89286 As String, ByVal P89287 As Double, ByVal P89288 As Double, ByVal P89289 As Object) As DataTable
    ' 
    ' 23005 - Valida Parametros de Descontos - Itens - Versão: 0
    ' 
        On Error GoTo Err_R23005
        Dim sCurrComponent as String

        Dim QueryC554447 As New Object
        Dim RsC554447 As New Object
        Dim C554447DataType() As Short
        Dim RsC554447_EOF As Boolean
        Dim C555022 As Object
        Dim C555022DataType As Short
        Dim C555025DataType() As Short
        Dim C555026 As Double
        Dim C555026DataType As Short
        Dim QueryC555028 As New Object
        Dim RsC555028 As New Object
        Dim C555028DataType() As Short
        Dim RsC555028_EOF As Boolean
        Dim C555029 As Object
        Dim C555029DataType As Short
        Dim C555030 As Object
        Dim C555030DataType As Short
        Dim C555031 As Object
        Dim C555031DataType As Short
        Dim C555032 As Object
        Dim C555032DataType As Short
        Dim C555033 As Object
        Dim C555033DataType As Short
        Dim C555034 As Object
        Dim C555034DataType As Short
        Dim C555035 As Boolean
        Dim C555035DataType As Short
        Dim C555036 As Double
        Dim C555036DataType As Short
        Dim C555037 As Object
        Dim C555037DataType As Short
        Dim C555038 As Object
        Dim C555038DataType As Short
        Dim C555039 As Object
        Dim C555039DataType As Short
        Dim C555040 As Boolean
        Dim C555040DataType As Short
        Dim C555041 As Double
        Dim C555041DataType As Short
        Dim C555042 As Object
        Dim C555042DataType As Short
        Dim C555043 As Boolean
        Dim C555043DataType As Short
        Dim C555044 As Double
        Dim C555044DataType As Short
        Dim C555045 As Object
        Dim C555045DataType As Short
        Dim C555046 As Boolean
        Dim C555046DataType As Short
        Dim C555047 As Double
        Dim C555047DataType As Short
        Dim C555048 As Object
        Dim C555048DataType As Short
        Dim C555050 As Boolean
        Dim C555050DataType As Short
        Dim C555051 As Double
        Dim C555051DataType As Short
        Dim C555052 As Object
        Dim C555052DataType As Short
        Dim C555053 As Boolean
        Dim C555053DataType As Short
        Dim C555054 As Double
        Dim C555054DataType As Short
        Dim C555055 As Object
        Dim C555055DataType As Short
        Dim C555056 As Boolean
        Dim C555056DataType As Short
        Dim QueryC555058 As New Object
        Dim C555058 As Integer
        Dim C555058DataType As Short
        Dim C555287 As Boolean
        Dim C555287DataType As Short
        Dim C555985 As Object
        Dim C555985DataType As Short
        Dim C555986 As Boolean
        Dim C555986DataType As Short
        Dim QueryC555987 As New Object
        Dim RsC555987 As New Object
        Dim C555987DataType() As Short
        Dim RsC555987_EOF As Boolean
        Dim C555998 As Double
        Dim C555998DataType As Short
        Dim C556000 As Object
        Dim C556000DataType As Short
        Dim C556027 As Boolean
        Dim C556027DataType As Short
        Dim QueryC556028 As New Object
        Dim C556028 As Integer
        Dim C556028DataType As Short
        Dim C556029 As Boolean
        Dim C556029DataType As Short
        Dim C556030 As Boolean
        Dim C556030DataType As Short
        Dim C556031 As Object
        Dim C556031DataType As Short
        Dim QueryC556032 As New Object
        Dim C556032 As Integer
        Dim C556032DataType As Short
        Dim QueryC556348 As New Object
        Dim RsC556348 As New Object
        Dim C556348DataType() As Short
        Dim RsC556348_EOF As Boolean
        Dim C556349 As Double
        Dim C556349DataType As Short
        Dim C556351 As Object
        Dim C556351DataType As Short
        Dim C556352 As Object
        Dim C556352DataType As Short
        Dim C556353 As Object
        Dim C556353DataType As Short
        Dim C556355 As Object
        Dim C556355DataType As Short
        Dim C556356 As Object
        Dim C556356DataType As Short
        Dim C556357 As Object
        Dim C556357DataType As Short
        Dim C556395 As Object
        Dim C556395DataType As Short
        Dim C556396 As Object
        Dim C556396DataType As Short
        Dim QueryC556397 As New Object
        Dim C556397 As Integer
        Dim C556397DataType As Short
        Dim QueryC556398 As New Object
        Dim C556398 As Integer
        Dim C556398DataType As Short
        Dim C556483 As Double
        Dim C556483DataType As Short
        Dim C556484 As Object
        Dim C556484DataType As Short
        Dim C556485 As Boolean
        Dim C556485DataType As Short
        P89215 = IIf(IsDBNull(P89215), 0, P89215)

        ReDim ReturnDatatype(0)

        GoTo Comp_C555022

    Comp_C554447:

        ' SelPorcAvista
        sCurrComponent = "SelPorcAvista"
        QueryC554447 = con.CreateCommand()
        QueryC554447.CommandText = QueryC554447.CommandText & " " & "SELECT NVL(PORC_DESC_DM01,0)"
        QueryC554447.CommandText = QueryC554447.CommandText & " " & "FROM WF_PARAM_TAB_PRECO"
        QueryC554447.CommandText = QueryC554447.CommandText & " " & "WHERE SEQ = " & _ 
ConvertValue(P89115, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC554447.CommandText = QueryC554447.CommandText & " " & "AND Tabela_Preco_Venda = " & _ 
ConvertValue(P89117, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        RsC554447 = QueryC554447.ExecuteReader()
        Dim iC554447 As Short
        ReDim C554447DataType(RsC554447.FieldCount)
        For iC554447 = 0 to RsC554447.FieldCount - 1
            Select Case RsC554447.GetDataTypeName(iC554447).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C554447DataType(iC554447 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C554447DataType(iC554447 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C554447DataType(iC554447 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC554447
        RsC554447_EOF = Not RsC554447.Read()

        GoTo Comp_C555026

    Comp_C555022:

        ' vRet
        sCurrComponent = "vRet"
        C555022 = System.DBNull.Value
        C555022DataType = 1
        GoTo Comp_C555038

    Comp_C555025:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C555025 As DataTable = New DataTable()
        tb_C555025.Columns.Add("vRet" & "")
        Dim row_C555025 As DataRow = tb_C555025.NewRow()
        row_C555025(0) = C555022
        tb_C555025.Rows.Add(row_C555025)
        R23005 = tb_C555025
        ReDim C555025DataType(1)
        C555025DataType(1) = C555022DataType
        ReturnDataType = C555025DataType

        GoTo Exit_R23005

    Comp_C555026:

        ' PreçoAVista
        sCurrComponent = "PreçoAVista"
        C555026 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)  - ((fn_ConvertValueCompiled(RsC554447(0), C554447DataType(1), AKB_DecimalPoint, True) /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) ), 2)
        C555026DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555037

    Comp_C555028:

        ' SelPorcDesc
        sCurrComponent = "SelPorcDesc"
        QueryC555028 = con.CreateCommand()
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "SELECT nvl(porc_desc_01,0),"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "nvl(porc_desc_02,0),"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "nvl(porc_desc_03,0),"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "nvl(porc_desc_04,0),"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "nvl(porc_desc_05,0),"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "nvl(porc_desc_06,0),"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "nvl(Aplic_Red_ICMS,0)"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "FROM WF_PARAM_TAB_PRECO"
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "WHERE SEQ = " & _ 
ConvertValue(P89115, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555028.CommandText = QueryC555028.CommandText & " " & "AND tabela_preco_venda = " & _ 
ConvertValue(P89117, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        RsC555028 = QueryC555028.ExecuteReader()
        Dim iC555028 As Short
        ReDim C555028DataType(RsC555028.FieldCount)
        For iC555028 = 0 to RsC555028.FieldCount - 1
            Select Case RsC555028.GetDataTypeName(iC555028).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C555028DataType(iC555028 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C555028DataType(iC555028 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C555028DataType(iC555028 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC555028
        RsC555028_EOF = Not RsC555028.Read()

        GoTo Comp_C555034

    Comp_C555029:

        ' PorcDesc6
        sCurrComponent = "PorcDesc6"
        C555029DataType = 0
        C555029 = RsC555028(5)
        C555029DataType = C555028Datatype(6)

        GoTo Comp_C555985

    Comp_C555030:

        ' PorcDesc5
        sCurrComponent = "PorcDesc5"
        C555030DataType = 0
        C555030 = RsC555028(4)
        C555030DataType = C555028Datatype(5)

        GoTo Comp_C555029

    Comp_C555031:

        ' PorcDesc4
        sCurrComponent = "PorcDesc4"
        C555031DataType = 0
        C555031 = RsC555028(3)
        C555031DataType = C555028Datatype(4)

        GoTo Comp_C555030

    Comp_C555032:

        ' PorcDesc2
        sCurrComponent = "PorcDesc2"
        C555032DataType = 0
        C555032 = RsC555028(1)
        C555032DataType = C555028Datatype(2)

        GoTo Comp_C555033

    Comp_C555033:

        ' PorcDesc3
        sCurrComponent = "PorcDesc3"
        C555033DataType = 0
        C555033 = RsC555028(2)
        C555033DataType = C555028Datatype(3)

        GoTo Comp_C555031

    Comp_C555034:

        ' PorcDesc1
        sCurrComponent = "PorcDesc1"
        C555034DataType = 0
        C555034 = RsC555028(0)
        C555034DataType = C555028Datatype(1)

        GoTo Comp_C555032

    Comp_C555035:

        ' TemDesc1?
        sCurrComponent = "TemDesc1?"
        C555035 = (fn_ConvertValueCompiled(C555034, C555034DataType, AKB_DecimalPoint, False) > 0)
        C555035DataType = AKBTypeConst.cAKBTypeLogical
        If C555035 Then
            GoTo Comp_C555036
        Else
            GoTo Comp_C555056
        End If

    Comp_C555036:

        ' CalcPorc1
        sCurrComponent = "CalcPorc1"
        C555036 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(C555034, C555034DataType, AKB_DecimalPoint, True)  /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)), 2)
        C555036DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555039

    Comp_C555037:

        ' attPreço
        sCurrComponent = "attPreço"
        C555037DataType = 4
        C555038 = fn_ConvertValueCompiled(C555026 , 1, AKB_DecimalPoint)
        C555037 = True
        GoTo Comp_C555035

    Comp_C555038:

        ' vPrecoCalc
        sCurrComponent = "vPrecoCalc"
        C555038 = P89148 & " "
        C555038DataType = 1
        GoTo Comp_C555028

    Comp_C555039:

        ' attPreço1
        sCurrComponent = "attPreço1"
        C555039DataType = 4
        C555038 = fn_ConvertValueCompiled(C555036 , 1, AKB_DecimalPoint)
        C555039 = True
        GoTo Comp_C555040

    Comp_C555040:

        ' TemDesc2?
        sCurrComponent = "TemDesc2?"
        C555040 = (fn_ConvertValueCompiled(C555032, C555032DataType, AKB_DecimalPoint, False) > 0)
        C555040DataType = AKBTypeConst.cAKBTypeLogical
        If C555040 Then
            GoTo Comp_C555041
        Else
            GoTo Comp_C555056
        End If

    Comp_C555041:

        ' CalcPorc2
        sCurrComponent = "CalcPorc2"
        C555041 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(C555032, C555032DataType, AKB_DecimalPoint, True)  /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)), 2)
        C555041DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555042

    Comp_C555042:

        ' attPreço2
        sCurrComponent = "attPreço2"
        C555042DataType = 4
        C555038 = fn_ConvertValueCompiled(C555041 , 1, AKB_DecimalPoint)
        C555042 = True
        GoTo Comp_C555043

    Comp_C555043:

        ' TemDesc3?
        sCurrComponent = "TemDesc3?"
        C555043 = (fn_ConvertValueCompiled(C555033, C555033DataType, AKB_DecimalPoint, False) > 0)
        C555043DataType = AKBTypeConst.cAKBTypeLogical
        If C555043 Then
            GoTo Comp_C555044
        Else
            GoTo Comp_C555056
        End If

    Comp_C555044:

        ' CalcPorc3
        sCurrComponent = "CalcPorc3"
        C555044 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(C555033, C555033DataType, AKB_DecimalPoint, True)  /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)), 2)
        C555044DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555045

    Comp_C555045:

        ' attPreço3
        sCurrComponent = "attPreço3"
        C555045DataType = 4
        C555038 = fn_ConvertValueCompiled(C555044 , 1, AKB_DecimalPoint)
        C555045 = True
        GoTo Comp_C555046

    Comp_C555046:

        ' TemDesc4?
        sCurrComponent = "TemDesc4?"
        C555046 = (fn_ConvertValueCompiled(C555031, C555031DataType, AKB_DecimalPoint, False) > 0)
        C555046DataType = AKBTypeConst.cAKBTypeLogical
        If C555046 Then
            GoTo Comp_C555047
        Else
            GoTo Comp_C555056
        End If

    Comp_C555047:

        ' CalcPorc4
        sCurrComponent = "CalcPorc4"
        C555047 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(C555031, C555031DataType, AKB_DecimalPoint, True)  /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)), 2)
        C555047DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555048

    Comp_C555048:

        ' attPreço4
        sCurrComponent = "attPreço4"
        C555048DataType = 4
        C555038 = fn_ConvertValueCompiled(C555047 , 1, AKB_DecimalPoint)
        C555048 = True
        GoTo Comp_C555050

    Comp_C555050:

        ' TemDesc5?
        sCurrComponent = "TemDesc5?"
        C555050 = (fn_ConvertValueCompiled(C555030, C555030DataType, AKB_DecimalPoint, False) > 0)
        C555050DataType = AKBTypeConst.cAKBTypeLogical
        If C555050 Then
            GoTo Comp_C555051
        Else
            GoTo Comp_C555056
        End If

    Comp_C555051:

        ' CalcPorc5
        sCurrComponent = "CalcPorc5"
        C555051 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(C555030, C555030DataType, AKB_DecimalPoint, True)  /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)), 2)
        C555051DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555052

    Comp_C555052:

        ' attPreço5
        sCurrComponent = "attPreço5"
        C555052DataType = 4
        C555038 = fn_ConvertValueCompiled(C555051 , 1, AKB_DecimalPoint)
        C555052 = True
        GoTo Comp_C555053

    Comp_C555053:

        ' TemDesc6?
        sCurrComponent = "TemDesc6?"
        C555053 = (fn_ConvertValueCompiled(C555029, C555029DataType, AKB_DecimalPoint, False) > 0)
        C555053DataType = AKBTypeConst.cAKBTypeLogical
        If C555053 Then
            GoTo Comp_C555054
        Else
            GoTo Comp_C555056
        End If

    Comp_C555054:

        ' CalcPorc6
        sCurrComponent = "CalcPorc6"
        C555054 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(C555029, C555029DataType, AKB_DecimalPoint, True)  /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)), 2)
        C555054DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555055

    Comp_C555055:

        ' attPreço6
        sCurrComponent = "attPreço6"
        C555055DataType = 4
        C555038 = fn_ConvertValueCompiled(C555054 , 1, AKB_DecimalPoint)
        C555055 = True
        GoTo Comp_C555056

    Comp_C555056:

        ' PreçoOk?
        sCurrComponent = "PreçoOk?"
        C555056 = (fn_ConvertValueCompiled(P89149, 1, AKB_DecimalPoint, False) >= fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, False))
        C555056DataType = AKBTypeConst.cAKBTypeLogical
        If C555056 Then
            GoTo Comp_C556029
        Else
            GoTo Comp_C556483
        End If

    Comp_C555058:

        ' UpPreçoCalc
        sCurrComponent = "UpPreçoCalc"
        QueryC555058 = con.CreateCommand()
        QueryC555058.CommandText = QueryC555058.CommandText & " " & "UPDATE wf_pre_pedido_itens SET"
        QueryC555058.CommandText = QueryC555058.CommandText & " " & "SEQ_PARAM_TABELA = " & _ 
ConvertValue(P89115, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ","
        QueryC555058.CommandText = QueryC555058.CommandText & " " & "VALOR_UNIT_CALC_DA_SEQ = " & _ 
ConvertValue(C555038, C555038DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ","
        QueryC555058.CommandText = QueryC555058.CommandText & " " & "OBS_BLOQ_PARAM_TAB = 'OK'"
        QueryC555058.CommandText = QueryC555058.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P89116, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555058.CommandText = QueryC555058.CommandText & " " & "AND SEQ_ITEM = " & _ 
ConvertValue(P89147, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C555058 = QueryC555058.ExecuteNonQuery()
        C555058DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555025

    Comp_C555287:

        ' PedAVista?
        sCurrComponent = "PedAVista?"
        C555287 = (fn_ConvertValueCompiled(P89215, 4, AKB_DecimalPoint, False) = 1)
        C555287DataType = AKBTypeConst.cAKBTypeLogical
        If C555287 Then
            GoTo Comp_C554447
        Else
            GoTo Comp_C555035
        End If

    Comp_C555985:

        ' AplicICMS
        sCurrComponent = "AplicICMS"
        C555985DataType = 0
        C555985 = RsC555028(6)
        C555985DataType = C555028Datatype(7)

        GoTo Comp_C555986

    Comp_C555986:

        ' AplicaICMS?
        sCurrComponent = "AplicaICMS?"
        C555986 = (fn_ConvertValueCompiled(C555985, C555985DataType, AKB_DecimalPoint, False) = 1)
        C555986DataType = AKBTypeConst.cAKBTypeLogical
        If C555986 Then
            GoTo Comp_C555987
        Else
            GoTo Comp_C556027
        End If

    Comp_C555987:

        ' SelPorcRedICMS
        sCurrComponent = "SelPorcRedICMS"
        QueryC555987 = con.CreateCommand()
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "select nvl(wf_estados.porc_red_icms_ecommerce,0)"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "from wf_pre_pedido, wf_pessoas, wf_cidades, wf_estados"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "where id_prepedido = " & _ 
ConvertValue(P89116, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and wf_pessoas.cod_pessoa = wf_pre_pedido.cod_cliente"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and wf_pessoas.codigo_cidade = wf_cidades.codigo_cidade"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and wf_estados.sigla_estado = wf_cidades.sigla_estado"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and " & _ 
ConvertValue(P89283, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 1"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & ""
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "UNION"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & ""
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "select nvl(UF.porc_red_icms_ecommerce,0)"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "from WF_UF_PARAM_TABELA PA, wf_estados UF"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "where PA.Tabela_Preco_Venda = " & _ 
ConvertValue(P89117, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and PA.Seq = " & _ 
ConvertValue(P89115, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and PA.Sigla_Estado = " & _ 
ConvertValue(P89285, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and PA.Sigla_Estado = UF.Sigla_Estado"
        QueryC555987.CommandText = QueryC555987.CommandText & " " & "and " & _ 
ConvertValue(P89283, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 2"
        RsC555987 = QueryC555987.ExecuteReader()
        Dim iC555987 As Short
        ReDim C555987DataType(RsC555987.FieldCount)
        For iC555987 = 0 to RsC555987.FieldCount - 1
            Select Case RsC555987.GetDataTypeName(iC555987).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C555987DataType(iC555987 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C555987DataType(iC555987 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C555987DataType(iC555987 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC555987
        RsC555987_EOF = Not RsC555987.Read()

        GoTo Comp_C555998

    Comp_C555998:

        ' CalcDescICMS
        sCurrComponent = "CalcDescICMS"
        C555998 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - ((fn_ConvertValueCompiled(RsC555987(0), C555987DataType(1), AKB_DecimalPoint, True)   /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)), 2)
        C555998DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556000

    Comp_C556000:

        ' attPreçoICMS
        sCurrComponent = "attPreçoICMS"
        C556000DataType = 4
        C555038 = fn_ConvertValueCompiled(C555998 , 1, AKB_DecimalPoint)
        C556000 = True
        GoTo Comp_C556027

    Comp_C556027:

        ' TPREGRA=2?
        sCurrComponent = "TPREGRA=2?"
        C556027 = (fn_ConvertValueCompiled(P89283, 1, AKB_DecimalPoint, False) = 2)
        C556027DataType = AKBTypeConst.cAKBTypeLogical
        If C556027 Then
            GoTo Comp_C554447
        Else
            GoTo Comp_C555287
        End If

    Comp_C556028:

        ' UpObsPreço
        sCurrComponent = "UpObsPreço"
        QueryC556028 = con.CreateCommand()
        QueryC556028.CommandText = QueryC556028.CommandText & " " & "UPDATE wf_pre_pedido_itens SET"
        QueryC556028.CommandText = QueryC556028.CommandText & " " & "OBS_BLOQ_PARAM_TAB = 'PL Menor',"
        QueryC556028.CommandText = QueryC556028.CommandText & " " & "SEQ_PARAM_TABELA = " & _ 
ConvertValue(P89115, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ","
        QueryC556028.CommandText = QueryC556028.CommandText & " " & "VALOR_UNIT_CALC_DA_SEQ = " & _ 
ConvertValue(C555038, C555038DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556028.CommandText = QueryC556028.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P89116, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556028.CommandText = QueryC556028.CommandText & " " & "AND SEQ_ITEM = " & _ 
ConvertValue(P89147, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C556028 = QueryC556028.ExecuteNonQuery()
        C556028DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555025

    Comp_C556029:

        ' TPREGRA=1?a
        sCurrComponent = "TPREGRA=1?a"
        C556029 = (fn_ConvertValueCompiled(P89283, 1, AKB_DecimalPoint, False) = 1)
        C556029DataType = AKBTypeConst.cAKBTypeLogical
        If C556029 Then
            GoTo Comp_C555058
        Else
            GoTo Comp_C556031
        End If

    Comp_C556030:

        ' TPREGRA=1?A1
        sCurrComponent = "TPREGRA=1?A1"
        C556030 = (fn_ConvertValueCompiled(P89283, 1, AKB_DecimalPoint, False) = 1)
        C556030DataType = AKBTypeConst.cAKBTypeLogical
        If C556030 Then
            GoTo Comp_C556028
        Else
            GoTo Comp_C556031
        End If

    Comp_C556031:

        ' Sequencial
        sCurrComponent = "Sequencial"
        C556031DataType = 1
        C556031 = ObjTable_NewIdentity (con, "WF_TAB_PRC_LIQ_CLIENTE")
        GoTo Comp_C556348

    Comp_C556032:

        ' INS1
        sCurrComponent = "INS1"
        QueryC556032 = con.CreateCommand()
        QueryC556032.CommandText = QueryC556032.CommandText & " " & "INSERT INTO AKBUSER01.WF_TAB_PRC_LIQ_CLIENTE ( WF_TAB_PRC_LIQ_CLIENTE.Tabela_Preco_Venda , WF_TAB_PRC_LIQ_CLIENTE.Sigla_Prod , WF_TAB_PRC_LIQ_CLIENTE.Acesso , WF_TAB_PRC_LIQ_CLIENTE.Cod_Embalagem , WF_TAB_PRC_LIQ_CLIENTE.Cod_Cliente , WF_TAB_PRC_LIQ_CLIENTE.Sequencial , WF_TAB_PRC_LIQ_CLIENTE.Preco_Unit_Bruto , WF_TAB_PRC_LIQ_CLIENTE.Porc_Desc , WF_TAB_PRC_LIQ_CLIENTE.Preco_Unit_Liquido , WF_TAB_PRC_LIQ_CLIENTE.Seq_Param , WF_TAB_PRC_LIQ_CLIENTE.Preco_Unit_Liq_Prazo , WF_TAB_PRC_LIQ_CLIENTE.Descr_Acesso , WF_TAB_PRC_LIQ_CLIENTE.Descr_Embalagem , WF_TAB_PRC_LIQ_CLIENTE.Descr_Auxiliar ) VALUES( " & _ 
ConvertValue(P89117, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P89286, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P89287, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P89288, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P89289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556031, C556031DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P89148, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , 0 , " & _ 
ConvertValue(C555038, C555038DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P89115, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556351, C556351DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556352, C556352DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556353, C556353DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556357, C556357DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        C556032 = QueryC556032.ExecuteNonQuery()
        C556032DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C555025

    Comp_C556348:

        ' SelDescr
        sCurrComponent = "SelDescr"
        QueryC556348 = con.CreateCommand()
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "SELECT SUBSTR(A.Descr_Acesso,1,200), "
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "DECODE(ECV.Cod_Embalagem_Int_Emb,NULL,TRIM(E.DESCR_EMBALAGEM), TRIM(E.DESCR_EMBALAGEM) ||' c/ ' || ECV.Fator_Conv_Cpr_Inter || ' '||E1.DESCR_EMBALAGEM),"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "(SELECT EIO.Descr_Espec_Item_Operac"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & " FROM WF_PADRAO_PRODUTO_ITENS PPI, WF_ESPEC_ITENS_OPER EIO"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "  WHERE PPI.Padrao_Produto = A.Padrao_Produto"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "  AND PPI.Item_Oper = 137"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "  AND PPI.Espec_Item_Operac = EIO.Espec_Item_Operac) DESCR_CODCOMERC,"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "(SELECT EIO.Descr_Espec_Item_Operac"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & " FROM WF_PADRAO_PRODUTO_ITENS PPI, WF_ESPEC_ITENS_OPER EIO"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "  WHERE PPI.Padrao_Produto = A.Padrao_Produto"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "  AND PPI.Item_Oper = 139"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "  AND PPI.Espec_Item_Operac = EIO.Espec_Item_Operac) LARGURA"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "FROM WF_ACESSOS A, WF_TABELA_EMBALAGEM E, WF_TABELA_EMBALAGEM E1, WF_EMB_COMP_VDA_PROD ECV"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "WHERE A.SIGLA_PROD = " & _ 
ConvertValue(P89286, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "AND A.ACESSO = " & _ 
ConvertValue(P89287, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "AND ECV.SIGLA_PROD = A.SIGLA_PROD"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "AND ECV.ACESSO = A.ACESSO"
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "AND ECV.Cod_Embalagem = " & _ 
ConvertValue(P89288, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "AND E.Cod_Embalagem = " & _ 
ConvertValue(P89288, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556348.CommandText = QueryC556348.CommandText & " " & "AND ECV.Cod_Embalagem_Int_Emb = E1.Cod_Embalagem (+)"
        RsC556348 = QueryC556348.ExecuteReader()
        Dim iC556348 As Short
        ReDim C556348DataType(RsC556348.FieldCount)
        For iC556348 = 0 to RsC556348.FieldCount - 1
            Select Case RsC556348.GetDataTypeName(iC556348).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556348DataType(iC556348 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556348DataType(iC556348 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556348DataType(iC556348 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556348
        RsC556348_EOF = Not RsC556348.Read()

        GoTo Comp_C556349

    Comp_C556349:

        ' PrcLiqPrazo
        sCurrComponent = "PrcLiqPrazo"
        C556349 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True)  + ((fn_ConvertValueCompiled(RsC554447(0), C554447DataType(1), AKB_DecimalPoint, True) /100) * fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) ), 2)
        C556349DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556351

    Comp_C556351:

        ' vPrcLiqPrazo
        sCurrComponent = "vPrcLiqPrazo"
        C556351 = C556349 
        C556351DataType = 1
        GoTo Comp_C556395

    Comp_C556352:

        ' DescrAcesso
        sCurrComponent = "DescrAcesso"
        C556352DataType = 0
        C556352 = RsC556348(0)
        C556352DataType = C556348Datatype(1)

        GoTo Comp_C556353

    Comp_C556353:

        ' DescrEmb
        sCurrComponent = "DescrEmb"
        C556353DataType = 0
        C556353 = RsC556348(1)
        C556353DataType = C556348Datatype(2)

        GoTo Comp_C556355

    Comp_C556355:

        ' CodComercial
        sCurrComponent = "CodComercial"
        C556355DataType = 0
        C556355 = RsC556348(2)
        C556355DataType = C556348Datatype(3)

        GoTo Comp_C556356

    Comp_C556356:

        ' Largura
        sCurrComponent = "Largura"
        C556356DataType = 0
        C556356 = RsC556348(3)
        C556356DataType = C556348Datatype(4)

        GoTo Comp_C556357

    Comp_C556357:

        ' Conc
        sCurrComponent = "Conc"
        C556357DataType = 3
        C556357 = C556355  & C556356 
        GoTo Comp_C556032

    Comp_C556395:

        ' SysDate
        sCurrComponent = "SysDate"
        C556395DataType = 2
        C556395 = CDate(Format(GetDate(con), "dd/MM/yyyy HH:mm:ss"))
        GoTo Comp_C556396

    Comp_C556396:

        ' Usuario
        sCurrComponent = "Usuario"
        C556396DataType = 3
        C556396 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C556397

    Comp_C556397:

        ' EXC1
        sCurrComponent = "EXC1"
        QueryC556397 = con.CreateCommand()
        QueryC556397.CommandText = QueryC556397.CommandText & " " & "DELETE WF_TAB_VIRTUAL_CLIENTE WHERE Tabela_Preco_Venda = " & _ 
ConvertValue(P89117, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND Cod_Cliente = " & _ 
ConvertValue(P89289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C556397 = QueryC556397.ExecuteNonQuery()
        C556397DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556398

    Comp_C556398:

        ' INS2
        sCurrComponent = "INS2"
        QueryC556398 = con.CreateCommand()
        QueryC556398.CommandText = QueryC556398.CommandText & " " & "INSERT INTO AKBUSER01.WF_TAB_VIRTUAL_CLIENTE ( WF_TAB_VIRTUAL_CLIENTE.Tabela_Preco_Venda , WF_TAB_VIRTUAL_CLIENTE.Cod_Cliente , WF_TAB_VIRTUAL_CLIENTE.Data_Hora_Ult_Proces , WF_TAB_VIRTUAL_CLIENTE.Usuario_Ultimo_Proces ) VALUES( " & _ 
ConvertValue(P89117, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(P89289, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556395, C556395DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C556396, C556396DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        C556398 = QueryC556398.ExecuteNonQuery()
        C556398DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556352

    Comp_C556483:

        ' Menos1Cent
        sCurrComponent = "Menos1Cent"
        C556483 = Math.Round(fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, True) - 0.01, 2)
        C556483DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C556484

    Comp_C556484:

        ' attPreço7
        sCurrComponent = "attPreço7"
        C556484DataType = 4
        C555038 = fn_ConvertValueCompiled(C556483 , 1, AKB_DecimalPoint)
        C556484 = True
        GoTo Comp_C556485

    Comp_C556485:

        ' PreçoOkArredondado?
        sCurrComponent = "PreçoOkArredondado?"
        C556485 = (fn_ConvertValueCompiled(P89149, 1, AKB_DecimalPoint, False) >= fn_ConvertValueCompiled(C555038, C555038DataType, AKB_DecimalPoint, False))
        C556485DataType = AKBTypeConst.cAKBTypeLogical
        If C556485 Then
            GoTo Comp_C556029
        Else
            GoTo Comp_C556030
        End If

    Exit_R23005:

        Exit Function

    Err_R23005:
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
