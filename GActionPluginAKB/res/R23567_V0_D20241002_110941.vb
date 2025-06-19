Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R23567

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

    Public Shared Function R23567(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P91857 As Double, ByVal P91861 As Double) As DataTable
    ' 
    ' 23567 - Calcula Total Faturado do CNPJ do Tipo Empresa - Versão: 0
    ' 
        On Error GoTo Err_R23567
        Dim sCurrComponent as String

        Dim QueryC578754 As New Object
        Dim RsC578754 As New Object
        Dim C578754DataType() As Short
        Dim RsC578754_EOF As Boolean
        Dim C578755 As Object
        Dim C578755DataType As Short
        Dim C578756 As Object
        Dim C578756DataType As Short
        Dim QueryC578757 As New Object
        Dim RsC578757 As New Object
        Dim C578757DataType() As Short
        Dim RsC578757_EOF As Boolean
        Dim C578758DataType() As Short
        Dim QueryC578760 As New Object
        Dim RsC578760 As New Object
        Dim C578760DataType() As Short
        Dim RsC578760_EOF As Boolean
        Dim C578761 As Object
        Dim C578761DataType As Short
        Dim C578762 As Object
        Dim C578762DataType As Short
        Dim C578763 As Object
        Dim C578763DataType As Short
        Dim C578764 As Object
        Dim C578764DataType As Short
        Dim C578765 As Object
        Dim C578765DataType As Short
        Dim C578766 As Object
        Dim C578766DataType As Short
        Dim C578767 As Object
        Dim C578767DataType As Short
        Dim QueryC578768 As New Object
        Dim RsC578768 As New Object
        Dim C578768DataType() As Short
        Dim RsC578768_EOF As Boolean
        Dim C578769 As Boolean
        Dim C578769DataType As Short
        Dim C578770 As Object
        Dim C578770DataType As Short
        Dim C578771 As Boolean
        Dim C578771DataType As Short
        Dim C578773 As Object
        Dim C578773DataType As Short
        Dim C578774 As Double
        Dim C578774DataType As Short
        Dim C578775 As Boolean
        Dim C578775DataType As Short
        Dim C578776 As Short
        Dim C578776DataType As Short
        Dim C578776ReturnDataType() As Short

        Dim QueryC578777 As New Object
        Dim C578777 As Integer
        Dim C578777DataType As Short
        Dim C578778 As Object
        Dim C578778DataType As Short
        Dim C579031 As Object
        Dim C579031DataType As Short
        Dim C579032 As Boolean
        Dim C579032DataType As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C578767

    Comp_C578754:

        ' Ano
        sCurrComponent = "Ano"
        QueryC578754 = con.CreateCommand()
        QueryC578754.CommandText = QueryC578754.CommandText & " " & "SELECT TO_CHAR(TO_DATE(" & _ 
ConvertValue(C578762, C578762DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ),'YYYY') FROM DUAL"
        RsC578754 = QueryC578754.ExecuteReader()
        Dim iC578754 As Short
        ReDim C578754DataType(RsC578754.FieldCount)
        For iC578754 = 0 to RsC578754.FieldCount - 1
            Select Case RsC578754.GetDataTypeName(iC578754).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C578754DataType(iC578754 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C578754DataType(iC578754 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C578754DataType(iC578754 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC578754
        RsC578754_EOF = Not RsC578754.Read()

        GoTo Comp_C578755

    Comp_C578755:

        ' DataInicial
        sCurrComponent = "DataInicial"
        C578755 = "01/01/" & RsC578754(0) & " "
        C578755DataType = 2
        GoTo Comp_C578756

    Comp_C578756:

        ' DataFinal
        sCurrComponent = "DataFinal"
        C578756 = "31/12/" & RsC578754(0) & " "
        C578756DataType = 2
        GoTo Comp_C578757

    Comp_C578757:

        ' TpEmpToTFat
        sCurrComponent = "TpEmpToTFat"
        QueryC578757 = con.CreateCommand()
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "SELECT WF_TIPO_EMPRESA.Sigla_Empresa,WF_Pess_Juridica_Tipo_Emp.Validade_Inicial,WF_Pess_Juridica_Tipo_Emp.Validade_Final,"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "(SELECT NVL(SUM(NF.Valor_Nota),0)"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  FROM WF_NF_SAIDA NF, WF_TIPO_VENDA TV, WF_NATUREZA_OPERACAO NO"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  WHERE NF.Cod_Pessoa_Estab_Juridico = (" & _ 
ConvertValue(P91857, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND  NF.Cod_Nat_Oper = NO.Cod_Nat_Oper "
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND  NO.Tp_Venda = TV.Tp_Venda "
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND NF.Nota_Cancelada = 0"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND  TV.E_Venda = 1 "
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND NF.Cod_Cliente = WF_Pess_Juridica_Tipo_Emp.Cod_Pessoa"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND WF_Pess_Juridica_Tipo_Emp.Validade_Inicial <= NF.DT_EMISSAO"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND (WF_Pess_Juridica_Tipo_Emp.Validade_Final >= NF.DT_EMISSAO OR WF_Pess_Juridica_Tipo_Emp.Validade_Final IS NULL)"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & ") TOTALFAT, NVL(WF_TIPO_EMPRESA.Limite_Faturamento_Anual,0)"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  FROM WF_Pess_Juridica_Tipo_Emp, WF_TIPO_EMPRESA"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  WHERE WF_Pess_Juridica_Tipo_Emp.Cod_Pessoa = " & _ 
ConvertValue(C578761, C578761DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND WF_Pess_Juridica_Tipo_Emp.Validade_Inicial <= " & _ 
ConvertValue(C578755, C578755DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND (WF_Pess_Juridica_Tipo_Emp.Validade_Final >= " & _ 
ConvertValue(C578756, C578756DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " OR WF_Pess_Juridica_Tipo_Emp.Validade_Final IS NULL)"
        QueryC578757.CommandText = QueryC578757.CommandText & " " & "  AND WF_TIPO_EMPRESA.Tipo_de_Empresa = WF_Pess_Juridica_Tipo_Emp.Tipo_de_Empresa"
        RsC578757 = QueryC578757.ExecuteReader()
        Dim iC578757 As Short
        ReDim C578757DataType(RsC578757.FieldCount)
        For iC578757 = 0 to RsC578757.FieldCount - 1
            Select Case RsC578757.GetDataTypeName(iC578757).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C578757DataType(iC578757 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C578757DataType(iC578757 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C578757DataType(iC578757 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC578757
        RsC578757_EOF = Not RsC578757.Read()

        GoTo Comp_C579031

    Comp_C578758:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C578758 As DataTable = New DataTable()
        tb_C578758.Columns.Add("Ret" & "")
        Dim row_C578758 As DataRow = tb_C578758.NewRow()
        row_C578758(0) = C578767
        tb_C578758.Rows.Add(row_C578758)
        R23567 = tb_C578758
        ReDim C578758DataType(1)
        C578758DataType(1) = C578767DataType
        ReturnDataType = C578758DataType

        GoTo Exit_R23567

    Comp_C578760:

        ' Sel Prepedido
        sCurrComponent = "Sel Prepedido"
        QueryC578760 = con.CreateCommand()
        QueryC578760.CommandText = QueryC578760.CommandText & " " & "SELECT WF_PRE_PEDIDO.Cod_Cliente , WF_PRE_PEDIDO.Data , WF_PRE_PEDIDO.VALOR_TOTAL_ITENS FROM AKBUSER01.WF_PRE_PEDIDO WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P91861, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        RsC578760 = QueryC578760.ExecuteReader()
        Dim iC578760 As Short
        ReDim C578760DataType(RsC578760.FieldCount)
        For iC578760 = 0 to RsC578760.FieldCount - 1
            Select Case RsC578760.GetDataTypeName(iC578760).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C578760DataType(iC578760 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C578760DataType(iC578760 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C578760DataType(iC578760 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC578760
        RsC578760_EOF = Not RsC578760.Read()

        GoTo Comp_C578761

    Comp_C578761:

        ' CodCliente
        sCurrComponent = "CodCliente"
        C578761DataType = 0
        C578761 = RsC578760(0)
        C578761DataType = C578760Datatype(1)

        GoTo Comp_C578762

    Comp_C578762:

        ' DataPedido
        sCurrComponent = "DataPedido"
        C578762DataType = 0
        C578762 = RsC578760(1)
        C578762DataType = C578760Datatype(2)

        GoTo Comp_C578770

    Comp_C578763:

        ' TipoEmpr
        sCurrComponent = "TipoEmpr"
        C578763DataType = 0
        C578763 = RsC578757(0)
        C578763DataType = C578757Datatype(1)

        GoTo Comp_C578764

    Comp_C578764:

        ' ValidInicial
        sCurrComponent = "ValidInicial"
        C578764DataType = 0
        C578764 = RsC578757(1)
        C578764DataType = C578757Datatype(2)

        GoTo Comp_C578765

    Comp_C578765:

        ' ValidFinal
        sCurrComponent = "ValidFinal"
        C578765DataType = 0
        C578765 = RsC578757(2)
        C578765DataType = C578757Datatype(3)

        GoTo Comp_C578766

    Comp_C578766:

        ' TotalFat
        sCurrComponent = "TotalFat"
        C578766DataType = 0
        C578766 = RsC578757(3)
        C578766DataType = C578757Datatype(4)

        GoTo Comp_C578773

    Comp_C578767:

        ' Ret
        sCurrComponent = "Ret"
        C578767 = 0
        C578767DataType = 1
        GoTo Comp_C578768

    Comp_C578768:

        ' FlagEstab
        sCurrComponent = "FlagEstab"
        QueryC578768 = con.CreateCommand()
        QueryC578768.CommandText = QueryC578768.CommandText & " " & "SELECT WF_ESTAB_JURIDICO_PARAM.Controle_Lim_Fatur_Tp_Empresa FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM WHERE WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P91857, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        RsC578768 = QueryC578768.ExecuteReader()
        Dim iC578768 As Short
        ReDim C578768DataType(RsC578768.FieldCount)
        For iC578768 = 0 to RsC578768.FieldCount - 1
            Select Case RsC578768.GetDataTypeName(iC578768).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C578768DataType(iC578768 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C578768DataType(iC578768 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C578768DataType(iC578768 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC578768
        RsC578768_EOF = Not RsC578768.Read()

        GoTo Comp_C578769

    Comp_C578769:

        ' FlagEstab=0?
        sCurrComponent = "FlagEstab=0?"
        C578769 = (fn_ConvertValueCompiled(RsC578768(0), C578768DataType(1), AKB_DecimalPoint, False) = 0)
        C578769DataType = AKBTypeConst.cAKBTypeLogical
        If C578769 Then
            GoTo Comp_C578758
        Else
            GoTo Comp_C578760
        End If

    Comp_C578770:

        ' ValorPrePed
        sCurrComponent = "ValorPrePed"
        C578770DataType = 0
        C578770 = RsC578760(2)
        C578770DataType = C578760Datatype(3)

        GoTo Comp_C578754

    Comp_C578771:

        ' LimiteAnualt=0?
        sCurrComponent = "LimiteAnualt=0?"
        C578771 = (fn_ConvertValueCompiled(C578773, C578773DataType, AKB_DecimalPoint, False) = 0)
        C578771DataType = AKBTypeConst.cAKBTypeLogical
        If C578771 Then
            GoTo Comp_C578758
        Else
            GoTo Comp_C578774
        End If

    Comp_C578773:

        ' LimiteAnual
        sCurrComponent = "LimiteAnual"
        C578773DataType = 0
        C578773 = RsC578757(4)
        C578773DataType = C578757Datatype(5)

        GoTo Comp_C578771

    Comp_C578774:

        ' TotalFatValorPP
        sCurrComponent = "TotalFatValorPP"
        C578774 = fn_ConvertValueCompiled(C578766, C578766DataType, AKB_DecimalPoint, True) + fn_ConvertValueCompiled(C578770, C578770DataType, AKB_DecimalPoint, True)
        C578774DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C578775

    Comp_C578775:

        ' NaoEstoura?
        sCurrComponent = "NaoEstoura?"
        C578775 = (fn_ConvertValueCompiled(C578774, C578774DataType, AKB_DecimalPoint, False) <= fn_ConvertValueCompiled(C578773, C578773DataType, AKB_DecimalPoint, False))
        C578775DataType = AKBTypeConst.cAKBTypeLogical
        If C578775 Then
            GoTo Comp_C578758
        Else
            GoTo Comp_C578776
        End If

    Comp_C578776:

        ' MSG1
        sCurrComponent = "MSG1"
        Dim tb_C578776 As DataTable = New DataTable()
        tb_C578776.Columns.Add("C578776")
        Dim row_C578776 As DataRow = tb_C578776.NewRow()
        row_C578776(0) = "Pré-pedido " & _ 
P91861 & " não pode ser gerado, limite fatur anual do Tipo de Empresa do cliente ultrapassa." & Chr(13) & "" & Chr(10) & "Liberação somente por supervisor pela tela de autorização." & Chr(13) & "" & Chr(10) & "Valor Pré-Ped: " & _ 
C578770 & " " & Chr(13) & "" & Chr(10) & "Valor Fat ano: " & _ 
C578766 & " " & Chr(13) & "" & Chr(10) & "Limite Anual: " & _ 
C578773 & ""
        tb_C578776.Rows.Add(row_C578776)
        R23567 = tb_C578776
        ReDim C578776ReturnDataType(1)
        C578776ReturnDataType(1) = AKBTypeConst.cAKBTypeString
        C578776DataType = AKBTypeConst.cAKBTypeString
        ReturnDataType = C578776ReturnDataType
        GoTo Err_R23567

        GoTo Comp_C578777

    Comp_C578777:

        ' ATUAL1
        sCurrComponent = "ATUAL1"
        QueryC578777 = con.CreateCommand()
        QueryC578777.CommandText = QueryC578777.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO SET WF_PRE_PEDIDO.Bloqueado_Limite_Tp_Empresa = 1 WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P91861, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        C578777 = QueryC578777.ExecuteNonQuery()
        C578777DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C578778

    Comp_C578778:

        ' attRet
        sCurrComponent = "attRet"
        C578778DataType = 4
        C578767 = fn_ConvertValueCompiled(1, 1, AKB_DecimalPoint)
        C578778 = True
        GoTo Comp_C578758

    Comp_C579031:

        ' FIM_TpEmpToTFat
        sCurrComponent = "FIM_TpEmpToTFat"
        C579031DataType = 4
        C579031 = RsC578757_EOF
        GoTo Comp_C579032

    Comp_C579032:

        ' FIM_TpEmpToTFat = 1
        sCurrComponent = "FIM_TpEmpToTFat = 1"
        C579032 = (fn_ConvertValueCompiled(C579031, C579031DataType, AKB_DecimalPoint, False) = 1)
        C579032DataType = AKBTypeConst.cAKBTypeLogical
        If C579032 Then
            GoTo Comp_C578758
        Else
            GoTo Comp_C578763
        End If

    Exit_R23567:

        Exit Function

    Err_R23567:
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
