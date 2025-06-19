Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R4557

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

    Public Shared Function R4557(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByVal P12596 As Double, ByVal P20438 As Object, ByVal P20439 As Object, ByVal P22403 As Object, ByVal P26981 As Object, ByVal P56293 As Object, ByVal P68164 As Object, ByVal P76804 As Double, ByVal P78985 As Object, ByVal P88682 As Object, ByVal P89539 As Object) As DataTable
    ' 
    ' 4557 - Gerar Pedido do Pré Pedido - 1 Pedido - Versão: 0
    ' 
        'On Error GoTo Err_R4557
        Dim sCurrComponent as String

        Dim QueryC54109 As New Object
        Dim C54109 As Integer
        Dim C54109DataType As Short
        Dim QueryC54110 As New Object
        Dim C54110 As Integer
        Dim C54110DataType As Short
        Dim QueryC54111 As New Object
        Dim C54111 As Integer
        Dim C54111DataType As Short
        Dim QueryC54114 As New Object
        Dim C54114 As Integer
        Dim C54114DataType As Short
        Dim C54129 As Boolean
        Dim C54129DataType As Short
        Dim C54130 As Boolean
        Dim C54130DataType As Short
        Dim C56122 As DataTable
        Dim C56122CurrentRow As Integer
        Dim C56122DataType() As Short
        Dim QueryC56123 As New Object
        Dim RsC56123 As New Object
        Dim C56123DataType() As Short
        Dim RsC56123_EOF As Boolean
        Dim C56125 As Object
        Dim C56125DataType As Short
        Dim C56126 As Object
        Dim C56126DataType As Short
        Dim C56127 As Object
        Dim C56127DataType As Short
        Dim C56128 As Boolean
        Dim C56128DataType As Short
        Dim QueryC56132 As New Object
        Dim C56132 As Integer
        Dim C56132DataType As Short
        Dim QueryC56133 As New Object
        Dim C56133 As Integer
        Dim C56133DataType As Short
        Dim QueryC56135 As New Object
        Dim C56135 As Integer
        Dim C56135DataType As Short
        Dim QueryC56139 As New Object
        Dim RsC56139 As New Object
        Dim C56139DataType() As Short
        Dim RsC56139_EOF As Boolean
        Dim C56141DataType() As Short
        Dim C56907 As DataTable
        Dim C56907CurrentRow As Integer
        Dim C56907DataType() As Short
        Dim C56908 As Object
        Dim C56908DataType As Short
        Dim C56909 As DataTable
        Dim C56909CurrentRow As Integer
        Dim C56909DataType() As Short
        Dim C56910 As DataTable
        Dim C56910CurrentRow As Integer
        Dim C56910DataType() As Short
        Dim C56912 As Boolean
        Dim C56912DataType As Short
        Dim C56913 As Boolean
        Dim C56913DataType As Short
        Dim C56914 As Boolean
        Dim C56914DataType As Short
        Dim QueryC56916 As New Object
        Dim C56916 As Integer
        Dim C56916DataType As Short
        Dim C56917 As Object
        Dim C56917DataType As Short
        Dim C56918 As Object
        Dim C56918DataType As Short
        Dim C56919 As Object
        Dim C56919DataType As Short
        Dim C56920 As Object
        Dim C56920DataType As Short
        Dim C56921 As Object
        Dim C56921DataType As Short
        Dim QueryC57793 As New Object
        Dim RsC57793 As New Object
        Dim C57793DataType() As Short
        Dim RsC57793_EOF As Boolean
        Dim QueryC57925 As New Object
        Dim RsC57925 As New Object
        Dim C57925DataType() As Short
        Dim RsC57925_EOF As Boolean
        Dim C57926 As Object
        Dim C57926DataType As Short
        Dim C57927 As Boolean
        Dim C57927DataType As Short
        Dim C57928 As Object
        Dim C57928DataType As Short
        Dim C57929 As Object
        Dim C57929DataType As Short
        Dim C57930 As Object
        Dim C57930DataType As Short
        Dim C57931 As Object
        Dim C57931DataType As Short
        Dim QueryC57932 As New Object
        Dim C57932 As Integer
        Dim C57932DataType As Short
        Dim C75152 As DataTable
        Dim C75152CurrentRow As Integer
        Dim C75152DataType() As Short
        Dim C75153 As Object
        Dim C75153DataType As Short
        Dim C75154 As Object
        Dim C75154DataType As Short
        Dim C90278 As Object
        Dim C90278DataType As Short
        Dim C90279 As Boolean
        Dim C90279DataType As Short
        Dim QueryC90281 As New Object
        Dim C90281 As Integer
        Dim C90281DataType As Short
        Dim C90293 As Object
        Dim C90293DataType As Short
        Dim C90294 As Object
        Dim C90294DataType As Short
        Dim C90297 As Object
        Dim C90297DataType As Short
        Dim C92593 As Object
        Dim C92593DataType As Short
        Dim C92594 As DataTable
        Dim C92594CurrentRow As Integer
        Dim C92594DataType() As Short
        Dim C92595 As Boolean
        Dim C92595DataType As Short
        Dim C92631 As DataTable
        Dim C92631CurrentRow As Integer
        Dim C92631DataType() As Short
        Dim C92633 As Object
        Dim C92633DataType As Short
        Dim C92634 As Boolean
        Dim C92634DataType As Short
        Dim C92637 As DataTable
        Dim C92637CurrentRow As Integer
        Dim C92637DataType() As Short
        Dim C92640 As Boolean
        Dim C92640DataType As Short
        Dim C92643 As DataTable
        Dim C92643CurrentRow As Integer
        Dim C92643DataType() As Short
        Dim C92644 As Boolean
        Dim C92644DataType As Short
        Dim C92645 As DataTable
        Dim C92645CurrentRow As Integer
        Dim C92645DataType() As Short
        Dim C92646 As Boolean
        Dim C92646DataType As Short
        Dim C92647 As DataTable
        Dim C92647CurrentRow As Integer
        Dim C92647DataType() As Short
        Dim C92649 As DataTable
        Dim C92649CurrentRow As Integer
        Dim C92649DataType() As Short
        Dim C92650 As Object
        Dim C92650DataType As Short
        Dim C92652 As DataTable
        Dim C92652CurrentRow As Integer
        Dim C92652DataType() As Short
        Dim C92653 As Boolean
        Dim C92653DataType As Short
        Dim C96452 As Short
        Dim C96452DataType As Short
        Dim C96452ReturnDataType() As Short

        Dim C96549 As Boolean
        Dim C96549DataType As Short
        Dim C96552 As Boolean
        Dim C96552DataType As Short
        Dim C96561 As Object
        Dim C96561DataType As Short
        Dim C96571 As Boolean
        Dim C96571DataType As Short
        Dim C96583 As Boolean
        Dim C96583DataType As Short
        Dim C96585 As Boolean
        Dim C96585DataType As Short
        Dim C96588 As Boolean
        Dim C96588DataType As Short
        Dim QueryC96667 As New Object
        Dim RsC96667 As New Object
        Dim C96667DataType() As Short
        Dim RsC96667_EOF As Boolean
        Dim C96668 As Object
        Dim C96668DataType As Short
        Dim C96669 As Object
        Dim C96669DataType As Short
        Dim QueryC97960 As New Object
        Dim RsC97960 As New Object
        Dim C97960DataType() As Short
        Dim RsC97960_EOF As Boolean
        Dim C97961 As Object
        Dim C97961DataType As Short
        Dim C97962 As Boolean
        Dim C97962DataType As Short
        Dim C97963 As Short
        Dim C97963DataType As Short
        Dim C97963ReturnDataType() As Short

        Dim C97964 As Object
        Dim C97964DataType As Short
        Dim C97965 As Boolean
        Dim C97965DataType As Short
        Dim C97966DataType() As Short
        Dim C97967 As Boolean
        Dim C97967DataType As Short
        Dim C97968 As Object
        Dim C97968DataType As Short
        Dim C97969 As Object
        Dim C97969DataType As Short
        Dim C97970 As Boolean
        Dim C97970DataType As Short
        Dim C97971 As Object
        Dim C97971DataType As Short
        Dim C98015 As Object
        Dim C98015DataType As Short
        Dim C100317 As DataTable
        Dim C100317CurrentRow As Integer
        Dim C100317DataType() As Short
        Dim C100318 As Boolean
        Dim C100318DataType As Short
        Dim C109297 As Boolean
        Dim C109297DataType As Short
        Dim QueryC119833 As New Object
        Dim C119833 As Integer
        Dim C119833DataType As Short
        Dim QueryC123427 As New Object
        Dim C123427 As Integer
        Dim C123427DataType As Short
        Dim C128119 As DataTable
        Dim C128119CurrentRow As Integer
        Dim C128119DataType() As Short
        Dim C130729 As Object
        Dim C130729DataType As Short
        Dim C130730 As Object
        Dim C130730DataType As Short
        Dim C130731 As Object
        Dim C130731DataType As Short
        Dim C130732 As Object
        Dim C130732DataType As Short
        Dim C130734 As Object
        Dim C130734DataType As Short
        Dim C130738 As DataTable
        Dim C130738CurrentRow As Integer
        Dim C130738DataType() As Short
        Dim C130739 As Boolean
        Dim C130739DataType As Short
        Dim C136195 As DataTable
        Dim C136195CurrentRow As Integer
        Dim C136195DataType() As Short
        Dim C136196 As Boolean
        Dim C136196DataType As Short
        Dim C137208DataType() As Short
        Dim C138014 As Object
        Dim C138014DataType As Short
        Dim C148812 As Object
        Dim C148812DataType As Short
        Dim C153120 As Boolean
        Dim C153120DataType As Short
        Dim C153121 As DataTable
        Dim C153121CurrentRow As Integer
        Dim C153121DataType() As Short
        Dim C159496 As Boolean
        Dim C159496DataType As Short
        Dim QueryC159602 As New Object
        Dim RsC159602 As New Object
        Dim C159602DataType() As Short
        Dim RsC159602_EOF As Boolean
        Dim C159603 As Boolean
        Dim C159603DataType As Short
        Dim QueryC164218 As New Object
        Dim RsC164218 As New Object
        Dim C164218DataType() As Short
        Dim RsC164218_EOF As Boolean
        Dim C164220 As Boolean
        Dim C164220DataType As Short
        Dim C164221 As Object
        Dim C164221DataType As Short
        Dim C164222 As Short
        Dim C164222DataType As Short
        Dim C164222ReturnDataType() As Short

        Dim C164223DataType() As Short
        Dim QueryC166350 As New Object
        Dim C166350 As Integer
        Dim C166350DataType As Short
        Dim QueryC166352 As New Object
        Dim C166352 As Integer
        Dim C166352DataType As Short
        Dim QueryC166353 As New Object
        Dim C166353 As Integer
        Dim C166353DataType As Short
        Dim QueryC174609 As New Object
        Dim RsC174609 As New Object
        Dim C174609DataType() As Short
        Dim RsC174609_EOF As Boolean
        Dim C174610 As Boolean
        Dim C174610DataType As Short
        Dim QueryC174612 As New Object
        Dim C174612 As Integer
        Dim C174612DataType As Short
        Dim C174615 As Object
        Dim C174615DataType As Short
        Dim C174933 As DataTable
        Dim C174933CurrentRow As Integer
        Dim C174933DataType() As Short
        Dim C175718 As Boolean
        Dim C175718DataType As Short
        Dim QueryC175721 As New Object
        Dim C175721 As Integer
        Dim C175721DataType As Short
        Dim C176563 As DataTable
        Dim C176563CurrentRow As Integer
        Dim C176563DataType() As Short
        Dim QueryC176765 As New Object
        Dim C176765 As Integer
        Dim C176765DataType As Short
        Dim C178051 As DataTable
        Dim C178051CurrentRow As Integer
        Dim C178051DataType() As Short
        Dim C178053 As Boolean
        Dim C178053DataType As Short
        Dim QueryC180507 As New Object
        Dim RsC180507 As New Object
        Dim C180507DataType() As Short
        Dim RsC180507_EOF As Boolean
        Dim C180508 As Object
        Dim C180508DataType As Short
        Dim C180509 As Object
        Dim C180509DataType As Short
        Dim C180510 As Boolean
        Dim C180510DataType As Short
        Dim C206505 As Short
        Dim C206505DataType As Short
        Dim C206505ReturnDataType() As Short

        Dim C213564 As DataTable
        Dim C213564CurrentRow As Integer
        Dim C213564DataType() As Short
        Dim C217998 As Object
        Dim C217998DataType As Short
        Dim C217999 As Object
        Dim C217999DataType As Short
        Dim QueryC218740 As New Object
        Dim RsC218740 As New Object
        Dim C218740DataType() As Short
        Dim RsC218740_EOF As Boolean
        Dim C233469 As DataTable
        Dim C233469CurrentRow As Integer
        Dim C233469DataType() As Short
        Dim C233557 As Boolean
        Dim C233557DataType As Short
        Dim C233558 As Boolean
        Dim C233558DataType As Short
        Dim C233559DataType() As Short
        Dim QueryC246209 As New Object
        Dim C246209 As Integer
        Dim C246209DataType As Short
        Dim QueryC252514 As New Object
        Dim C252514 As Integer
        Dim C252514DataType As Short
        Dim C265444 As Boolean
        Dim C265444DataType As Short
        Dim QueryC265481 As New Object
        Dim RsC265481 As New Object
        Dim C265481DataType() As Short
        Dim RsC265481_EOF As Boolean
        Dim C265482 As Object
        Dim C265482DataType As Short
        Dim C265483 As Boolean
        Dim C265483DataType As Short
        Dim QueryC265488 As New Object
        Dim C265488 As Integer
        Dim C265488DataType As Short
        Dim C266026 As DataTable
        Dim C266026CurrentRow As Integer
        Dim C266026DataType() As Short
        Dim QueryC283840 As New Object
        Dim RsC283840 As New Object
        Dim C283840DataType() As Short
        Dim RsC283840_EOF As Boolean
        Dim C283841 As Object
        Dim C283841DataType As Short
        Dim C283843 As Boolean
        Dim C283843DataType As Short
        Dim C284023 As Object
        Dim C284023DataType As Short
        Dim QueryC310322 As New Object
        Dim RsC310322 As New Object
        Dim C310322DataType() As Short
        Dim RsC310322_EOF As Boolean
        Dim C310328 As Boolean
        Dim C310328DataType As Short
        Dim C310329 As Short
        Dim C310329DataType As Short
        Dim C310329ReturnDataType() As Short

        Dim C310330DataType() As Short
        Dim QueryC315473 As New Object
        Dim RsC315473 As New Object
        Dim C315473DataType() As Short
        Dim RsC315473_EOF As Boolean
        Dim C315844 As DataTable
        Dim C315844CurrentRow As Integer
        Dim C315844DataType() As Short
        Dim QueryC315845 As New Object
        Dim RsC315845 As New Object
        Dim C315845DataType() As Short
        Dim RsC315845_EOF As Boolean
        Dim C315847 As Boolean
        Dim C315847DataType As Short
        Dim QueryC317977 As New Object
        Dim RsC317977 As New Object
        Dim C317977DataType() As Short
        Dim RsC317977_EOF As Boolean
        Dim C317978 As Boolean
        Dim C317978DataType As Short
        Dim C339223 As Boolean
        Dim C339223DataType As Short
        Dim C339248 As Boolean
        Dim C339248DataType As Short
        Dim C339250 As Boolean
        Dim C339250DataType As Short
        Dim C339251 As Boolean
        Dim C339251DataType As Short
        Dim C339252 As Boolean
        Dim C339252DataType As Short
        Dim C339255 As Boolean
        Dim C339255DataType As Short
        Dim QueryC341912 As New Object
        Dim RsC341912 As New Object
        Dim C341912DataType() As Short
        Dim RsC341912_EOF As Boolean
        Dim C341913 As Object
        Dim C341913DataType As Short
        Dim C347147 As DataTable
        Dim C347147CurrentRow As Integer
        Dim C347147DataType() As Short
        Dim C347148 As Boolean
        Dim C347148DataType As Short
        Dim QueryC350047 As New Object
        Dim C350047 As Integer
        Dim C350047DataType As Short
        Dim C350049 As Boolean
        Dim C350049DataType As Short
        Dim QueryC350050 As New Object
        Dim C350050 As Integer
        Dim C350050DataType As Short
        Dim C357368 As Object
        Dim C357368DataType As Short
        Dim C357369 As Boolean
        Dim C357369DataType As Short
        Dim QueryC357370 As New Object
        Dim C357370 As Integer
        Dim C357370DataType As Short
        Dim QueryC359024 As New Object
        Dim C359024 As Integer
        Dim C359024DataType As Short
        Dim C359545 As DataTable
        Dim C359545CurrentRow As Integer
        Dim C359545DataType() As Short
        Dim C363384 As Object
        Dim C363384DataType As Short
        Dim C363386 As Boolean
        Dim C363386DataType As Short
        Dim C363387 As Object
        Dim C363387DataType As Short
        Dim C367573 As Object
        Dim C367573DataType As Short
        Dim C367574 As Object
        Dim C367574DataType As Short
        Dim C367575 As Object
        Dim C367575DataType As Short
        Dim C367578DataType() As Short
        Dim C367579 As Boolean
        Dim C367579DataType As Short
        Dim C367580 As Short
        Dim C367580DataType As Short
        Dim C367580ReturnDataType() As Short

        Dim C367737 As Boolean
        Dim C367737DataType As Short
        Dim C367790 As DataTable
        Dim C367790CurrentRow As Integer
        Dim C367790DataType() As Short
        Dim C367999 As Object
        Dim C367999DataType As Short
        Dim C368000 As Boolean
        Dim C368000DataType As Short
        Dim C368016 As Object
        Dim C368016DataType As Short
        Dim C368020 As Boolean
        Dim C368020DataType As Short
        Dim C368021 As Object
        Dim C368021DataType As Short
        Dim C368027 As Boolean
        Dim C368027DataType As Short
        Dim QueryC368405 As New Object
        Dim RsC368405 As New Object
        Dim C368405DataType() As Short
        Dim RsC368405_EOF As Boolean
        Dim C368408 As Boolean
        Dim C368408DataType As Short
        Dim C368409 As Short
        Dim C368409DataType As Short
        Dim C368409ReturnDataType() As Short

        Dim C368438DataType() As Short
        Dim C368898 As Object
        Dim C368898DataType As Short
        Dim C368899 As Boolean
        Dim C368899DataType As Short
        Dim C371802 As Boolean
        Dim C371802DataType As Short
        Dim QueryC371804 As New Object
        Dim C371804 As Integer
        Dim C371804DataType As Short
        Dim C371816 As Object
        Dim C371816DataType As Short
        Dim C371817 As Object
        Dim C371817DataType As Short
        Dim QueryC371818 As New Object
        Dim RsC371818 As New Object
        Dim C371818DataType() As Short
        Dim RsC371818_EOF As Boolean
        Dim C371819 As Object
        Dim C371819DataType As Short
        Dim C371832 As Boolean
        Dim C371832DataType As Short
        Dim C372920 As DataTable
        Dim C372920CurrentRow As Integer
        Dim C372920DataType() As Short
        Dim C374705 As DataTable
        Dim C374705CurrentRow As Integer
        Dim C374705DataType() As Short
        Dim C374707 As Boolean
        Dim C374707DataType As Short
        Dim C374708 As DataTable
        Dim C374708CurrentRow As Integer
        Dim C374708DataType() As Short
        Dim C374709 As Boolean
        Dim C374709DataType As Short
        Dim C378418 As DataTable
        Dim C378418CurrentRow As Integer
        Dim C378418DataType() As Short
        Dim C381761 As Boolean
        Dim C381761DataType As Short
        Dim C381762 As Short
        Dim C381762DataType As Short
        Dim C381762ReturnDataType() As Short

        Dim C381763 As Object
        Dim C381763DataType As Short
        Dim C384492 As Boolean
        Dim C384492DataType As Short
        Dim C386925 As DataTable
        Dim C386925CurrentRow As Integer
        Dim C386925DataType() As Short
        Dim C387531 As DataTable
        Dim C387531CurrentRow As Integer
        Dim C387531DataType() As Short
        Dim C391471 As Boolean
        Dim C391471DataType As Short
        Dim C391632 As DataTable
        Dim C391632CurrentRow As Integer
        Dim C391632DataType() As Short
        Dim C392962 As DataTable
        Dim C392962CurrentRow As Integer
        Dim C392962DataType() As Short
        Dim C394307 As DataTable
        Dim C394307CurrentRow As Integer
        Dim C394307DataType() As Short
        Dim QueryC394339 As New Object
        Dim RsC394339 As New Object
        Dim C394339DataType() As Short
        Dim RsC394339_EOF As Boolean
        Dim C394340 As Boolean
        Dim C394340DataType As Short
        Dim C394341 As Object
        Dim C394341DataType As Short
        Dim C394342 As Object
        Dim C394342DataType As Short
        Dim C394368 As DataTable
        Dim C394368CurrentRow As Integer
        Dim C394368DataType() As Short
        Dim C394369 As Boolean
        Dim C394369DataType As Short
        Dim C394370DataType() As Short
        Dim C394371 As Object
        Dim C394371DataType As Short
        Dim C394392 As DataTable
        Dim C394392CurrentRow As Integer
        Dim C394392DataType() As Short
        Dim C405241 As Boolean
        Dim C405241DataType As Short
        Dim QueryC412007 As New Object
        Dim C412007 As Integer
        Dim C412007DataType As Short
        Dim QueryC412045 As New Object
        Dim C412045 As Integer
        Dim C412045DataType As Short
        Dim C414346 As Object
        Dim C414346DataType As Short
        Dim C414347 As Object
        Dim C414347DataType As Short
        Dim C414348 As Boolean
        Dim C414348DataType As Short
        Dim QueryC414349 As New Object
        Dim RsC414349 As New Object
        Dim C414349DataType() As Short
        Dim RsC414349_EOF As Boolean
        Dim C414350 As Object
        Dim C414350DataType As Short
        Dim C414351 As Object
        Dim C414351DataType As Short
        Dim QueryC414371 As New Object
        Dim C414371 As Integer
        Dim C414371DataType As Short
        Dim C418089 As Boolean
        Dim C418089DataType As Short
        Dim C418090 As Short
        Dim C418090DataType As Short
        Dim C418090ReturnDataType() As Short

        Dim QueryC418177 As New Object
        Dim RsC418177 As New Object
        Dim C418177DataType() As Short
        Dim RsC418177_EOF As Boolean
        Dim C418178 As Object
        Dim C418178DataType As Short
        Dim C418179DataType() As Short
        Dim QueryC418180 As New Object
        Dim C418180 As Integer
        Dim C418180DataType As Short
        Dim QueryC418181 As New Object
        Dim C418181 As Integer
        Dim C418181DataType As Short
        Dim C418182 As Object
        Dim C418182DataType As Short
        Dim C418191 As Object
        Dim C418191DataType As Short
        Dim QueryC419420 As New Object
        Dim RsC419420 As New Object
        Dim C419420DataType() As Short
        Dim RsC419420_EOF As Boolean
        Dim C419441 As Boolean
        Dim C419441DataType As Short
        Dim QueryC419442 As New Object
        Dim C419442 As Integer
        Dim C419442DataType As Short
        Dim C420246 As Boolean
        Dim C420246DataType As Short
        Dim C423441 As Boolean
        Dim C423441DataType As Short
        Dim QueryC423442 As New Object
        Dim RsC423442 As New Object
        Dim C423442DataType() As Short
        Dim RsC423442_EOF As Boolean
        Dim C423446 As Object
        Dim C423446DataType As Short
        Dim C423448 As Object
        Dim C423448DataType As Short
        Dim QueryC425563 As New Object
        Dim C425563 As Integer
        Dim C425563DataType As Short
        Dim C425566 As Object
        Dim C425566DataType As Short
        Dim C432039 As Boolean
        Dim C432039DataType As Short
        Dim C433219 As Object
        Dim C433219DataType As Short
        Dim C433221 As Object
        Dim C433221DataType As Short
        Dim C433222 As Boolean
        Dim C433222DataType As Short
        Dim QueryC433223 As New Object
        Dim RsC433223 As New Object
        Dim C433223DataType() As Short
        Dim RsC433223_EOF As Boolean
        Dim C433225 As Object
        Dim C433225DataType As Short
        Dim C438134 As Object
        Dim C438134DataType As Short
        Dim C438135 As Object
        Dim C438135DataType As Short
        Dim C439144 As Object
        Dim C439144DataType As Short
        Dim C439145 As Object
        Dim C439145DataType As Short
        Dim C439146 As Boolean
        Dim C439146DataType As Short
        Dim QueryC439147 As New Object
        Dim C439147 As Integer
        Dim C439147DataType As Short
        Dim C439225 As Object
        Dim C439225DataType As Short
        Dim C439226 As Object
        Dim C439226DataType As Short
        Dim C467451 As Boolean
        Dim C467451DataType As Short
        Dim QueryC470976 As New Object
        Dim RsC470976 As New Object
        Dim C470976DataType() As Short
        Dim RsC470976_EOF As Boolean
        Dim C470977 As Boolean
        Dim C470977DataType As Short
        Dim C470978 As Object
        Dim C470978DataType As Short
        Dim QueryC470979 As New Object
        Dim C470979 As Integer
        Dim C470979DataType As Short
        Dim C471477 As DataTable
        Dim C471477CurrentRow As Integer
        Dim C471477DataType() As Short
        Dim C471989 As Boolean
        Dim C471989DataType As Short
        Dim C474187 As Object
        Dim C474187DataType As Short
        Dim C474188 As Object
        Dim C474188DataType As Short
        Dim C474189 As Object
        Dim C474189DataType As Short
        Dim QueryC474763 As New Object
        Dim C474763 As Integer
        Dim C474763DataType As Short
        Dim C477022 As DataTable
        Dim C477022CurrentRow As Integer
        Dim C477022DataType() As Short
        Dim C477023 As Boolean
        Dim C477023DataType As Short
        Dim QueryC477036 As New Object
        Dim RsC477036 As New Object
        Dim C477036DataType() As Short
        Dim RsC477036_EOF As Boolean
        Dim C477037 As Boolean
        Dim C477037DataType As Short
        Dim C480958 As Boolean
        Dim C480958DataType As Short
        Dim C480959 As Object
        Dim C480959DataType As Short
        Dim QueryC482983 As New Object
        Dim RsC482983 As New Object
        Dim C482983DataType() As Short
        Dim RsC482983_EOF As Boolean
        Dim C482993 As Object
        Dim C482993DataType As Short
        Dim C482994 As Object
        Dim C482994DataType As Short
        Dim C482995 As Boolean
        Dim C482995DataType As Short
        Dim QueryC482996 As New Object
        Dim C482996 As Integer
        Dim C482996DataType As Short
        Dim QueryC482998 As New Object
        Dim RsC482998 As New Object
        Dim C482998DataType() As Short
        Dim RsC482998_EOF As Boolean
        Dim C483007 As Object
        Dim C483007DataType As Short
        Dim C483008 As Object
        Dim C483008DataType As Short
        Dim C483009 As Object
        Dim C483009DataType As Short
        Dim QueryC483010 As New Object
        Dim RsC483010 As New Object
        Dim C483010DataType() As Short
        Dim RsC483010_EOF As Boolean
        Dim C483011 As Object
        Dim C483011DataType As Short
        Dim C483012 As Boolean
        Dim C483012DataType As Short
        Dim C483013 As Boolean
        Dim C483013DataType As Short
        Dim QueryC483014 As New Object
        Dim C483014 As Integer
        Dim C483014DataType As Short
        Dim QueryC483015 As New Object
        Dim C483015 As Integer
        Dim C483015DataType As Short
        Dim QueryC483016 As New Object
        Dim C483016 As Integer
        Dim C483016DataType As Short
        Dim QueryC483018 As New Object
        Dim RsC483018 As New Object
        Dim C483018DataType() As Short
        Dim RsC483018_EOF As Boolean
        Dim C483019 As Object
        Dim C483019DataType As Short
        Dim C483020 As Object
        Dim C483020DataType As Short
        Dim C483021 As Boolean
        Dim C483021DataType As Short
        Dim C483022 As Short
        Dim C483022DataType As Short
        Dim C483022ReturnDataType() As Short

        Dim C483023DataType() As Short
        Dim QueryC483025 As New Object
        Dim C483025 As Integer
        Dim C483025DataType As Short
        Dim C483026 As Object
        Dim C483026DataType As Short
        Dim C483027 As Boolean
        Dim C483027DataType As Short
        Dim C486084 As Boolean
        Dim C486084DataType As Short
        Dim C486385 As Boolean
        Dim C486385DataType As Short
        Dim C486673 As Boolean
        Dim C486673DataType As Short
        Dim QueryC502254 As New Object
        Dim RsC502254 As New Object
        Dim C502254DataType() As Short
        Dim RsC502254_EOF As Boolean
        Dim QueryC502255 As New Object
        Dim RsC502255 As New Object
        Dim C502255DataType() As Short
        Dim RsC502255_EOF As Boolean
        Dim C502256 As Boolean
        Dim C502256DataType As Short
        Dim QueryC502257 As New Object
        Dim C502257 As Integer
        Dim C502257DataType As Short
        Dim QueryC508135 As New Object
        Dim RsC508135 As New Object
        Dim C508135DataType() As Short
        Dim RsC508135_EOF As Boolean
        Dim C508137 As Boolean
        Dim C508137DataType As Short
        Dim C508138 As Object
        Dim C508138DataType As Short
        Dim QueryC508139 As New Object
        Dim RsC508139 As New Object
        Dim C508139DataType() As Short
        Dim RsC508139_EOF As Boolean
        Dim C508140 As Object
        Dim C508140DataType As Short
        Dim C508141 As Object
        Dim C508141DataType As Short
        Dim C508142 As Boolean
        Dim C508142DataType As Short
        Dim C508143 As Object
        Dim C508143DataType As Short
        Dim C508144DataType() As Short
        Dim C508145 As Object
        Dim C508145DataType As Short
        Dim C508146 As Object
        Dim C508146DataType As Short
        Dim C508147 As Object
        Dim C508147DataType As Short
        Dim C508148 As Object
        Dim C508148DataType As Short
        Dim QueryC512122 As New Object
        Dim RsC512122 As New Object
        Dim C512122DataType() As Short
        Dim RsC512122_EOF As Boolean
        Dim QueryC512123 As New Object
        Dim RsC512123 As New Object
        Dim C512123DataType() As Short
        Dim RsC512123_EOF As Boolean
        Dim C512124 As Boolean
        Dim C512124DataType As Short
        Dim QueryC512125 As New Object
        Dim C512125 As Integer
        Dim C512125DataType As Short
        Dim QueryC512820 As New Object
        Dim RsC512820 As New Object
        Dim C512820DataType() As Short
        Dim RsC512820_EOF As Boolean
        Dim QueryC512821 As New Object
        Dim RsC512821 As New Object
        Dim C512821DataType() As Short
        Dim RsC512821_EOF As Boolean
        Dim C512822 As Boolean
        Dim C512822DataType As Short
        Dim QueryC512823 As New Object
        Dim C512823 As Integer
        Dim C512823DataType As Short
        Dim C516410 As Object
        Dim C516410DataType As Short
        Dim QueryC516411 As New Object
        Dim RsC516411 As New Object
        Dim C516411DataType() As Short
        Dim RsC516411_EOF As Boolean
        Dim C516412 As Object
        Dim C516412DataType As Short
        Dim C516413 As Object
        Dim C516413DataType As Short
        Dim QueryC516414 As New Object
        Dim RsC516414 As New Object
        Dim C516414DataType() As Short
        Dim RsC516414_EOF As Boolean
        Dim C516415 As Boolean
        Dim C516415DataType As Short
        Dim QueryC516416 As New Object
        Dim RsC516416 As New Object
        Dim C516416DataType() As Short
        Dim RsC516416_EOF As Boolean
        Dim C516417 As Object
        Dim C516417DataType As Short
        Dim C516418 As Short
        Dim C516418DataType As Short
        Dim C516418ReturnDataType() As Short

        Dim C516419DataType() As Short
        Dim C516420 As Boolean
        Dim C516420DataType As Short
        Dim C516421 As Boolean
        Dim C516421DataType As Short
        Dim QueryC517971 As New Object
        Dim RsC517971 As New Object
        Dim C517971DataType() As Short
        Dim RsC517971_EOF As Boolean
        Dim C517972 As Object
        Dim C517972DataType As Short
        Dim QueryC519010 As New Object
        Dim RsC519010 As New Object
        Dim C519010DataType() As Short
        Dim RsC519010_EOF As Boolean
        Dim C519011 As Object
        Dim C519011DataType As Short
        Dim C519012 As Object
        Dim C519012DataType As Short
        Dim C519013 As Boolean
        Dim C519013DataType As Short
        Dim C519014 As Object
        Dim C519014DataType As Short
        Dim QueryC524142 As New Object
        Dim RsC524142 As New Object
        Dim C524142DataType() As Short
        Dim RsC524142_EOF As Boolean
        Dim C524143 As Boolean
        Dim C524143DataType As Short
        Dim QueryC524144 As New Object
        Dim C524144 As Integer
        Dim C524144DataType As Short
        Dim C532561 As Boolean
        Dim C532561DataType As Short
        Dim QueryC539128 As New Object
        Dim RsC539128 As New Object
        Dim C539128DataType() As Short
        Dim RsC539128_EOF As Boolean
        Dim C539130 As Boolean
        Dim C539130DataType As Short
        Dim C539465 As Object
        Dim C539465DataType As Short
        Dim C539466 As Boolean
        Dim C539466DataType As Short
        Dim QueryC539467 As New Object
        Dim RsC539467 As New Object
        Dim C539467DataType() As Short
        Dim RsC539467_EOF As Boolean
        Dim C539468 As DataTable
        Dim C539468CurrentRow As Integer
        Dim C539468DataType() As Short
        Dim C539541 As Boolean
        Dim C539541DataType As Short
        Dim QueryC539543 As New Object
        Dim C539543 As Integer
        Dim C539543DataType As Short
        Dim C540756 As Object
        Dim C540756DataType As Short
        Dim C540774 As Boolean
        Dim C540774DataType As Short
        Dim QueryC540775 As New Object
        Dim RsC540775 As New Object
        Dim C540775DataType() As Short
        Dim RsC540775_EOF As Boolean
        Dim C540776 As Object
        Dim C540776DataType As Short
        Dim C540777 As Object
        Dim C540777DataType As Short
        Dim C540778 As Boolean
        Dim C540778DataType As Short
        Dim C540779 As Boolean
        Dim C540779DataType As Short
        Dim C540780 As Object
        Dim C540780DataType As Short
        Dim C540781 As DataTable
        Dim C540781CurrentRow As Integer
        Dim C540781DataType() As Short
        Dim C540782 As Boolean
        Dim C540782DataType As Short
        Dim C540786 As Boolean
        Dim C540786DataType As Short
        Dim C540909 As Boolean
        Dim C540909DataType As Short
        Dim C541305 As Boolean
        Dim C541305DataType As Short
        Dim C541306 As Boolean
        Dim C541306DataType As Short
        Dim C545148 As DataTable
        Dim C545148CurrentRow As Integer
        Dim C545148DataType() As Short
        Dim C546207 As DataTable
        Dim C546207CurrentRow As Integer
        Dim C546207DataType() As Short
        Dim C546213 As Boolean
        Dim C546213DataType As Short
        Dim C546214DataType() As Short
        Dim C549977 As DataTable
        Dim C549977CurrentRow As Integer
        Dim C549977DataType() As Short
        Dim C551587 As Boolean
        Dim C551587DataType As Short
        Dim C551588 As Short
        Dim C551588DataType As Short
        Dim C551588ReturnDataType() As Short

        Dim C551589 As Object
        Dim C551589DataType As Short
        Dim QueryC551599 As New Object
        Dim RsC551599 As New Object
        Dim C551599DataType() As Short
        Dim RsC551599_EOF As Boolean
        Dim C551600 As Boolean
        Dim C551600DataType As Short
        Dim QueryC555064 As New Object
        Dim RsC555064 As New Object
        Dim C555064DataType() As Short
        Dim RsC555064_EOF As Boolean
        Dim C555065 As Boolean
        Dim C555065DataType As Short
        Dim C555066 As DataTable
        Dim C555066CurrentRow As Integer
        Dim C555066DataType() As Short
        Dim QueryC555067 As New Object
        Dim RsC555067 As New Object
        Dim C555067DataType() As Short
        Dim RsC555067_EOF As Boolean
        Dim C555068 As Boolean
        Dim C555068DataType As Short
        Dim C555069 As Short
        Dim C555069DataType As Short
        Dim C555069ReturnDataType() As Short

        Dim QueryC555070 As New Object
        Dim RsC555070 As New Object
        Dim C555070DataType() As Short
        Dim RsC555070_EOF As Boolean
        Dim C555071 As Boolean
        Dim C555071DataType As Short
        Dim QueryC555080 As New Object
        Dim RsC555080 As New Object
        Dim C555080DataType() As Short
        Dim RsC555080_EOF As Boolean
        Dim C555081 As Boolean
        Dim C555081DataType As Short
        Dim C555428 As Object
        Dim C555428DataType As Short
        Dim QueryC556382 As New Object
        Dim RsC556382 As New Object
        Dim C556382DataType() As Short
        Dim RsC556382_EOF As Boolean
        Dim C556383 As Boolean
        Dim C556383DataType As Short
        Dim QueryC563342 As New Object
        Dim C563342 As Integer
        Dim C563342DataType As Short
        Dim C578759 As DataTable
        Dim C578759CurrentRow As Integer
        Dim C578759DataType() As Short
        Dim C578779 As Boolean
        Dim C578779DataType As Short
        Dim C579427 As Object
        Dim C579427DataType As Short
        P20438 = IIf(IsDBNull(P20438), 0, P20438)

        P20439 = IIf(IsDBNull(P20439), 0, P20439)

        P22403 = IIf(IsDBNull(P22403), 0, P22403)

        P56293 = IIf(IsDBNull(P56293), 1, P56293)

        P78985 = IIf(IsDBNull(P78985), 0, P78985)

        P88682 = IIf(IsDBNull(P88682), 0, P88682)

        P89539 = IIf(IsDBNull(P89539), 0, P89539)

        ReDim ReturnDatatype(0)

        GoTo Comp_C517971

    Comp_C54109:

        ' InPedido
        sCurrComponent = "InPedido"
        QueryC54109 = con.CreateCommand()
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & ""
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "( WF_PEDIDO.Tp_Produto , WF_PEDIDO.Pedido , WF_PEDIDO.Dt_Pedido , WF_PEDIDO.Cod_Cliente , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Cod_Repres , WF_PEDIDO.Num_Ped_Empresa , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Num_Ped_Cliente , WF_PEDIDO.Prazo_Entr_Cliente , WF_PEDIDO.Entrega_Parcial , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Obs_Comercial , WF_PEDIDO.Custo_Financ_no_Fat , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Flag_Pos_Ped , WF_PEDIDO.Porc_Custo_Financ , WF_PEDIDO.Cod_Pessoa_Estab_Juridico , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Cod_Pagto , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Tabela_Preco_Venda , WF_PEDIDO.Tabela_Preco_Custo , WF_PEDIDO.Total_Itens , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Porc_Custo_Financ_Incluso , WF_PEDIDO.Data_Base_Faturamento , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Sequencia_Atual , WF_PEDIDO.Cod_Cliente2 , WF_PEDIDO.GENERO , WF_PEDIDO.Aplicacao , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.AVISO_PEDIDO_ANTECIPADO , WF_PEDIDO.Pedido_Nota_Manual , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Cod_Cliente3 , WF_PEDIDO.Codigo_Moeda , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.EDI_PEDIDO , WF_PEDIDO.EDI_REPRESENTANTE , WF_PEDIDO.Departamento , WF_PEDIDO.EDI_DATA_ENTRADA_PEDIDO, WF_PEDIDO.Valor_Seguro , WF_PEDIDO.Valor_Frete ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Despesas_Acessorias , WF_PEDIDO.Local_Embarque , WF_PEDIDO.Local_Destino ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Gastos_Embarque , WF_PEDIDO.CODIGO_VIA,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Porc_Pg_Antec_Prod , WF_PEDIDO.Valor_Antec_Prod ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Pagto_Antecipado , WF_PEDIDO.Bloquear_Producao ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Bloquear_Faturamento , WF_PEDIDO.Porc_Pg_Antec_Fat ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Valor_Antec_Fat , WF_PEDIDO.CODIGO_DETALHE_PAGTO, "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.COD_FRETE_EXP, WF_PEDIDO.BONIFICACAO, WF_PEDIDO.Dias_Entrega,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Tipo_Compose, WF_PEDIDO.Num_Conta_Bancaria, "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Cod_Pes_Estab_Juridico_Benef,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Cod_Agencia_Beneficiaria , WF_PEDIDO.Possui_Red_ICMS,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Nao_Possui_Red_ICMS, WF_PEDIDO.COTACAO_MOEDA,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Cod_Pessoa_Estab_Juridico2, WF_PEDIDO.Codigo_Moeda2, "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.Num_Conta_Bancaria_Benef,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.UF_Embarque,WF_PEDIDO.Pedido_Site,WF_PEDIDO.Cod_Pessoa_Celula , WF_PEDIDO.Tp_Pessoa_Celula,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.LIBERADO_CREDITO,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.DATA_MANUT_CREDITO,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PEDIDO.USUARIO_MANUT_CREDITO, WF_PEDIDO.Celula , WF_PEDIDO.Cod_Transp_Entrega, WF_PEDIDO.Cod_Frete_Transp_Entreg , WF_PEDIDO.Id_Pedido_PDA , WF_PEDIDO.Cod_Ind_Pres_Comprador,WF_PEDIDO.Cod_Redesp_Entrega , WF_PEDIDO.Cod_Frete_Red_Entr,WF_PEDIDO.Cod_Ferramenta_Venda,WF_PEDIDO.Cod_Marcador_Origem ,WF_PEDIDO.Cod_Acao_Marketing , WF_PEDIDO.Qtd_Dias_Venc_Risco_Sac , WF_PEDIDO.Porc_CF_Mes_Risco_Sac , WF_PEDIDO.Estab_Juridico_Faturar, WF_PEDIDO.Pedido_Ecommerce,WF_PEDIDO.E_Commerce_Internacional , WF_PEDIDO.Tipo_Frete_Fedex,WF_PEDIDO.Portal_BrasilCacau , WF_PEDIDO.Portal_CacauShow , WF_PEDIDO.Portal_Kopenhagen , WF_PEDIDO.Portal_Lindt , WF_PEDIDO.CODIGO_SERVICO , WF_PEDIDO.Portal_Bauducco, WF_PEDIDO.E_Comm_Preco_Atacado  ) "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & ""
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "SELECT "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Tp_Produto  , WF_PRE_PEDIDO.Pedido  , SYSDATE , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Cod_Cliente  , WF_PRE_PEDIDO.Cod_Repres  ,  "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "TIRA_ACENTO(WF_PRE_PEDIDO.Num_Ped_Empresa ) , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "TIRA_ACENTO(WF_PRE_PEDIDO.Num_Ped_Cliente ) , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Prazo_Entr_Cliente  , WF_PRE_PEDIDO.Entrega_Parcial , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Obs_Comercial  , WF_PRE_PEDIDO.Custo_Financ_no_Fat  , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "" & _ 
ConvertValue(C367999, C367999DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , NVL(WF_PRE_PEDIDO.Porc_Custo_Financ,0 ) ,  WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico  , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Cod_Pagto , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Tabela_Preco_Venda  , WF_PRE_PEDIDO.Tabela_Preco_Custo , WF_PRE_PEDIDO.VALOR_TOTAL_ITENS , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Porc_Custo_Financ_Incluso  , 1 , 1 , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Cod_Cliente2 , WF_PRE_PEDIDO.GENERO  , WF_PRE_PEDIDO.Aplicacao  , 0, " & _ 
ConvertValue(C519012, C519012DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Cod_Cliente_Entrega , NVL(WF_TABELA_PRECO_VENDA.Codigo_Moeda,WF_PRE_PEDIDO.Codigo_Moeda ) , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.EDI_PEDIDO , WF_PRE_PEDIDO.EDI_REPRESENTANTE , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "NVL(WF_PRE_PEDIDO.Departamento," & _ 
ConvertValue(P68164, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") , WF_PRE_PEDIDO.Dt_Pedido ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Valor_Seguro , WF_PRE_PEDIDO.Valor_Frete ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Despesas_Acessorias , WF_PRE_PEDIDO.Local_Embarque ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Local_Destino , WF_PRE_PEDIDO.Gastos_Embarque ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.CODIGO_VIA , WF_PRE_PEDIDO.Porc_Pg_Antec_Prod ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Valor_Antec_Prod , WF_PRE_PEDIDO.Pagto_Antecipado ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Bloquear_Producao , WF_PRE_PEDIDO.Bloquear_Faturamento , "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Porc_Pg_Antec_Fat , WF_PRE_PEDIDO.Valor_Antec_Fat ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.CODIGO_DETALHE_PAGTO, WF_PRE_PEDIDO.COD_FRETE_EXP, "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.BONIFICACAO , WF_PRE_PEDIDO.Dias_Entrega, "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Tipo_Compose, "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "DECODE(WF_CLIENTES.Num_Conta_Bancaria , NULL , WF_PRE_PEDIDO.Num_Conta_Bancaria , WF_CLIENTES.Num_Conta_Bancaria) ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Cod_Pes_Estab_Juridico_Benef, "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Cod_Agencia_Beneficiaria , WF_PRE_PEDIDO.Possui_Red_ICMS,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Nao_Possui_Red_ICMS, NVL(" & _ 
ConvertValue(RsC218740(0), C218740DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",0),"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico2, WF_PRE_PEDIDO.Codigo_Moeda2,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.Num_Conta_Bancaria_Benef, WF_PRE_PEDIDO.UF_Embarque,WF_PRE_PEDIDO.PEDIDO_SITE,WF_CELULAS.Cod_Pessoa_Gestor  , 'F' ,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "TRIM(WF_PRE_PEDIDO.ANALISE_CREDITO),"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.DATA_MANUT_CRED,"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WF_PRE_PEDIDO.USUARIO_MANUT_CRED , WF_CELULAS.Celula ,WF_PRE_PEDIDO.Cod_Transp_Entrega, WF_PRE_PEDIDO.Cod_Frete_Trans_Entreg , WF_PRE_PEDIDO.Id_Pedido_PDA , " & _ 
ConvertValue(C433221, C433221DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , WF_PRE_PEDIDO.Cod_Redesp_Entrega , WF_PRE_PEDIDO.Cod_Frete_Redesp_Entrega , WF_PRE_PEDIDO.Cod_Ferramenta_Venda , WF_PRE_PEDIDO.Cod_Marcador_Origem , WF_PRE_PEDIDO.Cod_Acao_Marketing , " & _ 
ConvertValue(C508145, C508145DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C508146, C508146DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C517972, C517972DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ",  WF_PRE_PEDIDO.pedido_ecommerce, nvl(wf_pre_pedido.E_COMMERCE_INTERNACIONAL,0), nvl(WF_PRE_PEDIDO.TIPO_FRETE_FEDEX,0), NVL(WF_PRE_PEDIDO.PORTAL_BRASILCACAU,0), NVL(WF_PRE_PEDIDO.PORTAL_CACAUSHOW,0), NVL(WF_PRE_PEDIDO.PORTAL_KOPENHAGEN,0), NVL(WF_PRE_PEDIDO.PORTAL_LINDT,0), WF_PRE_PEDIDO.CODIGO_SERVICO, NVL(WF_PRE_PEDIDO.Portal_Bauducco,0),  NVL(WF_PRE_PEDIDO.E_COMM_PRECO_ATACADO,0)"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & ""
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO ,  AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_TAB_PRECO_CUSTO , AKBUSER01.WF_REPRESENTANTE, AKBUSER01.WF_CELULAS , AKBUSER01.WF_CLIENTES "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & ""
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "AND  WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda (+) = WF_PRE_PEDIDO.Tabela_Preco_Venda"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "AND WF_TAB_PRECO_CUSTO.Tabela_Preco_Custo (+) = WF_PRE_PEDIDO.Tabela_Preco_Custo "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "AND WF_REPRESENTANTE.Cod_Repres = WF_PRE_PEDIDO.Cod_Repres "
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "AND WF_REPRESENTANTE.Celula = WF_CELULAS.Celula (+)"
        QueryC54109.CommandText = QueryC54109.CommandText & " " & "AND WF_CLIENTES.Cod_Cliente = WF_PRE_PEDIDO.Cod_Cliente "
        QueryC54109.Transaction = txn
        C54109 = QueryC54109.ExecuteNonQuery()
        C54109DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C502256

    Comp_C54110:

        ' InPedItens
        sCurrComponent = "InPedItens"
        QueryC54110 = con.CreateCommand()
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_ITENS "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "( WF_PEDIDO_ITENS.Tp_Produto , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Pedido , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Seq_Item , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Sigla_Prod , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Acesso , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Cod_Embalagem , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Prazo_Entrega_Comercial,  "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Prazo_Entr_Cliente , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Porc_Desconto ,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Pedido , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Atendida , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Tipo_Ped , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Flag_Pos_Ped , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Preco_Unit , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Preco_Digitado , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Qualid_Item_Estoque , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Ultimo_Nivel_Visao , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.GENERO , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Aplicacao , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Min_Item_Unid1 , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.CFO, WF_PEDIDO_ITENS.Numero_compose, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Produto_Exclusivo,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Qtde_Dias_Evento , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Id_Colecao ,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Sigla_Prod2 , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Acesso2,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Cod_Embalagem2 ,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.PRECO_SEM_DESCONTO, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.PORC_ADICIONAL_QTDE_PEDIDO,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.OC_Cliente, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Seq_OC_Cliente,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PEDIDO_ITENS.Prazo_Entrega)"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & ""
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "SELECT WF_PRE_PEDIDO.Tp_Produto  , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO.Pedido,  "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "AkbUser01.Seq_WF_Pedido_Itens.Nextval , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod  , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Prazo_Entrega, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Prazo_Entr_Cliente , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Porc_Desconto  , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Qtde_Pedido  , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "0 , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Tipo_Ped  , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "" & _ 
ConvertValue(C367999, C367999DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "NVL(WF_PRE_PEDIDO_ITENS.Preco_Unit, 0), "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Preco_Digitado,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Qualid_Item_Estoque , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "'', "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.GENERO, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Aplicacao, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Qtde_Min_Item_Unid1, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.CFO, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "NVL(WF_PRE_PEDIDO_ITENS.Numero_Compose, 0), "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Produto_Exclusivo,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "NVL(WF_PRE_PEDIDO_ITENS.Qtde_Dias_Evento, 0) , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Id_Colecao , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod2 , "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso2,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem2 ,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.PRECO_SEM_DESCONTO, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "NVL(WF_PRE_PEDIDO_ITENS.PORC_ADICIONAL_QTDE_PEDIDO,0),"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.OC_Cliente, "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Seq_OC_Cliente ,"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Prazo_Entrega"
        QueryC54110.CommandText = QueryC54110.CommandText & " " & ""
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO , AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "AND  WF_PRE_PEDIDO.Id_PrePedido = WF_PRE_PEDIDO_ITENS.Id_PrePedido "
        QueryC54110.CommandText = QueryC54110.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.INATIVO_PDA = 0 "
        QueryC54110.Transaction = txn
        C54110 = QueryC54110.ExecuteNonQuery()
        C54110DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C54114

    Comp_C54111:

        ' InPedSequencia
        sCurrComponent = "InPedSequencia"
        QueryC54111 = con.CreateCommand()
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_SEQ "
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "( WF_PEDIDO_SEQ.Tp_Produto , WF_PEDIDO_SEQ.Pedido , WF_PEDIDO_SEQ.Seq , "
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "WF_PEDIDO_SEQ.Prazo_Entrega_Fabrica , WF_PEDIDO_SEQ.Porc_Desc_Ped , "
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "WF_PEDIDO_SEQ.Porc_Comissao ,  WF_PEDIDO_SEQ.OBS ) "
        QueryC54111.CommandText = QueryC54111.CommandText & " " & ""
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "SELECT WF_PRE_PEDIDO.Tp_Produto, WF_PRE_PEDIDO.Pedido, 1, MAX(WF_PRE_PEDIDO_ITENS.Prazo_Entrega), 0, "
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "NVL (WF_PRE_PEDIDO.Comissao, 0) , WF_PRE_PEDIDO.Obs_Comercial "
        QueryC54111.CommandText = QueryC54111.CommandText & " " & ""
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO, AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC54111.CommandText = QueryC54111.CommandText & " " & ""
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = WF_PRE_PEDIDO_ITENS.Id_PrePedido"
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC54111.CommandText = QueryC54111.CommandText & " " & ""
        QueryC54111.CommandText = QueryC54111.CommandText & " " & "GROUP BY  WF_PRE_PEDIDO.Tp_Produto, WF_PRE_PEDIDO.Pedido, NVL (WF_PRE_PEDIDO.Comissao, 0) , WF_PRE_PEDIDO.Obs_Comercial"
        QueryC54111.Transaction = txn
        C54111 = QueryC54111.ExecuteNonQuery()
        C54111DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C233469

    Comp_C54114:

        ' InItensSeq
        sCurrComponent = "InItensSeq"
        QueryC54114 = con.CreateCommand()
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_SEQ_ITENS "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "( WF_PED_SEQ_ITENS.Item_do_Pedido , WF_PED_SEQ_ITENS.Tp_Produto , WF_PED_SEQ_ITENS.Pedido , "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "WF_PED_SEQ_ITENS.Seq , WF_PED_SEQ_ITENS.Tp_Produto2 , WF_PED_SEQ_ITENS.Pedido2 , "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "WF_PED_SEQ_ITENS.Seq_Item , WF_PED_SEQ_ITENS.Qtde_Reservada , "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "WF_PED_SEQ_ITENS.Quantidade_Pedida_Original ) "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "SELECT Akbuser01.seq_wf_ped_seq_itens.nextval, WF_PEDIDO_ITENS.Tp_Produto, WF_PEDIDO_ITENS.Pedido , 1, "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "                  WF_PEDIDO_ITENS.Tp_Produto, WF_PEDIDO_ITENS.Pedido, WF_PEDIDO_ITENS.Seq_Item , 0, "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "                  WF_PEDIDO_ITENS.Qtde_Pedido                    "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "FROM  AKBUSER01.WF_PEDIDO_ITENS "
        QueryC54114.CommandText = QueryC54114.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC54114.Transaction = txn
        C54114 = QueryC54114.ExecuteNonQuery()
        C54114DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56135

    Comp_C54129:

        ' Begin
        sCurrComponent = "Begin"
        txn = con.BeginTransaction
        C54129 = True
        C54129DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C56122

    Comp_C54130:

        ' Commit
        sCurrComponent = "Commit"
        txn.Commit()
        C54130 = True
        C54130DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C265444

    Comp_C56122:

        ' Restrições
        sCurrComponent = "Restrições"
        C56122 = clsRuleDynamicallyCompiled_R4625.R4625(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C56127), C56127, System.DBNull.Value), IIf(Not IsDBNull(C56125), C56125, System.DBNull.Value), IIf(Not IsDBNull(C92650), C92650, System.DBNull.Value))
        C56122CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C56122) Then
          iColumns = C56122.Columns.Count
        End If
        ReDim C56122DataType(iColumns)
        For iCol = 1 To iColumns
            C56122DataType(iCol) = clsRuleDynamicallyCompiled_R4625.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C56128

    Comp_C56123:

        ' DadosPedido
        sCurrComponent = "DadosPedido"
        QueryC56123 = con.CreateCommand()
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "SELECT "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Tp_Produto , "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL( WF_PRE_PEDIDO.Porc_Custo_Financ , 0) , "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL( WF_PRE_PEDIDO.PORC_DESCONTO , 0) , "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Pedido ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_PRE_PEDIDO.Porc_Custo_Financ_Incluso , 0)  , "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Tabela_Preco_Venda ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_PRE_PEDIDO.Tabela_Preco_Custo , 0) , "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Pagto ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Prazo_Entr_Cliente , "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Cliente , "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Num_Ped_Empresa,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Data,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Repres ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_PRE_PEDIDO.Porc_Pg_Antec_Prod, 0 ) ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_PRE_PEDIDO.Comissao, 0) ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Bloquear_Producao ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Bloquear_Faturamento ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_PRE_PEDIDO.Porc_Pg_Antec_Fat, 0) ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_PRE_PEDIDO.Dias_Entrega,0),"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_TABELA_PRECO_VENDA.Codigo_Moeda,0) MoedaVda,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_TAB_PRECO_CUSTO.Codigo_Moeda,0) MoedaCusto,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "TRIM(WF_PRE_PEDIDO.Departamento),"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "TRIM (DECODE(WF_PRE_PEDIDO.ANALISE_CREDITO,NULL,DECODE("
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "  (SELECT WF_PRE_PEDIDO.DATA_MANUT_CRED FROM AKBUSER01.WF_PRE_PEDIDO  "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WHERE WF_PRE_PEDIDO.DATA_MANUT_CRED > =to_date('01/08/2012','dd/mm/yyyy')"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "AND  WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "    ),NULL,'Crédito em Análise', 1), WF_PRE_PEDIDO.ANALISE_CREDITO)) StatusAnalise,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.DATA_MANUT_CRED,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.USUARIO_MANUT_CRED ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_ESTAB_JURIDICO_PARAM.LIBERA_CRED_PRE_PED ,0) ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Cliente2,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "NVL(WF_PRE_PEDIDO.Cod_Ind_Pres_Comprador,0),"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Ferramenta_Venda ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Marcador_Origem ,"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WF_PRE_PEDIDO.Cod_Acao_Marketing , WF_PRE_PEDIDO.CODIGO_SERVICO "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & ""
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO , AKBUSER01.WF_TABELA_PRECO_VENDA, AKBUSER01.WF_TAB_PRECO_CUSTO, AKBUSER01.WF_ESTAB_JURIDICO_PARAM "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & ""
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "AND        WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico  = WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico "
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "AND        WF_PRE_PEDIDO.Tabela_Preco_Venda = WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda (+)"
        QueryC56123.CommandText = QueryC56123.CommandText & " " & "AND        WF_PRE_PEDIDO.Tabela_Preco_Custo = WF_TAB_PRECO_CUSTO.Tabela_Preco_Custo (+)"
        QueryC56123.Transaction = txn
        RsC56123 = QueryC56123.ExecuteReader()
        Dim iC56123 As Short
        ReDim C56123DataType(RsC56123.FieldCount)
        For iC56123 = 0 to RsC56123.FieldCount - 1
            Select Case RsC56123.GetDataTypeName(iC56123).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56123DataType(iC56123 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56123DataType(iC56123 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56123DataType(iC56123 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56123
        RsC56123_EOF = Not RsC56123.Read()

        GoTo Comp_C56125

    Comp_C56125:

        ' %CustoFinanceiro
        sCurrComponent = "%CustoFinanceiro"
        C56125DataType = 0
        C56125DataType = C56123Datatype(2)
        If C56125DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(1)) Then
          C56125 = Strings.RTrim(RsC56123(1))
        Else 
          C56125 = RsC56123(1)
        End If 

        GoTo Comp_C56127

    Comp_C56126:

        ' TpProduto
        sCurrComponent = "TpProduto"
        C56126DataType = 0
        C56126 = RsC56123(0)
        C56126DataType = C56123Datatype(1)
        If C56126DataType = AKBTypeConst.cAKBTypeString Then
          C56126 = IIF(IsDBNull(C56126), C56126, Strings.RTrim(C56126))
        End If 

        GoTo Comp_C56908

    Comp_C56127:

        ' Desconto
        sCurrComponent = "Desconto"
        C56127DataType = 0
        C56127DataType = C56123Datatype(3)
        If C56127DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(2)) Then
          C56127 = Strings.RTrim(RsC56123(2))
        Else 
          C56127 = RsC56123(2)
        End If 

        GoTo Comp_C56126

    Comp_C56128:

        ' Restrições=1
        sCurrComponent = "Restrições=1"
        C56128 = (fn_ConvertValueCompiled(C56122.rows(C56122CurrentRow - 1)(0), C56122DataType(1), AKB_DecimalPoint, False) = 1)
        C56128DataType = AKBTypeConst.cAKBTypeLogical
        If C56128 Then
            GoTo Comp_C347147
        Else
            GoTo Comp_C96668
        End If

    Comp_C56132:

        ' InPedObs
        sCurrComponent = "InPedObs"
        QueryC56132 = con.CreateCommand()
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_OBSERV "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & ""
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "(WF_PED_OBSERV.Tp_Produto , WF_PED_OBSERV.Pedido , "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "WF_PED_OBSERV.Departamento , WF_PED_OBSERV.Observacao, WF_PED_OBSERV.SEQUENCIA, WF_PED_OBSERV.Data_Geracao, WF_PED_OBSERV.Usuario_Geracao , WF_PED_OBSERV.Destinatario  ) "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & ""
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "SELECT " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "WF_PRE_PEDIDO_OBS.Departamento , WF_PRE_PEDIDO_OBS.Obs, "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "akbUser01.seq_Wf_Ped_Observ.nextval, "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "WF_PRE_PEDIDO_OBS.DATA_GERACAO , WF_PRE_PEDIDO_OBS.USUARIO_GERACAO , "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "WF_PRE_PEDIDO_OBS.Destinatario  "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & ""
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_OBS "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & ""
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "WHERE WF_PRE_PEDIDO_OBS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "AND  WF_PRE_PEDIDO_OBS.Obs  IS NOT NULL"
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "AND  WF_PRE_PEDIDO_OBS.INATIVO_PDA =0"
        QueryC56132.CommandText = QueryC56132.CommandText & " " & "AND  NVL(WF_PRE_PEDIDO_OBS.Destinatario,0) <> 'Sistema'"
        QueryC56132.CommandText = QueryC56132.CommandText & " " & ""
        QueryC56132.CommandText = QueryC56132.CommandText & " " & ""
        QueryC56132.Transaction = txn
        C56132 = QueryC56132.ExecuteNonQuery()
        C56132DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C56133

    Comp_C56133:

        ' InPedDesc
        sCurrComponent = "InPedDesc"
        QueryC56133 = con.CreateCommand()
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_SEQ_DESCONTO "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & ""
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "( WF_PED_SEQ_DESCONTO.Tp_Produto , WF_PED_SEQ_DESCONTO.Pedido , "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "WF_PED_SEQ_DESCONTO.Seq , WF_PED_SEQ_DESCONTO.Tipo_Desconto , "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "WF_PED_SEQ_DESCONTO.Porc_Desconto ) "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & ""
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "SELECT " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ," & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , 1, "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "WF_PRE_PEDIDO_DESCONTO.Tipo_Desconto , WF_PRE_PEDIDO_DESCONTO.PORC_DESCONTO "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & ""
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_DESCONTO "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & ""
        QueryC56133.CommandText = QueryC56133.CommandText & " " & "WHERE WF_PRE_PEDIDO_DESCONTO.Id_PrePedido  = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56133.CommandText = QueryC56133.CommandText & " " & ""
        QueryC56133.Transaction = txn
        C56133 = QueryC56133.ExecuteNonQuery()
        C56133DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C166350

    Comp_C56135:

        ' InItemDesc
        sCurrComponent = "InItemDesc"
        QueryC56135 = con.CreateCommand()
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "INSERT INTO AKBUSER01.WF_PED_ITENS_DESC "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & ""
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "( WF_PED_ITENS_DESC.Tp_Produto , WF_PED_ITENS_DESC.Pedido , "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "WF_PED_ITENS_DESC.Seq_Item , WF_PED_ITENS_DESC.Tipo_Desconto , "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "WF_PED_ITENS_DESC.Porc_Desconto ) "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & ""
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "SELECT /*+ CURSOR_SHARING_EXACT */  " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ," & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "   , WF_PEDIDO_ITENS.Seq_Item , "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "                  WF_PRE_PED_ITENS_DESCONTO.Tipo_Desconto , WF_PRE_PED_ITENS_DESCONTO.PORC_DESCONTO "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & ""
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "FROM AKBUSER01.WF_PRE_PED_ITENS_DESCONTO , AKBUSER01.WF_PEDIDO_ITENS, AKBUSER01.WF_PRE_PEDIDO_ITENS, AKBUSER01.WF_PRE_PEDIDO "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & ""
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "WHERE WF_PRE_PED_ITENS_DESCONTO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PED_ITENS_DESCONTO.Id_PrePedido"
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PRE_PEDIDO_ITENS.Seq_Item = WF_PRE_PED_ITENS_DESCONTO.Seq_Item "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & ""
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PRE_PEDIDO_ITENS.Acesso = WF_PEDIDO_ITENS.Acesso "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PRE_PEDIDO_ITENS.Sigla_Prod = WF_PEDIDO_ITENS.Sigla_Prod "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PRE_PEDIDO_ITENS.Cod_Embalagem = WF_PEDIDO_ITENS.Cod_Embalagem "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "AND (WF_PRE_PEDIDO_ITENS.Prazo_Entrega = WF_PEDIDO_ITENS.Prazo_Entrega"
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "OR WF_PRE_PEDIDO_ITENS.Prazo_Entrega = WF_PEDIDO_ITENS.Prazo_Entrega_Comercial)"
        QueryC56135.CommandText = QueryC56135.CommandText & " " & ""
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido"
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PRE_PED_ITENS_DESCONTO.INATIVO_PDA =0"
        QueryC56135.CommandText = QueryC56135.CommandText & " " & ""
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PEDIDO_ITENS.Pedido = WF_PRE_PEDIDO.Pedido "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PEDIDO_ITENS.Tp_Produto = WF_PRE_PEDIDO.Tp_Produto "
        QueryC56135.CommandText = QueryC56135.CommandText & " " & "And WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC56135.Transaction = txn
        C56135 = QueryC56135.ExecuteNonQuery()
        C56135DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C412007

    Comp_C56139:

        ' Zona
        sCurrComponent = "Zona"
        QueryC56139 = con.CreateCommand()
        QueryC56139.CommandText = QueryC56139.CommandText & " " & "SELECT  WF_REPRES_ZONA.Cod_Zona"
        QueryC56139.CommandText = QueryC56139.CommandText & " " & ""
        QueryC56139.CommandText = QueryC56139.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO ,  AKBUSER01.WF_REPRES_ZONA "
        QueryC56139.CommandText = QueryC56139.CommandText & " " & ""
        QueryC56139.CommandText = QueryC56139.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56139.CommandText = QueryC56139.CommandText & " " & "AND  WF_REPRES_ZONA.Cod_Repres = WF_PRE_PEDIDO.Cod_Repres"
        QueryC56139.CommandText = QueryC56139.CommandText & " " & ""
        QueryC56139.Transaction = txn
        RsC56139 = QueryC56139.ExecuteReader()
        Dim iC56139 As Short
        ReDim C56139DataType(RsC56139.FieldCount)
        For iC56139 = 0 to RsC56139.FieldCount - 1
            Select Case RsC56139.GetDataTypeName(iC56139).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C56139DataType(iC56139 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C56139DataType(iC56139 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C56139DataType(iC56139 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC56139
        RsC56139_EOF = Not RsC56139.Read()

        GoTo Comp_C56123

    Comp_C56141:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C56141 As DataTable = New DataTable()
        tb_C56141.Columns.Add(C96561  & "")
        tb_C56141.Columns.Add(C90294  & "")
        ReDim ReturnDatatype(RsC96667.FieldCount)
        Do Until RsC96667_EOF
            Dim row_C56141 As DataRow = tb_C56141.NewRow()
            Dim iC56141 As Short
            For iC56141 = 0 To 1
                row_C56141(iC56141) = RsC96667(iC56141)
                ReturnDatatype(iC56141 + 1) = C96667DataType(iC56141)
            Next iC56141
            tb_C56141.Rows.Add(row_C56141)
            RsC96667_EOF = Not RsC96667.Read()
        Loop
        R4557 = tb_C56141

        GoTo Exit_R4557

    Comp_C56907:

        ' RDesconto
        sCurrComponent = "RDesconto"
        C56907 = clsRuleDynamicallyCompiled_R817.R817(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), fn_ConvertValueCompiled( "1", 1, AKB_DecimalPoint))
        C56907CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C56907) Then
          iColumns = C56907.Columns.Count
        End If
        ReDim C56907DataType(iColumns)
        For iCol = 1 To iColumns
            C56907DataType(iCol) = clsRuleDynamicallyCompiled_R817.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C56912

    Comp_C56908:

        ' Pedido
        sCurrComponent = "Pedido"
        C56908DataType = 0
        C56908DataType = C56123Datatype(4)
        If C56908DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(3)) Then
          C56908 = Strings.RTrim(RsC56123(3))
        Else 
          C56908 = RsC56123(3)
        End If 

        GoTo Comp_C90297

    Comp_C56909:

        ' RTotalItens
        sCurrComponent = "RTotalItens"
        'C56909 = clsRuleDynamicallyCompiled_R1814.R1814(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56918), C56918, System.DBNull.Value), IIf(Not IsDBNull(C56919), C56919, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C56917), C56917, System.DBNull.Value), IIf(Not IsDBNull(C56921), C56921, System.DBNull.Value), IIf(Not IsDBNull(C56920), C56920, System.DBNull.Value), System.DBNull.Value)
        C56909CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C56909) Then
          iColumns = C56909.Columns.Count
        End If
        ReDim C56909DataType(iColumns)
        For iCol = 1 To iColumns
            'C56909DataType(iCol) = clsRuleDynamicallyCompiled_R1814.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C56913

    Comp_C56910:

        ' RSequencia
        sCurrComponent = "RSequencia"
        C56910 = clsRuleDynamicallyCompiled_R3181.R3181(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), fn_ConvertValueCompiled( "0", 4, AKB_DecimalPoint))
        C56910CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C56910) Then
          iColumns = C56910.Columns.Count
        End If
        ReDim C56910DataType(iColumns)
        For iCol = 1 To iColumns
            C56910DataType(iCol) = clsRuleDynamicallyCompiled_R3181.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C56914

    Comp_C56912:

        ' RDesconto=1
        sCurrComponent = "RDesconto=1"
        C56912 = (fn_ConvertValueCompiled(C56907.rows(C56907CurrentRow - 1)(0), C56907DataType(1), AKB_DecimalPoint, False) = 1)
        C56912DataType = AKBTypeConst.cAKBTypeLogical
        If C56912 Then
            GoTo Comp_C178051
        Else
            GoTo Comp_C96585
        End If

    Comp_C56913:

        ' RTotalItens=1
        sCurrComponent = "RTotalItens=1"
        C56913 = (fn_ConvertValueCompiled(C56909.rows(C56909CurrentRow - 1)(0), C56909DataType(1), AKB_DecimalPoint, False) = 1)
        C56913DataType = AKBTypeConst.cAKBTypeLogical
        If C56913 Then
            GoTo Comp_C56910
        Else
            GoTo Comp_C96585
        End If

    Comp_C56914:

        ' RSequencia=1
        sCurrComponent = "RSequencia=1"
        C56914 = (fn_ConvertValueCompiled(C56910.rows(C56910CurrentRow - 1)(0), C56910DataType(1), AKB_DecimalPoint, False) = 1)
        C56914DataType = AKBTypeConst.cAKBTypeLogical
        If C56914 Then
            GoTo Comp_C92594
        Else
            GoTo Comp_C96585
        End If

    Comp_C56916:

        ' Rejeitado
        sCurrComponent = "Rejeitado"
        QueryC56916 = con.CreateCommand()
        QueryC56916.CommandText = QueryC56916.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO SET WF_PRE_PEDIDO.Rejeitado = 1 WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC56916.Transaction = txn
        C56916 = QueryC56916.ExecuteNonQuery()
        C56916DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C96667

    Comp_C56917:

        ' Custo_nos_Itens
        sCurrComponent = "Custo_nos_Itens"
        C56917DataType = 0
        C56917DataType = C56123Datatype(5)
        If C56917DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(4)) Then
          C56917 = Strings.RTrim(RsC56123(4))
        Else 
          C56917 = RsC56123(4)
        End If 

        GoTo Comp_C56918

    Comp_C56918:

        ' TabVenda
        sCurrComponent = "TabVenda"
        C56918DataType = 0
        C56918DataType = C56123Datatype(6)
        If C56918DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(5)) Then
          C56918 = Strings.RTrim(RsC56123(5))
        Else 
          C56918 = RsC56123(5)
        End If 

        GoTo Comp_C56919

    Comp_C56919:

        ' TabServiço
        sCurrComponent = "TabServiço"
        C56919DataType = 0
        C56919DataType = C56123Datatype(7)
        If C56919DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(6)) Then
          C56919 = Strings.RTrim(RsC56123(6))
        Else 
          C56919 = RsC56123(6)
        End If 

        GoTo Comp_C56920

    Comp_C56920:

        ' CndPagto
        sCurrComponent = "CndPagto"
        C56920DataType = 0
        C56920DataType = C56123Datatype(8)
        If C56920DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(7)) Then
          C56920 = Strings.RTrim(RsC56123(7))
        Else 
          C56920 = RsC56123(7)
        End If 

        GoTo Comp_C75153

    Comp_C56921:

        ' PzCliente
        sCurrComponent = "PzCliente"
        C56921DataType = 0
        C56921 = RsC57793(0)
        C56921DataType = C57793Datatype(1)
        If C56921DataType = AKBTypeConst.cAKBTypeString Then
          C56921 = IIF(IsDBNull(C56921), C56921, Strings.RTrim(C56921))
        End If 

        GoTo Comp_C414346

    Comp_C57793:

        ' SelData
        sCurrComponent = "SelData"
        QueryC57793 = con.CreateCommand()
        QueryC57793.CommandText = QueryC57793.CommandText & " " & "SELECT "
        QueryC57793.CommandText = QueryC57793.CommandText & " " & "to_date(nvl( min(wf_pre_pedido_itens.prazo_entrega), wf_pre_pedido.dt_pedido))"
        QueryC57793.CommandText = QueryC57793.CommandText & " " & ""
        QueryC57793.CommandText = QueryC57793.CommandText & " " & "FROM wf_pre_pedido, WF_PRE_PEDIDO_itens"
        QueryC57793.CommandText = QueryC57793.CommandText & " " & ""
        QueryC57793.CommandText = QueryC57793.CommandText & " " & "WHERE wf_pre_pedido.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC57793.CommandText = QueryC57793.CommandText & " " & "and wf_pre_pedido.Id_PrePedido = wf_pre_pedido_itens.Id_PrePedido"
        QueryC57793.CommandText = QueryC57793.CommandText & " " & ""
        QueryC57793.CommandText = QueryC57793.CommandText & " " & "group by wf_pre_pedido.dt_pedido"
        QueryC57793.CommandText = QueryC57793.CommandText & " " & ""
        QueryC57793.CommandText = QueryC57793.CommandText & " " & ""
        QueryC57793.Transaction = txn
        RsC57793 = QueryC57793.ExecuteReader()
        Dim iC57793 As Short
        ReDim C57793DataType(RsC57793.FieldCount)
        For iC57793 = 0 to RsC57793.FieldCount - 1
            Select Case RsC57793.GetDataTypeName(iC57793).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C57793DataType(iC57793 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C57793DataType(iC57793 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C57793DataType(iC57793 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC57793
        RsC57793_EOF = Not RsC57793.Read()

        GoTo Comp_C56921

    Comp_C57925:

        ' SelDadosClientes
        sCurrComponent = "SelDadosClientes"
        QueryC57925 = con.CreateCommand()
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "SELECT  "
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "WF_PRE_PEDIDO.Cod_Transp ,"
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "WF_PRE_PEDIDO.Cod_Redespacho ,"
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "WF_PRE_PEDIDO.Cod_Frete ,"
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "WF_PRE_PEDIDO.Cod_Frete_Redesp ,"
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "WF_CLIENTES.Cod_Zona , WF_PRE_PEDIDO.Departamento"
        QueryC57925.CommandText = QueryC57925.CommandText & " " & ""
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "FROM  AKBUSER01.WF_CLIENTES , AKBUSER01.WF_PRE_PEDIDO "
        QueryC57925.CommandText = QueryC57925.CommandText & " " & ""
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC57925.CommandText = QueryC57925.CommandText & " " & "AND  WF_CLIENTES.Cod_Cliente = WF_PRE_PEDIDO.Cod_Cliente"
        QueryC57925.CommandText = QueryC57925.CommandText & " " & ""
        QueryC57925.Transaction = txn
        RsC57925 = QueryC57925.ExecuteReader()
        Dim iC57925 As Short
        ReDim C57925DataType(RsC57925.FieldCount)
        For iC57925 = 0 to RsC57925.FieldCount - 1
            Select Case RsC57925.GetDataTypeName(iC57925).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C57925DataType(iC57925 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C57925DataType(iC57925 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C57925DataType(iC57925 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC57925
        RsC57925_EOF = Not RsC57925.Read()

        GoTo Comp_C57926

    Comp_C57926:

        ' EofDadosClientes
        sCurrComponent = "EofDadosClientes"
        C57926DataType = 4
        C57926 = RsC57925_EOF
        GoTo Comp_C57927

    Comp_C57927:

        ' EofDadosClientes=1
        sCurrComponent = "EofDadosClientes=1"
        C57927 = (fn_ConvertValueCompiled(C57926, C57926DataType, AKB_DecimalPoint, False) = 1)
        C57927DataType = AKBTypeConst.cAKBTypeLogical
        If C57927 Then
            GoTo Comp_C56907
        Else
            GoTo Comp_C57928
        End If

    Comp_C57928:

        ' Transportadora
        sCurrComponent = "Transportadora"
        C57928DataType = 0
        C57928 = RsC57925(0)
        C57928DataType = C57925Datatype(1)
        If C57928DataType = AKBTypeConst.cAKBTypeString Then
          C57928 = IIF(IsDBNull(C57928), C57928, Strings.RTrim(C57928))
        End If 

        GoTo Comp_C57931

    Comp_C57929:

        ' FreteRedespacho
        sCurrComponent = "FreteRedespacho"
        C57929DataType = 0
        C57929DataType = C57925Datatype(4)
        If C57929DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC57925(3)) Then
          C57929 = Strings.RTrim(RsC57925(3))
        Else 
          C57929 = RsC57925(3)
        End If 

        GoTo Comp_C418182

    Comp_C57930:

        ' Frete
        sCurrComponent = "Frete"
        C57930DataType = 0
        C57930DataType = C57925Datatype(3)
        If C57930DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC57925(2)) Then
          C57930 = Strings.RTrim(RsC57925(2))
        Else 
          C57930 = RsC57925(2)
        End If 

        GoTo Comp_C57929

    Comp_C57931:

        ' Redespacho
        sCurrComponent = "Redespacho"
        C57931DataType = 0
        C57931DataType = C57925Datatype(2)
        If C57931DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC57925(1)) Then
          C57931 = Strings.RTrim(RsC57925(1))
        Else 
          C57931 = RsC57925(1)
        End If 

        GoTo Comp_C57930

    Comp_C57932:

        ' UpPedido
        sCurrComponent = "UpPedido"
        QueryC57932 = con.CreateCommand()
        QueryC57932.CommandText = QueryC57932.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO SET WF_PEDIDO.Cod_Redespacho = " & _ 
ConvertValue(C57931, C57931DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC57932.CommandText = QueryC57932.CommandText & " " & "                                                 WF_PEDIDO.Cod_Frete_Redesp = " & _ 
ConvertValue(C57929, C57929DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC57932.CommandText = QueryC57932.CommandText & " " & "                                                 WF_PEDIDO.Cod_Transp = " & _ 
ConvertValue(C57928, C57928DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " , "
        QueryC57932.CommandText = QueryC57932.CommandText & " " & "                                                 WF_PEDIDO.Cod_Frete = " & _ 
ConvertValue(C57930, C57930DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC57932.CommandText = QueryC57932.CommandText & " " & "WHERE WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC57932.CommandText = QueryC57932.CommandText & " " & "AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC57932.Transaction = txn
        C57932 = QueryC57932.ExecuteNonQuery()
        C57932DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C477036

    Comp_C75152:

        ' Crédito
        sCurrComponent = "Crédito"
        C75152 = clsRuleDynamicallyCompiled_R1480.R1480(con, txn, msg, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C75153), C75153, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value), IIf(Not IsDBNull(P20438), P20438, System.DBNull.Value), System.DBNull.Value)
        C75152CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C75152) Then
          iColumns = C75152.Columns.Count
        End If
        ReDim C75152DataType(iColumns)
        For iCol = 1 To iColumns
            C75152DataType(iCol) = clsRuleDynamicallyCompiled_R1480.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C109297

    Comp_C75153:

        ' Cliente
        sCurrComponent = "Cliente"
        C75153DataType = 0
        C75153DataType = C56123Datatype(10)
        If C75153DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(9)) Then
          C75153 = Strings.RTrim(RsC56123(9))
        Else 
          C75153 = RsC56123(9)
        End If 

        GoTo Comp_C75154

    Comp_C75154:

        ' Estabelecimento
        sCurrComponent = "Estabelecimento"
        C75154DataType = 0
        C75154DataType = C56123Datatype(11)
        If C75154DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(10)) Then
          C75154 = Strings.RTrim(RsC56123(10))
        Else 
          C75154 = RsC56123(10)
        End If 

        GoTo Comp_C92593

    Comp_C90278:

        ' IdPedido
        sCurrComponent = "IdPedido"
        C90278DataType = 1
        C90278 = ObjTable_NewIdentity (con, "WF_PEDIDO")
        GoTo Comp_C90293

    Comp_C90279:

        ' Pedido=0
        sCurrComponent = "Pedido=0"
        C90279 = (( fn_ConvertValueCompiled(C56908, C56908DataType, AKB_DecimalPoint, False) <=0 ) or ( fn_ConvertValueCompiled(C56908, C56908DataType, AKB_DecimalPoint, False) Is System.DBNull.Value ))
        C90279DataType = AKBTypeConst.cAKBTypeLogical
        If C90279 Then
            GoTo Comp_C90278
        Else
            GoTo Comp_C368000
        End If

    Comp_C90281:

        ' Up_PrePed
        sCurrComponent = "Up_PrePed"
        QueryC90281 = con.CreateCommand()
        QueryC90281.CommandText = QueryC90281.CommandText & " " & "Update wf_Pre_Pedido set wf_Pre_Pedido.Pedido = " & _ 
ConvertValue(C90278, C90278DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC90281.CommandText = QueryC90281.CommandText & " " & "where  wf_Pre_Pedido.ID_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC90281.Transaction = txn
        C90281 = QueryC90281.ExecuteNonQuery()
        C90281DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C368000

    Comp_C90293:

        ' AtribPedido
        sCurrComponent = "AtribPedido"
        C90293DataType = 4
        C90294 = fn_ConvertValueCompiled(C90278 , 1, AKB_DecimalPoint)
        C90293 = True
        GoTo Comp_C90281

    Comp_C90294:

        ' vPedido
        sCurrComponent = "vPedido"
        C90294 = System.DBNull.Value
        C90294DataType = 1
        GoTo Comp_C96561

    Comp_C90297:

        ' AtribPed
        sCurrComponent = "AtribPed"
        C90297DataType = 4
        C90294 = fn_ConvertValueCompiled(C56908 , 1, AKB_DecimalPoint)
        C90297 = True
        GoTo Comp_C56917

    Comp_C92593:

        ' Ped_Emp
        sCurrComponent = "Ped_Emp"
        C92593DataType = 0
        C92593DataType = C56123Datatype(12)
        If C92593DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(11)) Then
          C92593 = Strings.RTrim(RsC56123(11))
        Else 
          C92593 = RsC56123(11)
        End If 

        GoTo Comp_C357368

    Comp_C92594:

        ' RPed_Emp
        sCurrComponent = "RPed_Emp"
        C92594 = clsRuleDynamicallyCompiled_R1918.R1918(con, txn, msg, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C92593), C92593, System.DBNull.Value))
        C92594CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92594) Then
          iColumns = C92594.Columns.Count
        End If
        ReDim C92594DataType(iColumns)
        For iCol = 1 To iColumns
            C92594DataType(iCol) = clsRuleDynamicallyCompiled_R1918.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C92595

    Comp_C92595:

        ' RPed_Emp=1
        sCurrComponent = "RPed_Emp=1"
        C92595 = (fn_ConvertValueCompiled(C92594.rows(C92594CurrentRow - 1)(0), C92594DataType(1), AKB_DecimalPoint, False) = 1)
        C92595DataType = AKBTypeConst.cAKBTypeLogical
        If C92595 Then
            GoTo Comp_C92631
        Else
            GoTo Comp_C96585
        End If

    Comp_C92631:

        ' RTab_Preco
        sCurrComponent = "RTab_Preco"
        C92631 = clsRuleDynamicallyCompiled_R1533.R1533(con, txn, IIf(Not IsDBNull(C56918), C56918, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C56919), C56919, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C92633), C92633, System.DBNull.Value))
        C92631CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92631) Then
          iColumns = C92631.Columns.Count
        End If
        ReDim C92631DataType(iColumns)
        For iCol = 1 To iColumns
            C92631DataType(iCol) = clsRuleDynamicallyCompiled_R1533.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C92634

    Comp_C92633:

        ' Dt_Pedido
        sCurrComponent = "Dt_Pedido"
        C92633DataType = 0
        C92633DataType = C56123Datatype(13)
        If C92633DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(12)) Then
          C92633 = Strings.RTrim(RsC56123(12))
        Else 
          C92633 = RsC56123(12)
        End If 

        GoTo Comp_C92650

    Comp_C92634:

        ' RTab_Preco=1
        sCurrComponent = "RTab_Preco=1"
        C92634 = (fn_ConvertValueCompiled(C92631.rows(C92631CurrentRow - 1)(0), C92631DataType(1), AKB_DecimalPoint, False) = 1)
        C92634DataType = AKBTypeConst.cAKBTypeLogical
        If C92634 Then
            GoTo Comp_C92637
        Else
            GoTo Comp_C96585
        End If

    Comp_C92637:

        ' R_Itens
        sCurrComponent = "R_Itens"
        C92637 = clsRuleDynamicallyCompiled_R6446.R6446(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), System.DBNull.Value)
        C92637CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92637) Then
          iColumns = C92637.Columns.Count
        End If
        ReDim C92637DataType(iColumns)
        For iCol = 1 To iColumns
            C92637DataType(iCol) = clsRuleDynamicallyCompiled_R6446.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C92640

    Comp_C92640:

        ' RItens=1
        sCurrComponent = "RItens=1"
        C92640 = (fn_ConvertValueCompiled(C92637.rows(C92637CurrentRow - 1)(0), C92637DataType(1), AKB_DecimalPoint, False) = 1)
        C92640DataType = AKBTypeConst.cAKBTypeLogical
        If C92640 Then
            GoTo Comp_C92643
        Else
            GoTo Comp_C96585
        End If

    Comp_C92643:

        ' RProd_X_Ped
        sCurrComponent = "RProd_X_Ped"
        C92643 = clsRuleDynamicallyCompiled_R4954.R4954(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value))
        C92643CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92643) Then
          iColumns = C92643.Columns.Count
        End If
        ReDim C92643DataType(iColumns)
        For iCol = 1 To iColumns
            C92643DataType(iCol) = clsRuleDynamicallyCompiled_R4954.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C92644

    Comp_C92644:

        ' RProd_X_Ped=1
        sCurrComponent = "RProd_X_Ped=1"
        C92644 = (fn_ConvertValueCompiled(C92643.rows(C92643CurrentRow - 1)(0), C92643DataType(1), AKB_DecimalPoint, False) = 1)
        C92644DataType = AKBTypeConst.cAKBTypeLogical
        If C92644 Then
            GoTo Comp_C92645
        Else
            GoTo Comp_C96585
        End If

    Comp_C92645:

        ' RPed_Duplicados
        sCurrComponent = "RPed_Duplicados"
        C92645 = clsRuleDynamicallyCompiled_R3125.R3125(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value))
        C92645CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92645) Then
          iColumns = C92645.Columns.Count
        End If
        ReDim C92645DataType(iColumns)
        For iCol = 1 To iColumns
            C92645DataType(iCol) = clsRuleDynamicallyCompiled_R3125.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C92646

    Comp_C92646:

        ' RPed_Duplicados=1
        sCurrComponent = "RPed_Duplicados=1"
        C92646 = (fn_ConvertValueCompiled(C92645.rows(C92645CurrentRow - 1)(0), C92645DataType(1), AKB_DecimalPoint, False) = 1)
        C92646DataType = AKBTypeConst.cAKBTypeLogical
        If C92646 Then
            GoTo Comp_C92647
        Else
            GoTo Comp_C96585
        End If

    Comp_C92647:

        ' Premio
        sCurrComponent = "Premio"
        C92647 = clsRuleDynamicallyCompiled_R6033.R6033(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value), IIf(Not IsDBNull(C92633), C92633, System.DBNull.Value), IIf(Not IsDBNull(C418182), C418182, System.DBNull.Value), IIf(Not IsDBNull(C92650), C92650, System.DBNull.Value))
        C92647CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92647) Then
          iColumns = C92647.Columns.Count
        End If
        ReDim C92647DataType(iColumns)
        For iCol = 1 To iColumns
            C92647DataType(iCol) = clsRuleDynamicallyCompiled_R6033.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C92649

    Comp_C92649:

        ' Cad_Inativo
        sCurrComponent = "Cad_Inativo"
        C92649 = clsRuleDynamicallyCompiled_R4852.R4852(con, txn, IIf(Not IsDBNull(C75153), C75153, System.DBNull.Value), IIf(Not IsDBNull(C92650), C92650, System.DBNull.Value))
        C92649CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92649) Then
          iColumns = C92649.Columns.Count
        End If
        ReDim C92649DataType(iColumns)
        For iCol = 1 To iColumns
            C92649DataType(iCol) = clsRuleDynamicallyCompiled_R4852.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C159496

    Comp_C92650:

        ' Representante
        sCurrComponent = "Representante"
        C92650DataType = 0
        C92650DataType = C56123Datatype(14)
        If C92650DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(13)) Then
          C92650 = Strings.RTrim(RsC56123(13))
        Else 
          C92650 = RsC56123(13)
        End If 

        GoTo Comp_C130732

    Comp_C92652:

        ' Valida_Acesso
        sCurrComponent = "Valida_Acesso"
        C92652 = clsRuleDynamicallyCompiled_R6286.R6286(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value))
        C92652CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C92652) Then
          iColumns = C92652.Columns.Count
        End If
        ReDim C92652DataType(iColumns)
        For iCol = 1 To iColumns
            C92652DataType(iCol) = clsRuleDynamicallyCompiled_R6286.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C92653

    Comp_C92653:

        ' Valida_Acesso=1
        sCurrComponent = "Valida_Acesso=1"
        C92653 = (fn_ConvertValueCompiled(C92652.rows(C92652CurrentRow - 1)(0), C92652DataType(1), AKB_DecimalPoint, False) = 1)
        C92653DataType = AKBTypeConst.cAKBTypeLogical
        If C92653 Then
            GoTo Comp_C100317
        Else
            GoTo Comp_C96585
        End If

    Comp_C96452:

        ' MsgInforma
        sCurrComponent = "MsgInforma"
        Dim row_C96452 As DataRow = msg.NewRow()
        row_C96452(0) = "O Pedido foi gerado com sucesso!" & Chr(13) & "" & Chr(10) & "Nº Pedido: " & _ 
C90294 & ""
        msg.Rows.Add(row_C96452)
        C96452DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C96667

    Comp_C96549:

        ' NaoAbre_Trans=1
        sCurrComponent = "NaoAbre_Trans=1"
        C96549 = (( fn_ConvertValueCompiled(P20438, 4, AKB_DecimalPoint, False) = 1 ))
        C96549DataType = AKBTypeConst.cAKBTypeLogical
        If C96549 Then
            GoTo Comp_C56122
        Else
            GoTo Comp_C339223
        End If

    Comp_C96552:

        ' NaoFecha_trans=1
        sCurrComponent = "NaoFecha_trans=1"
        C96552 = (( fn_ConvertValueCompiled(P20439, 4, AKB_DecimalPoint, False) = 1 ))
        C96552DataType = AKBTypeConst.cAKBTypeLogical
        If C96552 Then
            GoTo Comp_C265444
        Else
            GoTo Comp_C339248
        End If

    Comp_C96561:

        ' vRet
        sCurrComponent = "vRet"
        C96561 = 1
        C96561DataType = 4
        GoTo Comp_C138014

    Comp_C96571:

        ' Roll
        sCurrComponent = "Roll"
        txn.Rollback()
        C96571 = True
        C96571DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C56916

    Comp_C96583:

        ' Naofecha_Trans2=1
        sCurrComponent = "Naofecha_Trans2=1"
        C96583 = (( fn_ConvertValueCompiled(P20439, 4, AKB_DecimalPoint, False) = 1 ))
        C96583DataType = AKBTypeConst.cAKBTypeLogical
        If C96583 Then
            GoTo Comp_C56916
        Else
            GoTo Comp_C339250
        End If

    Comp_C96585:

        ' NaoFecha_Trans3=1
        sCurrComponent = "NaoFecha_Trans3=1"
        C96585 = (( fn_ConvertValueCompiled(P20439, 4, AKB_DecimalPoint, False) = 1 ))
        C96585DataType = AKBTypeConst.cAKBTypeLogical
        If C96585 Then
            GoTo Comp_C96669
        Else
            GoTo Comp_C339251
        End If

    Comp_C96588:

        ' Rollback
        sCurrComponent = "Rollback"
        txn.Rollback()
        C96588 = True
        C96588DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C96669

    Comp_C96667:

        ' Sel_Ret
        sCurrComponent = "Sel_Ret"
        QueryC96667 = con.CreateCommand()
        QueryC96667.CommandText = QueryC96667.CommandText & " " & "Select " & _ 
ConvertValue(C96561, C96561DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC96667.CommandText = QueryC96667.CommandText & " " & "from dual"
        QueryC96667.Transaction = txn
        RsC96667 = QueryC96667.ExecuteReader()
        Dim iC96667 As Short
        ReDim C96667DataType(RsC96667.FieldCount)
        For iC96667 = 0 to RsC96667.FieldCount - 1
            Select Case RsC96667.GetDataTypeName(iC96667).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C96667DataType(iC96667 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C96667DataType(iC96667 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C96667DataType(iC96667 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC96667
        RsC96667_EOF = Not RsC96667.Read()

        GoTo Comp_C174933

    Comp_C96668:

        ' AttFalse
        sCurrComponent = "AttFalse"
        C96668DataType = 4
        C96561 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C96668 = True
        GoTo Comp_C96583

    Comp_C96669:

        ' AttFase2
        sCurrComponent = "AttFase2"
        C96669DataType = 4
        C96561 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C96669 = True
        GoTo Comp_C56916

    Comp_C97960:

        ' Sel_Itens
        sCurrComponent = "Sel_Itens"
        QueryC97960 = con.CreateCommand()
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "SELECT WF_PRE_PEDIDO_ITENS.Sigla_Prod  , "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                  WF_PRE_PEDIDO_ITENS.Acesso , "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                  WF_PRE_PEDIDO_ITENS.Cod_Embalagem , "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                  WF_PRE_PEDIDO_ITENS.Tipo_Ped,"
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                  WF_PRE_PEDIDO_ITENS.Prazo_Entrega, "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                  Count(*),"
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                  WF_PRE_PEDIDO_ITENS.Id_PrePedido "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO , AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & ""
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "AND  WF_PRE_PEDIDO.Id_PrePedido = WF_PRE_PEDIDO_ITENS.Id_PrePedido "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.INATIVO_PDA =0"
        QueryC97960.CommandText = QueryC97960.CommandText & " " & ""
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "Group By WF_PRE_PEDIDO.Pedido,  WF_PRE_PEDIDO_ITENS.Sigla_Prod  , WF_PRE_PEDIDO_ITENS.Acesso , "
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                    WF_PRE_PEDIDO_ITENS.Cod_Embalagem , WF_PRE_PEDIDO_ITENS.Tipo_Ped,"
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "                    WF_PRE_PEDIDO_ITENS.Prazo_Entrega,WF_PRE_PEDIDO_ITENS.Id_PrePedido"
        QueryC97960.CommandText = QueryC97960.CommandText & " " & ""
        QueryC97960.CommandText = QueryC97960.CommandText & " " & "Having Count(*) > 1"
        QueryC97960.Transaction = txn
        RsC97960 = QueryC97960.ExecuteReader()
        Dim iC97960 As Short
        ReDim C97960DataType(RsC97960.FieldCount)
        For iC97960 = 0 to RsC97960.FieldCount - 1
            Select Case RsC97960.GetDataTypeName(iC97960).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C97960DataType(iC97960 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C97960DataType(iC97960 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C97960DataType(iC97960 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC97960
        RsC97960_EOF = Not RsC97960.Read()

        GoTo Comp_C97968

    Comp_C97961:

        ' Count_Itens
        sCurrComponent = "Count_Itens"
        C97961DataType = 0
        C97961DataType = C97960Datatype(6)
        If C97961DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC97960(5)) Then
          C97961 = Strings.RTrim(RsC97960(5))
        Else 
          C97961 = RsC97960(5)
        End If 

        GoTo Comp_C97962

    Comp_C97962:

        ' Count>1
        sCurrComponent = "Count>1"
        C97962 = (fn_ConvertValueCompiled(C97961, C97961DataType, AKB_DecimalPoint, False) > 1)
        C97962DataType = AKBTypeConst.cAKBTypeLogical
        If C97962 Then
            GoTo Comp_C180508
        Else
            GoTo Comp_C97971
        End If

    Comp_C97963:

        ' MsgItens
        sCurrComponent = "MsgItens"
        Dim row_C97963 As DataRow = msg.NewRow()
        row_C97963(0) = "O item abaixo foi informado em duplicidade e o seu prazo de entrega não foi informado, ou então é o mesmo para ambos." & Chr(13) & "" & Chr(10) & "Acesso: " & _ 
C98015 & " " & Chr(13) & "" & Chr(10) & "Favor verificar. O pedido não será salvo."
        msg.Rows.Add(row_C97963)
        C97963DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C97964

    Comp_C97964:

        ' AttvRet=0
        sCurrComponent = "AttvRet=0"
        C97964DataType = 4
        C96561 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C97964 = True
        GoTo Comp_C97965

    Comp_C97965:

        ' Nao_Trans=1
        sCurrComponent = "Nao_Trans=1"
        C97965 = (( fn_ConvertValueCompiled(P20439, 4, AKB_DecimalPoint, False) = 1 ))
        C97965DataType = AKBTypeConst.cAKBTypeLogical
        If C97965 Then
            GoTo Comp_C97966
        Else
            GoTo Comp_C339252
        End If

    Comp_C97966:

        ' RET2
        sCurrComponent = "RET2"
        Dim tb_C97966 As DataTable = New DataTable()
        tb_C97966.Columns.Add("vRet" & "")
        Dim row_C97966 As DataRow = tb_C97966.NewRow()
        row_C97966(0) = C96561
        tb_C97966.Rows.Add(row_C97966)
        R4557 = tb_C97966
        ReDim C97966DataType(1)
        C97966DataType(1) = C96561DataType
        ReturnDataType = C97966DataType

        GoTo Exit_R4557

    Comp_C97967:

        ' Cancela
        sCurrComponent = "Cancela"
        txn.Rollback()
        C97967 = True
        C97967DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C97966

    Comp_C97968:

        ' Primeiro
        sCurrComponent = "Primeiro"
        C97968DataType = 4

        GoTo Comp_C97969

    Comp_C97969:

        ' EOF
        sCurrComponent = "EOF"
        C97969DataType = 4
        C97969 = RsC97960_EOF
        GoTo Comp_C97970

    Comp_C97970:

        ' EOF=1
        sCurrComponent = "EOF=1"
        C97970 = (fn_ConvertValueCompiled(C97969, C97969DataType, AKB_DecimalPoint, False) = 1)
        C97970DataType = AKBTypeConst.cAKBTypeLogical
        If C97970 Then
            GoTo Comp_C315845
        Else
            GoTo Comp_C97961
        End If

    Comp_C97971:

        ' Next
        sCurrComponent = "Next"
        C97971DataType = 4
        RsC97960_EOF = Not RsC97960.Read()
        C97971 = True

        GoTo Comp_C97969

    Comp_C98015:

        ' Acesso
        sCurrComponent = "Acesso"
        C98015DataType = 0
        C98015DataType = C97960Datatype(2)
        If C98015DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC97960(1)) Then
          C98015 = Strings.RTrim(RsC97960(1))
        Else 
          C98015 = RsC97960(1)
        End If 

        GoTo Comp_C97963

    Comp_C100317:

        ' Tipo_Desc
        sCurrComponent = "Tipo_Desc"
        C100317 = clsRuleDynamicallyCompiled_R6668.R6668(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C418182), C418182, System.DBNull.Value), IIf(Not IsDBNull(C92650), C92650, System.DBNull.Value), IIf(Not IsDBNull(C75153), C75153, System.DBNull.Value), IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value))
        C100317CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C100317) Then
          iColumns = C100317.Columns.Count
        End If
        ReDim C100317DataType(iColumns)
        For iCol = 1 To iColumns
            C100317DataType(iCol) = clsRuleDynamicallyCompiled_R6668.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C100318

    Comp_C100318:

        ' Tipo_Desc=1
        sCurrComponent = "Tipo_Desc=1"
        C100318 = (fn_ConvertValueCompiled(C100317.rows(C100317CurrentRow - 1)(0), C100317DataType(1), AKB_DecimalPoint, False) = 1)
        C100318DataType = AKBTypeConst.cAKBTypeLogical
        If C100318 Then
            GoTo Comp_C153121
        Else
            GoTo Comp_C96585
        End If

    Comp_C109297:

        ' Mensagem?=0
        sCurrComponent = "Mensagem?=0"
        C109297 = (fn_ConvertValueCompiled(P22403, 4, AKB_DecimalPoint, False) = 1)
        C109297DataType = AKBTypeConst.cAKBTypeLogical
        If C109297 Then
            GoTo Comp_C96452
        Else
            GoTo Comp_C96667
        End If

    Comp_C119833:

        ' Up_Gerou_Ped
        sCurrComponent = "Up_Gerou_Ped"
        QueryC119833 = con.CreateCommand()
        QueryC119833.CommandText = QueryC119833.CommandText & " " & "Update AKBUSER01.WF_PRE_PEDIDO set WF_PRE_PEDIDO.Gerou_Pedido = 1"
        QueryC119833.CommandText = QueryC119833.CommandText & " " & ""
        QueryC119833.CommandText = QueryC119833.CommandText & " " & "where WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC119833.CommandText = QueryC119833.CommandText & " " & ""
        QueryC119833.Transaction = txn
        C119833 = QueryC119833.ExecuteNonQuery()
        C119833DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C96552

    Comp_C123427:

        ' DataUltCpa
        sCurrComponent = "DataUltCpa"
        QueryC123427 = con.CreateCommand()
        QueryC123427.CommandText = QueryC123427.CommandText & " " & "UPDATE  WF_CLIENTES C SET C.Data_Ultima_Compra = (SELECT  SYSDATE FROM dual)"
        QueryC123427.CommandText = QueryC123427.CommandText & " " & "WHERE C.COD_CLIENTE = " & _ 
ConvertValue(C75153, C75153DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC123427.CommandText = QueryC123427.CommandText & " " & ""
        QueryC123427.CommandText = QueryC123427.CommandText & " " & ""
        QueryC123427.CommandText = QueryC123427.CommandText & " " & ""
        QueryC123427.Transaction = txn
        C123427 = QueryC123427.ExecuteNonQuery()
        C123427DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C414347

    Comp_C128119:

        ' GerarEspelho
        sCurrComponent = "GerarEspelho"
        C128119 = clsRuleDynamicallyCompiled_R1889.R1889(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value))
        C128119CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C128119) Then
          iColumns = C128119.Columns.Count
        End If
        ReDim C128119DataType(iColumns)
        For iCol = 1 To iColumns
            C128119DataType(iCol) = clsRuleDynamicallyCompiled_R1889.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C213564

    Comp_C130729:

        ' BloqFat
        sCurrComponent = "BloqFat"
        C130729DataType = 0
        C130729DataType = C56123Datatype(18)
        If C130729DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(17)) Then
          C130729 = Strings.RTrim(RsC56123(17))
        Else 
          C130729 = RsC56123(17)
        End If 

        GoTo Comp_C130734

    Comp_C130730:

        ' %Comissao
        sCurrComponent = "%Comissao"
        C130730DataType = 0
        C130730DataType = C56123Datatype(16)
        If C130730DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(15)) Then
          C130730 = Strings.RTrim(RsC56123(15))
        Else 
          C130730 = RsC56123(15)
        End If 

        GoTo Comp_C317977

    Comp_C130731:

        ' BloqProd
        sCurrComponent = "BloqProd"
        C130731DataType = 0
        C130731DataType = C56123Datatype(17)
        If C130731DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(16)) Then
          C130731 = Strings.RTrim(RsC56123(16))
        Else 
          C130731 = RsC56123(16)
        End If 

        GoTo Comp_C130729

    Comp_C130732:

        ' %PagtoProd
        sCurrComponent = "%PagtoProd"
        C130732DataType = 0
        C130732DataType = C56123Datatype(15)
        If C130732DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(14)) Then
          C130732 = Strings.RTrim(RsC56123(14))
        Else 
          C130732 = RsC56123(14)
        End If 

        GoTo Comp_C130730

    Comp_C130734:

        ' %PagtoFat
        sCurrComponent = "%PagtoFat"
        C130734DataType = 0
        C130734DataType = C56123Datatype(19)
        If C130734DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(18)) Then
          C130734 = Strings.RTrim(RsC56123(18))
        Else 
          C130734 = RsC56123(18)
        End If 

        GoTo Comp_C148812

    Comp_C130738:

        ' PedidoAntec
        sCurrComponent = "PedidoAntec"
        C130738 = clsRuleDynamicallyCompiled_R7671.R7671(con, txn, IIf(Not IsDBNull(C56920), C56920, System.DBNull.Value), IIf(Not IsDBNull(C130732), C130732, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value), IIf(Not IsDBNull(C130730), C130730, System.DBNull.Value), IIf(Not IsDBNull(C75153), C75153, System.DBNull.Value), IIf(Not IsDBNull(C92650), C92650, System.DBNull.Value), IIf(Not IsDBNull(C92633), C92633, System.DBNull.Value), IIf(Not IsDBNull(C130731), C130731, System.DBNull.Value), IIf(Not IsDBNull(C130729), C130729, System.DBNull.Value), IIf(Not IsDBNull(C130734), C130734, System.DBNull.Value), fn_ConvertValueCompiled( "0", 1, AKB_DecimalPoint), IIf(Not IsDBNull(C148812), C148812, System.DBNull.Value), fn_ConvertValueCompiled( "1", 1, AKB_DecimalPoint), IIf(Not IsDBNull(C217999), C217999, System.DBNull.Value), IIf(Not IsDBNull(C217998), C217998, System.DBNull.Value), IIf(Not IsDBNull(C56918), C56918, System.DBNull.Value), System.DBNull.Value, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value))
        C130738CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C130738) Then
          iColumns = C130738.Columns.Count
        End If
        ReDim C130738DataType(iColumns)
        For iCol = 1 To iColumns
            C130738DataType(iCol) = clsRuleDynamicallyCompiled_R7671.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C130739

    Comp_C130739:

        ' PedAntec= ok?
        sCurrComponent = "PedAntec= ok?"
        C130739 = (fn_ConvertValueCompiled(C130738.rows(C130738CurrentRow - 1)(0), C130738DataType(1), AKB_DecimalPoint, False) = 1)
        C130739DataType = AKBTypeConst.cAKBTypeLogical
        If C130739 Then
            GoTo Comp_C128119
        Else
            GoTo Comp_C96585
        End If

    Comp_C136195:

        ' Valida_Data_Venc_Proforma
        sCurrComponent = "Valida_Data_Venc_Proforma"
        C136195 = clsRuleDynamicallyCompiled_R8677.R8677(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value), IIf(Not IsDBNull(P26981), P26981, System.DBNull.Value))
        C136195CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C136195) Then
          iColumns = C136195.Columns.Count
        End If
        ReDim C136195DataType(iColumns)
        For iCol = 1 To iColumns
            C136195DataType(iCol) = clsRuleDynamicallyCompiled_R8677.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C136196

    Comp_C136196:

        ' Vencida?
        sCurrComponent = "Vencida?"
        C136196 = (fn_ConvertValueCompiled(C136195.rows(C136195CurrentRow - 1)(0), C136195DataType(1), AKB_DecimalPoint, False) = 0)
        C136196DataType = AKBTypeConst.cAKBTypeLogical
        If C136196 Then
            GoTo Comp_C137208
        Else
            GoTo Comp_C159602
        End If

    Comp_C137208:

        ' RET3
        sCurrComponent = "RET3"
        Dim tb_C137208 As DataTable = New DataTable()
        tb_C137208.Columns.Add("Vencida?" & "")
        Dim row_C137208 As DataRow = tb_C137208.NewRow()
        row_C137208(0) = C136196
        tb_C137208.Rows.Add(row_C137208)
        R4557 = tb_C137208
        ReDim C137208DataType(1)
        C137208DataType(1) = C136196DataType
        ReturnDataType = C137208DataType

        GoTo Exit_R4557

    Comp_C138014:

        ' vExclusivo
        sCurrComponent = "vExclusivo"
        C138014 = 0
        C138014DataType = 4
        GoTo Comp_C508145

    Comp_C148812:

        ' Dias_Entregar
        sCurrComponent = "Dias_Entregar"
        C148812DataType = 0
        C148812DataType = C56123Datatype(20)
        If C148812DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(19)) Then
          C148812 = Strings.RTrim(RsC56123(19))
        Else 
          C148812 = RsC56123(19)
        End If 

        GoTo Comp_C217999

    Comp_C153120:

        ' Emb_Inativa?=0
        sCurrComponent = "Emb_Inativa?=0"
        C153120 = (fn_ConvertValueCompiled(C153121.rows(C153121CurrentRow - 1)(0), C153121DataType(1), AKB_DecimalPoint, False) = 1)
        C153120DataType = AKBTypeConst.cAKBTypeLogical
        If C153120 Then
            GoTo Comp_C130738
        Else
            GoTo Comp_C96585
        End If

    Comp_C153121:

        ' Emb_Inativa?
        sCurrComponent = "Emb_Inativa?"
        C153121 = clsRuleDynamicallyCompiled_R7367.R7367(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value))
        C153121CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C153121) Then
          iColumns = C153121.Columns.Count
        End If
        ReDim C153121DataType(iColumns)
        For iCol = 1 To iColumns
            C153121DataType(iCol) = clsRuleDynamicallyCompiled_R7367.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C153120

    Comp_C159496:

        ' Cada_Inativo=1
        sCurrComponent = "Cada_Inativo=1"
        C159496 = (fn_ConvertValueCompiled(C92649.rows(C92649CurrentRow - 1)(0), C92649DataType(1), AKB_DecimalPoint, False) = 1)
        C159496DataType = AKBTypeConst.cAKBTypeLogical
        If C159496 Then
            GoTo Comp_C92652
        Else
            GoTo Comp_C96585
        End If

    Comp_C159602:

        ' Sel_CliLib
        sCurrComponent = "Sel_CliLib"
        QueryC159602 = con.CreateCommand()
        QueryC159602.CommandText = QueryC159602.CommandText & " " & "SELECT COUNT(*)"
        QueryC159602.CommandText = QueryC159602.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES "
        QueryC159602.CommandText = QueryC159602.CommandText & " " & "WHERE WF_CLIENTES.Cod_Cliente = " & _ 
ConvertValue(C75153, C75153DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC159602.CommandText = QueryC159602.CommandText & " " & "AND WF_CLIENTES.Pre_Cadastro = 1"
        QueryC159602.CommandText = QueryC159602.CommandText & " " & "AND WF_CLIENTES.Pre_Cadastro_Liberado = 0 "
        QueryC159602.CommandText = QueryC159602.CommandText & " " & ""
        QueryC159602.Transaction = txn
        RsC159602 = QueryC159602.ExecuteReader()
        Dim iC159602 As Short
        ReDim C159602DataType(RsC159602.FieldCount)
        For iC159602 = 0 to RsC159602.FieldCount - 1
            Select Case RsC159602.GetDataTypeName(iC159602).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C159602DataType(iC159602 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C159602DataType(iC159602 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C159602DataType(iC159602 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC159602
        RsC159602_EOF = Not RsC159602.Read()

        GoTo Comp_C159603

    Comp_C159603:

        ' Não está lib?
        sCurrComponent = "Não está lib?"
        C159603 = (fn_ConvertValueCompiled(RsC159602(0), C159602DataType(1), AKB_DecimalPoint, False) = 1)
        C159603DataType = AKBTypeConst.cAKBTypeLogical
        If C159603 Then
            GoTo Comp_C206505
        Else
            GoTo Comp_C96549
        End If

    Comp_C164218:

        ' Codigo Moeda
        sCurrComponent = "Codigo Moeda"
        QueryC164218 = con.CreateCommand()
        QueryC164218.CommandText = QueryC164218.CommandText & " " & "SELECT WF_PRE_PEDIDO.Codigo_Moeda "
        QueryC164218.CommandText = QueryC164218.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC164218.CommandText = QueryC164218.CommandText & " " & "WHERE"
        QueryC164218.CommandText = QueryC164218.CommandText & " " & "WF_PRE_PEDIDO.Codigo_Moeda IS NOT NULL"
        QueryC164218.CommandText = QueryC164218.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164218.Transaction = txn
        RsC164218 = QueryC164218.ExecuteReader()
        Dim iC164218 As Short
        ReDim C164218DataType(RsC164218.FieldCount)
        For iC164218 = 0 to RsC164218.FieldCount - 1
            Select Case RsC164218.GetDataTypeName(iC164218).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C164218DataType(iC164218 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C164218DataType(iC164218 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C164218DataType(iC164218 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC164218
        RsC164218_EOF = Not RsC164218.Read()

        GoTo Comp_C164221

    Comp_C164220:

        ' FV_Codigo_Moeda = 1
        sCurrComponent = "FV_Codigo_Moeda = 1"
        C164220 = (fn_ConvertValueCompiled(C164221, C164221DataType, AKB_DecimalPoint, False) = 1)
        C164220DataType = AKBTypeConst.cAKBTypeLogical
        If C164220 Then
            GoTo Comp_C164222
        Else
            GoTo Comp_C418181
        End If

    Comp_C164221:

        ' FV_Codigo_Moeda
        sCurrComponent = "FV_Codigo_Moeda"
        C164221DataType = 4
        C164221 = RsC164218_EOF
        GoTo Comp_C164220

    Comp_C164222:

        ' MSG1
        sCurrComponent = "MSG1"
        Dim row_C164222 As DataRow = msg.NewRow()
        row_C164222(0) = "Não será possivel gerar o Pedido deste Pré - Pedido, pois não tem um código de moeda para ele."
        msg.Rows.Add(row_C164222)
        C164222DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C164223

    Comp_C164223:

        ' RET4
        sCurrComponent = "RET4"
        Dim tb_C164223 As DataTable = New DataTable()
        tb_C164223.Columns.Add("MSG1" & "")
        Dim row_C164223 As DataRow = tb_C164223.NewRow()
        row_C164223(0) = C164222
        tb_C164223.Rows.Add(row_C164223)
        R4557 = tb_C164223
        ReDim C164223DataType(1)
        C164223DataType(1) = C164222DataType
        ReturnDataType = C164223DataType

        GoTo Exit_R4557

    Comp_C166350:

        ' Ins_PedXEventos
        sCurrComponent = "Ins_PedXEventos"
        QueryC166350 = con.CreateCommand()
        QueryC166350.CommandText = QueryC166350.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_EVENTOS ( WF_PEDIDO_EVENTOS.Tp_Produto ,"
        QueryC166350.CommandText = QueryC166350.CommandText & " " & "WF_PEDIDO_EVENTOS.Pedido , "
        QueryC166350.CommandText = QueryC166350.CommandText & " " & "WF_PEDIDO_EVENTOS.Codigo_Evento )"
        QueryC166350.CommandText = QueryC166350.CommandText & " " & ""
        QueryC166350.CommandText = QueryC166350.CommandText & " " & "Select WF_PRE_PEDIDO.Tp_Produto, WF_PRE_PEDIDO.Pedido, WF_PRE_PEDIDO_EVENTOS.Codigo_Evento "
        QueryC166350.CommandText = QueryC166350.CommandText & " " & "From AKBUSER01.WF_PRE_PEDIDO_EVENTOS, AKBUSER01.WF_PRE_PEDIDO"
        QueryC166350.CommandText = QueryC166350.CommandText & " " & "Where WF_PRE_PEDIDO_EVENTOS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido "
        QueryC166350.CommandText = QueryC166350.CommandText & " " & "AND    WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC166350.Transaction = txn
        C166350 = QueryC166350.ExecuteNonQuery()
        C166350DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C166352

    Comp_C166352:

        ' Ins_Emails
        sCurrComponent = "Ins_Emails"
        QueryC166352 = con.CreateCommand()
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "INSERT INTO AKBUSER01.WF_EMAIL_PEDIDO_EVENTOS "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "( WF_EMAIL_PEDIDO_EVENTOS.Tp_Produto , "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_EMAIL_PEDIDO_EVENTOS.Pedido , "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_EMAIL_PEDIDO_EVENTOS.Codigo_Evento , "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_EMAIL_PEDIDO_EVENTOS.Assunto,"
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_EMAIL_PEDIDO_EVENTOS.Descricao) "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & ""
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "Select WF_PRE_PEDIDO.Tp_Produto, "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_PRE_PEDIDO.Pedido,"
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_EMAILS_PRE_PED_EVENTOS.Codigo_Evento , "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_EMAILS_PRE_PED_EVENTOS.Assunto , "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "WF_EMAILS_PRE_PED_EVENTOS.Descricao"
        QueryC166352.CommandText = QueryC166352.CommandText & " " & ""
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "From  AKBUSER01.WF_EMAILS_PRE_PED_EVENTOS, AKBUSER01.WF_PRE_PEDIDO "
        QueryC166352.CommandText = QueryC166352.CommandText & " " & ""
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "Where WF_PRE_PEDIDO.Id_PrePedido = WF_EMAILS_PRE_PED_EVENTOS.Id_PrePedido"
        QueryC166352.CommandText = QueryC166352.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC166352.Transaction = txn
        C166352 = QueryC166352.ExecuteNonQuery()
        C166352DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C166353

    Comp_C166353:

        ' Ins_Usuarios
        sCurrComponent = "Ins_Usuarios"
        QueryC166353 = con.CreateCommand()
        QueryC166353.CommandText = QueryC166353.CommandText & " " & "INSERT INTO AKBUSER01.WF_EMAILS_USUARIOS ( WF_EMAILS_USUARIOS.Id_Email , WF_EMAILS_USUARIOS.Tp_Produto , WF_EMAILS_USUARIOS.Pedido , WF_EMAILS_USUARIOS.Codigo_Evento , WF_EMAILS_USUARIOS.Cod_Pessoa ) "
        QueryC166353.CommandText = QueryC166353.CommandText & " " & ""
        QueryC166353.CommandText = QueryC166353.CommandText & " " & "Select  Seq_WF_Emails_Usuarios.Nextval, WF_PRE_PEDIDO.Tp_Produto, WF_PRE_PEDIDO.Pedido, WF_EMAILS_PRE_PED_USUARIO.Codigo_Evento, WF_EMAILS_PRE_PED_USUARIO.Cod_Pessoa "
        QueryC166353.CommandText = QueryC166353.CommandText & " " & "From AKBUSER01.WF_EMAILS_PRE_PED_USUARIO, AKBUSER01.WF_PRE_PEDIDO "
        QueryC166353.CommandText = QueryC166353.CommandText & " " & "Where WF_EMAILS_PRE_PED_USUARIO.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido"
        QueryC166353.CommandText = QueryC166353.CommandText & " " & "And WF_EMAILS_PRE_PED_USUARIO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC166353.Transaction = txn
        C166353 = QueryC166353.ExecuteNonQuery()
        C166353DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C176765

    Comp_C174609:

        ' Sel_Colec
        sCurrComponent = "Sel_Colec"
        QueryC174609 = con.CreateCommand()
        QueryC174609.CommandText = QueryC174609.CommandText & " " & "Select WF_PRE_PEDIDO_ITENS.Seq_Item , "
        QueryC174609.CommandText = QueryC174609.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod , "
        QueryC174609.CommandText = QueryC174609.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso ,"
        QueryC174609.CommandText = QueryC174609.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC174609.CommandText = QueryC174609.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC174609.CommandText = QueryC174609.CommandText & " " & "WHERE  WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174609.CommandText = QueryC174609.CommandText & " " & "and ( WF_PRE_PEDIDO_ITENS.Id_Colecao is null or WF_PRE_PEDIDO_ITENS.Id_Colecao = 0 )"
        QueryC174609.Transaction = txn
        RsC174609 = QueryC174609.ExecuteReader()
        Dim iC174609 As Short
        ReDim C174609DataType(RsC174609.FieldCount)
        For iC174609 = 0 to RsC174609.FieldCount - 1
            Select Case RsC174609.GetDataTypeName(iC174609).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C174609DataType(iC174609 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C174609DataType(iC174609 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C174609DataType(iC174609 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC174609
        RsC174609_EOF = Not RsC174609.Read()

        GoTo Comp_C512821

    Comp_C174610:

        ' Sel_Colec=Nulo
        sCurrComponent = "Sel_Colec=Nulo"
        C174610 = (fn_ConvertValueCompiled(C579427, C579427DataType, AKB_DecimalPoint, False) = 1)
        C174610DataType = AKBTypeConst.cAKBTypeLogical
        If C174610 Then
            GoTo Comp_C92633
        Else
            GoTo Comp_C357369
        End If

    Comp_C174612:

        ' UpColecaoVda
        sCurrComponent = "UpColecaoVda"
        QueryC174612 = con.CreateCommand()
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "Update AKBUSER01.WF_PRE_PEDIDO_ITENS set "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                 WF_PRE_PEDIDO_ITENS.Id_Colecao = "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "(SELECT WF_COLECAO_PRODUTOS.Id_Colecao "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= SYSDATE"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= SYSDATE OR WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "UNION ALL"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "SELECT WF_TABELA_PRECO_VENDA.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "FROM AKBUSER01.WF_TABELA_PRECO_VENDA "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.Id_Colecao IS NOT NULL"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND NOT EXISTS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "(SELECT *"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= SYSDATE"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= SYSDATE OR WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao)"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND NOT EXISTS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "(SELECT *"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Final < SYSDATE"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao)"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND (SELECT COUNT(WF_COLECAO_PRODUTOS.Id_Colecao)"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= SYSDATE"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= SYSDATE OR WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL )) > 1 "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "),"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso2 = "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "( SELECT WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  UNION ALL"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  Select WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  from AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  where WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " or "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "        WF_COLECAO_PRODUTOS.Data_Validade_Final is null )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and not exists ( SELECT WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "				   AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                  )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & " ),"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod2 = "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "( SELECT WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  UNION ALL"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  Select WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  from AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  where WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " or "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "        WF_COLECAO_PRODUTOS.Data_Validade_Final is null )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and not exists ( SELECT WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "				   AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                  )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & " ),"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & " WF_PRE_PEDIDO_ITENS.Cod_Embalagem2  = "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "( SELECT WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  UNION ALL"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  Select WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  from AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  where WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " or "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "        WF_COLECAO_PRODUTOS.Data_Validade_Final is null )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "  and not exists ( SELECT WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "				   AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "                  )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & " )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & ""
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "where WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "and ( WF_PRE_PEDIDO_ITENS.Id_Colecao is null or WF_PRE_PEDIDO_ITENS.Id_Colecao = 0 )"
        QueryC174612.CommandText = QueryC174612.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Acesso not IN (" & _ 
ConvertValue(C283841, C283841DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC174612.Transaction = txn
        C174612 = QueryC174612.ExecuteNonQuery()
        C174612DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C92633

    Comp_C174615:

        ' DataSistema
        sCurrComponent = "DataSistema"
        C174615DataType = 2
        C174615 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C265482

    Comp_C174933:

        ' Atualiza 1 Pedido Flag Ordem de Compra
        sCurrComponent = "Atualiza 1 Pedido Flag Ordem de Compra"
        C174933 = clsRuleDynamicallyCompiled_R9997.R9997(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value))
        C174933CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C174933) Then
          iColumns = C174933.Columns.Count
        End If
        ReDim C174933DataType(iColumns)
        For iCol = 1 To iColumns
            C174933DataType(iCol) = clsRuleDynamicallyCompiled_R9997.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C471989

    Comp_C175718:

        ' TabServ=Nulo
        sCurrComponent = "TabServ=Nulo"
        C175718 = (( fn_ConvertValueCompiled(C56919, C56919DataType, AKB_DecimalPoint, False)="" ) or ( fn_ConvertValueCompiled(C56919, C56919DataType, AKB_DecimalPoint, False)=0 ) or ( fn_ConvertValueCompiled(C56919, C56919DataType, AKB_DecimalPoint, False) Is System.DBNull.Value ))
        C175718DataType = AKBTypeConst.cAKBTypeLogical
        If C175718 Then
            GoTo Comp_C174612
        Else
            GoTo Comp_C175721
        End If

    Comp_C175721:

        ' UpColecaoCto
        sCurrComponent = "UpColecaoCto"
        QueryC175721 = con.CreateCommand()
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.Id_Colecao = "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "(SELECT WF_COLECAO_PRODUTOS.Id_Colecao "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem  "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= SYSDATE"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= SYSDATE OR WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "UNION ALL"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "SELECT WF_TABELA_PRECO_VENDA.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "FROM AKBUSER01.WF_TABELA_PRECO_VENDA "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.Id_Colecao IS NOT NULL"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND NOT EXISTS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "(SELECT *"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= SYSDATE"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= SYSDATE OR WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao)"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND NOT EXISTS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "(SELECT *"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Final < SYSDATE"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao)"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND (SELECT COUNT(WF_COLECAO_PRODUTOS.Id_Colecao)"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= SYSDATE"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= SYSDATE OR WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL )) > 1 "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "),"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "            "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso2 = "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "( SELECT WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem  "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  UNION ALL"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  SELECT WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem  "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " or "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "        WF_COLECAO_PRODUTOS.Data_Validade_Final is null )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND NOT EXISTS ( SELECT WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "				   AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                  )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & " ),"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod2 = "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "( SELECT WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem  "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  UNION ALL"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  SELECT WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " OR "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "        WF_COLECAO_PRODUTOS.Data_Validade_Final IS NULL)"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND NOT EXISTS ( SELECT WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "				   AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                  )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & " ),"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem2 = "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "( SELECT WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  UNION ALL"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  SELECT WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " or "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "        WF_COLECAO_PRODUTOS.Data_Validade_Final is null )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & ""
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "  AND NOT EXISTS  ( SELECT WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   FROM AKBUSER01.WF_TABELA_PRECO_VENDA , AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                   AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "				   AND Wf_Colecao_Produtos.Cod_Embalagem = Wf_Pre_Pedido_Itens.Cod_Embalagem"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "                  )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & " )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND ( WF_PRE_PEDIDO_ITENS.Id_Colecao IS NULL OR WF_PRE_PEDIDO_ITENS.Id_Colecao = 0 )"
        QueryC175721.CommandText = QueryC175721.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Acesso NOT IN (" & _ 
ConvertValue(C283841, C283841DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC175721.Transaction = txn
        C175721 = QueryC175721.ExecuteNonQuery()
        C175721DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C92633

    Comp_C176563:

        ' MarcaVerificado
        sCurrComponent = "MarcaVerificado"
        C176563 = clsRuleDynamicallyCompiled_R10096.R10096(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value))
        C176563CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C176563) Then
          iColumns = C176563.Columns.Count
        End If
        ReDim C176563DataType(iColumns)
        For iCol = 1 To iColumns
            C176563DataType(iCol) = clsRuleDynamicallyCompiled_R10096.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C359545

    Comp_C176765:

        ' Ins_PedVencs
        sCurrComponent = "Ins_PedVencs"
        QueryC176765 = con.CreateCommand()
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "INSERT INTO AKBUSER01.WF_PEDIDO_VENCTOS "
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "( WF_PEDIDO_VENCTOS.Tp_Produto , WF_PEDIDO_VENCTOS.Pedido , "
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "  WF_PEDIDO_VENCTOS.DATA , WF_PEDIDO_VENCTOS.DIAS ) "
        QueryC176765.CommandText = QueryC176765.CommandText & " " & ""
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "SELECT WF_PRE_PEDIDO.Tp_Produto  , WF_PRE_PEDIDO.Pedido, "
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "                  WF_PRE_PEDIDO_VENCIMENTOS.Data, WF_PRE_PEDIDO_VENCIMENTOS.Dias "
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO, AKBUSER01.WF_PRE_PEDIDO_VENCIMENTOS "
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = WF_PRE_PEDIDO_VENCIMENTOS.Id_PrePedido "
        QueryC176765.CommandText = QueryC176765.CommandText & " " & "AND        WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC176765.Transaction = txn
        C176765 = QueryC176765.ExecuteNonQuery()
        C176765DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C97960

    Comp_C178051:

        ' Verif_Preco
        sCurrComponent = "Verif_Preco"
        C178051 = clsRuleDynamicallyCompiled_R1880.R1880(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56918), C56918, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), System.DBNull.Value, IIf(Not IsDBNull(C56919), C56919, System.DBNull.Value))
        C178051CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C178051) Then
          iColumns = C178051.Columns.Count
        End If
        ReDim C178051DataType(iColumns)
        For iCol = 1 To iColumns
            C178051DataType(iCol) = clsRuleDynamicallyCompiled_R1880.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C178053

    Comp_C178053:

        ' Preco Ok?
        sCurrComponent = "Preco Ok?"
        C178053 = (fn_ConvertValueCompiled(C178051.rows(C178051CurrentRow - 1)(0), C178051DataType(1), AKB_DecimalPoint, False)  = 1)
        C178053DataType = AKBTypeConst.cAKBTypeLogical
        If C178053 Then
            GoTo Comp_C432039
        Else
            GoTo Comp_C96585
        End If

    Comp_C180507:

        ' Verifica Origem do Pre_Pedido
        sCurrComponent = "Verifica Origem do Pre_Pedido"
        QueryC180507 = con.CreateCommand()
        QueryC180507.CommandText = QueryC180507.CommandText & " " & "SELECT"
        QueryC180507.CommandText = QueryC180507.CommandText & " " & " WF_PRE_PEDIDO.Id_PrePedido "
        QueryC180507.CommandText = QueryC180507.CommandText & " " & "FROM"
        QueryC180507.CommandText = QueryC180507.CommandText & " " & " AKBUSER01.WF_PRE_PEDIDO "
        QueryC180507.CommandText = QueryC180507.CommandText & " " & "WHERE"
        QueryC180507.CommandText = QueryC180507.CommandText & " " & " WF_PRE_PEDIDO.Origem_OC = 1"
        QueryC180507.CommandText = QueryC180507.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(C180508, C180508DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC180507.Transaction = txn
        RsC180507 = QueryC180507.ExecuteReader()
        Dim iC180507 As Short
        ReDim C180507DataType(RsC180507.FieldCount)
        For iC180507 = 0 to RsC180507.FieldCount - 1
            Select Case RsC180507.GetDataTypeName(iC180507).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C180507DataType(iC180507 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C180507DataType(iC180507 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C180507DataType(iC180507 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC180507
        RsC180507_EOF = Not RsC180507.Read()

        GoTo Comp_C180509

    Comp_C180508:

        ' V_ID_PrePedidoItens
        sCurrComponent = "V_ID_PrePedidoItens"
        C180508DataType = 0
        C180508DataType = C97960Datatype(7)
        If C180508DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC97960(6)) Then
          C180508 = Strings.RTrim(RsC97960(6))
        Else 
          C180508 = RsC97960(6)
        End If 

        GoTo Comp_C180507

    Comp_C180509:

        ' FV_Verifica_Origem_PrePedido
        sCurrComponent = "FV_Verifica_Origem_PrePedido"
        C180509DataType = 4
        C180509 = RsC180507_EOF
        GoTo Comp_C180510

    Comp_C180510:

        ' FV_Verifica_Origem_PrePedido = 1
        sCurrComponent = "FV_Verifica_Origem_PrePedido = 1"
        C180510 = (fn_ConvertValueCompiled(C180509, C180509DataType, AKB_DecimalPoint, False) = 1)
        C180510DataType = AKBTypeConst.cAKBTypeLogical
        If C180510 Then
            GoTo Comp_C98015
        Else
            GoTo Comp_C97971
        End If

    Comp_C206505:

        ' msg_ErrPreCadastro
        sCurrComponent = "msg_ErrPreCadastro"
        Dim row_C206505 As DataRow = msg.NewRow()
        row_C206505(0) = "O pedido nº " & _ 
C56908 & " possui o cliente " & _ 
C75153 & " o qual trata-se de pré-cadastro " & Chr(13) & "" & Chr(10) & "que não encontra-se liberado." & Chr(13) & "" & Chr(10) & "Favor verificar!" & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Processo cancelado."
        msg.Rows.Add(row_C206505)
        C206505DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C164223

    Comp_C213564:

        ' Desc. PPR
        sCurrComponent = "Desc. PPR"
        C213564 = clsRuleDynamicallyCompiled_R11322.R11322(con, txn, IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value))
        C213564CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C213564) Then
          iColumns = C213564.Columns.Count
        End If
        ReDim C213564DataType(iColumns)
        For iCol = 1 To iColumns
            C213564DataType(iCol) = clsRuleDynamicallyCompiled_R11322.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C391632

    Comp_C217998:

        ' Moeda_TabCusto
        sCurrComponent = "Moeda_TabCusto"
        C217998DataType = 0
        C217998DataType = C56123Datatype(22)
        If C217998DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(21)) Then
          C217998 = Strings.RTrim(RsC56123(21))
        Else 
          C217998 = RsC56123(21)
        End If 

        GoTo Comp_C57793

    Comp_C217999:

        ' Moeda_TabVda
        sCurrComponent = "Moeda_TabVda"
        C217999DataType = 0
        C217999DataType = C56123Datatype(21)
        If C217999DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(20)) Then
          C217999 = Strings.RTrim(RsC56123(20))
        Else 
          C217999 = RsC56123(20)
        End If 

        GoTo Comp_C217998

    Comp_C218740:

        ' Sel_Cotacao
        sCurrComponent = "Sel_Cotacao"
        QueryC218740 = con.CreateCommand()
        QueryC218740.CommandText = QueryC218740.CommandText & " " & "SELECT NVL(WF_MOEDAS_ULTIMA_COTACAO.MINIMO,0 ) COTACAO"
        QueryC218740.CommandText = QueryC218740.CommandText & " " & "FROM  AKBUSER01.WF_MOEDAS_ULTIMA_COTACAO, AKBUSER01.WF_ESTAB_JURIDICO"
        QueryC218740.CommandText = QueryC218740.CommandText & " " & "WHERE WF_ESTAB_JURIDICO.Codigo_Moeda = WF_MOEDAS_ULTIMA_COTACAO.Codigo_Moeda_Valor (+)"
        QueryC218740.CommandText = QueryC218740.CommandText & " " & "AND        WF_ESTAB_JURIDICO.Estabelecimento_Default = 1"
        QueryC218740.CommandText = QueryC218740.CommandText & " " & "AND        ( WF_MOEDAS_ULTIMA_COTACAO.Codigo_Moeda =  " & _ 
ConvertValue(C217999, C217999DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " OR "
        QueryC218740.CommandText = QueryC218740.CommandText & " " & "                    WF_MOEDAS_ULTIMA_COTACAO.Codigo_Moeda = " & _ 
ConvertValue(C217998, C217998DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC218740.Transaction = txn
        RsC218740 = QueryC218740.ExecuteReader()
        Dim iC218740 As Short
        ReDim C218740DataType(RsC218740.FieldCount)
        For iC218740 = 0 to RsC218740.FieldCount - 1
            Select Case RsC218740.GetDataTypeName(iC218740).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C218740DataType(iC218740 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C218740DataType(iC218740 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C218740DataType(iC218740 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC218740
        RsC218740_EOF = Not RsC218740.Read()

        GoTo Comp_C90279

    Comp_C233469:

        ' VerificaComissao
        sCurrComponent = "VerificaComissao"
        C233469 = clsRuleDynamicallyCompiled_R4484.R4484(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), fn_ConvertValueCompiled( "1", 1, AKB_DecimalPoint), IIf(Not IsDBNull(C174615), C174615, System.DBNull.Value), IIf(Not IsDBNull(C130730), C130730, System.DBNull.Value), IIf(Not IsDBNull(C92650), C92650, System.DBNull.Value), IIf(Not IsDBNull(C75153), C75153, System.DBNull.Value), fn_ConvertValueCompiled( "1", 4, AKB_DecimalPoint), IIf(Not IsDBNull(C414346), C414346, System.DBNull.Value))
        C233469CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C233469) Then
          iColumns = C233469.Columns.Count
        End If
        ReDim C233469DataType(iColumns)
        For iCol = 1 To iColumns
            C233469DataType(iCol) = clsRuleDynamicallyCompiled_R4484.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C233557

    Comp_C233557:

        ' VerificaComissao=1
        sCurrComponent = "VerificaComissao=1"
        C233557 = (fn_ConvertValueCompiled(C233469.rows(C233469CurrentRow - 1)(0), C233469DataType(1), AKB_DecimalPoint, False) = 1)
        C233557DataType = AKBTypeConst.cAKBTypeLogical
        If C233557 Then
            GoTo Comp_C56132
        Else
            GoTo Comp_C339255
        End If

    Comp_C233558:

        ' Cancel
        sCurrComponent = "Cancel"
        txn.Rollback()
        C233558 = True
        C233558DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C233559

    Comp_C233559:

        ' RetCancel
        sCurrComponent = "RetCancel"
        Dim tb_C233559 As DataTable = New DataTable()
        tb_C233559.Columns.Add("Cancel" & "")
        Dim row_C233559 As DataRow = tb_C233559.NewRow()
        row_C233559(0) = C233558
        tb_C233559.Rows.Add(row_C233559)
        R4557 = tb_C233559
        ReDim C233559DataType(1)
        C233559DataType(1) = C233558DataType
        ReturnDataType = C233559DataType

        GoTo Exit_R4557

    Comp_C246209:

        ' Up_PrazoPrePedidoItens2
        sCurrComponent = "Up_PrazoPrePedidoItens2"
        QueryC246209 = con.CreateCommand()
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "Update AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "Set WF_PRE_PEDIDO_ITENS.Prazo_Entrega = "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "Decode ( "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "NVL( ( SELECT WF_EMB_COMP_VDA_PROD.Prazo_Entrega_Minimo"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "             FROM AKBUSER01.WF_EMB_COMP_VDA_PROD "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "             WHERE WF_PRE_PEDIDO_ITENS.Sigla_Prod = WF_EMB_COMP_VDA_PROD.Sigla_Prod "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "             AND WF_PRE_PEDIDO_ITENS.Acesso = WF_EMB_COMP_VDA_PROD.Acesso"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "             AND WF_PRE_PEDIDO_ITENS.Cod_Embalagem = WF_EMB_COMP_VDA_PROD.Cod_Embalagem"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "            AND WF_PRE_PEDIDO_ITENS.Tipo_Ped = WF_EMB_COMP_VDA_PROD.Tipo_Ped"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "           )"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "   ,0 ), "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "        0, WF_PRE_PEDIDO_ITENS.Prazo_Entrega, "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "         Sysdate + "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "     ( SELECT WF_EMB_COMP_VDA_PROD.Prazo_Entrega_Minimo"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "       FROM AKBUSER01.WF_EMB_COMP_VDA_PROD "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "       WHERE WF_PRE_PEDIDO_ITENS.Sigla_Prod = WF_EMB_COMP_VDA_PROD.Sigla_Prod "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "       AND WF_PRE_PEDIDO_ITENS.Acesso = WF_EMB_COMP_VDA_PROD.Acesso"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "       AND WF_PRE_PEDIDO_ITENS.Cod_Embalagem = WF_EMB_COMP_VDA_PROD.Cod_Embalagem"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "       AND WF_PRE_PEDIDO_ITENS.Tipo_Ped = WF_EMB_COMP_VDA_PROD.Tipo_Ped"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "     )"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & ")"
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "Where WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC246209.CommandText = QueryC246209.CommandText & " " & "AND " & _ 
ConvertValue(C341913, C341913DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " = 0"
        QueryC246209.Transaction = txn
        C246209 = QueryC246209.ExecuteNonQuery()
        C246209DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C519010

    Comp_C252514:

        ' Up_PrazoPrePedidoItens
        sCurrComponent = "Up_PrazoPrePedidoItens"
        QueryC252514 = con.CreateCommand()
        QueryC252514.CommandText = QueryC252514.CommandText & " " & "Update AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC252514.CommandText = QueryC252514.CommandText & " " & "Set WF_PRE_PEDIDO_ITENS.Prazo_Entrega = Sysdate"
        QueryC252514.CommandText = QueryC252514.CommandText & " " & "Where WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC252514.CommandText = QueryC252514.CommandText & " " & "And WF_PRE_PEDIDO_ITENS.Prazo_Entrega < Sysdate "
        QueryC252514.Transaction = txn
        C252514 = QueryC252514.ExecuteNonQuery()
        C252514DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C341912

    Comp_C265444:

        ' TpProduto=200
        sCurrComponent = "TpProduto=200"
        C265444 = (( fn_ConvertValueCompiled(C56126, C56126DataType, AKB_DecimalPoint, False) = 200 ) or ( fn_ConvertValueCompiled(C56126, C56126DataType, AKB_DecimalPoint, False) = 201 ))
        C265444DataType = AKBTypeConst.cAKBTypeLogical
        If C265444 Then
            GoTo Comp_C266026
        Else
            GoTo Comp_C368027
        End If

    Comp_C265481:

        ' Coligada
        sCurrComponent = "Coligada"
        QueryC265481 = con.CreateCommand()
        QueryC265481.CommandText = QueryC265481.CommandText & " " & "SELECT COUNT(WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa) FROM AKBUSER01.WF_ESTAB_JURID_COLIGADAS WHERE WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa = " & _ 
ConvertValue(C75153, C75153DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND  WF_ESTAB_JURID_COLIGADAS.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(C75154, C75154DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC265481.Transaction = txn
        RsC265481 = QueryC265481.ExecuteReader()
        Dim iC265481 As Short
        ReDim C265481DataType(RsC265481.FieldCount)
        For iC265481 = 0 to RsC265481.FieldCount - 1
            Select Case RsC265481.GetDataTypeName(iC265481).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C265481DataType(iC265481 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C265481DataType(iC265481 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C265481DataType(iC265481 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC265481
        RsC265481_EOF = Not RsC265481.Read()

        GoTo Comp_C419420

    Comp_C265482:

        ' Usuario
        sCurrComponent = "Usuario"
        C265482DataType = 3
        C265482 = Mid(con.ConnectionString, InStr(con.ConnectionString, "User Id=")+8, Instr(InStr(con.ConnectionString, "User Id="), con.ConnectionString, ";")-InStr(con.ConnectionString, "User Id=")-8)
        GoTo Comp_C90294

    Comp_C265483:

        ' Coligada=0?
        sCurrComponent = "Coligada=0?"
        C265483 = (fn_ConvertValueCompiled(RsC265481(0), C265481DataType(1), AKB_DecimalPoint, False) = 0)
        C265483DataType = AKBTypeConst.cAKBTypeLogical
        If C265483 Then
            GoTo Comp_C414348
        Else
            GoTo Comp_C265488
        End If

    Comp_C265488:

        ' ATUAL1
        sCurrComponent = "ATUAL1"
        QueryC265488 = con.CreateCommand()
        QueryC265488.CommandText = QueryC265488.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO SET WF_PEDIDO.Verificado = 1 ,"
        QueryC265488.CommandText = QueryC265488.CommandText & " " & " WF_PEDIDO.Usuario_Verificacao = " & _ 
ConvertValue(C265482, C265482DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  , "
        QueryC265488.CommandText = QueryC265488.CommandText & " " & "WF_PEDIDO.Data_Verificacao = " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC265488.CommandText = QueryC265488.CommandText & " " & "WHERE WF_PEDIDO.Tp_Produto = " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC265488.CommandText = QueryC265488.CommandText & " " & "AND  WF_PEDIDO.Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  "
        QueryC265488.Transaction = txn
        C265488 = QueryC265488.ExecuteNonQuery()
        C265488DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C414348

    Comp_C266026:

        ' LibPedManual
        sCurrComponent = "LibPedManual"
        'C266026 = clsRuleDynamicallyCompiled_R9166.R9166(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), System.DBNull.Value, System.DBNull.Value)
        C266026CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C266026) Then
          iColumns = C266026.Columns.Count
        End If
        ReDim C266026DataType(iColumns)
        For iCol = 1 To iColumns
            'C266026DataType(iCol) = clsRuleDynamicallyCompiled_R9166.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C109297

    Comp_C283840:

        ' SelItemMaisColeção
        sCurrComponent = "SelItemMaisColeção"
        QueryC283840 = con.CreateCommand()
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "SELECT WF_PRE_PEDIDO_ITENS.Acesso ,"
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "                  WF_PRE_PEDIDO_ITENS.Sigla_Prod ,  WF_PRE_PEDIDO_ITENS.Cod_Embalagem ,"
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "                  COUNT(WF_PRE_PEDIDO_ITENS.Seq_Item) "
        QueryC283840.CommandText = QueryC283840.CommandText & " " & ""
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_ITENS , AKBUSER01.WF_COLECAO_PRODUTOS "
        QueryC283840.CommandText = QueryC283840.CommandText & " " & ""
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "      AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "      AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "      AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "      AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final  >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "            OR WF_COLECAO_PRODUTOS.Data_Validade_Final  IS NULL )"
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "      AND  WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC283840.CommandText = QueryC283840.CommandText & " " & ""
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "GROUP BY WF_PRE_PEDIDO_ITENS.Acesso , WF_PRE_PEDIDO_ITENS.Sigla_Prod ,"
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "                         WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC283840.CommandText = QueryC283840.CommandText & " " & ""
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "HAVING COUNT(WF_PRE_PEDIDO_ITENS.Seq_Item)  > 1"
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "UNION ALL"
        QueryC283840.CommandText = QueryC283840.CommandText & " " & "SELECT 0, NULL, NULL, NULL FROM DUAL"
        QueryC283840.Transaction = txn
        RsC283840 = QueryC283840.ExecuteReader()
        Dim iC283840 As Short
        ReDim C283840DataType(RsC283840.FieldCount)
        For iC283840 = 0 to RsC283840.FieldCount - 1
            Select Case RsC283840.GetDataTypeName(iC283840).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C283840DataType(iC283840 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C283840DataType(iC283840 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C283840DataType(iC283840 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC283840
        RsC283840_EOF = Not RsC283840.Read()

        GoTo Comp_C283841

    Comp_C283841:

        ' ListaSel
        sCurrComponent = "ListaSel"
        C283841DataType = 5
        C283841 = ""
        Do Until RsC283840_EOF
            If RTrim(C283841) <> "" Then
                C283841 = C283841 & ", "
            End If
            C283841 = C283841 & ConvertValue(RsC283840(0), 0, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss")
            RsC283840_EOF = Not RsC283840.Read()
        Loop

        GoTo Comp_C284023

    Comp_C283843:

        ' Não existe itens com mais de 1 coleção?
        sCurrComponent = "Não existe itens com mais de 1 coleção?"
        C283843 = (fn_ConvertValueCompiled(C284023, C284023DataType, AKB_DecimalPoint, False) = 0)
        C283843DataType = AKBTypeConst.cAKBTypeLogical
        If C283843 Then
            GoTo Comp_C175718
        Else
            GoTo Comp_C350049
        End If

    Comp_C284023:

        ' ComprListaSel
        sCurrComponent = "ComprListaSel"
        C284023DataType = 1
        C284023 = Len(C283841 & "")
        GoTo Comp_C419441

    Comp_C310322:

        ' Repres_zona
        sCurrComponent = "Repres_zona"
        QueryC310322 = con.CreateCommand()
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "SELECT WF_CLIENTES.Cod_Zona "
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES, AKBUSER01.WF_REPRES_ZONA, AKBUSER01.WF_PRE_PEDIDO"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & ""
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "WHERE WF_REPRES_ZONA.Cod_Zona = WF_CLIENTES.Cod_Zona"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_REPRES_ZONA.Cod_Repres = WF_PRE_PEDIDO.Cod_Repres"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_CLIENTES.Cod_Cliente = WF_PRE_PEDIDO.Cod_Cliente "
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_PRE_PEDIDO.Cod_Cliente2 IS NULL"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & ""
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "UNION"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & ""
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "SELECT WF_CLIENTES.Cod_Zona "
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "FROM AKBUSER01.WF_CLIENTES, AKBUSER01.WF_REPRES_ZONA, AKBUSER01.WF_PRE_PEDIDO"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & ""
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "WHERE WF_REPRES_ZONA.Cod_Zona = WF_CLIENTES.Cod_Zona"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_REPRES_ZONA.Cod_Repres = WF_PRE_PEDIDO.Cod_Repres"
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_CLIENTES.Cod_Cliente = WF_PRE_PEDIDO.Cod_Cliente2 "
        QueryC310322.CommandText = QueryC310322.CommandText & " " & "AND WF_PRE_PEDIDO.Cod_Cliente2 IS NOT NULL"
        QueryC310322.Transaction = txn
        RsC310322 = QueryC310322.ExecuteReader()
        Dim iC310322 As Short
        ReDim C310322DataType(RsC310322.FieldCount)
        For iC310322 = 0 to RsC310322.FieldCount - 1
            Select Case RsC310322.GetDataTypeName(iC310322).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C310322DataType(iC310322 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C310322DataType(iC310322 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C310322DataType(iC310322 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC310322
        RsC310322_EOF = Not RsC310322.Read()

        GoTo Comp_C310328

    Comp_C310328:

        ' Existe_zona ?
        sCurrComponent = "Existe_zona ?"
        C310328 = (fn_ConvertValueCompiled(RsC310322(0), C310322DataType(1), AKB_DecimalPoint, False) <>0)
        C310328DataType = AKBTypeConst.cAKBTypeLogical
        If C310328 Then
            GoTo Comp_C164218
        Else
            GoTo Comp_C310329
        End If

    Comp_C310329:

        ' Repres
        sCurrComponent = "Repres"
        Dim row_C310329 As DataRow = msg.NewRow()
        row_C310329(0) = "Representante fora da zona do cliente."
        msg.Rows.Add(row_C310329)
        C310329DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C310330

    Comp_C310330:

        ' RET6
        sCurrComponent = "RET6"
        Dim tb_C310330 As DataTable = New DataTable()
        tb_C310330.Columns.Add("Existe_zona ?" & "")
        Dim row_C310330 As DataRow = tb_C310330.NewRow()
        row_C310330(0) = C310328
        tb_C310330.Rows.Add(row_C310330)
        R4557 = tb_C310330
        ReDim C310330DataType(1)
        C310330DataType(1) = C310328DataType
        ReturnDataType = C310330DataType

        GoTo Exit_R4557

    Comp_C315473:

        ' EmailConfPed
        sCurrComponent = "EmailConfPed"
        QueryC315473 = con.CreateCommand()
        QueryC315473.CommandText = QueryC315473.CommandText & " " & "SELECT WF_ESTAB_JURIDICO_PARAM.COD_COM_EMAIL_CONF_PED FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM WHERE WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(C75154, C75154DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC315473.Transaction = txn
        RsC315473 = QueryC315473.ExecuteReader()
        Dim iC315473 As Short
        ReDim C315473DataType(RsC315473.FieldCount)
        For iC315473 = 0 to RsC315473.FieldCount - 1
            Select Case RsC315473.GetDataTypeName(iC315473).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C315473DataType(iC315473 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C315473DataType(iC315473 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C315473DataType(iC315473 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC315473
        RsC315473_EOF = Not RsC315473.Read()

        GoTo Comp_C265481

    Comp_C315844:

        ' Insere Pedido Itens
        sCurrComponent = "Insere Pedido Itens"
        'C315844 = clsRuleDynamicallyCompiled_R14876.R14876(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value), IIf(Not IsDBNull(C367999), C367999, System.DBNull.Value))
        C315844CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C315844) Then
          iColumns = C315844.Columns.Count
        End If
        ReDim C315844DataType(iColumns)
        For iCol = 1 To iColumns
            'C315844DataType(iCol) = clsRuleDynamicallyCompiled_R14876.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C54114

    Comp_C315845:

        ' selPrePedidoCargaMaquina
        sCurrComponent = "selPrePedidoCargaMaquina"
        QueryC315845 = con.CreateCommand()
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "SELECT "
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "(SELECT COUNT(WF_CARGA_EST_RES_PED_INT.Sequencia) "
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "FROM AKBUSER01.WF_CARGA_EST_RES_PED_INT "
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "WHERE WF_CARGA_EST_RES_PED_INT.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ") +"
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "(SELECT COUNT(WF_CARGA_EST_RES_PRE_PED.Sequencia) "
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "FROM AKBUSER01.WF_CARGA_EST_RES_PRE_PED "
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "WHERE WF_CARGA_EST_RES_PRE_PED.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC315845.CommandText = QueryC315845.CommandText & " " & "FROM DUAL"
        QueryC315845.Transaction = txn
        RsC315845 = QueryC315845.ExecuteReader()
        Dim iC315845 As Short
        ReDim C315845DataType(RsC315845.FieldCount)
        For iC315845 = 0 to RsC315845.FieldCount - 1
            Select Case RsC315845.GetDataTypeName(iC315845).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C315845DataType(iC315845 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C315845DataType(iC315845 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C315845DataType(iC315845 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC315845
        RsC315845_EOF = Not RsC315845.Read()

        GoTo Comp_C368020

    Comp_C315847:

        ' Tem PrePedido na Carga Máquina?
        sCurrComponent = "Tem PrePedido na Carga Máquina?"
        C315847 = (fn_ConvertValueCompiled(RsC315845(0), C315845DataType(1), AKB_DecimalPoint, False) > 0)
        C315847DataType = AKBTypeConst.cAKBTypeLogical
        If C315847 Then
            GoTo Comp_C315844
        Else
            GoTo Comp_C420246
        End If

    Comp_C317977:

        ' BNDES
        sCurrComponent = "BNDES"
        QueryC317977 = con.CreateCommand()
        QueryC317977.CommandText = QueryC317977.CommandText & " " & "SELECT WF_CONDICAO_PGTO.E_Cartao_BNDES FROM AKBUSER01.WF_CONDICAO_PGTO WHERE WF_CONDICAO_PGTO.Cod_Pagto = " & _ 
ConvertValue(C56920, C56920DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC317977.Transaction = txn
        RsC317977 = QueryC317977.ExecuteReader()
        Dim iC317977 As Short
        ReDim C317977DataType(RsC317977.FieldCount)
        For iC317977 = 0 to RsC317977.FieldCount - 1
            Select Case RsC317977.GetDataTypeName(iC317977).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C317977DataType(iC317977 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C317977DataType(iC317977 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C317977DataType(iC317977 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC317977
        RsC317977_EOF = Not RsC317977.Read()

        GoTo Comp_C130731

    Comp_C317978:

        ' BNDES=0?
        sCurrComponent = "BNDES=0?"
        C317978 = (fn_ConvertValueCompiled(RsC317977(0), C317977DataType(1), AKB_DecimalPoint, False) = 0)
        C317978DataType = AKBTypeConst.cAKBTypeLogical
        If C317978 Then
            GoTo Comp_C545148
        Else
            GoTo Comp_C75152
        End If

    Comp_C339223:

        ' Iniciar_Transação?
        sCurrComponent = "Iniciar_Transação?"
        C339223 = (fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) =1)
        C339223DataType = AKBTypeConst.cAKBTypeLogical
        If C339223 Then
            GoTo Comp_C54129
        Else
            GoTo Comp_C56122
        End If

    Comp_C339248:

        ' INICIAR_TRANSAÇÃO?1
        sCurrComponent = "INICIAR_TRANSAÇÃO?1"
        C339248 = (fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) =1)
        C339248DataType = AKBTypeConst.cAKBTypeLogical
        If C339248 Then
            GoTo Comp_C54130
        Else
            GoTo Comp_C265444
        End If

    Comp_C339250:

        ' INICIAR_TRANSAÇÃO?2
        sCurrComponent = "INICIAR_TRANSAÇÃO?2"
        C339250 = (fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) =1)
        C339250DataType = AKBTypeConst.cAKBTypeLogical
        If C339250 Then
            GoTo Comp_C96571
        Else
            GoTo Comp_C56916
        End If

    Comp_C339251:

        ' INICIAR_TRANSAÇÃO?3
        sCurrComponent = "INICIAR_TRANSAÇÃO?3"
        C339251 = (fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) =1)
        C339251DataType = AKBTypeConst.cAKBTypeLogical
        If C339251 Then
            GoTo Comp_C96588
        Else
            GoTo Comp_C96669
        End If

    Comp_C339252:

        ' INICIAR_TRANSAÇÃO?4
        sCurrComponent = "INICIAR_TRANSAÇÃO?4"
        C339252 = (fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) =1)
        C339252DataType = AKBTypeConst.cAKBTypeLogical
        If C339252 Then
            GoTo Comp_C97967
        Else
            GoTo Comp_C97966
        End If

    Comp_C339255:

        ' INICIAR_TRANSAÇÃO?5
        sCurrComponent = "INICIAR_TRANSAÇÃO?5"
        C339255 = (fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) =1)
        C339255DataType = AKBTypeConst.cAKBTypeLogical
        If C339255 Then
            GoTo Comp_C233558
        Else
            GoTo Comp_C233559
        End If

    Comp_C341912:

        ' selCalcularPrazo
        sCurrComponent = "selCalcularPrazo"
        QueryC341912 = con.CreateCommand()
        QueryC341912.CommandText = QueryC341912.CommandText & " " & "SELECT NVL(WF_ESTAB_JURIDICO_PARAM.Calc_Prazo_Pre_Pedido ,0)"
        QueryC341912.CommandText = QueryC341912.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM , AKBUSER01.WF_PRE_PEDIDO "
        QueryC341912.CommandText = QueryC341912.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC341912.CommandText = QueryC341912.CommandText & " " & "AND WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico "
        QueryC341912.Transaction = txn
        RsC341912 = QueryC341912.ExecuteReader()
        Dim iC341912 As Short
        ReDim C341912DataType(RsC341912.FieldCount)
        For iC341912 = 0 to RsC341912.FieldCount - 1
            Select Case RsC341912.GetDataTypeName(iC341912).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C341912DataType(iC341912 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C341912DataType(iC341912 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C341912DataType(iC341912 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC341912
        RsC341912_EOF = Not RsC341912.Read()

        GoTo Comp_C341913

    Comp_C341913:

        ' vCalcularPrazo
        sCurrComponent = "vCalcularPrazo"
        C341913DataType = 0
        C341913 = RsC341912(0)
        C341913DataType = C341912Datatype(1)
        If C341913DataType = AKBTypeConst.cAKBTypeString Then
          C341913 = IIF(IsDBNull(C341913), C341913, Strings.RTrim(C341913))
        End If 

        GoTo Comp_C246209

    Comp_C347147:

        ' Verifica PrePedidoWeb
        sCurrComponent = "Verifica PrePedidoWeb"
        C347147 = clsRuleDynamicallyCompiled_R16162.R16162(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value))
        C347147CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C347147) Then
          iColumns = C347147.Columns.Count
        End If
        ReDim C347147DataType(iColumns)
        For iCol = 1 To iColumns
            C347147DataType(iCol) = clsRuleDynamicallyCompiled_R16162.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C347148

    Comp_C347148:

        ' Verifica PrePedidoWeb = 1
        sCurrComponent = "Verifica PrePedidoWeb = 1"
        C347148 = (fn_ConvertValueCompiled(C347147.rows(C347147CurrentRow - 1)(0), C347147DataType(1), AKB_DecimalPoint, False) = 1)
        C347148DataType = AKBTypeConst.cAKBTypeLogical
        If C347148 Then
            GoTo Comp_C374705
        Else
            GoTo Comp_C96668
        End If

    Comp_C350047:

        ' upColecaoTabelaVenda
        sCurrComponent = "upColecaoTabelaVenda"
        QueryC350047 = con.CreateCommand()
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.Id_Colecao = ("
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "SELECT WF_TABELA_PRECO_VENDA.Id_Colecao "
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS , AKBUSER01.WF_TABELA_PRECO_VENDA, AKBUSER01.WF_PRE_PEDIDO "
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND  WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND  ( WF_COLECAO_PRODUTOS.Data_Validade_Final  >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "    OR WF_COLECAO_PRODUTOS.Data_Validade_Final  IS NULL )"
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND WF_PRE_PEDIDO.Tabela_Preco_Venda = WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda"
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido )"
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND (WF_PRE_PEDIDO_ITENS.Id_Colecao IS NULL OR WF_PRE_PEDIDO_ITENS.Id_Colecao = 0)"
        QueryC350047.CommandText = QueryC350047.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Acesso IN (" & _ 
ConvertValue(C283841, C283841DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC350047.Transaction = txn
        C350047 = QueryC350047.ExecuteNonQuery()
        C350047DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C174612

    Comp_C350049:

        ' TABSERV=NULO1
        sCurrComponent = "TABSERV=NULO1"
        C350049 = (( fn_ConvertValueCompiled(C56919, C56919DataType, AKB_DecimalPoint, False)="" ) or ( fn_ConvertValueCompiled(C56919, C56919DataType, AKB_DecimalPoint, False)=0 ) or ( fn_ConvertValueCompiled(C56919, C56919DataType, AKB_DecimalPoint, False) Is System.DBNull.Value ))
        C350049DataType = AKBTypeConst.cAKBTypeLogical
        If C350049 Then
            GoTo Comp_C350047
        Else
            GoTo Comp_C350050
        End If

    Comp_C350050:

        ' upColecaoTabelaServico
        sCurrComponent = "upColecaoTabelaServico"
        QueryC350050 = con.CreateCommand()
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.Id_Colecao = ("
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "SELECT WF_TAB_PRECO_CUSTO.Id_Colecao"
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS , AKBUSER01.WF_TAB_PRECO_CUSTO , AKBUSER01.WF_PRE_PEDIDO "
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso "
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Data_Validade_Inicial <= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND ( WF_COLECAO_PRODUTOS.Data_Validade_Final  >= " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "    OR WF_COLECAO_PRODUTOS.Data_Validade_Final  IS NULL )"
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND WF_PRE_PEDIDO.Tabela_Preco_Custo = WF_TAB_PRECO_CUSTO.Tabela_Preco_Custo "
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND  WF_TAB_PRECO_CUSTO.Id_Colecao = WF_COLECAO_PRODUTOS.Id_Colecao"
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido )"
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND (WF_PRE_PEDIDO_ITENS.Id_Colecao IS NULL OR WF_PRE_PEDIDO_ITENS.Id_Colecao = 0)"
        QueryC350050.CommandText = QueryC350050.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Acesso IN (" & _ 
ConvertValue(C283841, C283841DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC350050.Transaction = txn
        C350050 = QueryC350050.ExecuteNonQuery()
        C350050DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C175721

    Comp_C357368:

        ' vDepto
        sCurrComponent = "vDepto"
        C357368DataType = 0
        C357368DataType = C56123Datatype(23)
        If C357368DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(22)) Then
          C357368 = Strings.RTrim(RsC56123(22))
        Else 
          C357368 = RsC56123(22)
        End If 

        GoTo Comp_C367574

    Comp_C357369:

        ' Pronta Entrega?
        sCurrComponent = "Pronta Entrega?"
        C357369 = (fn_ConvertValueCompiled(C357368, C357368DataType, AKB_DecimalPoint, False) = "COM31")
        C357369DataType = AKBTypeConst.cAKBTypeLogical
        If C357369 Then
            GoTo Comp_C359024
        Else
            GoTo Comp_C283840
        End If

    Comp_C357370:

        ' upColecaoProntaEntrega
        sCurrComponent = "upColecaoProntaEntrega"
        QueryC357370 = con.CreateCommand()
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS"
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.Id_Colecao = ("
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "SELECT WF_TABELA_PRECO_VENDA.Id_Colecao"
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "FROM AKBUSER01.WF_TABELA_PRECO_VENDA"
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "WHERE WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = " & _ 
ConvertValue(C56918, C56918DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "),"
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod2 = WF_PRE_PEDIDO_ITENS.Sigla_Prod, "
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso2 = WF_PRE_PEDIDO_ITENS.Acesso,"
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem2 = WF_PRE_PEDIDO_ITENS.Cod_Embalagem "
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC357370.CommandText = QueryC357370.CommandText & " " & "AND (WF_PRE_PEDIDO_ITENS.Id_Colecao IS NULL OR WF_PRE_PEDIDO_ITENS.Id_Colecao = 0) "
        QueryC357370.Transaction = txn
        C357370 = QueryC357370.ExecuteNonQuery()
        C357370DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C92633

    Comp_C359024:

        ' in_prodColecao
        sCurrComponent = "in_prodColecao"
        QueryC359024 = con.CreateCommand()
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "INSERT INTO AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "(WF_COLECAO_PRODUTOS.Id_Colecao , "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_COLECAO_PRODUTOS.Sigla_Prod , "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_COLECAO_PRODUTOS.Acesso , "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_COLECAO_PRODUTOS.Cod_Embalagem ,"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_COLECAO_PRODUTOS.Data_Validade_Inicial, "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_COLECAO_PRODUTOS.Data_Validade_Final)"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "SELECT WF_COLECAO.Id_Colecao , "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Sigla_Prod , "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Acesso , "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_PRE_PEDIDO_ITENS.Cod_Embalagem,"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_COLECAO.Data_Validade_Inicial , "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WF_COLECAO.Data_Validade_Final"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO_ITENS, "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AKBUSER01.WF_PRE_PEDIDO, "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AKBUSER01.WF_TABELA_PRECO_VENDA, "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AKBUSER01.WF_COLECAO"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = WF_PRE_PEDIDO.Id_PrePedido"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AND WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda = WF_PRE_PEDIDO.Tabela_Preco_Venda"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AND WF_COLECAO.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AND NOT EXISTS"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "(SELECT *"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC359024.CommandText = QueryC359024.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Id_Colecao = WF_TABELA_PRECO_VENDA.Id_Colecao) "
        QueryC359024.Transaction = txn
        C359024 = QueryC359024.ExecuteNonQuery()
        C359024DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C357370

    Comp_C359545:

        ' Calcula Custo Presumido do Item do Pedido
        sCurrComponent = "Calcula Custo Presumido do Item do Pedido"
        C359545 = clsRuleDynamicallyCompiled_R16456.R16456(con, txn, IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value))
        C359545CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C359545) Then
          iColumns = C359545.Columns.Count
        End If
        ReDim C359545DataType(iColumns)
        For iCol = 1 To iColumns
            C359545DataType(iCol) = clsRuleDynamicallyCompiled_R16456.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C392962

    Comp_C363384:

        ' varTransacao
        sCurrComponent = "varTransacao"
        C363384 = 1
        C363384DataType = 4
        GoTo Comp_C363386

    Comp_C363386:

        ' Ver Transação?
        sCurrComponent = "Ver Transação?"
        C363386 = (fn_ConvertValueCompiled(P20438, 4, AKB_DecimalPoint, False) = 0 and fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) = 1)
        C363386DataType = AKBTypeConst.cAKBTypeLogical
        If C363386 Then
            GoTo Comp_C56139
        Else
            GoTo Comp_C363387
        End If

    Comp_C363387:

        ' atZeroTransacao
        sCurrComponent = "atZeroTransacao"
        C363387DataType = 4
        C363384 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C363387 = True
        GoTo Comp_C56139

    Comp_C367573:

        ' UsuarioManutCred
        sCurrComponent = "UsuarioManutCred"
        C367573DataType = 0
        C367573DataType = C56123Datatype(26)
        If C367573DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(25)) Then
          C367573 = Strings.RTrim(RsC56123(25))
        Else 
          C367573 = RsC56123(25)
        End If 

        GoTo Comp_C368898

    Comp_C367574:

        ' StatusAnaliseCred
        sCurrComponent = "StatusAnaliseCred"
        C367574DataType = 0
        C367574DataType = C56123Datatype(24)
        If C367574DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(23)) Then
          C367574 = Strings.RTrim(RsC56123(23))
        Else 
          C367574 = RsC56123(23)
        End If 

        GoTo Comp_C371816

    Comp_C367575:

        ' DataManutCred
        sCurrComponent = "DataManutCred"
        C367575DataType = 0
        C367575DataType = C56123Datatype(25)
        If C367575DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(24)) Then
          C367575 = Strings.RTrim(RsC56123(24))
        Else 
          C367575 = RsC56123(24)
        End If 

        GoTo Comp_C367573

    Comp_C367578:

        ' RET7
        sCurrComponent = "RET7"
        Dim tb_C367578 As DataTable = New DataTable()
        tb_C367578.Columns.Add("Msg?" & "")
        Dim row_C367578 As DataRow = tb_C367578.NewRow()
        row_C367578(0) = C367737
        tb_C367578.Rows.Add(row_C367578)
        R4557 = tb_C367578
        ReDim C367578DataType(1)
        C367578DataType(1) = C367737DataType
        ReturnDataType = C367578DataType

        GoTo Exit_R4557

    Comp_C367579:

        ' CreditoLiberado?
        sCurrComponent = "CreditoLiberado?"
        C367579 = (fn_ConvertValueCompiled(C371816, C371816DataType, AKB_DecimalPoint, False) = "Liberado")
        C367579DataType = AKBTypeConst.cAKBTypeLogical
        If C367579 Then
            GoTo Comp_C315473
        Else
            GoTo Comp_C367737
        End If

    Comp_C367580:

        ' MsgAnalise
        sCurrComponent = "MsgAnalise"
        Dim row_C367580 As DataRow = msg.NewRow()
        row_C367580(0) = "Não é possível gerar pedido do pré pedido " & _ 
P12596 & "!" & Chr(13) & "" & Chr(10) & "O crédito encontra-se em processo de análise." & Chr(13) & "" & Chr(10) & ""
        msg.Rows.Add(row_C367580)
        C367580DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C367578

    Comp_C367737:

        ' Msg?
        sCurrComponent = "Msg?"
        C367737 = (fn_ConvertValueCompiled(P56293, 1, AKB_DecimalPoint, False) = 0)
        C367737DataType = AKBTypeConst.cAKBTypeLogical
        If C367737 Then
            GoTo Comp_C367578
        Else
            GoTo Comp_C367580
        End If

    Comp_C367790:

        ' Verifica Coleção dos Itens do Pedido
        sCurrComponent = "Verifica Coleção dos Itens do Pedido"
        C367790 = clsRuleDynamicallyCompiled_R9977.R9977(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value), fn_ConvertValueCompiled( "1", 4, AKB_DecimalPoint))
        C367790CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C367790) Then
          iColumns = C367790.Columns.Count
        End If
        ReDim C367790DataType(iColumns)
        For iCol = 1 To iColumns
            C367790DataType(iCol) = clsRuleDynamicallyCompiled_R9977.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C391471

    Comp_C367999:

        ' VFlagPed
        sCurrComponent = "VFlagPed"
        C367999 = 1
        C367999DataType = 1
        GoTo Comp_C418177

    Comp_C368000:

        ' CredPrePedLiberado?
        sCurrComponent = "CredPrePedLiberado?"
        C368000 = (fn_ConvertValueCompiled(C367574, C367574DataType, AKB_DecimalPoint, False)  = "Liberado")
        C368000DataType = AKBTypeConst.cAKBTypeLogical
        If C368000 Then
            GoTo Comp_C368016
        Else
            GoTo Comp_C54109
        End If

    Comp_C368016:

        ' AtrVlrFlagPed2
        sCurrComponent = "AtrVlrFlagPed2"
        C368016DataType = 4
        C367999 = fn_ConvertValueCompiled(2, 1, AKB_DecimalPoint)
        C368016 = True
        GoTo Comp_C54109

    Comp_C368020:

        ' CredPrePediLiberado?
        sCurrComponent = "CredPrePediLiberado?"
        C368020 = (fn_ConvertValueCompiled(C367574, C367574DataType, AKB_DecimalPoint, False)  = "Liberado")
        C368020DataType = AKBTypeConst.cAKBTypeLogical
        If C368020 Then
            GoTo Comp_C368021
        Else
            GoTo Comp_C315847
        End If

    Comp_C368021:

        ' AtrVlFlagPed2
        sCurrComponent = "AtrVlFlagPed2"
        C368021DataType = 4
        C367999 = fn_ConvertValueCompiled(2, 1, AKB_DecimalPoint)
        C368021 = True
        GoTo Comp_C315847

    Comp_C368027:

        ' CredPrePedLiberado??
        sCurrComponent = "CredPrePedLiberado??"
        C368027 = (fn_ConvertValueCompiled(C367574, C367574DataType, AKB_DecimalPoint, False)  = "Liberado")
        C368027DataType = AKBTypeConst.cAKBTypeLogical
        If C368027 Then
            GoTo Comp_C372920
        Else
            GoTo Comp_C317978
        End If

    Comp_C368405:

        ' SelRepresCliente
        sCurrComponent = "SelRepresCliente"
        QueryC368405 = con.CreateCommand()
        QueryC368405.CommandText = QueryC368405.CommandText & " " & "SELECT COUNT(*) "
        QueryC368405.CommandText = QueryC368405.CommandText & " " & "FROM AKBUSER01.WF_REPRESENTANTE_CLIENTE "
        QueryC368405.CommandText = QueryC368405.CommandText & " " & "WHERE WF_REPRESENTANTE_CLIENTE.Cod_Repres = " & _ 
ConvertValue(C92650, C92650DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368405.CommandText = QueryC368405.CommandText & " " & "AND WF_REPRESENTANTE_CLIENTE.Cod_Cliente =" & _ 
ConvertValue(C75153, C75153DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC368405.Transaction = txn
        RsC368405 = QueryC368405.ExecuteReader()
        Dim iC368405 As Short
        ReDim C368405DataType(RsC368405.FieldCount)
        For iC368405 = 0 to RsC368405.FieldCount - 1
            Select Case RsC368405.GetDataTypeName(iC368405).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C368405DataType(iC368405 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C368405DataType(iC368405 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C368405DataType(iC368405 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC368405
        RsC368405_EOF = Not RsC368405.Read()

        GoTo Comp_C414350

    Comp_C368408:

        ' Tem RepresCliente?
        sCurrComponent = "Tem RepresCliente?"
        C368408 = (fn_ConvertValueCompiled(C414347, C414347DataType, AKB_DecimalPoint, False) >0)
        C368408DataType = AKBTypeConst.cAKBTypeLogical
        If C368408 Then
            GoTo Comp_C423442
        Else
            GoTo Comp_C368409
        End If

    Comp_C368409:

        ' MSG2
        sCurrComponent = "MSG2"
        Dim row_C368409 As DataRow = msg.NewRow()
        row_C368409(0) = "Não é possivel gerar o pedido. " & Chr(13) & "" & Chr(10) & "O Representante " & _ 
C92650 & "  não está associado com o Cliente." & Chr(13) & "" & Chr(10) & "Cliente do pedido: " & _ 
C75153 & " ." & Chr(13) & "" & Chr(10) & "Cliente Final do pedido: " & _ 
C414346 & " ." & Chr(13) & "" & Chr(10) & "Por favor atualize o cadastro para prosseguir!"
        msg.Rows.Add(row_C368409)
        C368409DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C371832

    Comp_C368438:

        ' RET8
        sCurrComponent = "RET8"
        Dim tb_C368438 As DataTable = New DataTable()
        tb_C368438.Columns.Add("Tem RepresCliente?" & "")
        Dim row_C368438 As DataRow = tb_C368438.NewRow()
        row_C368438(0) = C368408
        tb_C368438.Rows.Add(row_C368438)
        R4557 = tb_C368438
        ReDim C368438DataType(1)
        C368438DataType(1) = C368408DataType
        ReturnDataType = C368438DataType

        GoTo Exit_R4557

    Comp_C368898:

        ' AnalisaCredPrePed
        sCurrComponent = "AnalisaCredPrePed"
        C368898DataType = 0
        C368898DataType = C56123Datatype(27)
        If C368898DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(26)) Then
          C368898 = Strings.RTrim(RsC56123(26))
        Else 
          C368898 = RsC56123(26)
        End If 

        GoTo Comp_C368899

    Comp_C368899:

        ' AnalisaCredPrePed?
        sCurrComponent = "AnalisaCredPrePed?"
        C368899 = (fn_ConvertValueCompiled(C368898, C368898DataType, AKB_DecimalPoint, False)  = 1)
        C368899DataType = AKBTypeConst.cAKBTypeLogical
        If C368899 Then
            GoTo Comp_C371802
        Else
            GoTo Comp_C315473
        End If

    Comp_C371802:

        ' AmostraGrátis?
        sCurrComponent = "AmostraGrátis?"
        C371802 = (fn_ConvertValueCompiled(C56920, C56920DataType, AKB_DecimalPoint, False)  = 201)
        C371802DataType = AKBTypeConst.cAKBTypeLogical
        If C371802 Then
            GoTo Comp_C371804
        Else
            GoTo Comp_C532561
        End If

    Comp_C371804:

        ' AtualLiberadoAmostraGratis
        sCurrComponent = "AtualLiberadoAmostraGratis"
        QueryC371804 = con.CreateCommand()
        QueryC371804.CommandText = QueryC371804.CommandText & " " & "UPDATE WF_PRE_PEDIDO"
        QueryC371804.CommandText = QueryC371804.CommandText & " " & ""
        QueryC371804.CommandText = QueryC371804.CommandText & " " & "SET WF_PRE_PEDIDO.ANALISE_CREDITO  = 'Liberado',"
        QueryC371804.CommandText = QueryC371804.CommandText & " " & ""
        QueryC371804.CommandText = QueryC371804.CommandText & " " & "  WF_PRE_PEDIDO.DATA_MANUT_CRED    = " & _ 
ConvertValue(C174615, C174615DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC371804.CommandText = QueryC371804.CommandText & " " & ""
        QueryC371804.CommandText = QueryC371804.CommandText & " " & "  WF_PRE_PEDIDO.USUARIO_MANUT_CRED =" & _ 
ConvertValue(C265482, C265482DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC371804.CommandText = QueryC371804.CommandText & " " & ""
        QueryC371804.CommandText = QueryC371804.CommandText & " " & "WHERE WF_PRE_PEDIDO.ID_PREPEDIDO   = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC371804.Transaction = txn
        C371804 = QueryC371804.ExecuteNonQuery()
        C371804DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C371818

    Comp_C371816:

        ' vStatusAnalise
        sCurrComponent = "vStatusAnalise"
        C371816 = System.DBNull.Value
        C371816DataType = 3
        GoTo Comp_C371817

    Comp_C371817:

        ' AtrVlrStatusAnalise
        sCurrComponent = "AtrVlrStatusAnalise"
        C371817DataType = 4
        C371816 = fn_ConvertValueCompiled(C367574 , 3, AKB_DecimalPoint)
        C371817 = True
        GoTo Comp_C367575

    Comp_C371818:

        ' SelStatusAnálise
        sCurrComponent = "SelStatusAnálise"
        QueryC371818 = con.CreateCommand()
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "SELECT "
        QueryC371818.CommandText = QueryC371818.CommandText & " " & ""
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "TRIM (DECODE(WF_PRE_PEDIDO.ANALISE_CREDITO,NULL,DECODE("
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "  (SELECT WF_PRE_PEDIDO.DATA_MANUT_CRED FROM AKBUSER01.WF_PRE_PEDIDO  "
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "WHERE WF_PRE_PEDIDO.DATA_MANUT_CRED > ='01/08/2012'"
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "AND  WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "    ),NULL,'Crédito em Análise', 1), WF_PRE_PEDIDO.ANALISE_CREDITO)) StatusAnalise"
        QueryC371818.CommandText = QueryC371818.CommandText & " " & ""
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO , AKBUSER01.WF_TABELA_PRECO_VENDA, AKBUSER01.WF_TAB_PRECO_CUSTO, AKBUSER01.WF_ESTAB_JURIDICO_PARAM "
        QueryC371818.CommandText = QueryC371818.CommandText & " " & ""
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "AND        WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico  = WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico "
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "AND        WF_PRE_PEDIDO.Tabela_Preco_Venda = WF_TABELA_PRECO_VENDA.Tabela_Preco_Venda (+)"
        QueryC371818.CommandText = QueryC371818.CommandText & " " & "AND        WF_PRE_PEDIDO.Tabela_Preco_Custo = WF_TAB_PRECO_CUSTO.Tabela_Preco_Custo (+)"
        QueryC371818.Transaction = txn
        RsC371818 = QueryC371818.ExecuteReader()
        Dim iC371818 As Short
        ReDim C371818DataType(RsC371818.FieldCount)
        For iC371818 = 0 to RsC371818.FieldCount - 1
            Select Case RsC371818.GetDataTypeName(iC371818).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C371818DataType(iC371818 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C371818DataType(iC371818 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C371818DataType(iC371818 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC371818
        RsC371818_EOF = Not RsC371818.Read()

        GoTo Comp_C371819

    Comp_C371819:

        ' AtrVlrStatus
        sCurrComponent = "AtrVlrStatus"
        C371819DataType = 4
        C371816 = fn_ConvertValueCompiled(RsC371818(0) , 3, AKB_DecimalPoint)
        C371819 = True
        GoTo Comp_C367579

    Comp_C371832:

        ' RolbackPedido
        sCurrComponent = "RolbackPedido"
        txn.Rollback()
        C371832 = True
        C371832DataType = AKBTypeConst.cAKBTypeLogical
        GoTo Comp_C368438

    Comp_C372920:

        ' ReservaPedProntaEntrega
        sCurrComponent = "ReservaPedProntaEntrega"
        'C372920 = clsRuleDynamicallyCompiled_R15662.R15662(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(P56293), P56293, System.DBNull.Value), System.DBNull.Value, System.DBNull.Value, System.DBNull.Value)
        C372920CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C372920) Then
          iColumns = C372920.Columns.Count
        End If
        ReDim C372920DataType(iColumns)
        For iCol = 1 To iColumns
            'C372920DataType(iCol) = clsRuleDynamicallyCompiled_R15662.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C109297

    Comp_C374705:

        ' Valida Cadastro Cliente
        sCurrComponent = "Valida Cadastro Cliente"
        C374705 = clsRuleDynamicallyCompiled_R17116.R17116(con, txn, IIf(Not IsDBNull(C75153), C75153, System.DBNull.Value), fn_ConvertValueCompiled( "S", 3, AKB_DecimalPoint))
        C374705CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C374705) Then
          iColumns = C374705.Columns.Count
        End If
        ReDim C374705DataType(iColumns)
        For iCol = 1 To iColumns
            C374705DataType(iCol) = clsRuleDynamicallyCompiled_R17116.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C374707

    Comp_C374707:

        ' Cadastro Cliente Incompleto?
        sCurrComponent = "Cadastro Cliente Incompleto?"
        C374707 = (fn_ConvertValueCompiled(C374705.rows(C374705CurrentRow - 1)(0), C374705DataType(1), AKB_DecimalPoint, False) = 0)
        C374707DataType = AKBTypeConst.cAKBTypeLogical
        If C374707 Then
            GoTo Comp_C96668
        Else
            GoTo Comp_C252514
        End If

    Comp_C374708:

        ' Valida Cadastro Transportadora
        sCurrComponent = "Valida Cadastro Transportadora"
        C374708 = clsRuleDynamicallyCompiled_R17116.R17116(con, txn, IIf(Not IsDBNull(C57928), C57928, System.DBNull.Value), fn_ConvertValueCompiled( "S", 3, AKB_DecimalPoint))
        C374708CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C374708) Then
          iColumns = C374708.Columns.Count
        End If
        ReDim C374708DataType(iColumns)
        For iCol = 1 To iColumns
            C374708DataType(iCol) = clsRuleDynamicallyCompiled_R17116.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C374709

    Comp_C374709:

        ' Cadastro Transp Incompleto?
        sCurrComponent = "Cadastro Transp Incompleto?"
        C374709 = (fn_ConvertValueCompiled(C374708.rows(C374708CurrentRow - 1)(0), C374708DataType(1), AKB_DecimalPoint, False) = 0)
        C374709DataType = AKBTypeConst.cAKBTypeLogical
        If C374709 Then
            GoTo Comp_C97964
        Else
            GoTo Comp_C56907
        End If

    Comp_C378418:

        ' Atualiza desconto Pedido Seq - PUD
        sCurrComponent = "Atualiza desconto Pedido Seq - PUD"
        C378418 = clsRuleDynamicallyCompiled_R17223.R17223(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), fn_ConvertValueCompiled( "1", 1, AKB_DecimalPoint))
        C378418CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C378418) Then
          iColumns = C378418.Columns.Count
        End If
        ReDim C378418DataType(iColumns)
        For iCol = 1 To iColumns
            C378418DataType(iCol) = clsRuleDynamicallyCompiled_R17223.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C405241

    Comp_C381761:

        ' Sem Transportadora?
        sCurrComponent = "Sem Transportadora?"
        C381761 = (fn_ConvertValueCompiled(C57928, C57928DataType, AKB_DecimalPoint, False) = 0 OR fn_ConvertValueCompiled(C381763, C381763DataType, AKB_DecimalPoint, False) < 1)
        C381761DataType = AKBTypeConst.cAKBTypeLogical
        If C381761 Then
            GoTo Comp_C381762
        Else
            GoTo Comp_C384492
        End If

    Comp_C381762:

        ' MSG3
        sCurrComponent = "MSG3"
        Dim row_C381762 As DataRow = msg.NewRow()
        row_C381762(0) = "O Pedido não poderá ser gerado, pois não existe transportadora associada." & Chr(13) & "" & Chr(10) & "Favor verificar." & Chr(13) & "" & Chr(10) & "Operação cancelada!"
        msg.Rows.Add(row_C381762)
        C381762DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C97964

    Comp_C381763:

        ' ComprTransp
        sCurrComponent = "ComprTransp"
        C381763DataType = 1
        C381763 = Len(C57928 & "")
        GoTo Comp_C381761

    Comp_C384492:

        ' Proforma?
        sCurrComponent = "Proforma?"
        C384492 = (fn_ConvertValueCompiled(P26981, 4, AKB_DecimalPoint, False) =1)
        C384492DataType = AKBTypeConst.cAKBTypeLogical
        If C384492 Then
            GoTo Comp_C56907
        Else
            GoTo Comp_C477037
        End If

    Comp_C386925:

        ' Atualiza Representante de Cliente Disponível
        sCurrComponent = "Atualiza Representante de Cliente Disponível"
        C386925 = clsRuleDynamicallyCompiled_R17526.R17526(con, txn, IIf(Not IsDBNull(C75153), C75153, System.DBNull.Value), IIf(Not IsDBNull(C92650), C92650, System.DBNull.Value), IIf(Not IsDBNull(C75154), C75154, System.DBNull.Value))
        C386925CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C386925) Then
          iColumns = C386925.Columns.Count
        End If
        ReDim C386925DataType(iColumns)
        For iCol = 1 To iColumns
            C386925DataType(iCol) = clsRuleDynamicallyCompiled_R17526.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C123427

    Comp_C387531:

        ' Atualiza tipo de pedido parametrizado
        sCurrComponent = "Atualiza tipo de pedido parametrizado"
        C387531 = clsRuleDynamicallyCompiled_R17504.R17504(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value))
        C387531CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C387531) Then
          iColumns = C387531.Columns.Count
        End If
        ReDim C387531DataType(iColumns)
        For iCol = 1 To iColumns
            C387531DataType(iCol) = clsRuleDynamicallyCompiled_R17504.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C176563

    Comp_C391471:

        ' Verifica Coleção dos Itens do Pedido = 1?
        sCurrComponent = "Verifica Coleção dos Itens do Pedido = 1?"
        C391471 = (fn_ConvertValueCompiled(C367790.rows(C367790CurrentRow - 1)(0), C367790DataType(1), AKB_DecimalPoint, False) = 1)
        C391471DataType = AKBTypeConst.cAKBTypeLogical
        If C391471 Then
            GoTo Comp_C387531
        Else
            GoTo Comp_C97964
        End If

    Comp_C391632:

        ' Atualiza Prazo Entrega conforme o Tipo de Pedido
        sCurrComponent = "Atualiza Prazo Entrega conforme o Tipo de Pedido"
        C391632 = clsRuleDynamicallyCompiled_R17770.R17770(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value))
        C391632CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C391632) Then
          iColumns = C391632.Columns.Count
        End If
        ReDim C391632DataType(iColumns)
        For iCol = 1 To iColumns
            C391632DataType(iCol) = clsRuleDynamicallyCompiled_R17770.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C467451

    Comp_C392962:

        ' Gravar Codigo Produto NF-e
        sCurrComponent = "Gravar Codigo Produto NF-e"
        C392962 = clsRuleDynamicallyCompiled_R17706.R17706(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value))
        C392962CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C392962) Then
          iColumns = C392962.Columns.Count
        End If
        ReDim C392962DataType(iColumns)
        For iCol = 1 To iColumns
            C392962DataType(iCol) = clsRuleDynamicallyCompiled_R17706.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C394392

    Comp_C394307:

        ' Atualiza Coleção no Pré-Pedido
        sCurrComponent = "Atualiza Coleção no Pré-Pedido"
        C394307 = clsRuleDynamicallyCompiled_R17756.R17756(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value))
        C394307CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C394307) Then
          iColumns = C394307.Columns.Count
        End If
        ReDim C394307DataType(iColumns)
        For iCol = 1 To iColumns
            C394307DataType(iCol) = clsRuleDynamicallyCompiled_R17756.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C54110

    Comp_C394339:

        ' ParamTab
        sCurrComponent = "ParamTab"
        QueryC394339 = con.CreateCommand()
        QueryC394339.CommandText = QueryC394339.CommandText & " " & "SELECT WF_TP_PRODUTO.Utiliza_Prazo_Param_Tab_Preco, WF_PRE_PEDIDO.Cod_Cliente"
        QueryC394339.CommandText = QueryC394339.CommandText & " " & "FROM WF_PRE_PEDIDO, WF_TP_PRODUTO"
        QueryC394339.CommandText = QueryC394339.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC394339.CommandText = QueryC394339.CommandText & " " & "AND WF_PRE_PEDIDO.Tp_Produto = WF_TP_PRODUTO.Tp_Produto"
        QueryC394339.Transaction = txn
        RsC394339 = QueryC394339.ExecuteReader()
        Dim iC394339 As Short
        ReDim C394339DataType(RsC394339.FieldCount)
        For iC394339 = 0 to RsC394339.FieldCount - 1
            Select Case RsC394339.GetDataTypeName(iC394339).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C394339DataType(iC394339 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C394339DataType(iC394339 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C394339DataType(iC394339 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC394339
        RsC394339_EOF = Not RsC394339.Read()

        GoTo Comp_C394341

    Comp_C394340:

        ' ParamTab=0?
        sCurrComponent = "ParamTab=0?"
        C394340 = (fn_ConvertValueCompiled(C394341, C394341DataType, AKB_DecimalPoint, False) = 0)
        C394340DataType = AKBTypeConst.cAKBTypeLogical
        If C394340 Then
            GoTo Comp_C367999
        Else
            GoTo Comp_C394368
        End If

    Comp_C394341:

        ' vParamTab
        sCurrComponent = "vParamTab"
        C394341DataType = 0
        C394341 = RsC394339(0)
        C394341DataType = C394339Datatype(1)
        If C394341DataType = AKBTypeConst.cAKBTypeString Then
          C394341 = IIF(IsDBNull(C394341), C394341, Strings.RTrim(C394341))
        End If 

        GoTo Comp_C546207

    Comp_C394342:

        ' vCodCliPP
        sCurrComponent = "vCodCliPP"
        C394342DataType = 0
        C394342DataType = C394339Datatype(2)
        If C394342DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC394339(1)) Then
          C394342 = Strings.RTrim(RsC394339(1))
        Else 
          C394342 = RsC394339(1)
        End If 

        GoTo Comp_C394340

    Comp_C394368:

        ' ProdExc
        sCurrComponent = "ProdExc"
        'C394368 = clsRuleDynamicallyCompiled_R17758.R17758(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value), IIf(Not IsDBNull(C394342), C394342, System.DBNull.Value))
        C394368CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C394368) Then
          iColumns = C394368.Columns.Count
        End If
        ReDim C394368DataType(iColumns)
        For iCol = 1 To iColumns
            'C394368DataType(iCol) = clsRuleDynamicallyCompiled_R17758.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C394371

    Comp_C394369:

        ' ProdExc=1?
        sCurrComponent = "ProdExc=1?"
        C394369 = (fn_ConvertValueCompiled(C394371, C394371DataType, AKB_DecimalPoint, False) = 1)
        C394369DataType = AKBTypeConst.cAKBTypeLogical
        If C394369 Then
            GoTo Comp_C367999
        Else
            GoTo Comp_C394370
        End If

    Comp_C394370:

        ' RET9
        sCurrComponent = "RET9"
        Dim tb_C394370 As DataTable = New DataTable()
        tb_C394370.Columns.Add("attProdExc" & "")
        Dim row_C394370 As DataRow = tb_C394370.NewRow()
        row_C394370(0) = C394371
        tb_C394370.Rows.Add(row_C394370)
        R4557 = tb_C394370
        ReDim C394370DataType(1)
        C394370DataType(1) = C394371DataType
        ReturnDataType = C394370DataType

        GoTo Exit_R4557

    Comp_C394371:

        ' attProdExc
        sCurrComponent = "attProdExc"
        C394371DataType = 4
        C394368.rows(C394368CurrentRow - 1)(0) = System.DBNull.Value
        C394371 = True
        GoTo Comp_C394369

    Comp_C394392:

        ' GravaProdExc
        sCurrComponent = "GravaProdExc"
        C394392 = clsRuleDynamicallyCompiled_R17759.R17759(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value), IIf(Not IsDBNull(C394342), C394342, System.DBNull.Value))
        C394392CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C394392) Then
          iColumns = C394392.Columns.Count
        End If
        ReDim C394392DataType(iColumns)
        For iCol = 1 To iColumns
            C394392DataType(iCol) = clsRuleDynamicallyCompiled_R17759.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C524142

    Comp_C405241:

        ' Proforma2?
        sCurrComponent = "Proforma2?"
        C405241 = (fn_ConvertValueCompiled(P26981, 4, AKB_DecimalPoint, False) = 1 OR fn_ConvertValueCompiled(RsC419420(0), C419420DataType(1), AKB_DecimalPoint, False) > 0)
        C405241DataType = AKBTypeConst.cAKBTypeLogical
        If C405241 Then
            GoTo Comp_C387531
        Else
            GoTo Comp_C367790
        End If

    Comp_C412007:

        ' upPrazoIndustrialFinalSemana
        sCurrComponent = "upPrazoIndustrialFinalSemana"
        QueryC412007 = con.CreateCommand()
        QueryC412007.CommandText = QueryC412007.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC412007.CommandText = QueryC412007.CommandText & " " & "SET WF_PEDIDO_ITENS.Prazo_Entrega = NEXT_DAY(WF_PEDIDO_ITENS.Prazo_Entrega,2) "
        QueryC412007.CommandText = QueryC412007.CommandText & " " & "WHERE TO_NUMBER(TO_CHAR(WF_PEDIDO_ITENS.Prazo_Entrega,'D')) IN (1,7)"
        QueryC412007.CommandText = QueryC412007.CommandText & " " & "AND (WF_PEDIDO_ITENS.Tp_Produto , WF_PEDIDO_ITENS.Pedido) IN"
        QueryC412007.CommandText = QueryC412007.CommandText & " " & "(SELECT WF_PRE_PEDIDO.Tp_Produto  , WF_PRE_PEDIDO.Pedido"
        QueryC412007.CommandText = QueryC412007.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC412007.CommandText = QueryC412007.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC412007.Transaction = txn
        C412007 = QueryC412007.ExecuteNonQuery()
        C412007DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C414371

    Comp_C412045:

        ' upQtdeMinMaxItemEstoque
        sCurrComponent = "upQtdeMinMaxItemEstoque"
        QueryC412045 = con.CreateCommand()
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "SET WF_PEDIDO_ITENS.Qtde_Min_Item_Unid1 = "
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "(SELECT WF_CLIENTES.Qtde_Min_Item_Estoque"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO, AKBUSER01.WF_CLIENTES"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "WHERE WF_PRE_PEDIDO.Cod_Cliente = WF_CLIENTES.Cod_Cliente"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "), "
        QueryC412045.CommandText = QueryC412045.CommandText & " " & ""
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "WF_PEDIDO_ITENS.Quant_Max_Item_Unid1 = "
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "(SELECT WF_CLIENTES.Qtde_Max_Item_Estoque"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO, AKBUSER01.WF_CLIENTES"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "WHERE WF_PRE_PEDIDO.Cod_Cliente = WF_CLIENTES.Cod_Cliente"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & ""
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "WHERE (WF_PEDIDO_ITENS.Tp_Produto , WF_PEDIDO_ITENS.Pedido) IN"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "(SELECT WF_PRE_PEDIDO.Tp_Produto  , WF_PRE_PEDIDO.Pedido"
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC412045.CommandText = QueryC412045.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC412045.Transaction = txn
        C412045 = QueryC412045.ExecuteNonQuery()
        C412045DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C378418

    Comp_C414346:

        ' CodCliFinal
        sCurrComponent = "CodCliFinal"
        C414346DataType = 0
        C414346DataType = C56123Datatype(28)
        If C414346DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(27)) Then
          C414346 = Strings.RTrim(RsC56123(27))
        Else 
          C414346 = RsC56123(27)
        End If 

        GoTo Comp_C474187

    Comp_C414347:

        ' CountRepCli
        sCurrComponent = "CountRepCli"
        C414347 = System.DBNull.Value
        C414347DataType = 1
        GoTo Comp_C265483

    Comp_C414348:

        ' CliFinalNULL?
        sCurrComponent = "CliFinalNULL?"
        C414348 = (fn_ConvertValueCompiled(C414346, C414346DataType, AKB_DecimalPoint, False) Is System.DBNull.Value)
        C414348DataType = AKBTypeConst.cAKBTypeLogical
        If C414348 Then
            GoTo Comp_C368405
        Else
            GoTo Comp_C414349
        End If

    Comp_C414349:

        ' SelRepCliFim
        sCurrComponent = "SelRepCliFim"
        QueryC414349 = con.CreateCommand()
        QueryC414349.CommandText = QueryC414349.CommandText & " " & "SELECT COUNT(*) "
        QueryC414349.CommandText = QueryC414349.CommandText & " " & "FROM AKBUSER01.WF_REPRESENTANTE_CLIENTE "
        QueryC414349.CommandText = QueryC414349.CommandText & " " & "WHERE WF_REPRESENTANTE_CLIENTE.Cod_Repres = " & _ 
ConvertValue(C92650, C92650DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC414349.CommandText = QueryC414349.CommandText & " " & "AND WF_REPRESENTANTE_CLIENTE.Cod_Cliente =" & _ 
ConvertValue(C414346, C414346DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC414349.Transaction = txn
        RsC414349 = QueryC414349.ExecuteReader()
        Dim iC414349 As Short
        ReDim C414349DataType(RsC414349.FieldCount)
        For iC414349 = 0 to RsC414349.FieldCount - 1
            Select Case RsC414349.GetDataTypeName(iC414349).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C414349DataType(iC414349 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C414349DataType(iC414349 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C414349DataType(iC414349 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC414349
        RsC414349_EOF = Not RsC414349.Read()

        GoTo Comp_C414351

    Comp_C414350:

        ' ATRIBUIVA1
        sCurrComponent = "ATRIBUIVA1"
        C414350DataType = 4
        C414347 = fn_ConvertValueCompiled(RsC368405(0) , 1, AKB_DecimalPoint)
        C414350 = True
        GoTo Comp_C368408

    Comp_C414351:

        ' ATRIBUIVA2
        sCurrComponent = "ATRIBUIVA2"
        C414351DataType = 4
        C414347 = fn_ConvertValueCompiled(RsC414349(0) , 1, AKB_DecimalPoint)
        C414351 = True
        GoTo Comp_C368408

    Comp_C414371:

        ' upPrazoComercialFinalSemana
        sCurrComponent = "upPrazoComercialFinalSemana"
        QueryC414371 = con.CreateCommand()
        QueryC414371.CommandText = QueryC414371.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS "
        QueryC414371.CommandText = QueryC414371.CommandText & " " & "SET WF_PEDIDO_ITENS.Prazo_Entrega_Comercial = NEXT_DAY(WF_PEDIDO_ITENS.Prazo_Entrega_Comercial,2) "
        QueryC414371.CommandText = QueryC414371.CommandText & " " & "WHERE TO_NUMBER(TO_CHAR(WF_PEDIDO_ITENS.Prazo_Entrega_Comercial,'D')) IN (1,7)"
        QueryC414371.CommandText = QueryC414371.CommandText & " " & "AND (WF_PEDIDO_ITENS.Tp_Produto , WF_PEDIDO_ITENS.Pedido) IN"
        QueryC414371.CommandText = QueryC414371.CommandText & " " & "(SELECT WF_PRE_PEDIDO.Tp_Produto  , WF_PRE_PEDIDO.Pedido"
        QueryC414371.CommandText = QueryC414371.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC414371.CommandText = QueryC414371.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC414371.Transaction = txn
        C414371 = QueryC414371.ExecuteNonQuery()
        C414371DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C412045

    Comp_C418089:

        ' Tem Depto?
        sCurrComponent = "Tem Depto?"
        C418089 = (fn_ConvertValueCompiled(C418178, C418178DataType, AKB_DecimalPoint, False) > 0)
        C418089DataType = AKBTypeConst.cAKBTypeLogical
        If C418089 Then
            GoTo Comp_C310322
        Else
            GoTo Comp_C418090
        End If

    Comp_C418090:

        ' MSG4
        sCurrComponent = "MSG4"
        Dim row_C418090 As DataRow = msg.NewRow()
        row_C418090(0) = "Não foi selecionado departamento responsável pela venda no pré-pedido." & Chr(13) & "" & Chr(10) & "O pedido " & _ 
C56908 & " não pode ser gerado." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Operação cancelada."
        msg.Rows.Add(row_C418090)
        C418090DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C418179

    Comp_C418177:

        ' selDepto
        sCurrComponent = "selDepto"
        QueryC418177 = con.CreateCommand()
        QueryC418177.CommandText = QueryC418177.CommandText & " " & "SELECT NVL(WF_PRE_PEDIDO.Departamento," & _ 
ConvertValue(P68164, 3, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC418177.CommandText = QueryC418177.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC418177.CommandText = QueryC418177.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC418177.Transaction = txn
        RsC418177 = QueryC418177.ExecuteReader()
        Dim iC418177 As Short
        ReDim C418177DataType(RsC418177.FieldCount)
        For iC418177 = 0 to RsC418177.FieldCount - 1
            Select Case RsC418177.GetDataTypeName(iC418177).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C418177DataType(iC418177 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C418177DataType(iC418177 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C418177DataType(iC418177 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC418177
        RsC418177_EOF = Not RsC418177.Read()

        GoTo Comp_C418191

    Comp_C418178:

        ' comprDepto
        sCurrComponent = "comprDepto"
        C418178DataType = 1
        C418178 = Len(C418191 & "")
        GoTo Comp_C418089

    Comp_C418179:

        ' RET10
        sCurrComponent = "RET10"
        Dim tb_C418179 As DataTable = New DataTable()
        tb_C418179.Columns.Add("Tem Depto?" & "")
        Dim row_C418179 As DataRow = tb_C418179.NewRow()
        row_C418179(0) = C418089
        tb_C418179.Rows.Add(row_C418179)
        R4557 = tb_C418179
        ReDim C418179DataType(1)
        C418179DataType(1) = C418089DataType
        ReturnDataType = C418179DataType

        GoTo Exit_R4557

    Comp_C418180:

        ' upDeptoPrePedido
        sCurrComponent = "upDeptoPrePedido"
        QueryC418180 = con.CreateCommand()
        QueryC418180.CommandText = QueryC418180.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO"
        QueryC418180.CommandText = QueryC418180.CommandText & " " & "SET WF_PRE_PEDIDO.Departamento = " & _ 
ConvertValue(C418191, C418191DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC418180.CommandText = QueryC418180.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC418180.CommandText = QueryC418180.CommandText & " " & "AND WF_PRE_PEDIDO.Departamento IS NULL"
        QueryC418180.Transaction = txn
        C418180 = QueryC418180.ExecuteNonQuery()
        C418180DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C174615

    Comp_C418181:

        ' upDeptoObs
        sCurrComponent = "upDeptoObs"
        QueryC418181 = con.CreateCommand()
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_OBS"
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "SET WF_PRE_PEDIDO_OBS.Departamento = " & _ 
ConvertValue(C418191, C418191DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "WHERE WF_PRE_PEDIDO_OBS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "AND WF_PRE_PEDIDO_OBS.Departamento = 'COM00'"
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "AND EXISTS"
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "(SELECT *"
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO"
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = WF_PRE_PEDIDO_OBS.Id_PrePedido"
        QueryC418181.CommandText = QueryC418181.CommandText & " " & "AND WF_PRE_PEDIDO.Departamento IS NULL)"
        QueryC418181.Transaction = txn
        C418181 = QueryC418181.ExecuteNonQuery()
        C418181DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C418180

    Comp_C418182:

        ' Depto
        sCurrComponent = "Depto"
        C418182DataType = 0
        C418182DataType = C57925Datatype(6)
        If C418182DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC57925(5)) Then
          C418182 = Strings.RTrim(RsC57925(5))
        Else 
          C418182 = RsC57925(5)
        End If 

        GoTo Comp_C57932

    Comp_C418191:

        ' vDeptoPrePedido
        sCurrComponent = "vDeptoPrePedido"
        C418191DataType = 0
        C418191 = RsC418177(0)
        C418191DataType = C418177Datatype(1)
        If C418191DataType = AKBTypeConst.cAKBTypeString Then
          C418191 = IIF(IsDBNull(C418191), C418191, Strings.RTrim(C418191))
        End If 

        GoTo Comp_C418178

    Comp_C419420:

        ' SelTabPreçoIgualEstabParam
        sCurrComponent = "SelTabPreçoIgualEstabParam"
        QueryC419420 = con.CreateCommand()
        QueryC419420.CommandText = QueryC419420.CommandText & " " & "SELECT COUNT(WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico) "
        QueryC419420.CommandText = QueryC419420.CommandText & " " & ""
        QueryC419420.CommandText = QueryC419420.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO , AKBUSER01.WF_PRE_PEDIDO "
        QueryC419420.CommandText = QueryC419420.CommandText & " " & "WHERE WF_ESTAB_JURIDICO.Tabela_Preco_Venda = WF_PRE_PEDIDO.Tabela_Preco_Venda "
        QueryC419420.CommandText = QueryC419420.CommandText & " " & "      AND  WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico = WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico"
        QueryC419420.CommandText = QueryC419420.CommandText & " " & "       AND  WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC419420.Transaction = txn
        RsC419420 = QueryC419420.ExecuteReader()
        Dim iC419420 As Short
        ReDim C419420DataType(RsC419420.FieldCount)
        For iC419420 = 0 to RsC419420.FieldCount - 1
            Select Case RsC419420.GetDataTypeName(iC419420).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C419420DataType(iC419420 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C419420DataType(iC419420 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C419420DataType(iC419420 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC419420
        RsC419420_EOF = Not RsC419420.Read()

        GoTo Comp_C174609

    Comp_C419441:

        ' TabPreçoVdaPrePedido = TabPreçoVdaEstab
        sCurrComponent = "TabPreçoVdaPrePedido = TabPreçoVdaEstab"
        C419441 = (fn_ConvertValueCompiled(RsC419420(0), C419420DataType(1), AKB_DecimalPoint, False) > 0)
        C419441DataType = AKBTypeConst.cAKBTypeLogical
        If C419441 Then
            GoTo Comp_C419442
        Else
            GoTo Comp_C283843
        End If

    Comp_C419442:

        ' UpdateColeçãoProdutoNOP
        sCurrComponent = "UpdateColeçãoProdutoNOP"
        QueryC419442 = con.CreateCommand()
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO_ITENS "
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "SET WF_PRE_PEDIDO_ITENS.Id_Colecao = "
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "(SELECT Max(WF_COLECAO_PRODUTOS.Id_Colecao )"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "FROM AKBUSER01.WF_COLECAO_PRODUTOS"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "WHERE WF_COLECAO_PRODUTOS.Sigla_Prod = WF_PRE_PEDIDO_ITENS.Sigla_Prod "
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Acesso = WF_PRE_PEDIDO_ITENS.Acesso"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "AND WF_COLECAO_PRODUTOS.Cod_Embalagem = WF_PRE_PEDIDO_ITENS.Cod_Embalagem  ),"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "            "
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "  WF_PRE_PEDIDO_ITENS.Acesso2 = WF_PRE_PEDIDO_ITENS.Acesso,"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "  WF_PRE_PEDIDO_ITENS.Sigla_Prod2 = WF_PRE_PEDIDO_ITENS.Sigla_Prod,"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "  WF_PRE_PEDIDO_ITENS.Cod_Embalagem2 = WF_PRE_PEDIDO_ITENS.Cod_Embalagem"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & ""
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "WHERE WF_PRE_PEDIDO_ITENS.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "AND ( WF_PRE_PEDIDO_ITENS.Id_Colecao IS NULL OR WF_PRE_PEDIDO_ITENS.Id_Colecao = 0 )"
        QueryC419442.CommandText = QueryC419442.CommandText & " " & "AND WF_PRE_PEDIDO_ITENS.Acesso NOT IN (" & _ 
ConvertValue(C283841, C283841DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ")"
        QueryC419442.Transaction = txn
        C419442 = QueryC419442.ExecuteNonQuery()
        C419442DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C92633

    Comp_C420246:

        ' TABPREÇOVDAPREPEDIDO = TABPREÇOVDAESTAB1
        sCurrComponent = "TABPREÇOVDAPREPEDIDO = TABPREÇOVDAESTAB1"
        C420246 = (fn_ConvertValueCompiled(RsC419420(0), C419420DataType(1), AKB_DecimalPoint, False) > 0)
        C420246DataType = AKBTypeConst.cAKBTypeLogical
        If C420246 Then
            GoTo Comp_C54110
        Else
            GoTo Comp_C394307
        End If

    Comp_C423441:

        ' TemTranspEntrega?
        sCurrComponent = "TemTranspEntrega?"
        C423441 = (fn_ConvertValueCompiled(C439225, C439225DataType, AKB_DecimalPoint, False) >0 AND fn_ConvertValueCompiled(C439144, C439144DataType, AKB_DecimalPoint, False) = 0 AND fn_ConvertValueCompiled(C439145, C439145DataType, AKB_DecimalPoint, False) = 0)
        C423441DataType = AKBTypeConst.cAKBTypeLogical
        If C423441 Then
            GoTo Comp_C425563
        Else
            GoTo Comp_C54111
        End If

    Comp_C423442:

        ' SelEntrega
        sCurrComponent = "SelEntrega"
        QueryC423442 = con.CreateCommand()
        QueryC423442.CommandText = QueryC423442.CommandText & " " & "SELECT NVL(Cod_Cliente_Entrega,0), NVL(Cod_Transp_Entrega,0), nvl(Cod_Transp,0), nvl(Cod_Frete, 0),"
        QueryC423442.CommandText = QueryC423442.CommandText & " " & " nvl(Cod_Frete_Trans_Entreg,0), nvl(Cod_Redesp_Entrega,0), nvl(Cod_Frete_Redesp_Entrega,0), nvl(Cod_Redespacho,0), nvl(Cod_Frete_Redesp,0)"
        QueryC423442.CommandText = QueryC423442.CommandText & " " & "FROM WF_PRE_PEDIDO"
        QueryC423442.CommandText = QueryC423442.CommandText & " " & "WHERE Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC423442.Transaction = txn
        RsC423442 = QueryC423442.ExecuteReader()
        Dim iC423442 As Short
        ReDim C423442DataType(RsC423442.FieldCount)
        For iC423442 = 0 to RsC423442.FieldCount - 1
            Select Case RsC423442.GetDataTypeName(iC423442).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C423442DataType(iC423442 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C423442DataType(iC423442 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C423442DataType(iC423442 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC423442
        RsC423442_EOF = Not RsC423442.Read()

        GoTo Comp_C423446

    Comp_C423446:

        ' vCodCliEntrega
        sCurrComponent = "vCodCliEntrega"
        C423446DataType = 0
        C423446 = RsC423442(0)
        C423446DataType = C423442Datatype(1)
        If C423446DataType = AKBTypeConst.cAKBTypeString Then
          C423446 = IIF(IsDBNull(C423446), C423446, Strings.RTrim(C423446))
        End If 

        GoTo Comp_C423448

    Comp_C423448:

        ' vCodTranspEntrega
        sCurrComponent = "vCodTranspEntrega"
        C423448DataType = 0
        C423448DataType = C423442Datatype(2)
        If C423448DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(1)) Then
          C423448 = Strings.RTrim(RsC423442(1))
        Else 
          C423448 = RsC423442(1)
        End If 

        GoTo Comp_C425566

    Comp_C425563:

        ' UpTranspEntrega
        sCurrComponent = "UpTranspEntrega"
        QueryC425563 = con.CreateCommand()
        QueryC425563.CommandText = QueryC425563.CommandText & " " & "UPDATE WF_PEDIDO"
        QueryC425563.CommandText = QueryC425563.CommandText & " " & "SET Cod_Redesp_Entrega = " & _ 
ConvertValue(C439225, C439225DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ,"
        QueryC425563.CommandText = QueryC425563.CommandText & " " & "Cod_Frete_Red_Entr = " & _ 
ConvertValue(C439226, C439226DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC425563.CommandText = QueryC425563.CommandText & " " & "WHERE PEDIDO = " & _ 
ConvertValue(C56908, C56908DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC425563.Transaction = txn
        C425563 = QueryC425563.ExecuteNonQuery()
        C425563DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C54111

    Comp_C425566:

        ' vCodTransp
        sCurrComponent = "vCodTransp"
        C425566DataType = 0
        C425566DataType = C423442Datatype(3)
        If C425566DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(2)) Then
          C425566 = Strings.RTrim(RsC423442(2))
        Else 
          C425566 = RsC423442(2)
        End If 

        GoTo Comp_C438134

    Comp_C432039:

        ' Não altera Preço Unit?
        sCurrComponent = "Não altera Preço Unit?"
        C432039 = (1=1)
        C432039DataType = AKBTypeConst.cAKBTypeLogical
        If C432039 Then
            GoTo Comp_C56910
        Else
            GoTo Comp_C56909
        End If

    Comp_C433219:

        ' IPC
        sCurrComponent = "IPC"
        C433219DataType = 0
        C433219DataType = C56123Datatype(29)
        If C433219DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(28)) Then
          C433219 = Strings.RTrim(RsC56123(28))
        Else 
          C433219 = RsC56123(28)
        End If 

        GoTo Comp_C433221

    Comp_C433221:

        ' vIPC
        sCurrComponent = "vIPC"
        C433221 = C433219 & " "
        C433221DataType = 1
        GoTo Comp_C433222

    Comp_C433222:

        ' vIPC<>0?
        sCurrComponent = "vIPC<>0?"
        C433222 = (fn_ConvertValueCompiled(C433221, C433221DataType, AKB_DecimalPoint, False) <> 0)
        C433222DataType = AKBTypeConst.cAKBTypeLogical
        If C433222 Then
            GoTo Comp_C136195
        Else
            GoTo Comp_C433223
        End If

    Comp_C433223:

        ' IPC_Repres
        sCurrComponent = "IPC_Repres"
        QueryC433223 = con.CreateCommand()
        QueryC433223.CommandText = QueryC433223.CommandText & " " & "SELECT NVL(Cod_Ind_Pres_Comprador,0)"
        QueryC433223.CommandText = QueryC433223.CommandText & " " & "FROM WF_REPRESENTANTE"
        QueryC433223.CommandText = QueryC433223.CommandText & " " & "WHERE Cod_Repres = " & _ 
ConvertValue(C92650, C92650DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC433223.Transaction = txn
        RsC433223 = QueryC433223.ExecuteReader()
        Dim iC433223 As Short
        ReDim C433223DataType(RsC433223.FieldCount)
        For iC433223 = 0 to RsC433223.FieldCount - 1
            Select Case RsC433223.GetDataTypeName(iC433223).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C433223DataType(iC433223 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C433223DataType(iC433223 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C433223DataType(iC433223 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC433223
        RsC433223_EOF = Not RsC433223.Read()

        GoTo Comp_C433225

    Comp_C433225:

        ' ATRIBUIVA3
        sCurrComponent = "ATRIBUIVA3"
        C433225DataType = 4
        C433221 = fn_ConvertValueCompiled(RsC433223(0) , 1, AKB_DecimalPoint)
        C433225 = True
        GoTo Comp_C136195

    Comp_C438134:

        ' vCodFrete
        sCurrComponent = "vCodFrete"
        C438134DataType = 0
        C438134DataType = C423442Datatype(4)
        If C438134DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(3)) Then
          C438134 = Strings.RTrim(RsC423442(3))
        Else 
          C438134 = RsC423442(3)
        End If 

        GoTo Comp_C438135

    Comp_C438135:

        ' vCodFreteTransEntreg
        sCurrComponent = "vCodFreteTransEntreg"
        C438135DataType = 0
        C438135DataType = C423442Datatype(5)
        If C438135DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(4)) Then
          C438135 = Strings.RTrim(RsC423442(4))
        Else 
          C438135 = RsC423442(4)
        End If 

        GoTo Comp_C439144

    Comp_C439144:

        ' vCodTranspEntRedes
        sCurrComponent = "vCodTranspEntRedes"
        C439144DataType = 0
        C439144DataType = C423442Datatype(6)
        If C439144DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(5)) Then
          C439144 = Strings.RTrim(RsC423442(5))
        Else 
          C439144 = RsC423442(5)
        End If 

        GoTo Comp_C439145

    Comp_C439145:

        ' vCodFreteEntrRedes
        sCurrComponent = "vCodFreteEntrRedes"
        C439145DataType = 0
        C439145DataType = C423442Datatype(7)
        If C439145DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(6)) Then
          C439145 = Strings.RTrim(RsC423442(6))
        Else 
          C439145 = RsC423442(6)
        End If 

        GoTo Comp_C439225

    Comp_C439146:

        ' TemTranspEntRedes?
        sCurrComponent = "TemTranspEntRedes?"
        C439146 = (fn_ConvertValueCompiled(C423446, C423446DataType, AKB_DecimalPoint, False) > 0 AND fn_ConvertValueCompiled(C423448, C423448DataType, AKB_DecimalPoint, False)  = 0 AND fn_ConvertValueCompiled(C438135, C438135DataType, AKB_DecimalPoint, False)  = 0)
        C439146DataType = AKBTypeConst.cAKBTypeLogical
        If C439146 Then
            GoTo Comp_C439147
        Else
            GoTo Comp_C423441
        End If

    Comp_C439147:

        ' UpTranspEntRedesp
        sCurrComponent = "UpTranspEntRedesp"
        QueryC439147 = con.CreateCommand()
        QueryC439147.CommandText = QueryC439147.CommandText & " " & "UPDATE WF_PEDIDO"
        QueryC439147.CommandText = QueryC439147.CommandText & " " & "SET COD_TRANSP_ENTREGA = " & _ 
ConvertValue(C425566, C425566DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ", "
        QueryC439147.CommandText = QueryC439147.CommandText & " " & "COD_FRETE_TRANSP_ENTREG = " & _ 
ConvertValue(C438134, C438134DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC439147.CommandText = QueryC439147.CommandText & " " & "Where Pedido = " & _ 
ConvertValue(C56908, C56908DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC439147.Transaction = txn
        C439147 = QueryC439147.ExecuteNonQuery()
        C439147DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C423441

    Comp_C439225:

        ' vTranspRedesp
        sCurrComponent = "vTranspRedesp"
        C439225DataType = 0
        C439225DataType = C423442Datatype(8)
        If C439225DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(7)) Then
          C439225 = Strings.RTrim(RsC423442(7))
        Else 
          C439225 = RsC423442(7)
        End If 

        GoTo Comp_C439226

    Comp_C439226:

        ' vFreteRedesp
        sCurrComponent = "vFreteRedesp"
        C439226DataType = 0
        C439226DataType = C423442Datatype(9)
        If C439226DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC423442(8)) Then
          C439226 = Strings.RTrim(RsC423442(8))
        Else 
          C439226 = RsC423442(8)
        End If 

        GoTo Comp_C439146

    Comp_C467451:

        ' Prazo de entrega atualizado?
        sCurrComponent = "Prazo de entrega atualizado?"
        C467451 = (fn_ConvertValueCompiled(C391632.rows(C391632CurrentRow - 1)(0), C391632DataType(1), AKB_DecimalPoint, False) = 1)
        C467451DataType = AKBTypeConst.cAKBTypeLogical
        If C467451 Then
            GoTo Comp_C556382
        Else
            GoTo Comp_C96585
        End If

    Comp_C470976:

        ' selParamEstab
        sCurrComponent = "selParamEstab"
        QueryC470976 = con.CreateCommand()
        QueryC470976.CommandText = QueryC470976.CommandText & " " & "SELECT WF_ESTAB_JURIDICO_PARAM.Trocar_Estab_Pedido , WF_ESTAB_JURIDICO_PARAM.Bloqueia_Gerar_Ped_Email_Cobr , WF_ESTAB_JURIDICO_PARAM.Cod_Comun_Email_DANFE "
        QueryC470976.CommandText = QueryC470976.CommandText & " " & "FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM"
        QueryC470976.CommandText = QueryC470976.CommandText & " " & "WHERE WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P76804, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC470976.Transaction = txn
        RsC470976 = QueryC470976.ExecuteReader()
        Dim iC470976 As Short
        ReDim C470976DataType(RsC470976.FieldCount)
        For iC470976 = 0 to RsC470976.FieldCount - 1
            Select Case RsC470976.GetDataTypeName(iC470976).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C470976DataType(iC470976 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C470976DataType(iC470976 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C470976DataType(iC470976 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC470976
        RsC470976_EOF = Not RsC470976.Read()

        GoTo Comp_C578759

    Comp_C470977:

        ' Alterar EstabPrepedido?
        sCurrComponent = "Alterar EstabPrepedido?"
        C470977 = (fn_ConvertValueCompiled(P76804, 1, AKB_DecimalPoint, False) <> 0 and fn_ConvertValueCompiled(C470978, C470978DataType, AKB_DecimalPoint, False) = 1)
        C470977DataType = AKBTypeConst.cAKBTypeLogical
        If C470977 Then
            GoTo Comp_C470979
        Else
            GoTo Comp_C482983
        End If

    Comp_C470978:

        ' vTrocarEstab
        sCurrComponent = "vTrocarEstab"
        C470978DataType = 0
        C470978 = RsC470976(0)
        C470978DataType = C470976Datatype(1)
        If C470978DataType = AKBTypeConst.cAKBTypeString Then
          C470978 = IIF(IsDBNull(C470978), C470978, Strings.RTrim(C470978))
        End If 

        GoTo Comp_C516410

    Comp_C470979:

        ' upEstabPrePedido
        sCurrComponent = "upEstabPrePedido"
        QueryC470979 = con.CreateCommand()
        QueryC470979.CommandText = QueryC470979.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO"
        QueryC470979.CommandText = QueryC470979.CommandText & " " & "SET WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P76804, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC470979.CommandText = QueryC470979.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC470979.CommandText = QueryC470979.CommandText & " " & "AND WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico <> " & _ 
ConvertValue(P76804, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC470979.Transaction = txn
        C470979 = QueryC470979.ExecuteNonQuery()
        C470979DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C482983

    Comp_C471477:

        ' GeraPedidoServiçoColigada
        sCurrComponent = "GeraPedidoServiçoColigada"
        'C471477 = clsRuleDynamicallyCompiled_R20307.R20307(con, txn, IIf(Not IsDBNull(C56908), C56908, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C363384), C363384, System.DBNull.Value))
        C471477CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C471477) Then
          iColumns = C471477.Columns.Count
        End If
        ReDim C471477DataType(iColumns)
        For iCol = 1 To iColumns
            'C471477DataType(iCol) = clsRuleDynamicallyCompiled_R20307.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C480959

    Comp_C471989:

        ' Europa?
        sCurrComponent = "Europa?"
        C471989 = ((fn_ConvertValueCompiled(C75154, C75154DataType, AKB_DecimalPoint, False) = 4 OR fn_ConvertValueCompiled(C75154, C75154DataType, AKB_DecimalPoint, False) = 21) AND fn_ConvertValueCompiled(C96561, C96561DataType, AKB_DecimalPoint, False) <> 0)
        C471989DataType = AKBTypeConst.cAKBTypeLogical
        If C471989 Then
            GoTo Comp_C471477
        Else
            GoTo Comp_C477023
        End If

    Comp_C474187:

        ' CodFerVda
        sCurrComponent = "CodFerVda"
        C474187DataType = 0
        C474187DataType = C56123Datatype(30)
        If C474187DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(29)) Then
          C474187 = Strings.RTrim(RsC56123(29))
        Else 
          C474187 = RsC56123(29)
        End If 

        GoTo Comp_C474188

    Comp_C474188:

        ' CodMarcOrig
        sCurrComponent = "CodMarcOrig"
        C474188DataType = 0
        C474188DataType = C56123Datatype(31)
        If C474188DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(30)) Then
          C474188 = Strings.RTrim(RsC56123(30))
        Else 
          C474188 = RsC56123(30)
        End If 

        GoTo Comp_C474189

    Comp_C474189:

        ' CodAcaoMKT
        sCurrComponent = "CodAcaoMKT"
        C474189DataType = 0
        C474189DataType = C56123Datatype(32)
        If C474189DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(31)) Then
          C474189 = Strings.RTrim(RsC56123(31))
        Else 
          C474189 = RsC56123(31)
        End If 

        GoTo Comp_C555428

    Comp_C474763:

        ' InsAcoesMKT
        sCurrComponent = "InsAcoesMKT"
        QueryC474763 = con.CreateCommand()
        QueryC474763.CommandText = QueryC474763.CommandText & " " & "INSERT INTO WF_PEDIDO_ACAO_MKT (Tp_Produto, Pedido, Cod_Acao_Marketing)"
        QueryC474763.CommandText = QueryC474763.CommandText & " " & "SELECT WF_PRE_PEDIDO.Tp_Produto , WF_PRE_PEDIDO.Pedido , WF_PRE_PED_ACOES_MKT.Cod_Acao_Marketing"
        QueryC474763.CommandText = QueryC474763.CommandText & " " & " FROM WF_PRE_PED_ACOES_MKT, WF_PRE_PEDIDO"
        QueryC474763.CommandText = QueryC474763.CommandText & " " & " WHERE WF_PRE_PED_ACOES_MKT.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC474763.CommandText = QueryC474763.CommandText & " " & " AND WF_PRE_PEDIDO.Id_PrePedido = WF_PRE_PED_ACOES_MKT.Id_PrePedido "
        QueryC474763.Transaction = txn
        C474763 = QueryC474763.ExecuteNonQuery()
        C474763DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C386925

    Comp_C477022:

        ' InsereNecessidadeOP
        sCurrComponent = "InsereNecessidadeOP"
        'C477022 = clsRuleDynamicallyCompiled_R20507.R20507(con, txn, IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value), IIf(Not IsDBNull(C56908), C56908, System.DBNull.Value), IIf(Not IsDBNull(C363384), C363384, System.DBNull.Value))
        C477022CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C477022) Then
          iColumns = C477022.Columns.Count
        End If
        ReDim C477022DataType(iColumns)
        For iCol = 1 To iColumns
            'C477022DataType(iCol) = clsRuleDynamicallyCompiled_R20507.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C56141

    Comp_C477023:

        ' FENIX?
        sCurrComponent = "FENIX?"
        C477023 = (fn_ConvertValueCompiled(C75154, C75154DataType, AKB_DecimalPoint, False) = 22 AND fn_ConvertValueCompiled(C96561, C96561DataType, AKB_DecimalPoint, False) <> 0)
        C477023DataType = AKBTypeConst.cAKBTypeLogical
        If C477023 Then
            GoTo Comp_C477022
        Else
            GoTo Comp_C56141
        End If

    Comp_C477036:

        ' FlagOProprio
        sCurrComponent = "FlagOProprio"
        QueryC477036 = con.CreateCommand()
        QueryC477036.CommandText = QueryC477036.CommandText & " " & "SELECT NVL(Transp_o_Proprio,0)"
        QueryC477036.CommandText = QueryC477036.CommandText & " " & "FROM WF_TRANSPORTADORA"
        QueryC477036.CommandText = QueryC477036.CommandText & " " & "WHERE Cod_Transp = " & _ 
ConvertValue(C57928, C57928DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC477036.Transaction = txn
        RsC477036 = QueryC477036.ExecuteReader()
        Dim iC477036 As Short
        ReDim C477036DataType(RsC477036.FieldCount)
        For iC477036 = 0 to RsC477036.FieldCount - 1
            Select Case RsC477036.GetDataTypeName(iC477036).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C477036DataType(iC477036 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C477036DataType(iC477036 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C477036DataType(iC477036 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC477036
        RsC477036_EOF = Not RsC477036.Read()

        GoTo Comp_C381763

    Comp_C477037:

        ' FlagOProprio=1?
        sCurrComponent = "FlagOProprio=1?"
        C477037 = (fn_ConvertValueCompiled(RsC477036(0), C477036DataType(1), AKB_DecimalPoint, False) = 1)
        C477037DataType = AKBTypeConst.cAKBTypeLogical
        If C477037 Then
            GoTo Comp_C56907
        Else
            GoTo Comp_C374708
        End If

    Comp_C480958:

        ' Erro pedido Fenix?
        sCurrComponent = "Erro pedido Fenix?"
        C480958 = (fn_ConvertValueCompiled(C480959, C480959DataType, AKB_DecimalPoint, False) = -1)
        C480958DataType = AKBTypeConst.cAKBTypeLogical
        If C480958 Then
            GoTo Comp_C480960
        Else
            GoTo Comp_C56141
        End If

    Comp_C480959:

        ' vPedidoFenix
        sCurrComponent = "vPedidoFenix"
        C480959DataType = 0
        C480959 = C471477.rows(C471477CurrentRow - 1)(1)
        C480959DataType = C471477DataType(2)
        GoTo Comp_C480958

    Comp_C480960:

        ' frmPedidoFenix
        sCurrComponent = "frmPedidoFenix"

        GoTo Comp_C56141

    Comp_C482983:

        ' SelFlagBloqueiaRedespacho
        sCurrComponent = "SelFlagBloqueiaRedespacho"
        QueryC482983 = con.CreateCommand()
        QueryC482983.CommandText = QueryC482983.CommandText & " " & "SELECT  WF_TP_PRODUTO.bloqueia_pedido_com_frete ,"
        QueryC482983.CommandText = QueryC482983.CommandText & " " & "NVL (WF_PRE_PEDIDO.Cod_Frete_Redesp,0)"
        QueryC482983.CommandText = QueryC482983.CommandText & " " & ""
        QueryC482983.CommandText = QueryC482983.CommandText & " " & "FROM AKBUSER01.WF_TP_PRODUTO , AKBUSER01.WF_PRE_PEDIDO "
        QueryC482983.CommandText = QueryC482983.CommandText & " " & "WHERE WF_TP_PRODUTO.Tp_Produto = WF_PRE_PEDIDO.Tp_Produto "
        QueryC482983.CommandText = QueryC482983.CommandText & " " & "AND  WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482983.Transaction = txn
        RsC482983 = QueryC482983.ExecuteReader()
        Dim iC482983 As Short
        ReDim C482983DataType(RsC482983.FieldCount)
        For iC482983 = 0 to RsC482983.FieldCount - 1
            Select Case RsC482983.GetDataTypeName(iC482983).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C482983DataType(iC482983 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C482983DataType(iC482983 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C482983DataType(iC482983 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC482983
        RsC482983_EOF = Not RsC482983.Read()

        GoTo Comp_C482994

    Comp_C482993:

        ' TemRedespacho
        sCurrComponent = "TemRedespacho"
        C482993DataType = 0
        C482993DataType = C482983Datatype(2)
        If C482993DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC482983(1)) Then
          C482993 = Strings.RTrim(RsC482983(1))
        Else 
          C482993 = RsC482983(1)
        End If 

        GoTo Comp_C486673

    Comp_C482994:

        ' BloqueiaPedidoRedespacho
        sCurrComponent = "BloqueiaPedidoRedespacho"
        C482994DataType = 0
        C482994 = RsC482983(0)
        C482994DataType = C482983Datatype(1)
        If C482994DataType = AKBTypeConst.cAKBTypeString Then
          C482994 = IIF(IsDBNull(C482994), C482994, Strings.RTrim(C482994))
        End If 

        GoTo Comp_C482993

    Comp_C482995:

        ' Bloqueia(Redespacho)?
        sCurrComponent = "Bloqueia(Redespacho)?"
        C482995 = (fn_ConvertValueCompiled(C482994, C482994DataType, AKB_DecimalPoint, False) = 1  AND fn_ConvertValueCompiled(C482993, C482993DataType, AKB_DecimalPoint, False) > 0)
        C482995DataType = AKBTypeConst.cAKBTypeLogical
        If C482995 Then
            GoTo Comp_C540774
        Else
            GoTo Comp_C483016
        End If

    Comp_C482996:

        ' Up_PrePedidoBloquea(Redespacho)
        sCurrComponent = "Up_PrePedidoBloquea(Redespacho)"
        QueryC482996 = con.CreateCommand()
        QueryC482996.CommandText = QueryC482996.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC482996.CommandText = QueryC482996.CommandText & " " & "SET WF_PRE_PEDIDO.Bloq_Redesp  = 1"
        QueryC482996.CommandText = QueryC482996.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482996.Transaction = txn
        C482996 = QueryC482996.ExecuteNonQuery()
        C482996DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C483010

    Comp_C482998:

        ' SelDadosSintegra
        sCurrComponent = "SelDadosSintegra"
        QueryC482998 = con.CreateCommand()
        QueryC482998.CommandText = QueryC482998.CommandText & " " & "SELECT wf_estab_juridico_param.usa_data_sintegra, NVL(wf_estab_juridico_param.Meses_Vcto_Sitengra,0),"
        QueryC482998.CommandText = QueryC482998.CommandText & " " & "NVL(WF_ESTAB_JURIDICO.utiliza_sintegraws,0)"
        QueryC482998.CommandText = QueryC482998.CommandText & " " & "FROM wf_estab_juridico_param, WF_PRE_PEDIDO, WF_ESTAB_JURIDICO"
        QueryC482998.CommandText = QueryC482998.CommandText & " " & "WHERE WF_PRE_PEDIDO.Cod_Pessoa_Estab_Juridico = wf_estab_juridico_param.Cod_Pessoa_Estab_Juridico"
        QueryC482998.CommandText = QueryC482998.CommandText & " " & "AND  WF_ESTAB_JURIDICO.Cod_Pessoa_Estab_Juridico = wf_estab_juridico_param.Cod_Pessoa_Estab_Juridico"
        QueryC482998.CommandText = QueryC482998.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC482998.Transaction = txn
        RsC482998 = QueryC482998.ExecuteReader()
        Dim iC482998 As Short
        ReDim C482998DataType(RsC482998.FieldCount)
        For iC482998 = 0 to RsC482998.FieldCount - 1
            Select Case RsC482998.GetDataTypeName(iC482998).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C482998DataType(iC482998 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C482998DataType(iC482998 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C482998DataType(iC482998 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC482998
        RsC482998_EOF = Not RsC482998.Read()

        GoTo Comp_C483008

    Comp_C483007:

        ' mesesVctoSintegra
        sCurrComponent = "mesesVctoSintegra"
        C483007DataType = 0
        C483007DataType = C482998Datatype(2)
        If C483007DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC482998(1)) Then
          C483007 = Strings.RTrim(RsC482998(1))
        Else 
          C483007 = RsC482998(1)
        End If 

        GoTo Comp_C539465

    Comp_C483008:

        ' usa_data_sintegra
        sCurrComponent = "usa_data_sintegra"
        C483008DataType = 0
        C483008 = RsC482998(0)
        C483008DataType = C482998Datatype(1)
        If C483008DataType = AKBTypeConst.cAKBTypeString Then
          C483008 = IIF(IsDBNull(C483008), C483008, Strings.RTrim(C483008))
        End If 

        GoTo Comp_C483007

    Comp_C483009:

        ' DtSistema
        sCurrComponent = "DtSistema"
        C483009DataType = 2
        C483009 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C482995

    Comp_C483010:

        ' dtSintegra
        sCurrComponent = "dtSintegra"
        QueryC483010 = con.CreateCommand()
        QueryC483010.CommandText = QueryC483010.CommandText & " " & "SELECT add_months(MAX(S.Data), " & _ 
ConvertValue(C483007, C483007DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "), C.SIGLA_ESTADO"
        QueryC483010.CommandText = QueryC483010.CommandText & " " & "FROM WF_CLIENTE_SINTEGRA_DT S, WF_PRE_PEDIDO P, WF_PESSOAS PES, WF_CIDADES C"
        QueryC483010.CommandText = QueryC483010.CommandText & " " & "WHERE S.Cod_Cliente (+) = P.COD_CLIENTE"
        QueryC483010.CommandText = QueryC483010.CommandText & " " & "AND P.ID_PREPEDIDO = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC483010.CommandText = QueryC483010.CommandText & " " & "AND P.COD_CLIENTE = PES.COD_PESSOA"
        QueryC483010.CommandText = QueryC483010.CommandText & " " & "AND PES.codigo_cidade = c.codigo_cidade"
        QueryC483010.CommandText = QueryC483010.CommandText & " " & "GROUP BY  C.SIGLA_ESTADO"
        QueryC483010.Transaction = txn
        RsC483010 = QueryC483010.ExecuteReader()
        Dim iC483010 As Short
        ReDim C483010DataType(RsC483010.FieldCount)
        For iC483010 = 0 to RsC483010.FieldCount - 1
            Select Case RsC483010.GetDataTypeName(iC483010).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C483010DataType(iC483010 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C483010DataType(iC483010 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C483010DataType(iC483010 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC483010
        RsC483010_EOF = Not RsC483010.Read()

        GoTo Comp_C483011

    Comp_C483011:

        ' dtSintegraNull
        sCurrComponent = "dtSintegraNull"
        C483011DataType = 4
        C483011 = IsDBNull (RsC483010(0))
        GoTo Comp_C540756

    Comp_C483012:

        ' UsaSintegra?
        sCurrComponent = "UsaSintegra?"
        C483012 = (fn_ConvertValueCompiled(C483008, C483008DataType, AKB_DecimalPoint, False) = 1)
        C483012DataType = AKBTypeConst.cAKBTypeLogical
        If C483012 Then
            GoTo Comp_C486385
        Else
            GoTo Comp_C483014
        End If

    Comp_C483013:

        ' dtSistema > dtSintegra?
        sCurrComponent = "dtSistema > dtSintegra?"
        C483013 = (fn_ConvertValueCompiled(C483009, C483009DataType, AKB_DecimalPoint, False) > fn_ConvertValueCompiled(RsC483010(0), C483010DataType(1), AKB_DecimalPoint, False))
        C483013DataType = AKBTypeConst.cAKBTypeLogical
        If C483013 Then
            GoTo Comp_C539466
        Else
            GoTo Comp_C483014
        End If

    Comp_C483014:

        ' flagSintegra=0
        sCurrComponent = "flagSintegra=0"
        QueryC483014 = con.CreateCommand()
        QueryC483014.CommandText = QueryC483014.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC483014.CommandText = QueryC483014.CommandText & " " & "SET WF_PRE_PEDIDO.Bloq_Sintegra = 0"
        QueryC483014.CommandText = QueryC483014.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC483014.Transaction = txn
        C483014 = QueryC483014.ExecuteNonQuery()
        C483014DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C483018

    Comp_C483015:

        ' bloqueiaPedido(Sintegra)
        sCurrComponent = "bloqueiaPedido(Sintegra)"
        QueryC483015 = con.CreateCommand()
        QueryC483015.CommandText = QueryC483015.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC483015.CommandText = QueryC483015.CommandText & " " & "SET WF_PRE_PEDIDO.Bloq_Sintegra = 1"
        QueryC483015.CommandText = QueryC483015.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC483015.Transaction = txn
        C483015 = QueryC483015.ExecuteNonQuery()
        C483015DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C483018

    Comp_C483016:

        ' BloqRedesp=0
        sCurrComponent = "BloqRedesp=0"
        QueryC483016 = con.CreateCommand()
        QueryC483016.CommandText = QueryC483016.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC483016.CommandText = QueryC483016.CommandText & " " & "SET WF_PRE_PEDIDO.Bloq_Redesp  = 0"
        QueryC483016.CommandText = QueryC483016.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC483016.Transaction = txn
        C483016 = QueryC483016.ExecuteNonQuery()
        C483016DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C483010

    Comp_C483018:

        ' SelBloqueio
        sCurrComponent = "SelBloqueio"
        QueryC483018 = con.CreateCommand()
        QueryC483018.CommandText = QueryC483018.CommandText & " " & "SELECT NVL(WF_PRE_PEDIDO.Bloq_Redesp,0), NVL(WF_PRE_PEDIDO.Bloq_Sintegra,0), NVL(WF_PRE_PEDIDO.Lib_Sint_Redesp,0) "
        QueryC483018.CommandText = QueryC483018.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC483018.CommandText = QueryC483018.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC483018.Transaction = txn
        RsC483018 = QueryC483018.ExecuteReader()
        Dim iC483018 As Short
        ReDim C483018DataType(RsC483018.FieldCount)
        For iC483018 = 0 to RsC483018.FieldCount - 1
            Select Case RsC483018.GetDataTypeName(iC483018).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C483018DataType(iC483018 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C483018DataType(iC483018 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C483018DataType(iC483018 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC483018
        RsC483018_EOF = Not RsC483018.Read()

        GoTo Comp_C483020

    Comp_C483019:

        ' BloqSint
        sCurrComponent = "BloqSint"
        C483019DataType = 0
        C483019DataType = C483018Datatype(2)
        If C483019DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC483018(1)) Then
          C483019 = Strings.RTrim(RsC483018(1))
        Else 
          C483019 = RsC483018(1)
        End If 

        GoTo Comp_C483026

    Comp_C483020:

        ' BloqRedesp
        sCurrComponent = "BloqRedesp"
        C483020DataType = 0
        C483020 = RsC483018(0)
        C483020DataType = C483018Datatype(1)
        If C483020DataType = AKBTypeConst.cAKBTypeString Then
          C483020 = IIF(IsDBNull(C483020), C483020, Strings.RTrim(C483020))
        End If 

        GoTo Comp_C483019

    Comp_C483021:

        ' Bloqueado?
        sCurrComponent = "Bloqueado?"
        C483021 = (fn_ConvertValueCompiled(C483020, C483020DataType, AKB_DecimalPoint, False) = 1 OR fn_ConvertValueCompiled(C483019, C483019DataType, AKB_DecimalPoint, False)  = 1)
        C483021DataType = AKBTypeConst.cAKBTypeLogical
        If C483021 Then
            GoTo Comp_C483027
        Else
            GoTo Comp_C483025
        End If

    Comp_C483022:

        ' MSG5
        sCurrComponent = "MSG5"
        Dim row_C483022 As DataRow = msg.NewRow()
        row_C483022(0) = "Não é possível gerar o Pré Pedido - " & _ 
P12596 & "." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Pré Pedido Bloqueado: tem redespacho ou sintegra vencido." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "Favor liberar na tela Liberação de Pré-Pedido com Redespacho ou Sintegra Vencido."
        msg.Rows.Add(row_C483022)
        C483022DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C483023

    Comp_C483023:

        ' RET11
        sCurrComponent = "RET11"
        Dim tb_C483023 As DataTable = New DataTable()
        tb_C483023.Columns.Add("Bloqueado?" & "")
        Dim row_C483023 As DataRow = tb_C483023.NewRow()
        row_C483023(0) = C483021
        tb_C483023.Rows.Add(row_C483023)
        R4557 = tb_C483023
        ReDim C483023DataType(1)
        C483023DataType(1) = C483021DataType
        ReturnDataType = C483023DataType

        GoTo Exit_R4557

    Comp_C483025:

        ' LiberadoSintRedesp
        sCurrComponent = "LiberadoSintRedesp"
        QueryC483025 = con.CreateCommand()
        QueryC483025.CommandText = QueryC483025.CommandText & " " & "UPDATE AKBUSER01.WF_PRE_PEDIDO "
        QueryC483025.CommandText = QueryC483025.CommandText & " " & "SET WF_PRE_PEDIDO.Lib_Sint_Redesp = 1"
        QueryC483025.CommandText = QueryC483025.CommandText & " " & "WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC483025.Transaction = txn
        C483025 = QueryC483025.ExecuteNonQuery()
        C483025DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C394339

    Comp_C483026:

        ' LiberadoSintRedes?
        sCurrComponent = "LiberadoSintRedes?"
        C483026DataType = 0
        C483026DataType = C483018Datatype(3)
        If C483026DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC483018(2)) Then
          C483026 = Strings.RTrim(RsC483018(2))
        Else 
          C483026 = RsC483018(2)
        End If 

        GoTo Comp_C483021

    Comp_C483027:

        ' Liberado?
        sCurrComponent = "Liberado?"
        C483027 = (fn_ConvertValueCompiled(C483026, C483026DataType, AKB_DecimalPoint, False) = 0)
        C483027DataType = AKBTypeConst.cAKBTypeLogical
        If C483027 Then
            GoTo Comp_C483022
        Else
            GoTo Comp_C394339
        End If

    Comp_C486084:

        ' Proforma3?
        sCurrComponent = "Proforma3?"
        C486084 = (fn_ConvertValueCompiled(P26981, 4, AKB_DecimalPoint, False) = 1)
        C486084DataType = AKBTypeConst.cAKBTypeLogical
        If C486084 Then
            GoTo Comp_C563342
        Else
            GoTo Comp_C482998
        End If

    Comp_C486385:

        ' ÉNulo?
        sCurrComponent = "ÉNulo?"
        C486385 = (fn_ConvertValueCompiled(C483011, C483011DataType, AKB_DecimalPoint, False) = 1)
        C486385DataType = AKBTypeConst.cAKBTypeLogical
        If C486385 Then
            GoTo Comp_C539466
        Else
            GoTo Comp_C483013
        End If

    Comp_C486673:

        ' Copia?
        sCurrComponent = "Copia?"
        C486673 = (fn_ConvertValueCompiled(P78985, 1, AKB_DecimalPoint, False) = 1)
        C486673DataType = AKBTypeConst.cAKBTypeLogical
        If C486673 Then
            GoTo Comp_C394339
        Else
            GoTo Comp_C486084
        End If

    Comp_C502254:

        ' FlagCartaoCred
        sCurrComponent = "FlagCartaoCred"
        QueryC502254 = con.CreateCommand()
        QueryC502254.CommandText = QueryC502254.CommandText & " " & "SELECT NVL(Cartao_Credito,0)"
        QueryC502254.CommandText = QueryC502254.CommandText & " " & "FROM WF_CONDICAO_PGTO"
        QueryC502254.CommandText = QueryC502254.CommandText & " " & "WHERE Cod_Pagto = " & _ 
ConvertValue(C56920, C56920DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC502254.Transaction = txn
        RsC502254 = QueryC502254.ExecuteReader()
        Dim iC502254 As Short
        ReDim C502254DataType(RsC502254.FieldCount)
        For iC502254 = 0 to RsC502254.FieldCount - 1
            Select Case RsC502254.GetDataTypeName(iC502254).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C502254DataType(iC502254 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C502254DataType(iC502254 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C502254DataType(iC502254 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC502254
        RsC502254_EOF = Not RsC502254.Read()

        GoTo Comp_C502255

    Comp_C502255:

        ' CBCartaoCred
        sCurrComponent = "CBCartaoCred"
        QueryC502255 = con.CreateCommand()
        QueryC502255.CommandText = QueryC502255.CommandText & " " & "SELECT MAX(NVL(Num_Conta_Bancaria,0))"
        QueryC502255.CommandText = QueryC502255.CommandText & " " & "FROM WF_CONTAS_BANCARIAS"
        QueryC502255.CommandText = QueryC502255.CommandText & " " & "WHERE Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P76804, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC502255.CommandText = QueryC502255.CommandText & " " & "AND Conta_Cartao_Credito = 1"
        QueryC502255.Transaction = txn
        RsC502255 = QueryC502255.ExecuteReader()
        Dim iC502255 As Short
        ReDim C502255DataType(RsC502255.FieldCount)
        For iC502255 = 0 to RsC502255.FieldCount - 1
            Select Case RsC502255.GetDataTypeName(iC502255).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C502255DataType(iC502255 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C502255DataType(iC502255 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C502255DataType(iC502255 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC502255
        RsC502255_EOF = Not RsC502255.Read()

        GoTo Comp_C512122

    Comp_C502256:

        ' FlagCartaoCred=0?
        sCurrComponent = "FlagCartaoCred=0?"
        C502256 = (fn_ConvertValueCompiled(RsC502254(0), C502254DataType(1), AKB_DecimalPoint, False) = 0)
        C502256DataType = AKBTypeConst.cAKBTypeLogical
        If C502256 Then
            GoTo Comp_C512124
        Else
            GoTo Comp_C502257
        End If

    Comp_C502257:

        ' UpdContaBanc
        sCurrComponent = "UpdContaBanc"
        QueryC502257 = con.CreateCommand()
        QueryC502257.CommandText = QueryC502257.CommandText & " " & "UPDATE WF_PEDIDO SET Num_Conta_Bancaria = " & _ 
ConvertValue(RsC502255(0), C502255DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " WHERE Tp_Produto = " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC502257.Transaction = txn
        C502257 = QueryC502257.ExecuteNonQuery()
        C502257DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C512124

    Comp_C508135:

        ' FlagRiscoSac
        sCurrComponent = "FlagRiscoSac"
        QueryC508135 = con.CreateCommand()
        QueryC508135.CommandText = QueryC508135.CommandText & " " & "SELECT NVL(E_Risco_Sacado,0)"
        QueryC508135.CommandText = QueryC508135.CommandText & " " & "FROM WF_CONDICAO_PGTO"
        QueryC508135.CommandText = QueryC508135.CommandText & " " & "WHERE Cod_Pagto = " & _ 
ConvertValue(C56920, C56920DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC508135.Transaction = txn
        RsC508135 = QueryC508135.ExecuteReader()
        Dim iC508135 As Short
        ReDim C508135DataType(RsC508135.FieldCount)
        For iC508135 = 0 to RsC508135.FieldCount - 1
            Select Case RsC508135.GetDataTypeName(iC508135).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C508135DataType(iC508135 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C508135DataType(iC508135 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C508135DataType(iC508135 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC508135
        RsC508135_EOF = Not RsC508135.Read()

        GoTo Comp_C508137

    Comp_C508137:

        ' FlagRS=0?
        sCurrComponent = "FlagRS=0?"
        C508137 = (fn_ConvertValueCompiled(RsC508135(0), C508135DataType(1), AKB_DecimalPoint, False) = 0)
        C508137DataType = AKBTypeConst.cAKBTypeLogical
        If C508137 Then
            GoTo Comp_C579427
        Else
            GoTo Comp_C508138
        End If

    Comp_C508138:

        ' SysDate
        sCurrComponent = "SysDate"
        C508138DataType = 2
        C508138 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C508139

    Comp_C508139:

        ' Taxas
        sCurrComponent = "Taxas"
        QueryC508139 = con.CreateCommand()
        QueryC508139.CommandText = QueryC508139.CommandText & " " & "SELECT NVL(Qtde_Dias_Vencimento,0), NVL(Custo_Financeiro_Mensal,0)"
        QueryC508139.CommandText = QueryC508139.CommandText & " " & "FROM WF_TAXA_RISCO_SACADO"
        QueryC508139.CommandText = QueryC508139.CommandText & " " & "WHERE Cod_Cliente = " & _ 
ConvertValue(C75153, C75153DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC508139.CommandText = QueryC508139.CommandText & " " & "AND Data_Validade_Inicial <= " & _ 
ConvertValue(C508138, C508138DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC508139.CommandText = QueryC508139.CommandText & " " & "AND ( Data_Validade_Final >= " & _ 
ConvertValue(C508138, C508138DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " OR Data_Validade_Final IS NULL)"
        QueryC508139.Transaction = txn
        RsC508139 = QueryC508139.ExecuteReader()
        Dim iC508139 As Short
        ReDim C508139DataType(RsC508139.FieldCount)
        For iC508139 = 0 to RsC508139.FieldCount - 1
            Select Case RsC508139.GetDataTypeName(iC508139).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C508139DataType(iC508139 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C508139DataType(iC508139 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C508139DataType(iC508139 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC508139
        RsC508139_EOF = Not RsC508139.Read()

        GoTo Comp_C508140

    Comp_C508140:

        ' DiasRS
        sCurrComponent = "DiasRS"
        C508140DataType = 0
        C508140 = RsC508139(0)
        C508140DataType = C508139Datatype(1)
        If C508140DataType = AKBTypeConst.cAKBTypeString Then
          C508140 = IIF(IsDBNull(C508140), C508140, Strings.RTrim(C508140))
        End If 

        GoTo Comp_C508141

    Comp_C508141:

        ' %CF_RS
        sCurrComponent = "%CF_RS"
        C508141DataType = 0
        C508141DataType = C508139Datatype(2)
        If C508141DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC508139(1)) Then
          C508141 = Strings.RTrim(RsC508139(1))
        Else 
          C508141 = RsC508139(1)
        End If 

        GoTo Comp_C508147

    Comp_C508142:

        ' DESVIO1
        sCurrComponent = "DESVIO1"
        C508142 = (fn_ConvertValueCompiled(C508141, C508141DataType, AKB_DecimalPoint, False) > 0 AND fn_ConvertValueCompiled(C508140, C508140DataType, AKB_DecimalPoint, False) > 0)
        C508142DataType = AKBTypeConst.cAKBTypeLogical
        If C508142 Then
            GoTo Comp_C174610
        Else
            GoTo Comp_C508143
        End If

    Comp_C508143:

        ' Ret1n
        sCurrComponent = "Ret1n"
        C508143 = 0
        C508143DataType = 4
        GoTo Comp_C508144

    Comp_C508144:

        ' RET12
        sCurrComponent = "RET12"
        Dim tb_C508144 As DataTable = New DataTable()
        tb_C508144.Columns.Add("Ret1n" & "")
        Dim row_C508144 As DataRow = tb_C508144.NewRow()
        row_C508144(0) = C508143
        tb_C508144.Rows.Add(row_C508144)
        R4557 = tb_C508144
        ReDim C508144DataType(1)
        C508144DataType(1) = C508143DataType
        ReturnDataType = C508144DataType

        GoTo Exit_R4557

    Comp_C508145:

        ' vDiasRS
        sCurrComponent = "vDiasRS"
        C508145 = 0
        C508145DataType = 1
        GoTo Comp_C508146

    Comp_C508146:

        ' v%CF_RS
        sCurrComponent = "v%CF_RS"
        C508146 = 0
        C508146DataType = 1
        GoTo Comp_C363384

    Comp_C508147:

        ' ATRIBUIVA4
        sCurrComponent = "ATRIBUIVA4"
        C508147DataType = 4
        C508145 = fn_ConvertValueCompiled(C508140 , 1, AKB_DecimalPoint)
        C508147 = True
        GoTo Comp_C508148

    Comp_C508148:

        ' ATRIBUIVA5
        sCurrComponent = "ATRIBUIVA5"
        C508148DataType = 4
        C508146 = fn_ConvertValueCompiled(C508141 , 1, AKB_DecimalPoint)
        C508148 = True
        GoTo Comp_C508142

    Comp_C512122:

        ' FlagPgtoVista
        sCurrComponent = "FlagPgtoVista"
        QueryC512122 = con.CreateCommand()
        QueryC512122.CommandText = QueryC512122.CommandText & " " & "SELECT NVL(Pagamento_a_Vista,0)"
        QueryC512122.CommandText = QueryC512122.CommandText & " " & "FROM WF_CONDICAO_PGTO"
        QueryC512122.CommandText = QueryC512122.CommandText & " " & "WHERE Cod_Pagto = " & _ 
ConvertValue(C56920, C56920DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC512122.Transaction = txn
        RsC512122 = QueryC512122.ExecuteReader()
        Dim iC512122 As Short
        ReDim C512122DataType(RsC512122.FieldCount)
        For iC512122 = 0 to RsC512122.FieldCount - 1
            Select Case RsC512122.GetDataTypeName(iC512122).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C512122DataType(iC512122 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C512122DataType(iC512122 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C512122DataType(iC512122 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC512122
        RsC512122_EOF = Not RsC512122.Read()

        GoTo Comp_C512123

    Comp_C512123:

        ' CBPgtoVista
        sCurrComponent = "CBPgtoVista"
        QueryC512123 = con.CreateCommand()
        QueryC512123.CommandText = QueryC512123.CommandText & " " & "SELECT MAX(NVL(Num_Conta_Bancaria,0))"
        QueryC512123.CommandText = QueryC512123.CommandText & " " & "FROM WF_CONTAS_BANCARIAS"
        QueryC512123.CommandText = QueryC512123.CommandText & " " & "WHERE Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(C517972, C517972DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC512123.CommandText = QueryC512123.CommandText & " " & "AND Usada_para_Cond_Pag_Vista = 1"
        QueryC512123.Transaction = txn
        RsC512123 = QueryC512123.ExecuteReader()
        Dim iC512123 As Short
        ReDim C512123DataType(RsC512123.FieldCount)
        For iC512123 = 0 to RsC512123.FieldCount - 1
            Select Case RsC512123.GetDataTypeName(iC512123).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C512123DataType(iC512123 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C512123DataType(iC512123 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C512123DataType(iC512123 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC512123
        RsC512123_EOF = Not RsC512123.Read()

        GoTo Comp_C508135

    Comp_C512124:

        ' FlagPgtoVista=0?
        sCurrComponent = "FlagPgtoVista=0?"
        C512124 = (fn_ConvertValueCompiled(RsC512122(0), C512122DataType(1), AKB_DecimalPoint, False) = 0)
        C512124DataType = AKBTypeConst.cAKBTypeLogical
        If C512124 Then
            GoTo Comp_C512822
        Else
            GoTo Comp_C512125
        End If

    Comp_C512125:

        ' UpContaBanc1
        sCurrComponent = "UpContaBanc1"
        QueryC512125 = con.CreateCommand()
        QueryC512125.CommandText = QueryC512125.CommandText & " " & "UPDATE WF_PEDIDO SET Num_Conta_Bancaria =  " & _ 
ConvertValue(RsC512123(0), C512123DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & "  WHERE Tp_Produto = " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC512125.Transaction = txn
        C512125 = QueryC512125.ExecuteNonQuery()
        C512125DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C512822

    Comp_C512820:

        ' FlagPedFunc
        sCurrComponent = "FlagPedFunc"
        QueryC512820 = con.CreateCommand()
        QueryC512820.CommandText = QueryC512820.CommandText & " " & "SELECT NVL(Pedido_de_Funcionario,0)"
        QueryC512820.CommandText = QueryC512820.CommandText & " " & "FROM WF_CONDICAO_PGTO"
        QueryC512820.CommandText = QueryC512820.CommandText & " " & "WHERE Cod_Pagto = " & _ 
ConvertValue(C56920, C56920DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC512820.Transaction = txn
        RsC512820 = QueryC512820.ExecuteReader()
        Dim iC512820 As Short
        ReDim C512820DataType(RsC512820.FieldCount)
        For iC512820 = 0 to RsC512820.FieldCount - 1
            Select Case RsC512820.GetDataTypeName(iC512820).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C512820DataType(iC512820 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C512820DataType(iC512820 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C512820DataType(iC512820 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC512820
        RsC512820_EOF = Not RsC512820.Read()

        GoTo Comp_C502254

    Comp_C512821:

        ' CBPedFunc
        sCurrComponent = "CBPedFunc"
        QueryC512821 = con.CreateCommand()
        QueryC512821.CommandText = QueryC512821.CommandText & " " & "SELECT MAX(NVL(Num_Conta_Bancaria,0))"
        QueryC512821.CommandText = QueryC512821.CommandText & " " & "FROM WF_CONTAS_BANCARIAS"
        QueryC512821.CommandText = QueryC512821.CommandText & " " & "WHERE Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P76804, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC512821.CommandText = QueryC512821.CommandText & " " & "AND Usada_para_Ped_Func = 1"
        QueryC512821.Transaction = txn
        RsC512821 = QueryC512821.ExecuteReader()
        Dim iC512821 As Short
        ReDim C512821DataType(RsC512821.FieldCount)
        For iC512821 = 0 to RsC512821.FieldCount - 1
            Select Case RsC512821.GetDataTypeName(iC512821).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C512821DataType(iC512821 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C512821DataType(iC512821 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C512821DataType(iC512821 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC512821
        RsC512821_EOF = Not RsC512821.Read()

        GoTo Comp_C512820

    Comp_C512822:

        ' FlagPedFunc=0?
        sCurrComponent = "FlagPedFunc=0?"
        C512822 = (fn_ConvertValueCompiled(RsC512820(0), C512820DataType(1), AKB_DecimalPoint, False) = 0)
        C512822DataType = AKBTypeConst.cAKBTypeLogical
        If C512822 Then
            GoTo Comp_C474763
        Else
            GoTo Comp_C512823
        End If

    Comp_C512823:

        ' UpContaBancPedFunc
        sCurrComponent = "UpContaBancPedFunc"
        QueryC512823 = con.CreateCommand()
        QueryC512823.CommandText = QueryC512823.CommandText & " " & "UPDATE WF_PEDIDO SET Num_Conta_Bancaria = " & _ 
ConvertValue(RsC512821(0), C512821DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " WHERE Tp_Produto = " & _ 
ConvertValue(C56126, C56126DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " AND Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC512823.Transaction = txn
        C512823 = QueryC512823.ExecuteNonQuery()
        C512823DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C474763

    Comp_C516410:

        ' BloqPedEmail
        sCurrComponent = "BloqPedEmail"
        C516410DataType = 0
        C516410DataType = C470976Datatype(2)
        If C516410DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC470976(1)) Then
          C516410 = Strings.RTrim(RsC470976(1))
        Else 
          C516410 = RsC470976(1)
        End If 

        GoTo Comp_C516412

    Comp_C516411:

        ' CodCliPrePed
        sCurrComponent = "CodCliPrePed"
        QueryC516411 = con.CreateCommand()
        QueryC516411.CommandText = QueryC516411.CommandText & " " & "SELECT Cod_Cliente"
        QueryC516411.CommandText = QueryC516411.CommandText & " " & "FROM WF_PRE_PEDIDO"
        QueryC516411.CommandText = QueryC516411.CommandText & " " & "WHERE Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC516411.Transaction = txn
        RsC516411 = QueryC516411.ExecuteReader()
        Dim iC516411 As Short
        ReDim C516411DataType(RsC516411.FieldCount)
        For iC516411 = 0 to RsC516411.FieldCount - 1
            Select Case RsC516411.GetDataTypeName(iC516411).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C516411DataType(iC516411 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C516411DataType(iC516411 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C516411DataType(iC516411 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC516411
        RsC516411_EOF = Not RsC516411.Read()

        GoTo Comp_C516414

    Comp_C516412:

        ' CodEmailDANFE
        sCurrComponent = "CodEmailDANFE"
        C516412DataType = 0
        C516412DataType = C470976Datatype(3)
        If C516412DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC470976(2)) Then
          C516412 = Strings.RTrim(RsC470976(2))
        Else 
          C516412 = RsC470976(2)
        End If 

        GoTo Comp_C551599

    Comp_C516413:

        ' FlagEmail
        sCurrComponent = "FlagEmail"
        C516413 = 1
        C516413DataType = 4
        GoTo Comp_C516421

    Comp_C516414:

        ' EmailDANFE
        sCurrComponent = "EmailDANFE"
        QueryC516414 = con.CreateCommand()
        QueryC516414.CommandText = QueryC516414.CommandText & " " & "SELECT NVL(COUNT(*),0)"
        QueryC516414.CommandText = QueryC516414.CommandText & " " & "FROM WF_COMUNIC"
        QueryC516414.CommandText = QueryC516414.CommandText & " " & "WHERE Cod_Pessoa = " & _ 
ConvertValue(RsC516411(0), C516411DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC516414.CommandText = QueryC516414.CommandText & " " & "AND Cod_Comunic = " & _ 
ConvertValue(C516412, C516412DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC516414.CommandText = QueryC516414.CommandText & " " & "AND Num_Comunicacao IS NOT NULL"
        QueryC516414.Transaction = txn
        RsC516414 = QueryC516414.ExecuteReader()
        Dim iC516414 As Short
        ReDim C516414DataType(RsC516414.FieldCount)
        For iC516414 = 0 to RsC516414.FieldCount - 1
            Select Case RsC516414.GetDataTypeName(iC516414).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C516414DataType(iC516414 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C516414DataType(iC516414 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C516414DataType(iC516414 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC516414
        RsC516414_EOF = Not RsC516414.Read()

        GoTo Comp_C516415

    Comp_C516415:

        ' EmailDANFE>0?
        sCurrComponent = "EmailDANFE>0?"
        C516415 = (fn_ConvertValueCompiled(RsC516414(0), C516414DataType(1), AKB_DecimalPoint, False) > 0)
        C516415DataType = AKBTypeConst.cAKBTypeLogical
        If C516415 Then
            GoTo Comp_C516416
        Else
            GoTo Comp_C516417
        End If

    Comp_C516416:

        ' EmailNormal
        sCurrComponent = "EmailNormal"
        QueryC516416 = con.CreateCommand()
        QueryC516416.CommandText = QueryC516416.CommandText & " " & "SELECT NVL(COUNT(*),0)"
        QueryC516416.CommandText = QueryC516416.CommandText & " " & "FROM WF_COMUNIC"
        QueryC516416.CommandText = QueryC516416.CommandText & " " & "WHERE Cod_Pessoa = " & _ 
ConvertValue(RsC516411(0), C516411DataType(1), DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC516416.CommandText = QueryC516416.CommandText & " " & "AND Cod_Comunic = 2 "
        QueryC516416.CommandText = QueryC516416.CommandText & " " & "AND Num_Comunicacao IS NOT NULL"
        QueryC516416.Transaction = txn
        RsC516416 = QueryC516416.ExecuteReader()
        Dim iC516416 As Short
        ReDim C516416DataType(RsC516416.FieldCount)
        For iC516416 = 0 to RsC516416.FieldCount - 1
            Select Case RsC516416.GetDataTypeName(iC516416).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C516416DataType(iC516416 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C516416DataType(iC516416 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C516416DataType(iC516416 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC516416
        RsC516416_EOF = Not RsC516416.Read()

        GoTo Comp_C516420

    Comp_C516417:

        ' ATRIBUIVA6
        sCurrComponent = "ATRIBUIVA6"
        C516417DataType = 4
        C516413 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C516417 = True
        GoTo Comp_C516418

    Comp_C516418:

        ' MSG6
        sCurrComponent = "MSG6"
        Dim row_C516418 As DataRow = msg.NewRow()
        row_C516418(0) = "Cliente " & _ 
RsC516411(0) & " do pre-pedido " & _ 
P12596 & " não tem os e-mail´s obrigatórios (2-Email e/ou E-mail DANFE). " & Chr(13) & "" & Chr(10) & "Pedido não será gerado."
        msg.Rows.Add(row_C516418)
        C516418DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C516419

    Comp_C516419:

        ' RET13
        sCurrComponent = "RET13"
        Dim tb_C516419 As DataTable = New DataTable()
        tb_C516419.Columns.Add("FlagEmail" & "")
        Dim row_C516419 As DataRow = tb_C516419.NewRow()
        row_C516419(0) = C516413
        tb_C516419.Rows.Add(row_C516419)
        R4557 = tb_C516419
        ReDim C516419DataType(1)
        C516419DataType(1) = C516413DataType
        ReturnDataType = C516419DataType

        GoTo Exit_R4557

    Comp_C516420:

        ' EmailNormal>0?
        sCurrComponent = "EmailNormal>0?"
        C516420 = (fn_ConvertValueCompiled(RsC516416(0), C516416DataType(1), AKB_DecimalPoint, False) > 0)
        C516420DataType = AKBTypeConst.cAKBTypeLogical
        If C516420 Then
            GoTo Comp_C551587
        Else
            GoTo Comp_C516417
        End If

    Comp_C516421:

        ' BloqPedEmail=1?
        sCurrComponent = "BloqPedEmail=1?"
        C516421 = (fn_ConvertValueCompiled(C516410, C516410DataType, AKB_DecimalPoint, False) = 1)
        C516421DataType = AKBTypeConst.cAKBTypeLogical
        If C516421 Then
            GoTo Comp_C516411
        Else
            GoTo Comp_C551600
        End If

    Comp_C517971:

        ' EstabFatur
        sCurrComponent = "EstabFatur"
        QueryC517971 = con.CreateCommand()
        QueryC517971.CommandText = QueryC517971.CommandText & " " & "SELECT DECODE(WF_ESTAB_JURIDICO_PARAM.Cod_Estab_Faturamento,NULL,WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico,WF_ESTAB_JURIDICO_PARAM.Cod_Estab_Faturamento)"
        QueryC517971.CommandText = QueryC517971.CommandText & " " & " FROM AKBUSER01.WF_ESTAB_JURIDICO_PARAM WHERE WF_ESTAB_JURIDICO_PARAM.Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P76804, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC517971.Transaction = txn
        RsC517971 = QueryC517971.ExecuteReader()
        Dim iC517971 As Short
        ReDim C517971DataType(RsC517971.FieldCount)
        For iC517971 = 0 to RsC517971.FieldCount - 1
            Select Case RsC517971.GetDataTypeName(iC517971).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C517971DataType(iC517971 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C517971DataType(iC517971 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C517971DataType(iC517971 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC517971
        RsC517971_EOF = Not RsC517971.Read()

        GoTo Comp_C517972

    Comp_C517972:

        ' vEstabFatur
        sCurrComponent = "vEstabFatur"
        C517972DataType = 0
        C517972 = RsC517971(0)
        C517972DataType = C517971Datatype(1)
        If C517972DataType = AKBTypeConst.cAKBTypeString Then
          C517972 = IIF(IsDBNull(C517972), C517972, Strings.RTrim(C517972))
        End If 

        GoTo Comp_C555064

    Comp_C519010:

        ' TpPrePed
        sCurrComponent = "TpPrePed"
        QueryC519010 = con.CreateCommand()
        QueryC519010.CommandText = QueryC519010.CommandText & " " & "SELECT WF_PRE_PEDIDO.Tp_Produto FROM AKBUSER01.WF_PRE_PEDIDO WHERE WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC519010.Transaction = txn
        RsC519010 = QueryC519010.ExecuteReader()
        Dim iC519010 As Short
        ReDim C519010DataType(RsC519010.FieldCount)
        For iC519010 = 0 to RsC519010.FieldCount - 1
            Select Case RsC519010.GetDataTypeName(iC519010).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C519010DataType(iC519010 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C519010DataType(iC519010 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C519010DataType(iC519010 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC519010
        RsC519010_EOF = Not RsC519010.Read()

        GoTo Comp_C519011

    Comp_C519011:

        ' vTPPrePed
        sCurrComponent = "vTPPrePed"
        C519011DataType = 0
        C519011 = RsC519010(0)
        C519011DataType = C519010Datatype(1)
        If C519011DataType = AKBTypeConst.cAKBTypeString Then
          C519011 = IIF(IsDBNull(C519011), C519011, Strings.RTrim(C519011))
        End If 

        GoTo Comp_C519012

    Comp_C519012:

        ' PedManual
        sCurrComponent = "PedManual"
        C519012 = 0
        C519012DataType = 4
        GoTo Comp_C519013

    Comp_C519013:

        ' vTPPrePed=200?
        sCurrComponent = "vTPPrePed=200?"
        C519013 = (fn_ConvertValueCompiled(C519011, C519011DataType, AKB_DecimalPoint, False) = 200)
        C519013DataType = AKBTypeConst.cAKBTypeLogical
        If C519013 Then
            GoTo Comp_C519014
        Else
            GoTo Comp_C218740
        End If

    Comp_C519014:

        ' ATRIBUIVA7
        sCurrComponent = "ATRIBUIVA7"
        C519014DataType = 4
        C519012 = fn_ConvertValueCompiled(1, 4, AKB_DecimalPoint)
        C519014 = True
        GoTo Comp_C218740

    Comp_C524142:

        ' CountKIT
        sCurrComponent = "CountKIT"
        QueryC524142 = con.CreateCommand()
        QueryC524142.CommandText = QueryC524142.CommandText & " " & "SELECT NVL(COUNT(*),0)"
        QueryC524142.CommandText = QueryC524142.CommandText & " " & "FROM WF_PEDIDO_ITENS, WF_SIGLA_PROD_AGRUP"
        QueryC524142.CommandText = QueryC524142.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC524142.CommandText = QueryC524142.CommandText & " " & "AND WF_PEDIDO_ITENS.Sigla_Prod = WF_SIGLA_PROD_AGRUP.Sigla_Prod"
        QueryC524142.CommandText = QueryC524142.CommandText & " " & "AND WF_SIGLA_PROD_AGRUP.Sigla_de_Kit = 1"
        QueryC524142.Transaction = txn
        RsC524142 = QueryC524142.ExecuteReader()
        Dim iC524142 As Short
        ReDim C524142DataType(RsC524142.FieldCount)
        For iC524142 = 0 to RsC524142.FieldCount - 1
            Select Case RsC524142.GetDataTypeName(iC524142).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C524142DataType(iC524142 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C524142DataType(iC524142 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C524142DataType(iC524142 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC524142
        RsC524142_EOF = Not RsC524142.Read()

        GoTo Comp_C524143

    Comp_C524143:

        ' CountKIT=0?
        sCurrComponent = "CountKIT=0?"
        C524143 = (fn_ConvertValueCompiled(RsC524142(0), C524142DataType(1), AKB_DecimalPoint, False) = 0)
        C524143DataType = AKBTypeConst.cAKBTypeLogical
        If C524143 Then
            GoTo Comp_C57925
        Else
            GoTo Comp_C524144
        End If

    Comp_C524144:

        ' UpKitPedido
        sCurrComponent = "UpKitPedido"
        QueryC524144 = con.CreateCommand()
        QueryC524144.CommandText = QueryC524144.CommandText & " " & "UPDATE WF_PEDIDO SET Pedido_de_Kit = 1, Entrega_Parcial = 0 WHERE Pedido = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC524144.Transaction = txn
        C524144 = QueryC524144.ExecuteNonQuery()
        C524144DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C57925

    Comp_C532561:

        ' DESVIO2
        sCurrComponent = "DESVIO2"
        C532561 = (fn_ConvertValueCompiled(P78985, 1, AKB_DecimalPoint, False) = 1)
        C532561DataType = AKBTypeConst.cAKBTypeLogical
        If C532561 Then
            GoTo Comp_C371804
        Else
            GoTo Comp_C367579
        End If

    Comp_C539128:

        ' SelECommerce
        sCurrComponent = "SelECommerce"
        QueryC539128 = con.CreateCommand()
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "SELECT COUNT(*)"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "FROM WF_PRE_PEDIDO"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "WHERE WF_PRE_PEDIDO.id_prepedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "AND  NVL(WF_PRE_PEDIDO.pedido_ecommerce,0) = 1"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "and WF_PRE_PEDIDO.PORTAL_LINDT = 0"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "AND portal_brasilcacau = 0"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "AND portal_cacaushow = 0"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "AND portal_kopenhagen = 0"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "AND Portal_boticario = 0"
        QueryC539128.CommandText = QueryC539128.CommandText & " " & "AND Portal_Bauducco = 0"
        QueryC539128.Transaction = txn
        RsC539128 = QueryC539128.ExecuteReader()
        Dim iC539128 As Short
        ReDim C539128DataType(RsC539128.FieldCount)
        For iC539128 = 0 to RsC539128.FieldCount - 1
            Select Case RsC539128.GetDataTypeName(iC539128).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C539128DataType(iC539128 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C539128DataType(iC539128 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C539128DataType(iC539128 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC539128
        RsC539128_EOF = Not RsC539128.Read()

        GoTo Comp_C539130

    Comp_C539130:

        ' ECommerce+Avista?
        sCurrComponent = "ECommerce+Avista?"
        C539130 = (fn_ConvertValueCompiled(RsC539128(0), C539128DataType(1), AKB_DecimalPoint, False) > 0)
        C539130DataType = AKBTypeConst.cAKBTypeLogical
        If C539130 Then
            GoTo Comp_C549977
        Else
            GoTo Comp_C75152
        End If

    Comp_C539465:

        ' UtilizaSintegraWS
        sCurrComponent = "UtilizaSintegraWS"
        C539465DataType = 0
        C539465DataType = C482998Datatype(3)
        If C539465DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC482998(2)) Then
          C539465 = Strings.RTrim(RsC482998(2))
        Else 
          C539465 = RsC482998(2)
        End If 

        GoTo Comp_C483009

    Comp_C539466:

        ' SintegraWS?
        sCurrComponent = "SintegraWS?"
        C539466 = (fn_ConvertValueCompiled(C539465, C539465DataType, AKB_DecimalPoint, False) = 1)
        C539466DataType = AKBTypeConst.cAKBTypeLogical
        If C539466 Then
            GoTo Comp_C539467
        Else
            GoTo Comp_C483015
        End If

    Comp_C539467:

        ' CodCliente
        sCurrComponent = "CodCliente"
        QueryC539467 = con.CreateCommand()
        QueryC539467.CommandText = QueryC539467.CommandText & " " & "SELECT COD_CLIENTE FROM WF_PRE_PEDIDO"
        QueryC539467.CommandText = QueryC539467.CommandText & " " & "WHERE id_prepedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC539467.Transaction = txn
        RsC539467 = QueryC539467.ExecuteReader()
        Dim iC539467 As Short
        ReDim C539467DataType(RsC539467.FieldCount)
        For iC539467 = 0 to RsC539467.FieldCount - 1
            Select Case RsC539467.GetDataTypeName(iC539467).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C539467DataType(iC539467 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C539467DataType(iC539467 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C539467DataType(iC539467 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC539467
        RsC539467_EOF = Not RsC539467.Read()

        GoTo Comp_C539468

    Comp_C539468:

        ' ValidaSintegraWS
        sCurrComponent = "ValidaSintegraWS"
        'C539468 = clsRuleDynamicallyCompiled_R22550.R22550(con, txn, IIf(Not IsDBNull(RsC539467(0)), RsC539467(0), System.DBNull.Value), IIf(Not IsDBNull(P76804), P76804, System.DBNull.Value), System.DBNull.Value)
        C539468CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C539468) Then
          iColumns = C539468.Columns.Count
        End If
        ReDim C539468DataType(iColumns)
        For iCol = 1 To iColumns
            'C539468DataType(iCol) = clsRuleDynamicallyCompiled_R22550.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C539541

    Comp_C539541:

        ' IEAtiva?
        sCurrComponent = "IEAtiva?"
        C539541 = (fn_ConvertValueCompiled(C539468.rows(C539468CurrentRow - 1)(0), C539468DataType(1), AKB_DecimalPoint, False) = "Ativo")
        C539541DataType = AKBTypeConst.cAKBTypeLogical
        If C539541 Then
            GoTo Comp_C483018
        Else
            GoTo Comp_C541305
        End If

    Comp_C539543:

        ' upIEBloqueada
        sCurrComponent = "upIEBloqueada"
        QueryC539543 = con.CreateCommand()
        QueryC539543.CommandText = QueryC539543.CommandText & " " & "UPDATE wf_pre_pedido"
        QueryC539543.CommandText = QueryC539543.CommandText & " " & "SET IE_BLOQUEADA = 1"
        QueryC539543.CommandText = QueryC539543.CommandText & " " & "WHERE ID_PREPEDIDO = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC539543.Transaction = txn
        C539543 = QueryC539543.ExecuteNonQuery()
        C539543DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C483015

    Comp_C540756:

        ' SiglaEstado
        sCurrComponent = "SiglaEstado"
        C540756DataType = 0
        C540756DataType = C483010Datatype(2)
        If C540756DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC483010(1)) Then
          C540756 = Strings.RTrim(RsC483010(1))
        Else 
          C540756 = RsC483010(1)
        End If 

        GoTo Comp_C483012

    Comp_C540774:

        ' UtilizaSintegraWS?
        sCurrComponent = "UtilizaSintegraWS?"
        C540774 = (fn_ConvertValueCompiled(C539465, C539465DataType, AKB_DecimalPoint, False) = 1)
        C540774DataType = AKBTypeConst.cAKBTypeLogical
        If C540774 Then
            GoTo Comp_C540775
        Else
            GoTo Comp_C482996
        End If

    Comp_C540775:

        ' SelDtTransp
        sCurrComponent = "SelDtTransp"
        QueryC540775 = con.CreateCommand()
        QueryC540775.CommandText = QueryC540775.CommandText & " " & "SELECT add_months(MAX(T.Data), " & _ 
ConvertValue(C483007, C483007DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " ), C.SIGLA_ESTADO, P.cod_redespacho"
        QueryC540775.CommandText = QueryC540775.CommandText & " " & "FROM wf_transp_sintegra_dt T, WF_PRE_PEDIDO P, WF_PESSOAS PES, WF_CIDADES C"
        QueryC540775.CommandText = QueryC540775.CommandText & " " & "WHERE T.COD_TRANSP (+) = P.cod_redespacho"
        QueryC540775.CommandText = QueryC540775.CommandText & " " & "AND P.ID_PREPEDIDO = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC540775.CommandText = QueryC540775.CommandText & " " & "AND P.COD_CLIENTE = PES.COD_PESSOA"
        QueryC540775.CommandText = QueryC540775.CommandText & " " & "AND PES.codigo_cidade = c.codigo_cidade"
        QueryC540775.CommandText = QueryC540775.CommandText & " " & "GROUP BY  C.SIGLA_ESTADO,  P.cod_redespacho"
        QueryC540775.Transaction = txn
        RsC540775 = QueryC540775.ExecuteReader()
        Dim iC540775 As Short
        ReDim C540775DataType(RsC540775.FieldCount)
        For iC540775 = 0 to RsC540775.FieldCount - 1
            Select Case RsC540775.GetDataTypeName(iC540775).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C540775DataType(iC540775 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C540775DataType(iC540775 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C540775DataType(iC540775 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC540775
        RsC540775_EOF = Not RsC540775.Read()

        GoTo Comp_C540777

    Comp_C540776:

        ' SiglaEstadoRedesp
        sCurrComponent = "SiglaEstadoRedesp"
        C540776DataType = 0
        C540776DataType = C540775Datatype(2)
        If C540776DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC540775(1)) Then
          C540776 = Strings.RTrim(RsC540775(1))
        Else 
          C540776 = RsC540775(1)
        End If 

        GoTo Comp_C540780

    Comp_C540777:

        ' DtSintegraRedesp
        sCurrComponent = "DtSintegraRedesp"
        C540777DataType = 0
        C540777 = RsC540775(0)
        C540777DataType = C540775Datatype(1)
        If C540777DataType = AKBTypeConst.cAKBTypeString Then
          C540777 = IIF(IsDBNull(C540777), C540777, Strings.RTrim(C540777))
        End If 

        GoTo Comp_C540776

    Comp_C540778:

        ' DtSintegra?
        sCurrComponent = "DtSintegra?"
        C540778 = (fn_ConvertValueCompiled(C483008, C483008DataType, AKB_DecimalPoint, False) = 1)
        C540778DataType = AKBTypeConst.cAKBTypeLogical
        If C540778 Then
            GoTo Comp_C540909
        Else
            GoTo Comp_C483016
        End If

    Comp_C540779:

        ' DtRedesp?
        sCurrComponent = "DtRedesp?"
        C540779 = (fn_ConvertValueCompiled(C540777, C540777DataType, AKB_DecimalPoint, False)="" or fn_ConvertValueCompiled(C540777, C540777DataType, AKB_DecimalPoint, False) Is System.DBNull.Value)
        C540779DataType = AKBTypeConst.cAKBTypeLogical
        If C540779 Then
            GoTo Comp_C540781
        Else
            GoTo Comp_C540786
        End If

    Comp_C540780:

        ' CodRedesp
        sCurrComponent = "CodRedesp"
        C540780DataType = 0
        C540780DataType = C540775Datatype(3)
        If C540780DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC540775(2)) Then
          C540780 = Strings.RTrim(RsC540775(2))
        Else 
          C540780 = RsC540775(2)
        End If 

        GoTo Comp_C540778

    Comp_C540781:

        ' ValidaSintegraRedesp
        sCurrComponent = "ValidaSintegraRedesp"
        'C540781 = clsRuleDynamicallyCompiled_R22550.R22550(con, txn, IIf(Not IsDBNull(C540780), C540780, System.DBNull.Value), IIf(Not IsDBNull(P76804), P76804, System.DBNull.Value), fn_ConvertValueCompiled( "1", 4, AKB_DecimalPoint))
        C540781CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C540781) Then
          iColumns = C540781.Columns.Count
        End If
        ReDim C540781DataType(iColumns)
        For iCol = 1 To iColumns
            'C540781DataType(iCol) = clsRuleDynamicallyCompiled_R22550.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C540782

    Comp_C540782:

        ' Sintegra Ativo?
        sCurrComponent = "Sintegra Ativo?"
        C540782 = (fn_ConvertValueCompiled(C540781.rows(C540781CurrentRow - 1)(0), C540781DataType(1), AKB_DecimalPoint, False)  = "Ativo")
        C540782DataType = AKBTypeConst.cAKBTypeLogical
        If C540782 Then
            GoTo Comp_C483016
        Else
            GoTo Comp_C541306
        End If

    Comp_C540786:

        ' Vencido?
        sCurrComponent = "Vencido?"
        C540786 = (fn_ConvertValueCompiled(C483009, C483009DataType, AKB_DecimalPoint, False) > fn_ConvertValueCompiled(C540777, C540777DataType, AKB_DecimalPoint, False))
        C540786DataType = AKBTypeConst.cAKBTypeLogical
        If C540786 Then
            GoTo Comp_C540781
        Else
            GoTo Comp_C483016
        End If

    Comp_C540909:

        ' EstadoWS
        sCurrComponent = "EstadoWS"
        C540909 = (1 = 1)
        C540909DataType = AKBTypeConst.cAKBTypeLogical
        If C540909 Then
            GoTo Comp_C540779
        Else
            GoTo Comp_C482996
        End If

    Comp_C541305:

        ' ERRO?
        sCurrComponent = "ERRO?"
        C541305 = (fn_ConvertValueCompiled(C539468.rows(C539468CurrentRow - 1)(0), C539468DataType(1), AKB_DecimalPoint, False) = "ERRO")
        C541305DataType = AKBTypeConst.cAKBTypeLogical
        If C541305 Then
            GoTo Comp_C483023
        Else
            GoTo Comp_C539543
        End If

    Comp_C541306:

        ' erro2?
        sCurrComponent = "erro2?"
        C541306 = (fn_ConvertValueCompiled(C540781.rows(C540781CurrentRow - 1)(0), C540781DataType(1), AKB_DecimalPoint, False) = "ERRO")
        C541306DataType = AKBTypeConst.cAKBTypeLogical
        If C541306 Then
            GoTo Comp_C483023
        Else
            GoTo Comp_C482996
        End If

    Comp_C545148:

        ' PrazoEcommerce
        sCurrComponent = "PrazoEcommerce"
        C545148 = clsRuleDynamicallyCompiled_R22703.R22703(con, txn, IIf(Not IsDBNull(C90294), C90294, System.DBNull.Value), IIf(Not IsDBNull(C56126), C56126, System.DBNull.Value))
        C545148CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C545148) Then
          iColumns = C545148.Columns.Count
        End If
        ReDim C545148DataType(iColumns)
        For iCol = 1 To iColumns
            C545148DataType(iCol) = clsRuleDynamicallyCompiled_R22703.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C539128

    Comp_C546207:

        ' Valida End Danfe
        sCurrComponent = "Valida End Danfe"
        C546207 = clsRuleDynamicallyCompiled_R22737.R22737(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value))
        C546207CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C546207) Then
          iColumns = C546207.Columns.Count
        End If
        ReDim C546207DataType(iColumns)
        For iCol = 1 To iColumns
            C546207DataType(iCol) = clsRuleDynamicallyCompiled_R22737.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C546213

    Comp_C546213:

        ' validacao end danfe 1
        sCurrComponent = "validacao end danfe 1"
        C546213 = (fn_ConvertValueCompiled(C546207.rows(C546207CurrentRow - 1)(0), C546207DataType(1), AKB_DecimalPoint, False) = 1)
        C546213DataType = AKBTypeConst.cAKBTypeLogical
        If C546213 Then
            GoTo Comp_C394342
        Else
            GoTo Comp_C546214
        End If

    Comp_C546214:

        ' RET14
        sCurrComponent = "RET14"
        Dim tb_C546214 As DataTable = New DataTable()
        tb_C546214.Columns.Add("validacao end danfe 1" & "")
        Dim row_C546214 As DataRow = tb_C546214.NewRow()
        row_C546214(0) = C546213
        tb_C546214.Rows.Add(row_C546214)
        R4557 = tb_C546214
        ReDim C546214DataType(1)
        C546214DataType(1) = C546213DataType
        ReturnDataType = C546214DataType

        GoTo Exit_R4557

    Comp_C549977:

        ' ValidaECommerce
        sCurrComponent = "ValidaECommerce"
        'C549977 = clsRuleDynamicallyCompiled_R22831.R22831(con, txn, IIf(Not IsDBNull(C56908), C56908, System.DBNull.Value))
        C549977CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C549977) Then
          iColumns = C549977.Columns.Count
        End If
        ReDim C549977DataType(iColumns)
        For iCol = 1 To iColumns
            'C549977DataType(iCol) = clsRuleDynamicallyCompiled_R22831.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C109297

    Comp_C551587:

        ' Sintegra Cliente OK?
        sCurrComponent = "Sintegra Cliente OK?"
        C551587 = ((fn_ConvertValueCompiled(P88682, 4, AKB_DecimalPoint, False) = 1 AND fn_ConvertValueCompiled(P89539, 4, AKB_DecimalPoint, False) = 1)  OR fn_ConvertValueCompiled(P78985, 1, AKB_DecimalPoint, False) = 1 OR fn_ConvertValueCompiled(P26981, 4, AKB_DecimalPoint, False) = 1)
        C551587DataType = AKBTypeConst.cAKBTypeLogical
        If C551587 Then
            GoTo Comp_C470977
        Else
            GoTo Comp_C551588
        End If

    Comp_C551588:

        ' MsgSintegra
        sCurrComponent = "MsgSintegra"
        Dim row_C551588 As DataRow = msg.NewRow()
        row_C551588(0) = "O Sintegra do Cliente ou Redespacho ainda não esta atualizado, aguarde atualização ou chame o suporte para maiores informações." & Chr(13) & "" & Chr(10) & "" & Chr(13) & "" & Chr(10) & "O pedido não será gerado."
        msg.Rows.Add(row_C551588)
        C551588DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C551589

    Comp_C551589:

        ' ATRIBUIVA61
        sCurrComponent = "ATRIBUIVA61"
        C551589DataType = 4
        C516413 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C551589 = True
        GoTo Comp_C516419

    Comp_C551599:

        ' SintegraWS
        sCurrComponent = "SintegraWS"
        QueryC551599 = con.CreateCommand()
        QueryC551599.CommandText = QueryC551599.CommandText & " " & "SELECT NVL(WF_ESTAB_JURIDICO.utiliza_sintegraws,0)"
        QueryC551599.CommandText = QueryC551599.CommandText & " " & "from WF_ESTAB_JURIDICO"
        QueryC551599.CommandText = QueryC551599.CommandText & " " & "where Cod_Pessoa_Estab_Juridico = " & _ 
ConvertValue(P76804, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC551599.Transaction = txn
        RsC551599 = QueryC551599.ExecuteReader()
        Dim iC551599 As Short
        ReDim C551599DataType(RsC551599.FieldCount)
        For iC551599 = 0 to RsC551599.FieldCount - 1
            Select Case RsC551599.GetDataTypeName(iC551599).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C551599DataType(iC551599 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C551599DataType(iC551599 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C551599DataType(iC551599 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC551599
        RsC551599_EOF = Not RsC551599.Read()

        GoTo Comp_C516413

    Comp_C551600:

        ' ValidaSintegra?
        sCurrComponent = "ValidaSintegra?"
        C551600 = (fn_ConvertValueCompiled(RsC551599(0), C551599DataType(1), AKB_DecimalPoint, False) = 1)
        C551600DataType = AKBTypeConst.cAKBTypeLogical
        If C551600 Then
            GoTo Comp_C551587
        Else
            GoTo Comp_C470977
        End If

    Comp_C555064:

        ' SelParamTabVenda
        sCurrComponent = "SelParamTabVenda"
        QueryC555064 = con.CreateCommand()
        QueryC555064.CommandText = QueryC555064.CommandText & " " & "select nvl(MAX(tabela_base_geral),0) from wf_tabela_preco_venda, wf_pre_pedido"
        QueryC555064.CommandText = QueryC555064.CommandText & " " & "where wf_tabela_preco_venda.tabela_preco_venda = wf_pre_pedido.tabela_preco_venda"
        QueryC555064.CommandText = QueryC555064.CommandText & " " & "and wf_pre_pedido.id_prepedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555064.Transaction = txn
        RsC555064 = QueryC555064.ExecuteReader()
        Dim iC555064 As Short
        ReDim C555064DataType(RsC555064.FieldCount)
        For iC555064 = 0 to RsC555064.FieldCount - 1
            Select Case RsC555064.GetDataTypeName(iC555064).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C555064DataType(iC555064 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C555064DataType(iC555064 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C555064DataType(iC555064 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC555064
        RsC555064_EOF = Not RsC555064.Read()

        GoTo Comp_C555065

    Comp_C555065:

        ' NãoValidaParamTab
        sCurrComponent = "NãoValidaParamTab"
        C555065 = (fn_ConvertValueCompiled(RsC555064(0), C555064DataType(1), AKB_DecimalPoint, False) = 0)
        C555065DataType = AKBTypeConst.cAKBTypeLogical
        If C555065 Then
            GoTo Comp_C470976
        Else
            GoTo Comp_C555070
        End If

    Comp_C555066:

        ' Valida Preços
        sCurrComponent = "Valida Preços"
        C555066 = clsRuleDynamicallyCompiled_R23053.R23053(con, txn, IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value), fn_ConvertValueCompiled( "1", 1, AKB_DecimalPoint), System.DBNull.Value, System.DBNull.Value, System.DBNull.Value)
        C555066CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C555066) Then
          iColumns = C555066.Columns.Count
        End If
        ReDim C555066DataType(iColumns)
        For iCol = 1 To iColumns
            C555066DataType(iCol) = clsRuleDynamicallyCompiled_R23053.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C555067

    Comp_C555067:

        ' SelBloqPreço
        sCurrComponent = "SelBloqPreço"
        QueryC555067 = con.CreateCommand()
        QueryC555067.CommandText = QueryC555067.CommandText & " " & "select nvl(bloqueado_preco_fora,0)"
        QueryC555067.CommandText = QueryC555067.CommandText & " " & "from wf_pre_pedido"
        QueryC555067.CommandText = QueryC555067.CommandText & " " & "where id_prepedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555067.Transaction = txn
        RsC555067 = QueryC555067.ExecuteReader()
        Dim iC555067 As Short
        ReDim C555067DataType(RsC555067.FieldCount)
        For iC555067 = 0 to RsC555067.FieldCount - 1
            Select Case RsC555067.GetDataTypeName(iC555067).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C555067DataType(iC555067 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C555067DataType(iC555067 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C555067DataType(iC555067 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC555067
        RsC555067_EOF = Not RsC555067.Read()

        GoTo Comp_C555068

    Comp_C555068:

        ' BloqueadoPreço?
        sCurrComponent = "BloqueadoPreço?"
        C555068 = (fn_ConvertValueCompiled(RsC555067(0), C555067DataType(1), AKB_DecimalPoint, False) = 1)
        C555068DataType = AKBTypeConst.cAKBTypeLogical
        If C555068 Then
            GoTo Comp_C555069
        Else
            GoTo Comp_C470976
        End If

    Comp_C555069:

        ' MSG7
        sCurrComponent = "MSG7"
        Dim row_C555069 As DataRow = msg.NewRow()
        row_C555069(0) = "Pré pedido bloqueado para geração, pois existem itens com preço fora da parametrização."
        msg.Rows.Add(row_C555069)
        C555069DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C516419

    Comp_C555070:

        ' SelPedLiberado
        sCurrComponent = "SelPedLiberado"
        QueryC555070 = con.CreateCommand()
        QueryC555070.CommandText = QueryC555070.CommandText & " " & "select nvl(usuario_liber_preco,0)"
        QueryC555070.CommandText = QueryC555070.CommandText & " " & "from wf_pre_pedido"
        QueryC555070.CommandText = QueryC555070.CommandText & " " & "where id_prepedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555070.Transaction = txn
        RsC555070 = QueryC555070.ExecuteReader()
        Dim iC555070 As Short
        ReDim C555070DataType(RsC555070.FieldCount)
        For iC555070 = 0 to RsC555070.FieldCount - 1
            Select Case RsC555070.GetDataTypeName(iC555070).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C555070DataType(iC555070 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C555070DataType(iC555070 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C555070DataType(iC555070 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC555070
        RsC555070_EOF = Not RsC555070.Read()

        GoTo Comp_C555071

    Comp_C555071:

        ' PréPedidoLiberado?
        sCurrComponent = "PréPedidoLiberado?"
        C555071 = (fn_ConvertValueCompiled(RsC555070(0), C555070DataType(1), AKB_DecimalPoint, False) = 0)
        C555071DataType = AKBTypeConst.cAKBTypeLogical
        If C555071 Then
            GoTo Comp_C555080
        Else
            GoTo Comp_C470976
        End If

    Comp_C555080:

        ' SelBloqPreco
        sCurrComponent = "SelBloqPreco"
        QueryC555080 = con.CreateCommand()
        QueryC555080.CommandText = QueryC555080.CommandText & " " & "select nvl(bloqueado_preco_fora,0)"
        QueryC555080.CommandText = QueryC555080.CommandText & " " & "from wf_pre_pedido"
        QueryC555080.CommandText = QueryC555080.CommandText & " " & "where id_prepedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC555080.Transaction = txn
        RsC555080 = QueryC555080.ExecuteReader()
        Dim iC555080 As Short
        ReDim C555080DataType(RsC555080.FieldCount)
        For iC555080 = 0 to RsC555080.FieldCount - 1
            Select Case RsC555080.GetDataTypeName(iC555080).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C555080DataType(iC555080 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C555080DataType(iC555080 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C555080DataType(iC555080 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC555080
        RsC555080_EOF = Not RsC555080.Read()

        GoTo Comp_C555081

    Comp_C555081:

        ' PréPedidoBloqueado
        sCurrComponent = "PréPedidoBloqueado"
        C555081 = (fn_ConvertValueCompiled(RsC555080(0), C555080DataType(1), AKB_DecimalPoint, False) = 1)
        C555081DataType = AKBTypeConst.cAKBTypeLogical
        If C555081 Then
            GoTo Comp_C555069
        Else
            GoTo Comp_C555066
        End If

    Comp_C555428:

        ' CodServiço
        sCurrComponent = "CodServiço"
        C555428DataType = 0
        C555428DataType = C56123Datatype(33)
        If C555428DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC56123(32)) Then
          C555428 = Strings.RTrim(RsC56123(32))
        Else 
          C555428 = RsC56123(32)
        End If 

        GoTo Comp_C433219

    Comp_C556382:

        ' SelCountPedido
        sCurrComponent = "SelCountPedido"
        QueryC556382 = con.CreateCommand()
        QueryC556382.CommandText = QueryC556382.CommandText & " " & "SELECT COUNT(*) "
        QueryC556382.CommandText = QueryC556382.CommandText & " " & "FROM WF_PEDIDO"
        QueryC556382.CommandText = QueryC556382.CommandText & " " & "WHERE PEDIDO = " & _ 
ConvertValue(C90294, C90294DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC556382.Transaction = txn
        RsC556382 = QueryC556382.ExecuteReader()
        Dim iC556382 As Short
        ReDim C556382DataType(RsC556382.FieldCount)
        For iC556382 = 0 to RsC556382.FieldCount - 1
            Select Case RsC556382.GetDataTypeName(iC556382).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C556382DataType(iC556382 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C556382DataType(iC556382 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C556382DataType(iC556382 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC556382
        RsC556382_EOF = Not RsC556382.Read()

        GoTo Comp_C556383

    Comp_C556383:

        ' GravouPedido?
        sCurrComponent = "GravouPedido?"
        C556383 = (fn_ConvertValueCompiled(RsC556382(0), C556382DataType(1), AKB_DecimalPoint, False) > 0)
        C556383DataType = AKBTypeConst.cAKBTypeLogical
        If C556383 Then
            GoTo Comp_C119833
        Else
            GoTo Comp_C96585
        End If

    Comp_C563342:

        ' UpdateBloqueio
        sCurrComponent = "UpdateBloqueio"
        QueryC563342 = con.CreateCommand()
        QueryC563342.CommandText = QueryC563342.CommandText & " " & "UPDATE WF_PESSOAS SET cadastro_bloqueado = 0"
        QueryC563342.CommandText = QueryC563342.CommandText & " " & "WHERE exists (select * from wf_pre_pedido"
        QueryC563342.CommandText = QueryC563342.CommandText & " " & "where wf_pre_pedido.cod_cliente = WF_PESSOAS.cod_pessoa"
        QueryC563342.CommandText = QueryC563342.CommandText & " " & "AND WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12596, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " )"
        QueryC563342.Transaction = txn
        C563342 = QueryC563342.ExecuteNonQuery()
        C563342DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C483025

    Comp_C578759:

        ' FatTpEmpresa
        sCurrComponent = "FatTpEmpresa"
        C578759 = clsRuleDynamicallyCompiled_R23567.R23567(con, txn, IIf(Not IsDBNull(P76804), P76804, System.DBNull.Value), IIf(Not IsDBNull(P12596), P12596, System.DBNull.Value))
        C578759CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C578759) Then
          iColumns = C578759.Columns.Count
        End If
        ReDim C578759DataType(iColumns)
        For iCol = 1 To iColumns
            C578759DataType(iCol) = clsRuleDynamicallyCompiled_R23567.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C578779

    Comp_C578779:

        ' FatTpEmpresa=0?
        sCurrComponent = "FatTpEmpresa=0?"
        C578779 = (fn_ConvertValueCompiled(C578759.rows(C578759CurrentRow - 1)(0), C578759DataType(1), AKB_DecimalPoint, False) = 0)
        C578779DataType = AKBTypeConst.cAKBTypeLogical
        If C578779 Then
            GoTo Comp_C470978
        Else
            GoTo Comp_C516419
        End If

    Comp_C579427:

        ' FIM_Sel_Colec
        sCurrComponent = "FIM_Sel_Colec"
        C579427DataType = 4
        C579427 = RsC174609_EOF
        GoTo Comp_C174610

    Exit_R4557:

        Exit Function

    Err_R4557:
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
