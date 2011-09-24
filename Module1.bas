Attribute VB_Name = "Module1"
'''Varíaveis que manipulam registros nas tabelas do BD "players"
Global tabmusic As New ADODB.Recordset      'tabela musicas
Global tabgenero As New ADODB.Recordset     'tabela generos
Global tablexe As New ADODB.Recordset       'tabela listas de execuções
Global tablrepro As New ADODB.Recordset     'tabela listas de reproduções
Global tabautores As New ADODB.Recordset    'tabela autores
Global tabskin As New ADODB.Recordset       'tabela Skins
''''''''' Conectar a bd ''''''''''''''''''''''''''''
Global conectar As New ADODB.Connection
Global caminho As String
Global capsula As String

Option Explicit
Function abrir_banco()
            If conectar.State = 1 Then conectar.Close
                capsula = "Provider=microsoft.jet.oledb.4.0;data source="
                caminho = capsula + App.Path & "\Players.mdb"
                conectar.Open (caminho)
End Function
Function skin()
            Call abrir_banco
                If tabskin.State = 1 Then tabskin.Close
                
End Function
