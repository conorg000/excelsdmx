Attribute VB_Name = "SDMX"
'/* Copyright 2010,2014 Bank Of Italy
'*
'* Licensed under the EUPL, Version 1.1 or - as soon they
'* will be approved by the European Commission - subsequent
'* versions of the EUPL (the "Licence");
'* You may not use this work except in compliance with the
'* Licence.
'* You may obtain a copy of the Licence at:
'*
'*
'* http://ec.europa.eu/idabc/eupl
'*
'* Unless required by applicable law or agreed to in
'* writing, software distributed under the Licence is
'* distributed on an "AS IS" basis,
'* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either
'* express or implied.
'* See the Licence for the specific language governing
'* permissions and limitations under the Licence.
'*/

Public Sub getTimeSeries(dataflow As String)
'## environment for system call

    Dim ws As Excel.Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    Set ws = Excel.ActiveSheet
    ws.name = dataflow

'## clean-up sheet
    With Sheets(ws.name)
        .rows(2 & ":" & .rows.Count).ClearContents
    End With

'## get provider and query from cells. if not present prompt user
    Dim provider As String
    Dim query As String
    Dim startTime As String
    Dim endTime As String
    startTime = """"""  ' blank
    endTime = """"""    ' blank

    provider = "SPC"
    query = dataflow
    ' time period is not yet supported, but would be a pretty straightforward addition
    'startTime = "2000"
    'endTime = "2005"

'## build the command
    Dim command As String
    'Example cmd: curl -i -H "Accept: application/vnd.sdmx.data+csv" "https://stats.pacificdata.org/data-nsi/Rest/data/DF_POP_SUM/?startPeriod=2000&endPeriod=2005&format=csv"
    command = "curl -H ""Accept: application/vnd.sdmx.data+csv"" GET ""https://stats.pacificdata.org/data-nsi/Rest/data/" & query & "/?format=csv"""
'## execute the command
    Dim wsh As Object
    'We start command line here
    Set wsh = VBA.CreateObject("WScript.Shell")
    'Call command from cmd line
    Set objExec = wsh.Exec(command)
    Dim result As String
    Dim error As String
    Dim rc As Integer
    rc = objExec.Status
    result = objExec.StdOut.ReadAll
    error = objExec.StdErr.ReadAll
    If result = "" Then
        MsgBox Right(error, 1024)
        Exit Sub
    Else
'## parse results
        'MsgBox result
        Dim rows As Variant
        rows = Split(result, vbLf, -1)
        Dim tsSize As Integer
        tsSize = UBound(rows) + 1
        Dim i As Integer
        i = 0
        Do While i < tsSize
            Dim fields As Variant
            fields = Split(rows(i), ",", -1)
            Dim fieldSize As Integer
            fieldSize = UBound(fields) + 1
            Dim j As Integer
            j = 0
            Do While j < fieldSize
                Dim n As Integer
                Dim m As Integer
                n = 2 + i
                m = 1 + j
                ws.Cells(n, m).Value = fields(j)
                j = j + 1
            Loop
            i = i + 1
        Loop
    End If
End Sub
