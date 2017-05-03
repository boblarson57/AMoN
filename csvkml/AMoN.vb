Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Module AMoN
    Friend Const STR_CONN As String = "Server=datastorm.sws.uiuc.edu; Data Source=dbnew; Initial Catalog=nadpweb;integrated security=true; Min Pool Size=20;"

    ' Friend Const STR_SQL_SELECT_EMPTY_RS As String = "Select * from pptLoad where 1=2"
    Friend Const stDataDir As String = "\\carfree\INCOMINGPPT\"
    Friend Const ARCHIVEDIR As String = "\\nadp30\incomingppt\"
    Friend Const STAGEDIR As String = "\\nadppo1\pptincoming\"
    Friend Const targetDir As String = "\\nadppo1.sws.uiuc.edu\root3\amon\sites\"

    Sub Main()
        getsites()
    End Sub
    Public Sub CreateCSVfile(ByVal dtable As DataTable, ByVal strFilePath As String)
        Dim dt As DateTimeOffset
        Dim sw As New StreamWriter(strFilePath, False)
        Dim test As String
        Dim icolcount As Integer = dtable.Columns.Count
        For i As Integer = 0 To icolcount - 1
            sw.Write(dtable.Columns(i))
            If i < icolcount - 1 Then
                sw.Write(",")
            End If
        Next
        sw.Write(sw.NewLine)
        For Each drow As DataRow In dtable.Rows
            For i As Integer = 0 To icolcount - 1
                If Not Convert.IsDBNull(drow(i)) Then
                    If Right(dtable.Columns(i).ToString.ToUpper, 4) = "DATE" Then
                        dt = drow(i).ToString()
                        sw.Write(dt.ToUniversalTime.ToString("yyyy-MM-dd HH:mm"))
                        test = dt.ToString("yyyy-MM-dd HH:mm")
                    Else
                        sw.Write(drow(i).ToString())
                    End If
                    test = dtable.Columns(i).ToString + "= " + drow(i).ToString()
                End If
                If i < icolcount - 1 Then
                    sw.Write(",")
                End If
            Next
            sw.Write(sw.NewLine)
        Next
        sw.Close()
        dtable.WriteXml(Replace(strFilePath, "csv", "xml"))
    End Sub
    Public Function getsites()

        Dim connWeb1 As SqlConnection
        Dim myCommand As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable
        'create objects

        Try
            load("ALL", "XML", "AVE")
            load("ALL", "XML", "REP")
            load("ALL", "CSV", "AVE")
            load("ALL", "CSV", "REP")

            connWeb1 = New SqlConnection
            connWeb1.ConnectionString = STR_CONN
            connWeb1.Open()
            myCommand = New SqlCommand("select siteid from tblsites where network = 'AMoN' and status in ('A','a','I','i')", connWeb1)
            myCommand.CommandType = CommandType.Text

            reader = myCommand.ExecuteReader
            Do While reader.Read()
                load(reader("siteid"), "XML", "AVE")
                load(reader("siteid"), "XML", "REP")
                load(reader("siteid"), "CSV", "AVE")
                load(reader("siteid"), "CSV", "REP")
            Loop
            '    load("%")
            reader.Close()
            connWeb1.Close()
        Catch ex As Exception
            Console.WriteLine("Error: " + ex.Message + ex.StackTrace)
        Finally

        End Try

    End Function

    Public Sub load(ByVal siteID As String, ByVal format As String, ByVal aveORrep As String)

        Dim connWeb As SqlConnection
        Dim myCommand As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable
        'create objects

        Try
            connWeb = New SqlConnection
            connWeb.ConnectionString = STR_CONN
            connWeb.Open()
            If aveORrep = "AVE" Then
                myCommand = New SqlCommand("spAMONWebreportAverages", connWeb)
            Else
                myCommand = New SqlCommand("spAMONWebreport", connWeb)
            End If
            If siteID <> "ALL" Then
                myCommand.Parameters.Add("@site", SqlDbType.Char, 4)
                myCommand.Parameters("@site").Value = siteID
            End If
            If format = "XML" Then
                myCommand.Parameters.AddWithValue("@GMT", 0) 'keep local with offset
            Else
                myCommand.Parameters.AddWithValue("@GMT", 1) 'change to gmt
            End If

            myCommand.CommandType = CommandType.StoredProcedure

            reader = myCommand.ExecuteReader

            If reader.HasRows Then
                dt.Load(reader)
                dt.TableName = "AMoN"
                If aveORrep = "AVE" Then
                    If format = "XML" Then
                        CreateXMLfile(dt, targetDir + "xml\" + siteID + "-ave.xml")
                    Else
                        CreateCSVfile(dt, targetDir + "csv\" + siteID + "-ave.csv")
                    End If
                Else
                    If format = "XML" Then
                        CreateXMLfile(dt, targetDir + "xml\" + siteID + "-rep.xml")
                    Else
                        CreateCSVfile(dt, targetDir + "csv\" + siteID + "-rep.csv")
                    End If
                End If
            End If
            reader.Close()
            connWeb.Close()
        Catch ex As Exception
            Console.WriteLine("Error: " + ex.Message + ex.StackTrace)
        Finally

        End Try

    End Sub




    Public Sub CreateXMLfile(ByVal dtable As DataTable, ByVal strFilePath As String)

        dtable.WriteXml(strFilePath)
    End Sub



End Module
