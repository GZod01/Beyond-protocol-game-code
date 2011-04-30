Public Class PermSave
    'This class is responsible for saving the contents of an object...
    '  I use this class as opposed to a standard methodology throughout the program
    '  so that if we ever change the method of how an object is saved, we simply change the
    '  code here... every object will be responsible for producing what this class needs
    '  in order to save the object.

    Private moCN As OleDb.OleDbConnection

    Public Sub New()
        'go ahead and connect
        Dim sConnStr As String
        Dim oINI As New InitFile()
        Dim sUDL As String

        sUDL = oINI.GetString("SETTINGS", "Perm_Save_UDL", "")
        If sUDL <> "" Then
            sConnStr = "FILE NAME=" & sUDL
            moCN = New OleDb.OleDbConnection(sConnStr)
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If Not moCN Is Nothing Then
            moCN.Close()
        End If
        moCN = Nothing
        MyBase.Finalize()
    End Sub

    Public Function SaveObject(ByVal sData As String) As Boolean
        'Here, we save the object
        'sdata is currently assumed to be a SQL statement
        Dim bResult As Boolean
        Dim oComm As New OleDb.OleDbCommand(sData, moCN)

        On Error GoTo NormalExit

        bResult = False
        oComm.ExecuteNonQuery()
        bResult = True

NormalExit:
        oComm = Nothing
        Return bResult

    End Function
End Class
