Namespace Pop3
    Public Class Pop3LoginException
        Inherits Exception
        Private m_exceptionString As String

        Public Sub New()
            MyBase.New()
            m_exceptionString = Nothing
        End Sub

        Public Sub New(ByVal exceptionString As String)
            MyBase.New()
            m_exceptionString = exceptionString
        End Sub

        Public Sub New(ByVal exceptionString As String, ByVal ex As Exception)
            MyBase.New(exceptionString, ex)
        End Sub

        Public Overrides Function ToString() As String
            Return m_exceptionString
        End Function
    End Class
End Namespace
