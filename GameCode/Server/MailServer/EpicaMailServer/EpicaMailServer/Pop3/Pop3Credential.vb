Namespace Pop3
    ''' <summary>
    ''' Summary description for Credentials.
    ''' </summary>
    Public Class Pop3Credential
        Private m_user As String
        Private m_pass As String
        Private m_server As String

        Private m_sendStrings As String() = {"user", "pass"}

        Public ReadOnly Property SendStrings() As String()
            Get
                Return m_sendStrings
            End Get
        End Property

        Public Property User() As String
            Get
                Return m_user
            End Get
            Set(ByVal value As String)
                m_user = value
            End Set
        End Property

        Public Property Pass() As String
            Get
                Return m_pass
            End Get
            Set(ByVal value As String)
                m_pass = value
            End Set
        End Property

        Public Property Server() As String
            Get
                Return m_server
            End Get
            Set(ByVal value As String)
                m_server = value
            End Set
        End Property

        Public Sub New(ByVal user As String, ByVal pass As String, ByVal server As String)
            m_user = user
            m_pass = pass
            m_server = server
        End Sub

        Public Sub New()
            m_user = Nothing
            m_pass = Nothing
            m_server = Nothing
        End Sub
    End Class
End Namespace
