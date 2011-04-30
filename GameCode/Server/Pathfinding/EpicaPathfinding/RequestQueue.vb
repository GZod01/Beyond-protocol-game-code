Public Class RequestQueue
    Private Const ml_DEFAULT_QUEUE_SIZE As Int32 = 500

    Private muQueueData() As RequestData
    Private myQueueDataUsed() As Byte
    Private mlQueueDataUB As Int32 = -1

    Public Sub New()
        mlQueueDataUB = ml_DEFAULT_QUEUE_SIZE - 1
        ReDim muQueueData(mlQueueDataUB)
        ReDim myQueueDataUsed(mlQueueDataUB)
    End Sub

    Protected Overrides Sub Finalize()
        mlQueueDataUB = -1
        Erase muQueueData
        Erase myQueueDataUsed
        MyBase.Finalize()
    End Sub

    Public Sub AddNewRequest(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal lLocX As Int32, _
      ByVal lLocZ As Int32, ByVal iLocA As Int16, ByVal lLocID As Int32, ByVal iLocTypeID As Int32, _
      ByVal iModelID As Int16, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestA As Int16, _
      ByVal lDestID As Int32, ByVal iDestTypeID As Int16)
        Dim lIdx As Int32 = -1
        Dim X As Int32

        For X = 0 To mlQueueDataUB
            If myQueueDataUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            ReDim Preserve muQueueData(mlQueueDataUB + ml_DEFAULT_QUEUE_SIZE)
            ReDim Preserve myQueueDataUsed(mlQueueDataUB + ml_DEFAULT_QUEUE_SIZE)
            lIdx = mlQueueDataUB + 1
            mlQueueDataUB += ml_DEFAULT_QUEUE_SIZE
        End If

        With muQueueData(lIdx)
            .DestA = iDestA
            .DestID = lDestID
            .DestTypeID = iDestTypeID
            .DestX = lDestX
            .DestZ = lDestZ
            .LocA = iLocA
            .LocID = lLocID
			.LocTypeID = CShort(iLocTypeID)
            .LocX = lLocX
            .LocZ = lLocZ
            .ModelID = iModelID
            .ObjectID = lObjID
            .ObjTypeID = iObjTypeID
        End With
        myQueueDataUsed(lIdx) = 255
    End Sub

    Public Function GetNextRequest(ByRef bFound As Boolean) As RequestData
        bFound = False
        For X As Int32 = 0 To mlQueueDataUB
            If myQueueDataUsed(X) <> 0 Then
                myQueueDataUsed(X) = 0
                bFound = True
                Return muQueueData(X)
            End If
        Next X

        'Ok, now... see if we can clear up some space since nothing is left...
        If (mlQueueDataUB + 1) \ ml_DEFAULT_QUEUE_SIZE > 1 Then
            'Yes, we can... so compact the queue down to the default queue size
            mlQueueDataUB = ml_DEFAULT_QUEUE_SIZE - 1
            ReDim muQueueData(mlQueueDataUB)
            ReDim myQueueDataUsed(mlQueueDataUB)
        End If
    End Function
End Class

Public Structure RequestData
    Public ObjectID As Int32
    Public ObjTypeID As Int16
    Public LocX As Int32
    Public LocZ As Int32
    Public LocA As Int16
    Public LocID As Int32
    Public LocTypeID As Int16
    Public ModelID As Int16

    Public DestX As Int32
    Public DestZ As Int32
    Public DestA As Int16
    Public DestID As Int32
    Public DestTypeID As Int16
End Structure