'Option Strict On

'Public Class GuildRel
'	'Guild Rel can be another player or another guild
'	Public lEntityID As Int32
'	Public iEntityTypeID As Int16

'	Public sName As String = ""

'	Public yRelTowardsUs As Byte		'their rel towards us
'	Public yRelTowardsThem As Byte		'our rel towards them

'	Public lIcon As Int32		'icon for the player/guild

'	Public lWhoMadeFirstContact As Int32 = -1
'	Public lWhoFirstContactWasMadeWith As Int32 = -1
'	Public dtWhenFirstContactMade As Date = Date.MinValue		'in GMT
'	Public lLocationIDOfFC As Int32 = -1
'	Public iLocationTypeIDOfFC As Int16 = -1
'	Public lLocXOfFC As Int32 = 0
'	Public lLocZOfFC As Int32 = 0

'	Public sNotes As String = ""

'	'Guild-Specific
'	Public lKnownMemberCount As Int32
'	Public lLeaderID As Int32 = -1

'	'Player-Specific
'	Public lPlayerGuildID As Int32 = -1
'	Public yPlayerTitle As Byte = 0
'	Public bIsMale As Boolean = False

'	Private mbRequestedDetails As Boolean = False
'	Public Function RequestDetails() As Byte()
'		If mbRequestedDetails = False Then
'			mbRequestedDetails = True
'			'Ok request details
'			Dim yMsg(8) As Byte
'			Dim lPos As Int32 = 0
'			System.BitConverter.GetBytes(GlobalMessageCode.eGuildRequestDetails).CopyTo(yMsg, lPos) : lPos += 2
'			System.BitConverter.GetBytes(Me.lEntityID).CopyTo(yMsg, lPos) : lPos += 4
'			yMsg(lPos) = eyGuildRequestDetailsType.GuildRel : lPos += 1
'			System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, lPos) : lPos += 2
'			Return yMsg
'		End If
'		Return Nothing
'	End Function

'End Class
