VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_client2 
   Caption         =   "Client 2 (Lissens for Client 1)"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock Winsocket 
      Left            =   4560
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_setnew 
      Caption         =   "Set"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txt_newnick 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txt_dialog 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lbl_dialog 
      Caption         =   "Dialog of the Chat"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lbl_set 
      Caption         =   "Set Your NickName"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lbl_mynick 
      Caption         =   "Your NickName is:"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lbl_other 
      Caption         =   "You are talking to:"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frm_client2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'HOW TO ORGANISE COMMUNICTION IN WINSOCK ARTICAL HELPER PROJECT
'By: David Meirion Hughes. AKA CrAKiN-ShOt

'If you still are unsure about anythink in this code then please dont hessitatle to email me: crakinshot@hotmail.com

Option Explicit ' This is used to make sure we declare varables.
                ' With it, if you try to set a varable to a value but
                ' have not declared the varble then the compiler (VB)
                ' Will tell you that the varable has not been declared.
                

Private Enum DataTypes 'Start a New Enum Class
    Message = 0 ' Make a DataType called Message with the value of 0.
    NickName = 1 ' Make a DataType called NickName with the Value of 1
End Enum ' Stop the Enum Definition

Public Myname As String
Public othername As String

Private Sub Send_Data(dType As DataTypes, dData As String)
'Declare the New Sub Routine and make it do that Data Does not have to be sent
'We will also Declare the Type of Data as our ENUM DATATYPES
'USAGE EXAMPLE: call Send_data(Message,"hi There")

'as you cansee the Type is being set as Message when in fact its value is 0.
'like we defined in the ENUM definition

Dim oData As String ' the outgoing Data
Dim oType As String ' the outgoing data type
Dim oAll As String  ' All the Outgoing data including its type

oType = dType 'Make the outgoing varable the same as the one given to the sub routine

'Now to make sure the Type in the Packet takes up 2 Spaces we find out if the Type Number
'is less than 10. if it is then add the number 0 to the begining of it then add the
'single type digit.

If Len(oType) < 2 Then
    oType = "0" & dType
Else
    oType = dType
End If

oData = dData 'Make the  outgoing Varable the same as the one given to the sub routine

oAll = "¬" & oType & oData
'Now we add the ¬ symbol and then the DataType and then the Data
'An example of what this could look like for the the other client is this:
'"¬01David"


'Now we check if the socket if connected before we send
If Winsocket.State <> sckConnected Then
    MsgBox "ERROR: Not Connected", vbCritical, "No Connection"
    Exit Sub 'if it is then we exit the sub and DONT sent
End If

'if we are connected then the code will continue and will send the data Varable oALL
Winsocket.SendData (oAll)
End Sub

Public Sub Split_Packet(iData As String)
'This sub will look though the new packet and will find every instance
'of the character ¬ that is the begining of the out data packet
'if it finds more that one Data Packet then it splits the merged Data and sends
'it to be processed

Dim sPackS As Integer
Dim sPackE As Integer
Dim i As Integer
Dim j As Integer
Dim sLast As Integer
Dim sType As DataTypes
Dim sData As String
Dim sAllData As String

For i = 1 To Len(iData) 'loop around every character in the packet
    If Mid(iData, i, 1) = "¬" Then ' if we find the ¬ character then
        sPackS = i + 1 ' set the posistion of the data packets start
        For j = sPackS To Len(iData)
            If (j = Len(iData)) And Mid(iData, j, 1) <> "¬" Then
                'if we are at the end of the packet and there are no more
                'packets to split then we can do the current packet and then exit
                'the loop
                sPackE = Len(iData)
                'set the end of the packet found

                sAllData = Mid(iData, sPackS, sPackE)
                'rid out the found packet and store it
                
                If Len(sAllData) < 3 Then
                    sType = sAllData 'if the pack only holds a type and no data then only set the type
                Else
                    sType = Mid(sAllData, 1, 2) ' the Data type is the first 2 characters
                    sData = Mid(sAllData, 3, (Len(sAllData) - 2)) 'and the Data is everything afterward
                End If
                
                Call Incoming_Data(sType, sData) 'Run the found packet to the Incoming_Data Sub routine
                Exit Sub
                
            ElseIf Mid(iData, j, 1) = "¬" Then
                'if we find another ¬ in the packet then we will run what we have found
                'and continue on to the next packet afterwards
                sPackE = (j - 2)
                'set the end of the packet found
                sAllData = Mid(iData, sPackS, (sPackE - sPackS) + 2)
                
                If Len(sAllData) < 3 Then
                    sType = sAllData 'if the pack only holds a type and no data then only set the type
                Else
                    sType = Mid(sAllData, 1, 2) ' the Data type is the first 2 characters
                    sData = Mid(sAllData, 3, (Len(sAllData) - 2)) 'and the Data is everything afterward
                End If
                
                Call Incoming_Data(sType, sData) 'Run the found packet to the Incoming_Data Sub routine
                Exit For
            
            End If
        Next j
    End If
Next i
End Sub

Private Sub Incoming_Data(itype As DataTypes, iData As String)
'This sub is used to find out what the data is and how to act on it.
'so if is a message then send it to the dialog text box
'or if its the other persons name then save it and update the other person label

Select Case itype 'We will now look at the iType and find its number
    Case DataTypes.Message ' if it = the Message number then its a message
        txt_dialog.Text = txt_dialog.Text & iData & vbCrLf
    Case DataTypes.NickName ' if it = the Nickname number then its a nickname
        othername = iData
        lbl_other = "You are talking to: " & othername
End Select ' stop looking
End Sub

Private Sub cmd_setnew_Click()
Myname = txt_newnick
lbl_mynick.Caption = "Your NickName is: " & Myname
Send_Data NickName, Myname
End Sub

Private Sub Command1_Click()

txt_dialog.Text = txt_dialog.Text & Myname & ": " & Text1.Text & vbCrLf ' this sends this clients msg to its dialog box
Send_Data Message, Myname & ": " & Text1.Text ' This sends the message and the users nickname to the other person
Text1.Text = ""
End Sub

Private Sub Form_Load()
Myname = "Client 02"
txt_newnick = "Client 02"
Winsocket.LocalPort = 7293 ' set the port to lissen on
Winsocket.Listen 'listen for client1
End Sub

Private Sub Winsocket_ConnectionRequest(ByVal requestID As Long)
Winsocket.Close 'close anyconnections allready on...
Winsocket.Accept (requestID)
End Sub

Private Sub Winsocket_DataArrival(ByVal bytesTotal As Long)
Dim incdata As String 'make a var to store the data
Winsocket.GetData incdata, vbString 'put the data into the var

Call Split_Packet(incdata) 'send the var to be processed
End Sub
