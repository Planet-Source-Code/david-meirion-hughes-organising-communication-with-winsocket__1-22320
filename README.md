﻿<div align="center">

## Organising Communication with WinSocket


</div>

### Description

This artical will give the user a insight as to how data transactions should occur. it will cover how to send different types of data and then getting the recipent program to understand it and do something with that data and its type, such as setting nicknames and sending messages (in a chat program) also included is a simple chat program that uses the code in the artical and shows you the code in action. the code is very well commented and it aimed at begenner level to intermediate level coders. also included is a nice way to counter act teh effect of "Data Merging" that can occur over slow connections using TCP/IP with out using the slow "Send Complete" event in winsock.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-04-11 21:16:58
**By**             |[David Meirion Hughes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-meirion-hughes.md)
**Level**          |Beginner
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Organising182384112001\.zip](https://github.com/Planet-Source-Code/david-meirion-hughes-organising-communication-with-winsocket__1-22320/archive/master.zip)





### Source Code

<p><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#000000"><b>Organising
 Communication with WinSock</b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><b>[Introduction]
 </b><br>
 </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Hi.
 This is my first tutorial so it may be a little bad, but I hope I can help you
 out a little. This tutorial is for basic to intermediate coders and will explain
 a little on communication between 2 programs such as a chat program. Now you
 may have seen a lot of chat programs on here some are good some are just basic
 and ONLY do chat, in the latter the two programmes will be just sending text
 between one another. This tutorial will teach you how to organise your data
 packets and allow your chat program (or any other type of communication program)
 to do a lot more that just send text.<br>
 <br>
 Notice: I'm sorry that teh code is not indented properly. dam word doesnt copy
 and paste properly and I can add spaces using dream weaver so I hope you can
 make do. <br>
 </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Updated
 Data is in Blue: This Information is for new coders that dont quite understand
 what I'm doing and I think if I was in there shoes I would agree with them.
 The sections in blue will explain what exactly we are doing and how to do it.</font></p>
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF"><b>I
have included a fully documented example project that will help you see what is
going on in the code and how it all works. </b></font>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">What
 we are going to do. We are going to start off by learning how data packets are
 formed or better put, how you the coder should organise the information you
 tell your programs to tell each other. Then we will cover how to make your datatypes.
 These datatypes hold the key to the data and allow the program to understand
 exaclty what is has been given. thirdly we will cover how to make sending packets
 of data faster and easier in the long run. we will then learn how to unscramble
 merged data that can be a problems for both new and experenced Programmers.
 We will then learn how out programs can quickly and easilly decipher what it
 has been given. </font><br>
 <br>
 <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><font color="#0000FF">NOTE:
 the code in this artical is related to Visual Basic Version 5 and 6. it has
 not been tested for any other platform</font></font><br>
 <br>
 <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">
 </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">
 Okay so lets begin. First off let me explain how I organise my data packets</font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">
 <b>[Data Packet Structure] </b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">When
 my chat program sends data the data is in 2 parts, the Data Type and the Data.
 For example: </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">¬04HELLO!
 </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Part
 1 (Data Type): </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">The
 first two numbers tells the recipient program what the data is. Say if the number
 is 04 then it could mean, "Here is a message for you" or if its 07 the it could
 mean "My nick name is: " </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Part
 2 (Data) </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">This
 is the data that goes with the packet. What the recipient program does with
 it is dependant on the Data Type </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">You
 may have noticed the "¬" symbol. Don't worry its actually part
 of the data packet and not another instance of bad English / Typing. All will
 be relieved later on. <br>
 </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><b>[Setting
 the Types] </b><br>
 </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Okay
 so now we know how the packets are formed. Now we need to code the data types
 so that its easier later on to do stuff. Here is how you would set your data
 types in a module file (or whatever). </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">A
 Module file is a file that is used to store code. It allows coders to organise
 files alot easier and put certain code into certain, relivent files. To create
 a module file goto to the menu bar, choose "Project" then choose "Add
 module".<br>
 <br>
 </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">===================
 </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#666666">Public
 Enum DataTypes<br>
 MESSAGE = 0 <br>
 NICKNAME = 1<br>
 End enum </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><br>
 <br>
 =================== <br>
 <br>
 <font color="#0000FF">An Enum is like the type Boolean (True or false) with
 boolean if you choose TRUE then the number is set to 1 becuase that is how its
 defined. it you choose FALSE then the number is set to 0. With the Enum statement
 we can make our very own boolean type, type. <br>
 <br>
 Okay so you've entered in the code above, now we have a type called DataTypes,
 to test it hit the enter button and type<br>
 <br>
 <font color="#666666">DIM TESTVARABLE as DataTypes</font></font></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Now
 type in</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">
 <font color="#666666">TESTVARABLE = <font color="#0000FF">now a menu should
 come down and you can select which one you want and it will right it in for
 you. </font></font><br>
 <br>
 I'm not going to put a lot of types on here. My chat program has about 22 so
 far. But for this tutorial these two will do fine. Now I'm assuming you already
 know how to connect two computers together using winsock so I'm not going to
 go into that. If you need help then there are plenty of good examples and tutorials
 on the basics of winsock on this web site. </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><b>[Making
 a Fast Send Sub Routine] </b></font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">So
 we have our data types ready. Now we can code a really cool Sub that can really
 speed up sending data (coding wise) </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">===================
 </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#666666">Public
 Sub Send_Data(dType As DataTypes, Optional dData As String)<br>
 Dim oData As String <br>
 Dim oType As String <br>
 Dim oAll As String <br>
 <br>
 oType = dType<br>
 If Len(oType) < 2 Then <br>
 oType = "0" & dType<br>
 Else oType = dType <br>
 End If <br>
 <br>
 oData = dData <br>
 oAll = "¬" & oType & oData<br>
 <br>
 If WINSOCKCONTROL.State <> sckConnected <br>
 Then MsgBox "ERROR: Not Connected", vbCritical, "No Connection"<br>
 Exit Sub<br>
 End If <br>
 <br>
 WINSOCKCONTROL.SendData (oAll)<br>
 End Sub <br>
 </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><br>
 =================== <br>
 <br>
 Okay basically this is used so that if you want to send a message you can do
 it with one line of code and do it really fast. It also brings up a really cool
 menu for choosing the data type. It also makes sure the data type uses 2 characters
 so if the number for the type is less than 10 then it adds the character "0"
 to the beginning and then the single digit afterward. Its basically used so
 that its always 2 characters and can be easily ripped out of the data packet
 later on.<br>
 <br>
 Okay so you've sent the data. Now you need something to decipher it on the other
 side. But first I think its time I explained what the "¬" is for. Now when
 I started playing around with my chat program on the Internet I found that the
 data was getting merged together. Basically sometimes two data packets merged
 to look something like this 02Hi04David. When this happened my program went
 to find the data type "02" then sent the message "hi04david" which was annoying
 because the 04david was supposed to be a nickname and not a message. So any
 ways back to the point.<br>
 <br>
 <b>[Splitting merged data packets]</b><br>
 <br>
 I came up with the idea of adding a symbol to the beginning of all the packets
 then splitting the packets up after every "¬" symbol. It took a while to
 figure out but I managed it&#8230; so here the code to do it&#8230;. By the way there is
 a reference to the Incoming_Data Sub, which we will cover afterwards. <br>
 <br>
 =================== <br>
 <font color="#CCFF00"><br>
 <font color="#666666">Public Sub Split_Packet(iData As String) <br>
 Dim sPackS As Integer<br>
 Dim sPackE As Integer<br>
 Dim i As Integer<br>
 Dim j As Integer<br>
 Dim sLast As Integer<br>
 Dim sType As DataTypes<br>
 Dim sData As String<br>
 Dim sAllData As String<br>
 <br>
 For i = 1 To Len(iData) <br>
 <br>
 If Mid(iData, i, 1) = "¬" Then <br>
 sPackS = i + 1 <br>
 <br>
 For j = sPackS To Len(iData)<br>
 <br>
 If (j = Len(iData)) And Mid(iData, j, 1) <> "¬" Then <br>
 <br>
 sPackE = Len(iData) <br>
 sAllData = Mid(iData, sPackS, sPackE) '- (sPackS + 1)))<br>
 <br>
 If Len(sAllData) < 3 Then <br>
 sType = sAllData<br>
 Else <br>
 <br>
 sType = Mid(sAllData, 1, 2) <br>
 sData = Mid(sAllData, 3, (Len(sAllData) - 2))<br>
 <br>
 End If <br>
 Call incoming_data(sType, sData) <br>
 Exit Sub <br>
 <br>
 ElseIf Mid(iData, j, 1) = "¬" Then <br>
 <br>
 sPackE = (j - 2) <br>
 sAllData = Mid(iData, sPackS, (sPackE - sPackS) + 2)<br>
 <br>
 If Len(sAllData) < 3 Then<br>
 sType = sAllData<br>
 Else <br>
 <br>
 sType = Mid(sAllData, 1, 2)<br>
 sData = Mid(sAllData, 3, (Len(sAllData) - 2)) <br>
 End If <br>
 <br>
 Call incoming_data(sType, sData)<br>
 <br>
 Exit For<br>
 <br>
 End If <br>
 <br>
 Next j <br>
 <br>
 End If <br>
 <br>
 Next i <br>
 <br>
 End Sub </font><br>
 </font> </font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">===================</font></p>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><font color="#0000FF">The
 symbol "¬" is used by holding shift and then the button next to
 the number 1. It can be any symbol you wish, but I chose this symbol becuase
 I felt that it would not be used by the people testing the program and so I
 would be safe using it.<br>
 </font></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><br>
 Okay all this does it constantly loops around until it's found all the merged
 packets (if any) then send it to another sub to be processed. <br>
 <br>
 Now if you're actually an expert at Winsock and wonder why I just didn't just
 use the "send complete" event in Winsock. Well its because it kind of freezes
 up when u have 30 connections and it gets really slow doing it that way.<br>
 <br>
 <b>[Processing the Incoming Data]</b><br>
 <br>
 Okay we've now sorted the data now we need to do something with it. This is
 where incoming_data comes in. Basically all we do here is do a select case statement
 on the incoming data type. Then do something with the data.<br>
 <br>
 ===================<br>
 <br>
 <font color="#666666">Public Sub incoming_data(iType As DataTypes, iData As
 String) <br>
 <br>
 Select Case iType <br>
 Case DataTypes.MESSAGE <br>
 'send the data or message to the textbox <br>
 txt_dialog.Text = txt_dialog.Text & iData & vbCrLf <br>
 Case DataTypes.NICKNAME <br>
 'set the remote users nickname as the data <br>
 lbl_usernick.caption = idata <br>
 end select </font><br>
 <br>
 =================== <br>
 <br>
 <font color="#0000FF">So you now need to know the steps what you should do on
 a Data_Arival Event in winsock it is this....<br>
 >ON - DATA_ARRIVAL_WINSOCK><br>
 >SEND THE DATA TO ><br>
 >SPLIT_PACKET ><br>
 >SEND THE PACKETS TO><br>
 >INCOMING_DATA</font></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">
 ><br>
 <font color="#0000FF">>DO SOMETHING WITH THE DATA AND ITS TYPE. </font><br>
 <br>
 Now I mean this is quite basic what I've shown you here. But you can add new
 data types and do new things on that type of data. There is no limit to how
 many you want. Although don't go over 100 type if your using my code&#8230; come to
 think of it. If you manage to get over 100 individual types email me or post
 a comment because I don't believe its possible&#8230;. hehe&#8230;. I have multi channels
 and multi users and I've only got 22 types! Anyway I hope this tutorial has
 helped you a little. or if ya want to tell me how to do sothing proberly then
 please tell me becuase I've only recently started on VB<br>
 <br>
 Please leave a comment if you need any help or u would like to thank me or if
 you want to tell me that I'm wrong or if there is an error in the code. Thanks</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><br>
 <br>
 oh btw, I'll be uploading the Simple Chat 2.8 as soon as I comment the code.
 It uses what I have shown you above and a little more. ;) so watch out for it.
 I'm also doing something I havent seen on here. so keep ya eye out.<br>
 <br>
 Thanks for reading</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><br>
 <br>
 CrAKiN-ShOt <br>
 <a href="mailto:crakinshot@hotmail.com">crakinshot@hotmail.com </a></font></p>

