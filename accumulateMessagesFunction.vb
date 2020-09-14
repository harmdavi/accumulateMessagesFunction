'David Harmon 
'RCET0265
'Fall 2020
'accumulateMessageFunction
'https://github.com/harmdavi/accumulateMessagesFunction.git

' Lane Coleman helped a frazzled David complete this 


Option Strict On
Option Explicit On
Option Compare Text

Module accumulateMessagesFunction

    Sub Main()
        Dim message, userInput As String
        Dim clearMemory As Boolean
        Console.WriteLine($"Hello User! I am a message save device.
I have been programed to save any information you would like to be stored for you at any time!
Simply type what you would like and press enter. This will save your message. 
If you would like to recall all of your messages, type the word call and your saved messages will display in a message box.
If you would like to clear your messages, simply type clear.
If you would like to quit the message program press q at any time.")

        Do

            userInput = Console.ReadLine()
            If userInput = "call" Then
                MsgBox(message)
            ElseIf userInput = "clear" Then
                clearMemory = True
            ElseIf userInput = "q" Then
                Exit Sub

            End If
            message = accumulateMessagesFunction(userInput, clearMemory)

            clearMemory = False
        Loop



    End Sub
    Function accumulateMessagesFunction(ByVal incommingMessage As String, ByVal clear As Boolean) As String
        Static savedMessage As String

        If clear Then
            savedMessage = ""
        ElseIf incommingMessage = "call" Then
            'nothing is placed here so that the text box will not have the word call in it when displayed 
        Else
            savedMessage &= incommingMessage & vbNewLine
        End If
        Return savedMessage
    End Function

End Module
