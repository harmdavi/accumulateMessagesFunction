'David Harmon 
'RCET0265
'Fall 2020
'accumulateMessageFunction
'https://github.com/harmdavi/accumulateMessagesFunction.git

' Lane Coleman helped a frazzled David complete this 


Option Strict On
Option Explicit On
Option Compare Text

Module AccumulateMessagesFunction ' Solution, Project, Module, Method names all PascalCase - TJR

    Sub Main()
        Dim message, userInput As String
        Dim clearMemory As Boolean

        'format long string for readability- TJR
        Console.WriteLine($"Hello User! I am a message save device.
I have been programed to save any information you would like to be stored for you at any time!
Simply type what you would like and press enter. This will save your message. 
If you would like to recall all of your messages, type the word call and your saved messages will display in a message box.
If you would like to clear your messages, simply type clear.
If you would like to quit the message program press q at any time.")

        Do 'changed logic slightly - TJR
            userInput = Console.ReadLine()
            If userInput = "call" Then
                message = AccumulateMessagesFunction("", False)
                MsgBox(message)
            ElseIf userInput = "clear" Then
                'clearMemory = True
                message = AccumulateMessagesFunction("", True)
            ElseIf userInput = "q" Then
                Exit Sub
            Else
                message = AccumulateMessagesFunction(userInput, clearMemory)
            End If
            'clearMemory = False
        Loop

        'Remove extra blank lines - TJR

    End Sub

    'made more generic. suggestion only - TJR
    Function AccumulateMessagesFunction(ByVal incommingMessage As String, ByVal clear As Boolean) As String
        Static savedMessage As String
        If clear Then
            savedMessage = ""
            'ElseIf incommingMessage = "call" Then
            'nothing is placed here so that the text box will not have the word call in it when displayed 
        ElseIf incommingMessage <> "" Then
            savedMessage &= incommingMessage & vbNewLine
        End If
        Return savedMessage
    End Function

End Module
