Attribute VB_Name = "modMsgBoxAPI"
'/\/\/\/\/\/\/\ API MESSAGE BOX RE-USABLE MODULE /\/\/\/\/\
'Code by Andy McCurtin
'Do what you want with it
'And ENJOY!!!!
'Any probs e-mail
'andy_mccurtin@yahoo.com

'/\/\/\/\/\/\ BEFORE YOU READ ANY FURTHER!!! /\/\/\/\/\/\/\
'/\/\ Examples of how to use this Module are at the bottom
'/\/\ of the code, and included on the form.
'##########################################################
'# For anyone who isn't sure how the API declare works    #
'# itworks like this :-                                   #
'#                                                        #
'# [Public/Private] Declare Sub/Function name Lib _       #
'# "DLL name" [Alias "Alias name"] [(Argument List)] _    #
'# [As Type]                                              #
'#                                                        #
'# You don't need the underscore's(_) but they help make  #
'# your code easier to read.                              #
'#                                                        #
'# The stuff in the square brackets([ ]) is optional, the #
'# Slash (/) means that you choose one or the other.      #
'# The Argument List is list of arguments that may or may #
'# not be present.                                        #
'##########################################################

Option Explicit

'//This is the MessageBox API call
Public Declare Function MessageBox Lib "user32" Alias _
"MessageBoxA" (ByVal hwnd As Long, ByVal lpText As _
String, ByVal lpCaption As String, ByVal wType As Long) _
As Long
'##########################################################
'# The above breaks down like this :-                     #
'#                                                        #
'# hwnd = This is the handle to a window                  #
'#                                                        #
'# lpText = This is the message that appears in the       #
'#          MessageBox                                    #
'#                                                        #
'# lpCaption = The title of the MessageBox                #
'#                                                        #
'# wType = This is a number that represents the style of  #
'#         the MessageBox i.e. Icon or Buttons etc.       #
'##########################################################



'//These are the constants for the MessageBox buttons.
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_YESNOCANCEL = &H3&
Public Const MB_YESNO = &H4&
Public Const MB_RETRYCANCEL = &H5&


'//These are the constants for the MessageBox Icon.
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONCRITICAL = &H10&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONINFORMATION = &H40&
Public Const MB_ICONEXCLAMATION = &H30&


'//These are the constants for the MessageBox Modal (This
'//basically means the MessageBox gets the users attention.
'//How much attention is up to you, SYSTEMMODAL means that
'//the MessageBox stays visible until closed, above ALL other
'//applications.
'//TASKMODAL means that the program cannot continue however
'//the MessageBox will only be on top of the program not the
'//whole system as with SYSTEMMODAL (try them out to see the
'//difference firsthand).
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&


'//The Function(MessageBox) returns a number(Long) these are the
'//constants.
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7


'Using this COOL mosule
'----------------------
'##########################################################
'# To call this Function use the following :-             #
'# Call MessageBox(Form1.hwnd, "This is my MessageBox", _ #
'#  "This is a Test", MB_ICONQUESTION)                    #
'#                                                        #
'#                          OR                            #
'#                                                        #
'# Call MessageBox(Form1.hwnd, "This is my MessageBox", _ #
'#  "This is a Test", MB_OKCANCEL Or MB_ICONQUESTION )    #
'#                                                        #
'#                          OR                            #
'#                                                        #
'# Call MessageBox(Form1.hwnd, "This is my MessageBox", _ #
'#  "This is a Test", MB_OKCANCEL Or MB_ICONQUESTION _    #
'#  or MB_TASKMODAL)                                      #
'#                                                        #
'# I don't know why the API uses Or instead of the more   #
'# logical And, but it does.                              #
'#                                                        #
'# Again the underscore's(_) are not essential but make   #
'# the code more readable.                                #
'#                                                        #
'#         -------------------------------------          #
'#                                                        #
'# To respond to a response from the user use this :-     #
'#                                                        #
'# If MessageBox(Form1.hwnd, "Do you like this?", _       #
'#  "This is a Test", MB_YESNO Or MB_ICONQUESTION ) = _   #
'#   IDYES Then                                           #
'#      Call MessageBox(Form1.hwnd, "Thank you","Great" _ #
'#      MB_OK Or MB_ICONQUESTION )                        #
'# End IF                                                 #
'#                                                        #
'# You could also use a Case statment(shown on frmTest    #
'# that would be better if you have a few responses.      #
'##########################################################

