Attribute VB_Name = "modBitmapInMenu"
'/\/\/\/\/\/\/\ BITMAPS IN MENU RE-USABLE MODULE /\/\/\/\/\
'Code by Andy McCurtin
'Do what you want with it
'And ENJOY!!!!
'Any probs e-mail
'andy_mccurtin@yahoo.com

Option Explicit

'//This call accepts 1 argument, the window handle(hwnd)
'//e.g. Me.hwnd. It uses the windows handle to find a
'//'collection' of menu's from that form.
'//It returns a number if it finds any and a 0 if it doesn't
Public Declare Function GetMenu Lib "user32" (ByVal _
hwnd As Long) As Long

'//This accepts 2 arguments, the menu collection handle
'//(from above) and nPos this specifies the number of the
'//menu you are refering to (This always begins with 0)
Public Declare Function GetSubMenu Lib "user32" (ByVal _
hMenu As Long, ByVal nPos As Long) As Long

'//This accepts 2 arguments, the Sub Menu Id(from above)
'//and again nPos
Public Declare Function GetMenuItemID Lib "user32" (ByVal _
hMenu As Long, ByVal nPos As Long) As Long

'//This does the hard work, using info from the above 3 calls
'//With the above calls we now have a pointer to the menu
'// item that we want to add a picture to.  This function
'//returns a 1 if everything works
Public Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal _
wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal _
hBitmapChecked As Long) As Long




'//This function can easily be copied into a form
'//change frm for the forms name
'//Although the picture boxes have to be named you could
'//alter this using an array and run through the array
'//adding them to the menu
Public Function BitmapInMenu(frm As Form)
'//Variables
Dim Menu As Long
Dim SubMenu As Long
Dim MenuItemID As Long
Dim MenuItemID1 As Long
Dim Test As Long

'//Get the Menu collection ID for this form
Menu = GetMenu(frm.hwnd)

'//Get Sub Menu ID using value from above
SubMenu = GetSubMenu(Menu, 0)

'//Takes the SubMenuID from above and the position of the
'//menu item i.e. 0 is <<< Here it is, 1 is seperator,
'//2 is Exit etc
MenuItemID = GetMenuItemID(SubMenu, 0) 'mnuHere

MenuItemID1 = GetMenuItemID(SubMenu, 2) 'mnuExit

'//Takes picture from picTest and puts it in the menu next to
'//<<< Here it is
Test = SetMenuItemBitmaps(Menu, MenuItemID, 0, frm!picTest.Picture _
                , frm!picTest.Picture)

'//Takes picture from picExit and puts it in the menu next to
'//Exit
Test = SetMenuItemBitmaps(Menu, MenuItemID1, 0, frm!picExit.Picture _
                , frm!picExit.Picture)


'/////This code puts bitmaps in the sub menus
'//Get ID of Second Menu  (SubMenu)
SubMenu = GetSubMenu(Menu, 1)
'//Get ID of SubClass Sub Menu (Dummy)
SubMenu = GetSubMenu(SubMenu, 0)

    MenuItemID1 = GetMenuItemID(SubMenu, 0) 'Cool

'//Takes picture from picExit and puts it in the menu next to
'//Exit
Test = SetMenuItemBitmaps(SubMenu, MenuItemID1, 1, frm!picTest.Picture _
                , frm!picTest.Picture)

End Function
