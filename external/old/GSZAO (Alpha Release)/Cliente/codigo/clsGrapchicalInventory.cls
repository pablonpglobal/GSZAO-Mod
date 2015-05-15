VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************
'
'*****************************************************
'Last Modify Date: 05/09/09
'Last Modify Author: Parra
'Description: Adapted class for Directx8
'*****************************************************

Option Explicit

Dim InvSelectedItem As Byte             'Currently selected item
Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

Dim InvRect As RECT
Dim InvD3DRect As D3DRect

Dim Color(3) As Long

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

''
' Initialize Class
Private Sub Class_Initialize()
    Set InventoryWindow = frmMain.picInv
    InventoryWindow.AutoRedraw = True
    InventoryWindow.ScaleMode = 3
    
    InvRect.Top = 0: InvD3DRect.Y1 = 0
    InvRect.Left = 0: InvD3DRect.X1 = 0
    InvRect.bottom = 160: InvD3DRect.Y2 = 160
    InvRect.Right = 160: InvD3DRect.X2 = 160

    Engine.Graphics_ResetColor Color()

End Sub
' Retrieves the selected item index
Public Property Get SelectedItem() As Byte: SelectedItem = InvSelectedItem: End Property
' Retrieves the max hit of the selected item
Public Property Get MaxHit(ByVal slot As Byte) As Integer: MaxHit = UserInventory(slot).MaxHit: End Property
' Retrieves the min hit of the selected item
Public Property Get MinHit(ByVal slot As Byte) As Integer: MinHit = UserInventory(slot).MinHit: End Property
' Retrieves the defense of the selected item
Public Property Get Def(ByVal slot As Byte) As Integer: Def = UserInventory(slot).Def: End Property
' Retrieves the grh index of the selected item
Public Property Get GrhIndex(ByVal slot As Byte) As Integer: GrhIndex = UserInventory(slot).GrhIndex: End Property
' Retrieves the value of the selected item
Public Property Get Valor(ByVal slot As Byte) As Long: Valor = UserInventory(slot).Valor: End Property
' Retrieves the selected item's amount
Public Property Get Amount(ByVal slot As Byte) As Integer: Amount = UserInventory(slot).Amount: End Property
' Retrieves the selected item's object index
Public Property Get OBJIndex(ByVal slot As Byte) As Integer: OBJIndex = UserInventory(slot).OBJIndex: End Property
' Retrieves the selected item's object type
Public Property Get OBJType(ByVal slot As Byte) As Integer: OBJType = UserInventory(slot).OBJType: End Property
' Retrieves the selected item's name
Public Property Get ItemName(ByVal slot As Byte) As String: ItemName = UserInventory(slot).Name: End Property
' Retrieves True if the item at the given pos is eqiupped
Public Property Get Equipped(ByVal slot As Byte) As Boolean: Equipped = UserInventory(slot).Equipped: End Property

''
' Sets all data for a given inventory slot
Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Long, ByVal eName As String)
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
End Sub

''
' Selects the item clicked if it's valid and return's it's index
Private Function ClickItem(ByVal x As Integer, ByVal y As Integer) As Integer
    'Make sure it's within limits
    If (x \ 32) + (y \ 32) * (InventoryWindow.ScaleWidth \ 32) + 1 <= MAX_INVENTORY_SLOTS Then
        'Make sure slot isn't empty
        If UserInventory((x \ 32) + (y \ 32) * (InventoryWindow.ScaleWidth \ 32) + 1).GrhIndex Then
            ClickItem = (x \ 32) + (y \ 32) * (InventoryWindow.ScaleWidth \ 32) + 1
        Else
            ClickItem = 0
        End If
    End If
End Function

''
' Implements the mouse move event of the inventory picture box
' Displays a ToolTip of the item under the mouse
Public Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then Exit Sub

    If (x \ 32) + (y \ 32) * (InventoryWindow.ScaleWidth \ 32) + 1 <= MAX_INVENTORY_SLOTS Then
        InventoryWindow.ToolTipText = UserInventory((x \ 32) + (y \ 32) * (InventoryWindow.ScaleWidth \ 32) + 1).Name
    End If

End Sub

''
' Check outs which item was clicked
Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > InventoryWindow.Width Or y > InventoryWindow.Height Then Exit Sub
    
    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(x), CInt(y))
End Sub

Public Sub UpdateInventory()
'*****************************************************
'Date: 05/09/09
'Author: Parra
'Description: Updates the inventory only when needed
'*****************************************************

    Dim i As Long

    Dim x As Integer
    Dim y As Integer
    
    Engine.DirectX_Scene_BeginCustom 1, InvD3DRect, D3DCLEAR_TARGET
    
        For i = 1 To UBound(UserInventory) 'slots
            If UserInventory(i).Amount <> 0 Then
                x = ((i - 1) Mod (InventoryWindow.Width / 32)) * 32 + 2
                y = ((i - 1) \ (InventoryWindow.Width / 32)) * 32 + 2

                Engine.Graphics_Draw_Texture UserInventory(i).GrhIndex, x, y, 32, 32, 32, 32, Color()
                Engine.Fonts_Render 1, x, y - 1, x + 20, y + 15, CStr(UserInventory(i).Amount), -1

                If UserInventory(i).Equipped Then
                    Engine.Fonts_Render 1, x + 20, y + 20, x + 25, y + 25, "+", -65536
                End If
            End If
        Next i
        
    Engine.DirectX_Scene_EndCustom InvRect, frmMain.picInv.hWnd

End Sub
