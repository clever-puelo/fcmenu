Attribute VB_Name = "MiMsgBox1"
Option Explicit

Const MB_DEFBOTON1 = &H0&
Const MB_DEFBOTON2 = &H100&
Const MB_DEFBOTON3 = &H200&
Const MB_ICONOASTERISCO = &H40&
Const MB_ICONOEXCLAMACION = &H30&
Const MB_ICONOHAND = &H10&
Const MB_ICONOINFORMACION = MB_ICONOASTERISCO
Const MB_ICONOPREGUNTA = &H20&
Const MB_ICONOSTOP = MB_ICONOHAND
Const MB_OK = &H0&
Const MB_OKCANCEL = &H1&
Const MB_YESNO = &H4&
Const MB_YESNOCANCEL = &H3&
Const MB_ABORTRETRYIGNORE = &H2&
Const MB_RETRYCANCEL = &H5&


Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal _
hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType _
As Long) As Long

'------------------------------------------------------------------------------------------------------------------------
Function MiMsgBox(ByVal Ventana As Long, ByVal TextoCaja As String, _
                  ByVal Titulo As String, ByVal BotonesIconos As Long)
'------------------------------------------------------------------------------------------------------------------------

'si deseamos un Msgbox no Modal, bastará pasar como parametro Ventana=0

Dim Retval

Retval = MessageBox(Ventana, TextoCaja, Titulo, BotonesIconos)


End Function

