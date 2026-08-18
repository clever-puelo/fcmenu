Attribute VB_Name = "mdlRegistro"
Option Explicit

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Private Type FILETIME
            dwLowDateTime As Long
            dwHighDateTime As Long
End Type

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData() As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String) As Boolean
' Descripción:
'   Esta función borra una clave y devuelve true si pudo borrarla o false si no pudo
'
' Sintaxis:
'   variable = DeleteKey (ClaveRaíz, NombreDeClave)
'
'   ClaveRaíz debe ser HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   o HKEY_USERS
'
'   NombreDeClave es el nombre de la clave que queremos borrar,
'   puede incluir subclaves (por ejemplo "Clave1\SubClave1")
'
' Nota :
'   En W95 borrará todas las subclaves de la clave eliminada, en NT no se puede borrar
'   una clave que tenga subclaves


    Dim lRetVal As Long      'resultado de la función SetValueEx
    Dim hKey As Long         'handle de la clave abierta
    
    'para borrar una clave debe estar abierta
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    'borramos la clave
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
    DeleteKey = IIf(lRetVal = 0, True, False)
End Function
Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String) As Boolean
' Descripción:
'   Esta función borra un valor y devuelve true si pudo borrarlo y false si no pudo
'
' Sintaxis:
'   variable = DeleteValue (ClaveRaíz, NombreDeClave, NombreDeValor)
'
'   ClaveRaíz debe ser HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   o HKEY_USERS
'
'   NombreDeClave es el nombre de la clave que contiene el valor que queremos borrar,
'   puede incluir subclaves (por ejemplo "Clave1\SubClave1")
'
'   NombreDeValor es el nombre del valor que queremos borrar

       Dim lRetVal As Long      'resultado de la función SetValueEx
       Dim hKey As Long         'handle de la clave abierta

       'abrimos la clave especificada
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       'borramos el valor
       lRetVal = RegDeleteValue(hKey, sValueName)
       DeleteValue = IIf(lRetVal = 0, True, False)
       'la cerramos
       RegCloseKey (hKey)
End Function

Public Function EnumKey(lPredefinedKey As Long, sKeyName As String, vSubKeys As Variant) As Long
' Descripción:
'   Esta función busca todas las subclaves de una dada y forma una matriz con ellas en
'   el parámetro vSubKeys.
'
'   Si queremos saber las subclaves de una de las principales debemos dejar en blanco
'   el parámetro sKeyName
'
'   Devuelve el número de subclaves o -1 si hubo algún error
'
' Sintaxis:
'   variable = EnumKey (ClaveRaíz, NombreDeClave,SubClaves)
'
'   ClaveRaíz debe ser HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   o HKEY_USERS
'
'   NombreDeClave es el nombre de la clave cuyas subclaves queremos obtener,
'   puede incluir subclaves (por ejemplo "Clave1\SubClave1")
'
'   SubClaves es un variant que recogerá la matriz de subclaves obtenida

    Dim lRetVal As Long             'resultado de las funciones del API
    Dim hKey As Long                'handle de la clave abierta
    Dim sSubKeyName As String       'nombre de la subclave
    Dim lSubKeyLen As Long          'tamaño del nombre de la subclave
    Dim lMaxSubKeyLen As Long       'tamaño del nombre de subclave más grande
    Dim lNumSubKeys As Long         'número de subclaves existentes
    Dim ftLastWriteTime As FILETIME 'fecha última modif. de la clave (sólo NT)
    Dim lIndex As Long              'índice de la subclave
    Dim sSubClaves() As String      'matriz para contener las subclaves
    
    EnumKey = -1
    'si tenemos nombre de clave la abrimos
    If sKeyName <> "" Then
        lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        hKey = lPredefinedKey
    End If
    'obtenemos el nº de subclaves y el tamaño máximo de sus nombres
    lRetVal = RegQueryInfoKey(hKey, 0&, 0&, 0&, lNumSubKeys, lMaxSubKeyLen, 0&, 0&, 0&, 0&, 0&, ftLastWriteTime)
    If lRetVal = 0 Then
        'si no encontré subclaves
        If lNumSubKeys = 0 Then
            EnumKey = 0
        Else
            EnumKey = lNumSubKeys
            lNumSubKeys = lNumSubKeys - 1       ' va de 0 a n-1
            lMaxSubKeyLen = lMaxSubKeyLen + 1   ' dejar sitio para el 0 de fin de string en c
            'dimensionamos la matriz
            ReDim sSubClaves(lNumSubKeys)
            'recorremos las subclaves (en orden inverso, como dice la ayuda ¿?)
            For lIndex = lNumSubKeys To 0 Step -1
                lSubKeyLen = lMaxSubKeyLen
                sSubKeyName = String(lMaxSubKeyLen, 0)
                lRetVal = RegEnumKeyEx(hKey, lIndex, sSubKeyName, lSubKeyLen, 0&, 0&, 0&, ftLastWriteTime)
                If lRetVal = 0 Then
                    sSubClaves(lIndex) = Left$(sSubKeyName, lSubKeyLen)
                Else
                    EnumKey = -1
                End If
            Next lIndex
        End If
    Else
        EnumKey = -1
    End If
    'devolvemos el resultado
    vSubKeys = sSubClaves()
    'cerramos la clave
    If sKeyName <> "" Then RegCloseKey (hKey)
End Function

Public Function EnumValue(lPredefinedKey As Long, sKeyName As String, vValues As Variant) As Long
' Descripción:
'   Esta función busca todos los valores de una clave y forma una matriz con ellos
'   en el parámetro vValues.
'
'   Si queremos saber las subclaves de una de las principales debemos dejar en blanco
'   el parámetro sKeyName
'
'   Devuelve el número de valores o -1 si hubo algún error
'
' Sintaxis:
'   variable = EnumValue (ClaveRaíz, NombreDeClave,Valores)
'
'   ClaveRaíz debe ser HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   o HKEY_USERS
'
'   NombreDeClave es el nombre de la clave cuyos valores y datos queremos obtener,
'   puede incluir subclaves (por ejemplo "Clave1\SubClave1")
'
'   Valores es un variant que recogerá la matriz de valores y datos obtenida

    Dim lRetVal As Long             'resultado de las funciones del API
    Dim hKey As Long                'handle de la clave abierta
    Dim sValueName As String        'nombre del valor
    Dim lValueNameLen As Long       'tamaño del nombre del valor
    Dim lMaxValueNameLen As Long    'tamaño del nombre de valor más grande
    Dim lNumValues As Long          'número de valores existentes
    Dim bValueData(500) As Byte     'byte para obtener el dato del valor, no funciona?
    Dim lValueSize As Long          'longitud del array anterior
    Dim ftLastWriteTime As FILETIME 'fecha última modif. del valor (sólo NT)
    Dim lIndex As Long              'índice del valor
    Dim vValores() As String        'matriz para contener los valores
    
    EnumValue = -1
    'si tenemos nombre de clave la abrimos
    If sKeyName <> "" Then
        lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        hKey = lPredefinedKey
    End If
    
    'obtenemos el nº de valores y el tamaño máximo de sus nombres
    lRetVal = RegQueryInfoKey(hKey, 0&, 0&, 0&, 0&, 0&, 0&, lNumValues, lMaxValueNameLen, 0&, 0&, ftLastWriteTime)
    If lRetVal = 0 Then
        'si no encontré valores
        If lNumValues = 0 Then
            EnumValue = 0
        Else
            EnumValue = lNumValues
            lNumValues = lNumValues - 1 'va de 0 a n-1
            lMaxValueNameLen = lMaxValueNameLen + 1 'para que quepa el 0 de fin de cadena en C
            'dimensionamos la matriz
            ReDim vValores(lNumValues)
            'recorremos los valores (en orden inverso, como dice la ayuda ¿?)
            For lIndex = lNumValues To 0 Step -1
                lValueNameLen = lMaxValueNameLen
                sValueName = String(lMaxValueNameLen, 0)
                lValueSize = 500
                'no me funciona si no pongo un array de bytes para recoger el resultado,
                'aunque luego no me lo da¿?. Si pones una longitud menor de lo que ocupa
                'el dato tampoco funciona, por eso puse 500 bytes
                lRetVal = RegEnumValue(hKey, lIndex, sValueName, lValueNameLen, 0&, 0&, bValueData(), lValueSize)
                If lRetVal = 0 Then
                    vValores(lIndex) = Left$(sValueName, lValueNameLen)
                Else
                    EnumValue = -1
                End If
            Next lIndex
        End If
    Else
        EnumValue = -1
    End If
    'devolvemos el resultado
    vValues = vValores()
    'cerramos la clave
    If sKeyName <> "" Then RegCloseKey (hKey)

End Function


Private Function SetValueEx(ByVal hKey As Long, sValueName As String, ltype As Long, vValue As Variant) As Long
'
' Función utilizada por SetValue, no debemos llamarla directamente
'
    Dim lValue As Long
    Dim sValue As String

    Select Case ltype
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, ltype, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, ltype, lValue, 4)
    End Select

End Function





Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
'
' Función utilizada por QueryValue, no debemos llamarla directamente
'
    Dim cch As Long
    Dim lrc As Long
    Dim ltype As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError


    ' Determinar el tipo de datos y el tamaño que debemos leer

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, ltype, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case ltype
        ' Para strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, ltype, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If

        ' Para DWORDS (long)
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, ltype, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'no están soportados otros tipos
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String) As Boolean
' Descripción:
'   Esta función crea una nueva clave y devuelve true si pudo crearla o false si no pudo
'
' Sintaxis:
'   variable = CreateNewKey (ClaveRaíz, NombreDeClave)
'
'   ClaveRaíz debe ser igual a HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   o HKEY_USERS
'
'   NombreDeClave es el nombre de la clave que queremos crear,
'   puede incluir subclaves (por ejemplo "Clave1\SubClave1")

    
    
    Dim hNewKey As Long         'handle a la nueva clave
    Dim lRetVal As Long         'resultado de la función RegCreateKeyEx
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    CreateNewKey = IIf(lRetVal = 0, True, False)
    'cerramos la clave
    RegCloseKey (hNewKey)
End Function




Public Function SetValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long) As Boolean
' Descripción:
'   Esta función crea y/o modifica el dato contenido en un valor y devuelve true si lo
'   modificó o false si no pudo
'
' Sintaxis:
'   variable = SetValue (ClaveRaíz, NombreDeClave, NombreDeValor, NuevoDato, TipoDeDato)
'
'   ClaveRaíz debe ser HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   o HKEY_USERS
'
'   NombreDeClave es el nombre de la clave que contiene el valor que queremos recuperar,
'   puede incluir subclaves (por ejemplo "Clave1\SubClave1")
'
'   NombreDeValor es el nombre del valor que queremos crear o modificar

'   NuevoDato es el dato que queremos introducir en el valor
'
'   TipoDeDato debe ser REG_SZ (un string) o REG_DWORD (un long)

       Dim lRetVal As Long      'resultado de la función SetValueEx
       Dim hKey As Long         'handle de la clave abierta

       'abrimos la clave
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       'modificamos el dato del valor
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       SetValue = IIf(lRetVal = 0, True, False)
       'cerramos la clave
       RegCloseKey (hKey)

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Descripción:
'   Esta función devuelve los datos de un valor o Empty si no lo encontró
'
' Sintaxis:
'   variable = QueryValue(ClaveRaíz, NombreDeClave, NombreDeValor)
'
'   ClaveRaíz debe ser HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   o HKEY_USERS
'
'   NombreDeClave es el nombre de la clave que contiene el valor que queremos recuperar,
'   puede incluir subclaves (por ejemplo "Clave1\SubClave1")
'
'   NombreDeValor es el nombre del valor que queremos recuperar, si es null devolverá
'   el valor predeterminado de la clave (si existe)

       Dim lRetVal As Long      'resultado de las funciones del API
       Dim hKey As Long         'handle de la clave abierta
       Dim vValue As Variant    'datos del valor requerido

       'abrimos la clave
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       'obtenemos los datos del valor
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = vValue
       'cerramos la clave
       RegCloseKey (hKey)
End Function





