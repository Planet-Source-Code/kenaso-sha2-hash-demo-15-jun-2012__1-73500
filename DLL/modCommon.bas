Attribute VB_Name = "modCommon"
Option Explicit

' *****************************************************************************
' Global Constants
' *****************************************************************************
  Public Const DLL_NAME As String = "kiSHA2"

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long
  
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function
  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)



' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))
   
End Function
 
' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function QualifyPath(ByVal strPath As String, _
                   Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If
    
End Function

' ***************************************************************************
' Routine:       UnQualifyPath
'
' Description:   Removes a trailing character from the path
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to remove
'                          Default = "\"
'
' Returns:       Fully qualified path without a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function UnQualifyPath(ByVal strPath As String, _
                     Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        UnQualifyPath = Left$(strPath, Len(strPath) - 1)
    Else
        UnQualifyPath = strPath
    End If
    
End Function

' **************************************************************************
' Routine:       CalcProgress
'
' Description:   Calculates current amount of completion
'
' Parameters:    dblCurrAmt   - Current value
'                dblMaxAmount - Maximum value
'
' Returns:       percentage of progression
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function CalcProgress(ByVal dblCurrAmt As Double, _
                             ByVal dblMaxAmount As Double) As Long

    Dim lngPercent As Long    ' percentage to be calcuated
    Dim dblAmtLeft As Double  ' difference between current and max amount

    Const MAX_PERCENT As Long = 100
    
    ' Reset progress bar
    DoEvents
    If (dblCurrAmt <= 0) Or (dblMaxAmount <= 0) Then
        CalcProgress = 0
        Exit Function
    End If

    ' Make sure current value does
    ' not exceed maximum value
    If dblCurrAmt >= dblMaxAmount Then
        dblCurrAmt = dblMaxAmount
    End If

    ' Calculate percentage based on current
    ' value and maximum allowable value
    On Error Resume Next
    dblAmtLeft = dblMaxAmount - dblCurrAmt
    lngPercent = MAX_PERCENT - CLng((dblAmtLeft * MAX_PERCENT) / dblMaxAmount)
    On Error GoTo 0   ' Nullify this error trap

    ' Validate percentage so we
    ' do not exceed our bounds
    Select Case lngPercent
           Case Is < 0:           lngPercent = 0
           Case Is > MAX_PERCENT: lngPercent = MAX_PERCENT
    End Select

    CalcProgress = lngPercent  ' Return calculated percent
    
End Function

' ***************************************************************************
' Routine:       ByteArrayToString
'
' Description:   Converts a byte array to string data
'
' Parameters:    abytData - array of bytes
'
' Returns:       Data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function ByteArrayToString(ByRef abytData() As Byte) As String

    ByteArrayToString = StrConv(abytData(), vbUnicode)

End Function

' ***************************************************************************
' Routine:       StringToByteArray
'
' Description:   Converts string data to a byte array
'
' Parameters:    strData - Data string to be converted
'
' Returns:       byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function StringToByteArray(ByVal strData As String) As Byte()

     StringToByteArray = StrConv(strData, vbFromUnicode)

End Function

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer   ' Must be an integer

    On Error GoTo IsArrayInitialized_Exit
    
    IsArrayInitialized = 0  ' preset to FALSE
    
    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then
        
        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&
    
    End If
    
IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function


