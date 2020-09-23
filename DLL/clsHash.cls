VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:  clsHash.cls
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME  As String = "clsHash"
  
' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumHASH_ALGORITHM
      eHASH_SHA1        ' 0
      eHASH_SHA224      ' 1
      eHASH_SHA256      ' 2  Default
      eHASH_SHA384      ' 3
      eHASH_SHA512      ' 4
      eHASH_SHA512_224  ' 5  As per FIPS 180-4 (dtd March-2012)
      eHASH_SHA512_256  ' 6  As per FIPS 180-4 (dtd March-2012)
      eHASH_SHA512_320  ' 7  My creation as per FIPS 180-4 (dtd March-2012)
  End Enum
  
' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Long)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m lng HashMethod
' Variable name:     mlngHashMethod
'
' ***************************************************************************
  Private mblnRetLowercase As Boolean
  Private mlngHashRounds   As Long
  Private mlngHashMethod   As enumHASH_ALGORITHM

  Private WithEvents mobjSHA1 As cSHA1
Attribute mobjSHA1.VB_VarHelpID = -1
  Private WithEvents mobjSHA2 As cSHA2
Attribute mobjSHA2.VB_VarHelpID = -1


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

' Comment out Version Property if not used in this application
Public Property Get Version() As String
    ' Version of the DLL this class is associated with.
    Version = DLL_NAME & ".dll (tm) v" & App.Major & "." & App.Minor & "." & App.Revision & _
              " Copyright (c) 2004-2012 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = True
    mblnRetLowercase = blnData
End Property

Public Property Let HashMethod(ByVal lngHashMethod As enumHASH_ALGORITHM)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eHASH_SHA256
    mlngHashMethod = lngHashMethod
End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1

    Select Case lngData
           Case 1 To 10: mlngHashRounds = lngData   ' Good choice
           Case Else:    mlngHashRounds = 1         ' Use default value
    End Select
    
End Property

' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       HashString
'
' Description:   Function to create a unique hex string representation of 
'                the data passed.  See Readme.txt file for more information. 
' 
' Syntax:
'    mobjSHA.StopProcessing = False                 ' set the stop flag to FALSE 
'    mobjSHA.HashMethod = eHASH_SHA256              ' Set property to use hash method 
'    abytData() = StrConv("abc", vbFromUnicode)     ' convert string data to byte array 
'    txtData.Text = mobjSHA.HashString(abytData())  ' hash data and return as string 
' 
' Parameters:    abytData() - Data to be hashed in a byte array
'
' Returns:       Hashed data string in byte array 
' 
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 20-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array.  Added functionality
'              to return data in various formats.
' 04-APR-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in output format.
' ***************************************************************************
Public Function HashString(ByRef abytInput() As Byte) As Byte()
Attribute HashString.VB_Description = "Creates a hash output string based on string data input."

    Dim abytHash() As Byte
    
    On Error GoTo HashString_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytInput())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If
    
    Erase abytHash()  ' Always start with empty arrays
    ResetVariables    ' Verfiy hash objects are empty
    
    RaiseEvent HashProgress(0)   ' Reset progress bar
    
    Select Case mlngHashMethod
           
           Case eHASH_SHA1
                Set mobjSHA1 = New cSHA1   ' Instatiate class object
                With mobjSHA1
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashString(abytInput())
                End With

           Case eHASH_SHA224, eHASH_SHA256, eHASH_SHA384, eHASH_SHA512, _
                eHASH_SHA512_224, eHASH_SHA512_256, eHASH_SHA512_320
                
                Set mobjSHA2 = New cSHA2   ' Instatiate class object
                With mobjSHA2
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    .HashMethod = (mlngHashMethod - 1)
                    abytHash() = .HashString(abytInput())
                End With
    End Select
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Erase abytHash()
        GoTo HashString_CleanUp
    End If

HashString_CleanUp:
    HashString = abytHash()
           
    Erase abytHash()  ' Always empty arrays when not needed
    ResetVariables    ' Verfiy hash objects are empty
    
    RaiseEvent HashProgress(0)   ' Reset progress bar
    
    On Error GoTo 0
    Exit Function

HashString_Error:
    ErrorMsg MODULE_NAME, "HashString", Err.Description
    Erase abytHash()
    Resume HashString_CleanUp

End Function

' ***************************************************************************
' Routine:       HashFile
'
' Description:   Creates a hashed data string based on file contents.
'
' Parameters:    abytInput() - Data file location in byte array format
'
'                Ex:  abytInput() = StrConv("C:\Test Folder\Testfile.doc", vbFromUnicode)
'
' Returns:       Formatted hashed data
' 
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-AUG-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 20-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array.  Added functionality
'              to return data in various formats.
' 04-APR-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in output format.
' ***************************************************************************
Public Function HashFile(ByRef abytInput() As Byte) As Byte()
Attribute HashFile.VB_Description = "Creates a hash output string based on pathfilename."

    Dim abytHash() As Byte
    
    On Error GoTo HashFile_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytInput())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If
    
    Erase abytHash()  ' Always start with empty arrays
    ResetVariables    ' Verfiy hash objects are empty
    
    RaiseEvent HashProgress(0)   ' Reset progress bar
    
    Select Case mlngHashMethod
           
           Case eHASH_SHA1
                Set mobjSHA1 = New cSHA1   ' Instatiate class object
                With mobjSHA1
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    abytHash() = .HashFile(abytInput())
                End With

           Case eHASH_SHA224, eHASH_SHA256, eHASH_SHA384, eHASH_SHA512, _
                eHASH_SHA512_224, eHASH_SHA512_256, eHASH_SHA512_320
                
                Set mobjSHA2 = New cSHA2   ' Instatiate class object
                With mobjSHA2
                    .HashRounds = mlngHashRounds
                    .ReturnLowercase = mblnRetLowercase
                    .HashMethod = (mlngHashMethod - 1)
                    abytHash() = .HashFile(abytInput())
                End With
    End Select
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Erase abytHash()
        GoTo HashFile_CleanUp
    End If

HashFile_CleanUp:
    HashFile = abytHash()
           
    Erase abytHash()  ' Always empty arrays when not needed
    ResetVariables    ' Verfiy hash objects are empty
    
    On Error GoTo 0
    Exit Function

HashFile_Error:
    ErrorMsg MODULE_NAME, "HashFile", Err.Description
    Erase abytHash()
    Resume HashFile_CleanUp

End Function

Private Sub ResetVariables()

    ' Free class objects from memory
    Set mobjSHA1 = Nothing
    Set mobjSHA2 = Nothing

End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objHash = new cHash
                
    ResetVariables              ' Verify hash objects are empty
    
    ' Preset property values
    StopProcessing = False      ' Continue processing
    HashMethod = eHASH_SHA256   ' Default algorithm
    ReturnLowercase = False     ' Default return of hashed data
    HashRounds = 1              ' Set default number of passes

End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objHash = Nothing
            
    ResetVariables    ' Verfiy hash objects are empty

End Sub


' ***************************************************************************
' ****                      Events from subclasses                       ****
' ***************************************************************************
    
' Progress event is raised from the appropriate
' class and then raised from here to the form.

Private Sub mobjSHA1_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub

Private Sub mobjSHA2_HashProgress(ByVal lngProgress As Long)
    RaiseEvent HashProgress(lngProgress)
End Sub


