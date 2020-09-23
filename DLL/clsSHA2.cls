VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSHA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Project:       Secure Hash Algorithm-2
'                SHA-224, SHA-256, SHA-384, SHA-512, SHA-512/224, SHA-512/256
'
' Module:        clsSHA2.cls
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' Description:   The Secure Hash Algorithm (SHA) is required for use with
'                the Digital Signature Algorithm (DSA) as specified in the
'                Digital Signature Standard (DSS) and whenever a secure
'                hash algorithm is required for federal applications.  For
'                a message of length < 2^64 bits, this algorithm produces a
'                condensed representation of the message called a
'                message digest. The message digest is used during
'                generation of a signature for the message.  This
'                also used to compute a message digest for the received
'                version of the message during the process of verifying the
'                signature.  Any change to the message in transit will,
'                with very high probability, result in a different message
'                digest, and the signature will fail to verify.
'
'                These algorithm have been tested to be accurate in
'                accordance with FIPS 180-3 publication dated October 2008.
'                Also, additional test vectors by Aaron Gifford at
'                http://www.adg.us/computers/sha.html
'
'                According to FIPS 180-3 there are only two differences
'                between SHA-224 and SHA-256.
'
'                     1.  The initalizing values are different
'                     2.  Just the left most 224 bits (28 bytes) are saved
'
'                According to FIPS 180-3 there are only two differences
'                between SHA-384 and SHA-512.
'
'                     1.  The initalizing values are different
'                     2.  SHA-384 only uses the first six elements for the output.
'                         SHA-512 uses all eight elements for the output.
'
' REFERENCE:     NIST (National Institute of Standards and Technology)
'                (FIPS, Special Publications)
'                http://csrc.nist.gov/publications/PubsFIPS.html
'
'                FIPS 180-2 (Federal Information Processing Standards Publication)
'                dated 1-Aug-2002, with Change Notice 1, dated 25-Feb-2004
'                http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf
'
'                FIPS 180-3 (Federal Information Processing Standards Publication)
'                dated Oct-2008 (supercedes FIPS 180-2)
'                http://csrc.nist.gov/publications/fips/fips180-3/fips180-3_final.pdf
'
'                FIPS 180-4 (Federal Information Processing Standards Publication)
'                dated Mar-2012 (Supercedes FIPS-180-3)
'                http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
'
'                Examples of the implementation of the secure hash algorithms
'                SHA-1, SHA-224, SHA-256, SHA-384, SHA-512, SHA-512/224 and
'                SHA-512/256, can be found at:
'                http://csrc.nist.gov/groups/ST/toolkit/examples.html
'                http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA2_Additional.pdf
'
'                Aaron Gifford's additional test vectors
'                http://www.adg.us/computers/sha.html
'
' Warning:       Feb-2005:  SHA-1 has been compromised.  Recommended that
'                you do not use for password or document authentication.
'                http://www.schneier.com/blog/archives/2005/02/sha1_broken.html
'                http://csrc.nist.gov/groups/ST/toolkit/documents/shs/NISTHashComments-final.pdf
'
'                March 15, 2006:  The SHA-2 family of hash functions
'                (i.e., SHA-224, SHA-256, SHA-384 and SHA-512) may be used
'                by Federal agencies for all applications using secure hash
'                algorithms. Federal agencies should stop using SHA-1 for
'                digital signatures, digital time stamping and other
'                applications that require collision resistance as soon as
'                practical, and must use the SHA-2 family of hash functions
'                for these applications after 2010. After 2010, Federal
'                agencies may use SHA-1 only for the following applications:
'                    - hash-based message authentication codes (HMACs)
'                    - key derivation functions (KDFs)
'                    - random number generators (RNGs)
'                Regardless of use, NIST encourages application and protocol
'                designers to use the SHA-2 family of hash functions for all
'                new applications and protocols.
'                http://csrc.nist.gov/groups/ST/hash/policy.html
'
'                Export Control: Certain cryptographic devices and technical
'                data regarding them are subject to Federal export controls.
'                Exports of cryptographic modules implementing this standard
'                and technical data regarding them must comply with these
'                Federal regulations and be licensed by the Bureau of Export
'                Administration of the U.S. Department of Commerce.
'                Information about export regulations is available at:
'                http://www.bis.doc.gov/index.htm
'
' ===========================================================================
'
'  Secure Hash Algorithm-2 (SHA2) in accordance with ISO test vectors
'  Message Hash results using one (1) pass
'  All other outputs are my experiment.
'  Hash SHA-512/320 was created in accordance with instructions given
'  in FIPS 180-4, Section 5.3.6
'
'      "a"
'          SHA-224       ABD37534C7D9A2EFB9465DE931CD7055FFDB8879563AE98078D6D6D5
'          SHA-256       CA978112CA1BBDCAFAC231B39A23DC4DA786EFF8147C4E72B9807785AFEE48BB
'          SHA-384       54A59B9F22B0B80880D8427E548B7C23ABD873486E1F035DCE9CD697E85175033CAA88E6D57BC35EFAE0B5AFD3145F31
'          SHA-512       1F40FC92DA241694750979EE6CF582F2D5D7D28E18335DE05ABC54D0560E0F5302860C652BF08D560252AA5E74210546F369FBBBCE8C12CFC7957B2652FE9A75
'          SHA-512/224   D5CDB9CCC769A5121D4175F2BFDD13D6310E0D3D361EA75D82108327
'          SHA-512/256   455E518824BC0601F9FB858FF5C37D417D67C2F8E0DF2BABE4808858AEA830F8
'          SHA-512/320   409391CAB4F921FF20ADBF8591F6573A5D9F7BFFD121BDAECED691F986A7931FD8DEF6A8C9FDB2EA
'
'      "abc"
'          SHA-224       23097D223405D8228642A477BDA255B32AADBCE4BDA0B3F7E36C9DA7
'          SHA-256       BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD
'          SHA-384       CB00753F45A35E8BB5A03D699AC65007272C32AB0EDED1631A8B605A43FF5BED8086072BA1E7CC2358BAECA134C825A7
'          SHA-512       DDAF35A193617ABACC417349AE20413112E6FA4E89A97EA20A9EEEE64B55D39A2192992A274FC1A836BA3C23A3FEEBBD454D4423643CE80E2A9AC94FA54CA49F
'          SHA-512/224   4634270F707B6A54DAAE7530460842E20E37ED265CEEE9A43E8924AA
'          SHA-512/256   53048E2681941EF99B2E29B76B4C7DABE4C2D0C634FC6D46E0E2F13107E7AF23
'          SHA-512/320   0F7567ED5B9C77C089BE0D0F74EBD5DDA19FCC52DB4018D03036E6AE7FBFCEDA2567E1FEE10E37DC
'
'      "message digest"
'          SHA-224       2CB21C83AE2F004DE7E81C3C7019CBCB65B71AB656B22D6D0C39B8EB
'          SHA-256       F7846F55CF23E14EEBEAB5B4E1550CAD5B509E3348FBC4EFA3A1413D393CB650
'          SHA-384       473ED35167EC1F5D8E550368A3DB39BE54639F828868E9454C239FC8B52E3C61DBD0D8B4DE1390C256DCBB5D5FD99CD5
'          SHA-512       107DBF389D9E9F71A3A95F6C055B9251BC5268C2BE16D6C13492EA45B0199F3309E16455AB1E96118E8A905D5597B72038DDB372A89826046DE66687BB420E7C
'          SHA-512/224   AD1A4DB188FE57064F4F24609D2A83CD0AFB9B398EB2FCAEAAE2C564
'          SHA-512/256   0CF471FD17ED69D990DAF3433C89B16D63DEC1BB9CB42A6094604EE5D7B4E9FB
'          SHA-512/320   BA45D0F0098C7E1047DDBD02BEEC953FE03BFD6F14E83091579BF9A5EC9AADCC87AAF5987ACE02CA
'
'      "abcdefghijklmnopqrstuvwxyz"
'          SHA-224       45A5F72C39C5CFF2522EB3429799E49E5F44B356EF926BCF390DCCC2
'          SHA-256       71C480DF93D6AE2F1EFAD1447C66C9525E316218CF51FC8D9ED832F2DAF18B73
'          SHA-384       FEB67349DF3DB6F5924815D6C3DC133F091809213731FE5C7B5F4999E463479FF2877F5F2936FA63BB43784B12F3EBB4
'          SHA-512       4DBFF86CC2CA1BAE1E16468A05CB9881C97F1753BCE3619034898FAA1AABE429955A1BF8EC483D7421FE3C1646613A59ED5441FB0F321389F77F48A879C7B1F1
'          SHA-512/224   FF83148AA07EC30655C1B40AFF86141C0215FE2A54F767D3F38743D8
'          SHA-512/256   FC3189443F9C268F626AEA08A756ABE7B726B05F701CB08222312CCFD6710A26
'          SHA-512/320   82E68E855F9C9F05E2C1DE664336A126649383F450C727B5C466056F231804A66152DFAFE0D55DFD
'
'      "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
'          SHA-224       75388B16512776CC5DBA5DA1FD890150B0C6455CB4F58B1952522525
'          SHA-256       248D6A61D20638B8E5C026930C3E6039A33CE45964FF2167F6ECEDD419DB06C1
'          SHA-384       3391FDDDFC8DC7393707A65B1B4709397CF8B1D162AF05ABFE8F450DE5F36BC6B0455A8520BC4E6F5FE95B1FE3C8452B
'          SHA-512       204A8FC6DDA82F0A0CED7BEB8E08A41657C16EF468B228A8279BE331A703C33596FD15C13B1B07F9AA1D3BEA57789CA031AD85C7A71DD70354EC631238CA3445
'          SHA-512/224   E5302D6D54BB242275D1E7622D68DF6EB02DEDD13F564C13DBDA2174
'          SHA-512/256   BDE8E1F9F19BB9FD3406C90EC6BC47BD36D8ADA9F11880DBC8A22A7078B6A461
'          SHA-512/320   30F2DF31D6AA8ACDA222464BF24312A2C0CB3280D6D0B277ADFEF33522A59CF49561E522BE5D63F1
'
'      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
'          SHA-224       BFF72B4FCB7D75E5632900AC5F90D219E05E97A7BDE72E740DB393D9
'          SHA-256       DB4BFCBD4DA0CD85A60C3C37D3FBD8805C77F15FC6B1FDFE614EE0A7C8FDB4C0
'          SHA-384       1761336E3F7CBFE51DEB137F026F89E01A448E3B1FAFA64039C1464EE8732F11A5341A6F41E0C202294736ED64DB1A84
'          SHA-512       1E07BE23C26A86EA37EA810C8EC7809352515A970E9253C26F536CFC7A9996C45C8370583E0A78FA4A90041D71A4CEAB7423F19C71B9D5A3E01249F0BEBD5894
'          SHA-512/224   A8B4B9174B99FFC67D6F49BE9981587B96441051E16E6DD036B140D3
'          SHA-512/256   CDF1CC0EFFE26ECC0C13758F7B4A48E000615DF241284185C39EB05D355BB9C8
'          SHA-512/320   BD11188553E0E4D59BC8D889D024C5275403CE2153C5E3178F88F8FD443E23B936738794AF97F31B
'
'      "1234567890" replicate 8 times
'          SHA-224       B50AECBE4E9BB0B57BC5F3AE760A8E01DB24F203FB3CDCD13148046E
'          SHA-256       F371BC4A311F2B009EEF952DD83CA80E2B60026C8E935592D0F9C308453C813E
'          SHA-384       B12932B0627D1C060942F5447764155655BD4DA0C9AFA6DD9B9EF53129AF1B8FB0195996D2DE9CA0DF9D821FFEE67026
'          SHA-512       72EC1EF1124A45B047E8B7C75A932195135BB61DE24EC0D1914042246E0AEC3A2354E093D76F3048B456764346900CB130D2A4FD5DD16ABB5E30BCB850DEE843
'          SHA-512/224   AE988FAAA47E401A45F704D1272D99702458FEA2DDC6582827556DD2
'          SHA-512/256   2C9FDBC0C90BDD87612EE8455474F9044850241DC105B1E8B94B8DDF5FAC9148
'          SHA-512/320   19103023C00129F4C65B62D9799EDF8833A2276447757ABC957B3D4124B2DD2A7D34E986B70D13CD
'
'      1 million Letter "a"
'          SHA-224       20794655980C91D8BBB4C1EA97618A4BF03F42581948B2EE4EE7AD67
'          SHA-256       CDC76E5C9914FB9281A1C7E284D73E67F1809A48A497200E046D39CCC7112CD0
'          SHA-384       9D0E1809716474CB086E834E310A4A1CED149E9C00F248527972CEC5704C2A5B07B8B3DC38ECC4EBAE97DDD87F3D8985
'          SHA-512       E718483D0CE769644E2E42C7BC15B4638E1F98B13B2044285632A803AFA973EBDE0FF244877EA60A4CB0432CE577C31BEB009C5C2C49AA2E4EADB217AD8CC09B
'          SHA-512/224   37AB331D76F0D36DE422BD0EDEB22A28ACCD487B7A8453AE965DD287
'          SHA-512/256   9A59A052930187A97038CAE692F30708AA6491923EF5194394DC68D56C74FB21
'          SHA-512/320   5A19AABC267F08806F4FDD070423F9BCFF606481D7E0834A7D2C289DBEA840FF836258EC71E23035
'
'      1 million binary zeroes
'          SHA-224       3A5D74B68F14F3A4B2BE9289B8D370672D0B3D2F53BC303C59032DF3
'          SHA-256       D29751F2649B32FF572B5E0A9F541EA660A50F94FF0BEEDFB0B692B924CC8025
'          SHA-384       8A1979F9049B3FFF15EA3A43A4CF84C634FD14ACAD1C333FECB72C588B68868B66A994386DC0CD1687B9EE2E34983B81
'          SHA-512       CE044BC9FD43269D5BBC946CBEBC3BB711341115CC4ABDF2EDBC3FF2C57AD4B15DEB699BDA257FEA5AEF9C6E55FCF4CF9DC25A8C3CE25F2EFE90908379BFF7ED
'          SHA-512/224   7576F5B118A2DDC31AB05C641F04027FED5F1CBB65894D17EC664466
'          SHA-512/256   8B620FF17FD0414C7C3567704F9E275A5C37801720C75DC05CF81558E4A0F965
'          SHA-512/320   68DBEAF6CC9337DF8FACA863DABEAC6388CC60C267B30B41A4019398429D04558C377C761965CDC0
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
' 24-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              ReWrote this module
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote LoadWordArray() to handle both 32 and 64 bit processing
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 21-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Update reference to FIPS 180-3 and new test vectors
' 28-DEC-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Thanks to Alfred Hellmüller for notifying me that he had a
'                problem performing a hash with a specific file.  The hashed
'                results were not the same as with other hashing utilities.
'                I found that when a hex value is calculated from a number to
'                hex, a series of letters are created (hex).
'                    Ex:  File size = 5360 bytes -> 42880 bits -> A780 hex
'                The problem arose when converting the hex data back to a long
'                integer.
'                    Ex:  &hA780 -> -22656  (Bad)
'                By appending a trailing ampersand to the hex value the
'                conversion is corrected.
'                    Ex:  &hA780& -> 42880  (Good)
'                See Transform() routine.
'              - Added property HashRounds() to set number of rounds for hashing.
'              - Added property HashMethod() to determine hash algorithm.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while hashing data
'                (Transform_32() and Transform_64() routines).
'              - Removed obsolete code.  Updated documentation.
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' 02-May-2012  Kenneth Ives  kenaso@tx.rr.com 
'              Updated explanation of computing values for SHA-512/t.
'              See LoadConstArrays_64() routine.
' 20-May-2012  Kenneth Ives  kenaso@tx.rr.com 
'              Added hash algorithm SHA-512/320.  See LoadConstArrays_64()
'              and FormatOutput() routines.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module constants
' ***************************************************************************
  Private Const MODULE_NAME  As String = "clsSHA2"
  Private Const ZEROES_8     As String = "00000000"
  Private Const ZEROES_16    As String = ZEROES_8 & ZEROES_8
  Private Const ZEROES_32    As String = ZEROES_16 & ZEROES_16
  Private Const MAX_ROUNDS   As Long = 10
  Private Const MSG_LEN_32   As Long = 128          ' 128 bytes long
  Private Const MSG_LEN_64   As Long = 256          ' 256 bytes long
  Private Const KB_32        As Long = &H8000&      ' 32768

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumSHA2_ALGORITHM
      eSHA224       ' 0
      eSHA256       ' 1
      eSHA384       ' 2
      eSHA512       ' 3
      eSHA512_224   ' 4  As per FIPS 180-4 (dtd March-2012)
      eSHA512_256   ' 5  As per FIPS 180-4 (dtd March-2012)
      eSHA512_320   ' 6  My creation as per FIPS 180-4 (dtd March-2012)
  End Enum

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng Hash
' Variable name:     malngHash
'
' ***************************************************************************
  Private mblnRetLowercase As Boolean
  Private mcurMaxSize      As Currency
  Private mcurAccrued      As Currency
  Private mcurBitCount     As Currency
  Private mcurCurrBitCnt   As Currency
  Private mlngHashRounds   As Long
  Private malngHash()      As Long
  Private malngConst()     As Long
  Private mstrBitsInHex    As String
  Private mastrHash()      As String
  Private mastrConst()     As String
  Private mobjMath32       As cMath32
  Private mobjMath64       As cMath64
  Private mlngHashMethod   As enumSHA2_ALGORITHM


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let HashMethod(ByVal lngHashMethod As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = eSHA256
    Select Case lngHashMethod
           Case 0 To 6: mlngHashMethod = lngHashMethod
           Case Else:   mlngHashMethod = eSHA256
    End Select

End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngHashRounds = lngData   ' Good selection
           Case Else:            mlngHashRounds = 1         ' Default value
    End Select

End Property

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = False
    mblnRetLowercase = blnData
End Property


' *************************************************************************** 
' ****                      Methods                                      **** 
' *************************************************************************** 

' *************************************************************************** 
' Routine:       HashFile 
' 
' Description:   Function to create a unique hex string representation of 
'                the data passed.  Can process files larger than 2 gb.
'                See Readme.txt file for more information. 
' 
' Syntax:
'    With mobjSHA
'        abytData() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'        mabytHashed() = .HashFile(abytData())                        ' hash data and return as Byte array
'        txtData.Text = StrConv(mabytHashed(), vbUnicode)             ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - Path/filename in byte array to be hashed 
' 
' Returns:       Hashed data string in byte array 
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com 
'              Changed input parameter to byte array for security purposes 
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' *************************************************************************** 
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim hFile        As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    On Error GoTo HashFile_CleanUp

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Set objBigFiles = New cBigFiles  ' Instantiate class object
    ResetVariables                   ' Reset class variables
    Initialize                       ' load the work and constant arrays
    curFilePos = 0@                  ' set to first position in the file
    Erase abytData()                 ' Start with an empty array

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    With objBigFiles
        strSource = ByteArrayToString(abytFile())           ' convert byte array to string
        .CalcFileSize strSource, curAmtLeft, mstrBitsInHex  ' Get the size of the file
        mcurBitCount = curAmtLeft * 8                       ' Get total bit count for comparison

        ' Calculate amount (double for hex chars)
        ' plus number of hashing rounds requested
        mcurMaxSize = (curAmtLeft * 2) * mlngHashRounds

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        ' Determine which class object to instantiate
        Select Case mlngHashMethod

               ' 32-bit processing
               Case eSHA224, eSHA256
                    Set mobjMath32 = New cMath32
                    mstrBitsInHex = Right$(ZEROES_16 & mstrBitsInHex, 16)

               ' 64-bit processing
               Case eSHA384, eSHA512, eSHA512_224, eSHA512_256, eSHA512_320
                    Set mobjMath64 = New cMath64
                    mstrBitsInHex = Right$(ZEROES_32 & mstrBitsInHex, 32)
        End Select

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        Do
            Erase abytData()                              ' empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)

            ' Process the message data into the appropriate
            ' string length and perform the bit manipulations.
            If Not LoadWordArray(abytData()) Then
                gblnStopProcessing = True
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If

        Loop

        .API_CloseFile hFile    ' Close all files opened by this routine
    End With

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    ' Concatenate the state array into one data
    ' string and return the digital signature.
    HashFile = FormatOutput()

HashFile_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close all files opened by this routine
    ResetVariables                   ' Reset class variables
    Erase abytData()                 ' Always empty arrays when not needed
    Set objBigFiles = Nothing        ' Always free objects from memory

End Function

' *************************************************************************** 
' Routine:       HashString 
' 
' Description:   Function to create a unique hex string representation of 
'                the data passed.  See Readme.txt file for more information. 
' 
' Syntax:
'    With mobjSHA
'        abytData() = StrConv("abc", vbFromUnicode)      ' convert string data to byte array 
'        abytHashed() = .HashString(abytData())          ' hash data and return as Byte array
'        txtData.Text = StrConv(abytHashed(), vbUnicode) ' convert byte array to string data
'    End With
' 
' Parameters:    abytData() - data string in byte array to be hashed 
' 
' Returns:       Hashed data string in byte array 
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com 
'              Changed input parameter to byte array for security purposes 
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' *************************************************************************** 
Public Function HashString(ByRef abytData() As Byte) As Byte()

    On Error GoTo HashString_CleanUp

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If

    ResetVariables   ' Reset class variables
    Initialize       ' load work and constant arrays

    mcurAccrued = 0@                                              ' Init accrued progress
    mcurMaxSize = CCur((UBound(abytData) * 2@) * mlngHashRounds)  ' Calc max size
    mcurBitCount = CCur((UBound(abytData) + 1) * 8@)              ' Get the bit count for the message
    mstrBitsInHex = Hex$(mcurBitCount)                            ' Convert bit count into hex

    ' Determine which class object to instantiate
    Select Case mlngHashMethod

           ' 32-bit processing
           Case eSHA224, eSHA256
                Set mobjMath32 = New cMath32
                mstrBitsInHex = Right$(ZEROES_16 & mstrBitsInHex, 16)

           ' 64-bit processing
           Case eSHA384, eSHA512, eSHA512_224, eSHA512_256, eSHA512_320
                Set mobjMath64 = New cMath64
                mstrBitsInHex = Right$(ZEROES_32 & mstrBitsInHex, 32)
    End Select

    ' Process the message data into the appropriate
    ' string length and perform the bit manipulations.
    If LoadWordArray(abytData()) Then

        ' Concatenate the state array into one data
        ' string and return the digital signature.
        HashString = FormatOutput()
    Else
        gblnStopProcessing = True
    End If

HashString_CleanUp:
    ResetVariables   ' Reset class variables

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

' *************************************************************************** 
' Routine:       FormatOutput 
' 
' Description:   Concatenate hashed values into one string.
'
'                Output length in hex format:
'                    SHA-224     =  56 characters
'                    SHA-256     =  64 characters
'                    SHA-384     =  96 characters
'                    SHA-512     = 128 characters
'                    SHA-512/224 =  56 characters
'                    SHA-512/256 =  64 characters
'                    SHA-512/320 =  80 characters
' 
' Returns:       Concatenated hashed data in byte array 
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 06-SEP-2002  Kenneth Ives  kenaso@tx.rr.com 
'              Routine created 
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' *************************************************************************** 
Private Function FormatOutput() As Byte()

    ' Called by HashFile()
    '           HashString()

    Dim strHex     As String
    Dim intIndex   As Integer
    Dim intPointer As Integer

    strHex = Space$(256)   ' Preload output string with spaces
    intPointer = 1         ' Starting position in output string

    ' Concatenate the hashed values,
    ' in hex format, into one string
    Select Case mlngHashMethod

           Case eSHA224
                ' Save first 56 characters
                For intIndex = 0 To 6
                    Mid$(strHex, intPointer, 8) = Right$(ZEROES_8 & Hex$(malngHash(intIndex)), 8)
                    intPointer = intPointer + 8
                Next intIndex

           Case eSHA256
                ' Save first 64 characters
                For intIndex = 0 To 7
                    Mid$(strHex, intPointer, 8) = Right$(ZEROES_8 & Hex$(malngHash(intIndex)), 8)
                    intPointer = intPointer + 8
                Next intIndex

           Case eSHA384
                ' Save first 96 characters
                For intIndex = 0 To 5
                    Mid$(strHex, intPointer, 16) = mastrHash(intIndex)
                    intPointer = intPointer + 16
                Next intIndex

           Case eSHA512
                ' Save first 128 characters
                For intIndex = 0 To 7
                    Mid$(strHex, intPointer, 16) = mastrHash(intIndex)
                    intPointer = intPointer + 16
                Next intIndex

           Case eSHA512_224
                ' Capture first 64 characters
                For intIndex = 0 To 3
                    Mid$(strHex, intPointer, 16) = mastrHash(intIndex)
                    intPointer = intPointer + 16
                Next intIndex

                strHex = Left$(strHex, 56)  ' Save first 56 characters

           Case eSHA512_256
                ' Save first 64 characters
                For intIndex = 0 To 3
                    Mid$(strHex, intPointer, 16) = mastrHash(intIndex)
                    intPointer = intPointer + 16
                Next intIndex

           Case eSHA512_320
                ' Save first 80 characters
                For intIndex = 0 To 4
                    Mid$(strHex, intPointer, 16) = mastrHash(intIndex)
                    intPointer = intPointer + 16
                Next intIndex
    End Select

    If mblnRetLowercase Then
        strHex = LCase$(Trim$(strHex))   ' Convert string to lowercase
    Else
        strHex = UCase$(Trim$(strHex))   ' Convert string to uppercase
    End If

    FormatOutput = StringToByteArray(strHex)   ' convert string data to byte array

    RaiseEvent HashProgress(100)   ' Update progress bar

End Function

' *************************************************************************** 
' Routine:       LoadWordArray
' 
' Description:   All data is stored in Big_Endian format with the Most 
'                Significant Bit (MSB) first. 
' 
'                The message 'M' shall be padded before hash computation
'                begins. The purpose of this padding is to ensure that the
'                padded message is a multiple of 512 or 1024 bits, depending
'                on the algorithm.
'
'         32-Bit Format
'                Suppose the length of the message 'M', in bits, is 'l'
'                bits. Append the bit '1' to the end of the message,
'                followed by 'k' zero bits, where 'k' is the smallest
'                non-negative solution to the equation l+1+k=448 mod 512.
'                Then append the 64-bit block that is equal to the number
'                'l' expressed using a binary representation. The length of
'                the padded message should now be a multiple of 512 bits.
' 
'         64-Bit Format
'                Suppose the length of the message 'M', in bits, is 'l'
'                bits. Append the bit '1' to the end of the message,
'                followed by 'k' zero bits, where 'k' is the smallest
'                non-negative solution to the equation l+1+k=896 mod 1024.
'                Then append the 128-bit block that is equal to the number
'                'l' expressed using a binary representation. The length of
'                the padded message should now be a multiple of 1024 bits.
' 
' Parameters:    abytData() - message to be hashed in a byte array 
' 
' Returns:       TRUE if process completed successfully.
'                FALSE if user reqested to stop processing or an error
'                      occurred.
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 21-Feb-2008  Kenneth Ives  kenaso@tx.rr.com 
'              Rewrote routine 
' 20-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added flag to designate 32-bit or 64-bit processing
' *************************************************************************** 
Private Function LoadWordArray(ByRef abytData() As Byte) As Boolean

    ' Called by HashFile()
    '           HashString()

    Dim strMsg         As String
    Dim strHex         As String
    Dim strTemp        As String
    Dim lngIndex       As Long
    Dim lngPadding     As Long
    Dim lngMaxBytes    As Long
    Dim lngPosition    As Long
    Dim lngMsgLength   As Long
    Dim lngSpaceAlloc  As Long
    Dim lngMsgLenInHex As Long
    Dim bln32Bit       As Boolean
    Dim abytMsg()      As Byte

    Const ROUTINE_NAME As String = "LoadWordArray"
    
    On Error GoTo LoadWordArray_Error

    Select Case mlngHashMethod

           ' 32-bit processing
           Case eSHA224, eSHA256
                lngMsgLength = MSG_LEN_32
                lngSpaceAlloc = 512
                lngMaxBytes = 64
                bln32Bit = True

           ' 64-bit processing
           Case eSHA384, eSHA512, eSHA512_224, eSHA512_256, eSHA512_320
                lngMsgLength = MSG_LEN_64
                lngSpaceAlloc = 1024
                lngMaxBytes = 128
                bln32Bit = False
    End Select

    LoadWordArray = True  ' Preset flag to TRUE
    lngMsgLenInHex = Len(mstrBitsInHex)

    strMsg = ByteArrayToString(abytData())                    ' convert to string data
    mcurCurrBitCnt = mcurCurrBitCnt + CCur(Len(strMsg) * 8)   ' calc the ongoing bit count
    strTemp = vbNullString                                    ' Empty string variables
    strHex = vbNullString

    ' Incoming message is in equal multiples
    If Len(strMsg) Mod lngMaxBytes = 0 Then

        ' Process message in equal lengths. There
        ' is no need for any padding at this time.
        Do
            strTemp = vbNullString   ' Empty string variables
            strHex = vbNullString

            strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
            strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used

            ' See if we have some data to work with
            If Len(strTemp) > 0 Then

                ' the string of data is the exact length needed
                If Len(strTemp) = lngMaxBytes Then

                    strHex = Space$(lngSpaceAlloc)
                    lngPosition = 1

                    ' convert data to hex string
                    For lngIndex = 1 To Len(strTemp)
                        Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                        lngPosition = lngPosition + 2
                    Next lngIndex

                    If bln32Bit Then
                        Erase abytMsg()                                       ' Always start with an empty array
                        abytMsg() = mobjMath32.HexStringToByteArray(strHex)   ' Convert hex string to byte array
                        Transform_32 abytMsg()                                ' 32-bit processing
                    Else
                        Transform_64 strHex   ' 64-bit processing
                    End If

                End If

            Else
                Exit Do  ' exit Do..Loop
            End If

        Loop While Len(strMsg) > 0
    End If

    strTemp = vbNullString   ' Empty string variables
    strHex = vbNullString

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadWordArray_CleanUp
    End If

    ' if the length of the message is not
    ' in equal multiples
    If Len(strMsg) > 0 Then

        If Len(strMsg) >= lngMaxBytes Then

            Do
                strTemp = vbNullString   ' Empty string variables
                strHex = vbNullString

                strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
                strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used

                If Len(strTemp) = lngMaxBytes Then

                    strHex = Space$(lngSpaceAlloc)
                    lngPosition = 1

                    ' convert data to hex string
                    For lngIndex = 1 To Len(strTemp)
                        Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                        lngPosition = lngPosition + 2
                    Next lngIndex

                    If bln32Bit Then
                        Erase abytMsg()                                       ' Always start with an empty array
                        abytMsg() = mobjMath32.HexStringToByteArray(strHex)   ' Convert hex string to byte array
                        Transform_32 abytMsg()                                ' 32-bit processing
                    Else
                        Transform_64 strHex   ' 64-bit processing
                    End If

                Else
                    Exit Do  ' exit Do..Loop
                End If

            Loop

            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                GoTo LoadWordArray_CleanUp
            End If

        Else
            strTemp = Mid$(strMsg, 1, lngMaxBytes)   ' Capture chunk of data message
            strMsg = Mid$(strMsg, lngMaxBytes + 1)   ' Resize data message to what was not used
        End If

        ' the string of data is less than the length needed
        If (Len(strTemp) > 0) And (Len(strMsg) = 0) Then

            strTemp = strTemp & Chr$(128)
            strHex = Space$(lngSpaceAlloc)
            lngPosition = 1

            ' convert data to hex string
            For lngIndex = 1 To Len(strTemp)
                Mid$(strHex, lngPosition, 2) = Right$("00" & Hex$(Asc(Mid$(strTemp, lngIndex, 1))), 2)
                lngPosition = lngPosition + 2
            Next lngIndex

            strHex = Trim$(strHex)  ' remove all leading and trailing blanks
            strTemp = vbNullString

            ' if the len of the hex string plus the element length is less
            ' than or equal to the required record length then append the
            ' appropriate number of zeroes and the hex representation of
            ' the message length
            If (Len(strHex) < lngMsgLength) Then

                ' Calculate amount of padding needed
                lngPadding = lngMsgLength - Len(strHex)

                Select Case lngPadding

                       Case lngMsgLenInHex
                            ' If the amount of padding needed is the
                            ' exact length of the bit string then
                            ' append, process and leave
                            strHex = strHex & mstrBitsInHex

                       Case Is < lngMsgLenInHex

                            ' *** String 1 ***
                            ' Hex string is exact length needed.
                            ' Will need to create an extra string.
                            strHex = strHex & String$(lngPadding, "0")

                            If bln32Bit Then
                                Erase abytMsg()                                       ' Always start with an empty array
                                abytMsg() = mobjMath32.HexStringToByteArray(strHex)   ' Convert hex string to byte array
                                Transform_32 abytMsg()                                ' 32-bit processing
                            Else
                                Transform_64 strHex   ' 64-bit processing
                            End If

                            ' *** String 2 ***
                            ' Create a string of zeroes and the
                            ' bit count of the message length.
                            strHex = vbNullString
                            lngPadding = lngMsgLength - lngMsgLenInHex
                            strHex = String$(lngPadding, "0") & mstrBitsInHex

                       Case Is > lngMsgLenInHex
                            ' Adjust amount of padding to fill in
                            ' the gap between data and number of bits
                            lngPadding = lngMsgLength - (Len(strHex) + lngMsgLenInHex)
                            strHex = strHex & String$(lngPadding, "0") & mstrBitsInHex
                End Select

                If bln32Bit Then
                    Erase abytMsg()                                       ' Always start with an empty array
                    abytMsg() = mobjMath32.HexStringToByteArray(strHex)   ' Convert hex string to byte array
                    Transform_32 abytMsg()                                ' 32-bit processing
                Else
                    Transform_64 strHex   ' 64-bit processing
                End If

                GoTo LoadWordArray_CleanUp

            Else
                ' *** String 1 ***
                ' Hex string is exact length needed.
                ' Will need to create an extra string.
                If bln32Bit Then
                    Erase abytMsg()                                       ' Always start with an empty array
                    abytMsg() = mobjMath32.HexStringToByteArray(strHex)   ' Convert hex string to byte array
                    Transform_32 abytMsg()                                ' 32-bit processing
                Else
                    Transform_64 strHex   ' 64-bit processing
                End If

                ' *** String 2 ***
                ' Create a string of zeroes and the
                ' bit count of the message length.
                strHex = vbNullString
                lngPadding = lngMsgLength - lngMsgLenInHex
                strHex = String$(lngPadding, "0") & mstrBitsInHex

                If bln32Bit Then
                    Erase abytMsg()                                       ' Always start with an empty array
                    abytMsg() = mobjMath32.HexStringToByteArray(strHex)   ' Convert hex string to byte array
                    Transform_32 abytMsg()                                ' 32-bit processing
                Else
                    Transform_64 strHex   ' 64-bit processing
                End If

                GoTo LoadWordArray_CleanUp

            End If
        End If
    End If

    ' No data found in current work string
    If Len(strMsg) = 0 Then

        ' See if we have reached the end
        ' of the message file or string
        If mcurCurrBitCnt = mcurBitCount Then

            ' End of message indicator is the starting character followed by a
            ' string of zeroes and the hex representation of the message length
            lngPadding = lngMsgLength - (lngMsgLenInHex + 2)
            strHex = Hex$(128) & String$(lngPadding, "0") & mstrBitsInHex

            If bln32Bit Then
                Erase abytMsg()                                       ' Always start with an empty array
                abytMsg() = mobjMath32.HexStringToByteArray(strHex)   ' Convert hex string to byte array
                Transform_32 abytMsg()                                ' 32-bit processing
            Else
                Transform_64 strHex   ' 64-bit processing
            End If
        End If
    End If

LoadWordArray_CleanUp:
    Erase abytMsg()         ' Always empty arrays when not needed
    strMsg = vbNullString   ' Verify string variables are empty
    strTemp = vbNullString
    strHex = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        LoadWordArray = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LoadWordArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadWordArray_CleanUp

End Function

' *************************************************************************** 
' Routine:       Transform_32 (SHA-224 & SHA-256) 
' 
' Description:   Perform the main hashing functions here. 
' 
'                The definition of the initial hash value H(0) allows its 
'                implementation to be shared between the algorithms. That 
'                is, the left halves of the SHA-512 words of H(0) are the 
'                words of H(0) for SHA-256. Similarly for SHA-384, the 
'                right halves of the words of H(0) are the words for H(0) 
'                for SHA-224.
' 
'                SHA-224 H(0) =         c1059ed8     Right half of SHA-384 
'                SHA-384 H(0) = cbbb9d5dc1059ed8 
' 
'                SHA-256 H(0) = 6a09e667             Left half of SHA-512 
'                SHA-512 H(0) = 6a09e667f3bbc908 
' 
'                A similar implementation holds true for the constant array. 
' 
'                SHA-224 &
'                SHA-256 K(0) = 428a2f98             Left half of SHA-512 
'                SHA-512 K(0) = 428a2f98d728ae22
' 
' Parameters:    strMsg - 128 byte message string to be hashed
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com 
'              Passed padded message string for possible speed increase.
' 28-DEC-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for notifying me that he had a
'              problem performing a hash with a specific file.  The hashed
'              results were not the same as with other hashing utilities.
'              I found that when a hex value is calculated from a number to
'              hex, a series of letters are created (hex).
'
'                    Ex:  File size = 5360 bytes -> 42880 bits -> A780 hex
'
'              The problem arose when converting the hex data back to a long
'              integer.
'
'                    Ex:  &hA780 -> -22656  (Bad)
'
'              By appending a trailing ampersand to the hex value the
'              conversion is corrected.
'
'                    Ex:  &hA780& -> 42880  (Good)
'
'              Karl Peterson comments:
'              "The trailing ampersand tells VB it's a Long value (4 byte),
'              even for values that could be held in an Integer (2 byte).
'              Basically, if it's &h8000-&hFFFF, use the ampersand (actually,
'              it's never a bad idea), assuming you really do want to store
'              it in a 4-byte variable. Otherwise, VB strives to confuse, by
'              twiddling the sign bit for those higher values."
'              http://bytes.com/topic/visual-basic-net/answers/387151-ampersands-hexadecimal-numbers-whats-rule
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added RaiseEvents to track hash progress
' *************************************************************************** 
Private Sub Transform_32(ByRef abytMsg() As Byte)

    ' Called by LoadWordArray()

    Dim lngTmp1     As Long
    Dim lngTmp2     As Long
    Dim lngIndex    As Long
    Dim lngRounds   As Long
    Dim lngPointer  As Long
    Dim lngProgress As Long
    Dim H32()       As Long
    Dim alngMsg()   As Long

    ReDim H32(8)        ' Size temp work arrays
    ReDim alngMsg(64)
    lngPointer = 0      ' First element of incoming array

    With mobjMath32

        ' ******** Message preparation **************************
        For lngIndex = 0 To 15

            alngMsg(lngIndex) = .ByteArrayToLong(abytMsg(), lngPointer)
            lngPointer = lngPointer + 4

        Next lngIndex

        ' These 32-bit words must now be extended through the
        ' initial hashing phase for an additional forty-eight
        ' 32-bit words.
        For lngIndex = 16 To 63

            lngTmp1 = Gamma1_32(alngMsg(lngIndex - 2))
            lngTmp1 = .UnsignedAdd(lngTmp1, alngMsg(lngIndex - 7))
            lngTmp2 = Gamma0_32(alngMsg(lngIndex - 15))
            lngTmp1 = .UnsignedAdd(lngTmp1, lngTmp2)
            alngMsg(lngIndex) = .UnsignedAdd(lngTmp1, alngMsg(lngIndex - 16))

        Next lngIndex
        ' ******** End of message preparation ********************

        ' Set initial values for hash accumulators
        For lngIndex = 0 To 7
            H32(lngIndex) = malngHash(lngIndex)
        Next lngIndex

        ' Perform required number of iterations
        For lngRounds = 1 To mlngHashRounds

            ' Manipulate the data
            For lngIndex = 0 To 63

                lngTmp1 = .UnsignedAdd(H32(7), Sigma1_32(H32(4)))
                lngTmp1 = .UnsignedAdd(lngTmp1, Ch_32(H32(4), H32(5), H32(6)))
                lngTmp1 = .UnsignedAdd(lngTmp1, malngConst(lngIndex))
                lngTmp1 = .UnsignedAdd(lngTmp1, alngMsg(lngIndex))

                lngTmp2 = .UnsignedAdd(Sigma0_32(H32(0)), Maj_32(H32(0), H32(1), H32(2)))

                H32(7) = H32(6)
                H32(6) = H32(5)
                H32(5) = H32(4)
                H32(4) = .UnsignedAdd(H32(3), lngTmp1)
                H32(3) = H32(2)
                H32(2) = H32(1)
                H32(1) = H32(0)
                H32(0) = .UnsignedAdd(lngTmp1, lngTmp2)

            Next lngIndex

            ' 29-Jan-2010 Update progress bar. Separate
            '             steps for easier debugging
            mcurAccrued = mcurAccrued + MSG_LEN_32
            lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
            RaiseEvent HashProgress(lngProgress)

        Next lngRounds

        ' Add calculated values to what was started in this routine
        For lngIndex = 0 To 7
            malngHash(lngIndex) = .UnsignedAdd(malngHash(lngIndex), H32(lngIndex))
        Next lngIndex

    End With

End Sub

' *************************************************************************** 
' Routine:       Transform_64 (SHA-384 & SHA-512) 
' 
' Description:   Perform the main hashing functions here. 
' 
'                SHA-512 - The H(0-7) (initial work) values were obtained
'                by taking the first 64 bits of the fractional parts of
'                the SQUARE ROOTS of the first 8 prime numbers. 
' 
'                SHA-384 - The H(0-7) (initial work) values were obtained
'                by taking the first 64 bits of the fractional parts of
'                the SQUARE ROOTS of the 9th thru 16th prime numbers. 
' 
'                The K(0-79) (constants) values were obtained by taking the 
'                first 64 bits of the fractional parts of the CUBE ROOTS of
'                the first 80 prime numbers. 
' 
' Parameters:    strMsg - 256 byte message string to be hashed
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com 
'              Passed padded message string for possible speed increase.
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added RaiseEvents to track hash progress
' *************************************************************************** 
Private Sub Transform_64(ByVal strMsg As String)

    ' Called by LoadWordArray()

    Dim lngIndex    As Long
    Dim lngRounds   As Long
    Dim lngPointer  As Long
    Dim lngProgress As Long
    Dim strTmp1     As String
    Dim strTmp2     As String
    Dim H64()       As String
    Dim astrMsg()   As String

    ReDim H64(8)        ' Size temp work arrays
    ReDim astrMsg(80)
    lngPointer = 1      ' First position of incoming message string

    With mobjMath64

        ' ******** Message preparation **************************
        ' load work array with 64-bit words (16 chars per element)
        For lngIndex = 0 To 15

            astrMsg(lngIndex) = Mid$(strMsg, lngPointer, 16)
            lngPointer = lngPointer + 16

        Next lngIndex

        ' These 64-bit words must now be extended through the
        ' initial hashing phase to an additional sixty-four
        ' 64-bit words.
        For lngIndex = 16 To 79

            strTmp1 = .w64HexAdd(Gamma1_64(astrMsg(lngIndex - 2)), astrMsg(lngIndex - 7))
            strTmp1 = .w64HexAdd(strTmp1, Gamma0_64(astrMsg(lngIndex - 15)))
            astrMsg(lngIndex) = .w64HexAdd(strTmp1, astrMsg(lngIndex - 16))

        Next lngIndex
        ' ******** End of message preparation ********************

        ' Set initial values for hash accumulators
        For lngIndex = 0 To 7
            H64(lngIndex) = mastrHash(lngIndex)
        Next lngIndex

        ' Perform required number of iterations
        For lngRounds = 1 To mlngHashRounds

            ' Manipulate the data
            For lngIndex = 0 To 79

                strTmp1 = .w64HexAdd(H64(7), Sigma1_64(H64(4)))
                strTmp1 = .w64HexAdd(strTmp1, Ch_64(H64(4), H64(5), H64(6)))
                strTmp1 = .w64HexAdd(strTmp1, mastrConst(lngIndex))
                strTmp1 = .w64HexAdd(strTmp1, astrMsg(lngIndex))

                strTmp2 = Maj_64(H64(0), H64(1), H64(2))
                strTmp2 = .w64HexAdd(Sigma0_64(H64(0)), strTmp2)

                H64(7) = H64(6)
                H64(6) = H64(5)
                H64(5) = H64(4)
                H64(4) = .w64HexAdd(H64(3), strTmp1)
                H64(3) = H64(2)
                H64(2) = H64(1)
                H64(1) = H64(0)
                H64(0) = .w64HexAdd(strTmp1, strTmp2)

            Next lngIndex

            ' 29-Jan-2010 Update progress bar. Separate
            '             steps for easier debugging
            mcurAccrued = mcurAccrued + MSG_LEN_64
            lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
            RaiseEvent HashProgress(lngProgress)

        Next lngRounds

        ' Add calculated values to the initial
        ' values at the top of this routine
        For lngIndex = 0 To 7
            mastrHash(lngIndex) = .w64HexAdd(mastrHash(lngIndex), H64(lngIndex))
        Next lngIndex

    End With

End Sub

' *************************************************************************** 
' Routine:       Bit manipulation routines 
' 
' Description:   Perform the bit manipulating in accordance with FIPS 180-2 
'                dated 1-Aug-2002, with Change Notice 1, dated 25-Feb-2004 
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-Nov-2004  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routines 
' *************************************************************************** 

' *************************************************************************** 
' 32-Bit manipulation routines 
' *************************************************************************** 

Private Function Maj_32(ByVal XX As Long, _
                        ByVal YY As Long, _
                        ByVal ZZ As Long) As Long

    ' ( x And y) Xor ( x And z) Xor ( y And z)
    Maj_32 = (XX And YY) Xor (XX And ZZ) Xor (YY And ZZ)

End Function

Private Function Ch_32(ByVal XX As Long, _
                       ByVal YY As Long, _
                       ByVal ZZ As Long) As Long

    ' ( x And y) Xor (~x And z)
    Ch_32 = (XX And YY) Xor ((Not XX) And ZZ)

End Function

Private Function Sigma0_32(ByVal lngValue As Long) As Long

    If lngValue = 0 Then
        Sigma0_32 = 0
    Else
        ' ROTR 2(x) Xor ROTR 13(x) Xor ROTR 22(x)
        With mobjMath32
            Sigma0_32 = .w32Rotate(lngValue, -2) Xor _
                        .w32Rotate(lngValue, -13) Xor _
                        .w32Rotate(lngValue, -22)
        End With
    End If

End Function

Private Function Sigma1_32(ByVal lngValue As Long) As Long

    If lngValue = 0 Then
        Sigma1_32 = 0
    Else
        ' ROTR 6(x) Xor ROTR 11(x) Xor ROTR 25(x)
        With mobjMath32
            Sigma1_32 = .w32Rotate(lngValue, -6) Xor _
                        .w32Rotate(lngValue, -11) Xor _
                        .w32Rotate(lngValue, -25)
        End With
    End If

End Function

Private Function Gamma0_32(ByVal lngValue As Long) As Long

    If lngValue = 0 Then
        Gamma0_32 = 0
    Else
        ' ROTR 7(x) Xor ROTR 18(x) Xor SHR 3(x)
        With mobjMath32
            Gamma0_32 = .w32Rotate(lngValue, -7) Xor _
                        .w32Rotate(lngValue, -18) Xor _
                        .w32Shift(lngValue, -3)
        End With
    End If

End Function

Private Function Gamma1_32(ByVal lngValue As Long) As Long

    If lngValue = 0 Then
        Gamma1_32 = 0
    Else
        ' ROTR 17(x) Xor ROTR 19(x) Xor SHR 10(x)
        With mobjMath32
            Gamma1_32 = .w32Rotate(lngValue, -17) Xor _
                        .w32Rotate(lngValue, -19) Xor _
                        .w32Shift(lngValue, -10)
        End With
    End If

End Function


' *************************************************************************** 
' 64-Bit manipulation routines 
' *************************************************************************** 

Private Function Maj_64(ByVal XX As String, _
                        ByVal YY As String, _
                        ByVal ZZ As String) As String

    Dim strTemp1 As String
    Dim strTemp2 As String
    Dim strTemp3 As String
    Dim strTemp4 As String

    ' ( x And y) Xor ( x And z) Xor ( y And z)
    With mobjMath64
        strTemp1 = .w64Hex_AND(XX, YY)
        strTemp2 = .w64Hex_AND(XX, ZZ)
        strTemp3 = .w64Hex_AND(YY, ZZ)
        strTemp4 = .w64Hex_XoR(strTemp1, strTemp2)
        Maj_64 = .w64Hex_XoR(strTemp4, strTemp3)
    End With

End Function

Private Function Ch_64(ByVal XX As String, _
                       ByVal YY As String, _
                       ByVal ZZ As String) As String

    Dim strTemp1 As String
    Dim strTemp2 As String
    Dim strTemp3 As String

    ' ( x And y) Xor (~x And z)
    With mobjMath64
        strTemp1 = .w64Hex_AND(XX, YY)
        strTemp2 = .w64Hex_NOT(XX)
        strTemp3 = .w64Hex_AND(strTemp2, ZZ)
        Ch_64 = .w64Hex_XoR(strTemp1, strTemp3)
    End With

End Function

Private Function Sigma0_64(ByVal XX As String) As String

    Dim strRotR28 As String
    Dim strRotR34 As String
    Dim strRotR39 As String
    Dim strTemp   As String

    ' if all zeroes, return zeroes
    If StrComp(XX, ZEROES_16, vbBinaryCompare) = 0 Then
        Sigma0_64 = ZEROES_16
    Else
        ' ROTR 28(x) Xor ROTR 34(x) Xor ROTR 39(x)
        With mobjMath64
            strRotR28 = .w64Rotate(XX, -28)
            strRotR34 = .w64Rotate(XX, -34)
            strRotR39 = .w64Rotate(XX, -39)
            strTemp = .w64Hex_XoR(strRotR28, strRotR34)
            Sigma0_64 = .w64Hex_XoR(strTemp, strRotR39)
        End With
    End If

End Function

Private Function Sigma1_64(ByVal XX As String) As String

    Dim strRotR14 As String
    Dim strRotR18 As String
    Dim strRotR41 As String
    Dim strTemp   As String

    ' if all zeroes, return zeroes
    If StrComp(XX, ZEROES_16, vbBinaryCompare) = 0 Then
        Sigma1_64 = ZEROES_16
    Else
        ' ROTR 14(x) Xor ROTR 18(x) Xor ROTR 41(x)
        With mobjMath64
            strRotR14 = .w64Rotate(XX, -14)
            strRotR18 = .w64Rotate(XX, -18)
            strRotR41 = .w64Rotate(XX, -41)
            strTemp = .w64Hex_XoR(strRotR14, strRotR18)
            Sigma1_64 = .w64Hex_XoR(strTemp, strRotR41)
        End With
    End If

End Function

Private Function Gamma0_64(ByVal XX As String) As String

    Dim strRotR1   As String
    Dim strRotR8   As String
    Dim strShiftR7 As String
    Dim strTemp    As String

    ' if all zeroes, return zeroes
    If StrComp(XX, ZEROES_16, vbBinaryCompare) = 0 Then
        Gamma0_64 = ZEROES_16
    Else
        ' ROTR 1(x) Xor ROTR 8(x) Xor SHR 7(x)
        With mobjMath64
            strRotR1 = .w64Rotate(XX, -1)
            strRotR8 = .w64Rotate(XX, -8)
            strShiftR7 = .w64Shift(XX, -7)
            strTemp = .w64Hex_XoR(strRotR1, strRotR8)
            Gamma0_64 = .w64Hex_XoR(strTemp, strShiftR7)
        End With
    End If

End Function

Private Function Gamma1_64(ByVal XX As String) As String

    Dim strRotR19  As String
    Dim strRotR61  As String
    Dim strShiftR6 As String
    Dim strTemp    As String

    ' if all zeroes, return zeroes
    If StrComp(XX, ZEROES_16, vbBinaryCompare) = 0 Then
        Gamma1_64 = ZEROES_16
    Else
        ' ROTR 19(x) Xor ROTR 61(x) Xor SHR 6(x)
        With mobjMath64
            strRotR19 = .w64Rotate(XX, -19)
            strRotR61 = .w64Rotate(XX, -61)
            strShiftR6 = .w64Shift(XX, -6)
            strTemp = .w64Hex_XoR(strRotR19, strRotR61)
            Gamma1_64 = .w64Hex_XoR(strTemp, strShiftR6)
        End With
    End If

End Function

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the data to be processed.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= KB_32: GetBlockSize = KB_32
           Case Else:        GetBlockSize = CLng(curAmtLeft)
    End Select

End Function

' *************************************************************************** 
' Routine:       Initialize 
' 
' Description:   Preloads the arrays and constants with predetermined values 
'                to be used in the hashing calculations. 
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-Jan-2007  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' *************************************************************************** 
Private Sub Initialize()

    ' Called by HashFile()
    '           HashString()

    Select Case mlngHashMethod

           Case eSHA224, eSHA256
                LoadConstArrays_32

           Case eSHA384, eSHA512, eSHA512_224, eSHA512_256, eSHA512_320
                LoadConstArrays_64
    End Select

End Sub

' *************************************************************************** 
' Routine:       LoadConstArrays_32 
' 
' Description:   Preload the Work array and Constants array for 32-bit 
'                processing with predetermined values to be used in the 
'                hashing calculations. 
' 
'                The definition of the initial hash value H(0) allows its 
'                implementation to be shared between the algorithms. That 
'                is, the left halves of the SHA-512 words of H(0) are the 
'                words of H(0) for SHA-256. Similarly for SHA-384, the 
'                right halves of the words of H(0) are the words for H(0) 
'                for SHA-224. For example for H(0): 
' 
'                These words were obtained by taking the last thirty-two
'                bits (8 chars) of the fractional parts of the SQUARE ROOTS
'                of the 9th through 16th prime numbers.
'                SHA-224 H(0) =         c1059ed8      Right half of SHA-384 
'                SHA-384 H(0) = cbbb9d5dc1059ed8 
' 
'                These words were obtained by taking the first thirty-two
'                bits (8 chars) of the fractional parts of the SQUARE ROOTS
'                of the first eight prime numbers.
'                SHA-256 H(0) = 6a09e667               Left half of SHA-512 
'                SHA-512 H(0) = 6a09e667f3bbc908 
' 
'                A similar implementation holds true for the constant array. 
' 
'                These words represent the first thirty-two bits of the
'                fractional parts of the cube roots of the first sixty-four
'                prime numbers.
'                SHA-224, -256 K(0) = 428a2f98         Left half of SHA-512 
'                SHA-512       K(0) = 428a2f98d728ae22
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-Jan-2007  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' *************************************************************************** 
Private Sub LoadConstArrays_32()

    ' Called by Initialize()

    Dim lngIndex As Long
    Dim strData  As String
    Dim avntData As Variant

    strData = vbNullString
    avntData = Empty      ' Start with empty variants
    ReDim malngHash(8)    ' Size arrays
    ReDim malngConst(64)

    Select Case mlngHashMethod

           Case eSHA224
                ' These words were obtained by taking the last thirty-two
                ' bits (8 chars) of the fractional parts of the SQUARE ROOTS
                ' of the 9th through 16th prime numbers.
                ' As per FIPS 180-4, Section 5.3.2
                ' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
                malngHash(0) = &HC1059ED8
                malngHash(1) = &H367CD507
                malngHash(2) = &H3070DD17
                malngHash(3) = &HF70E5939
                malngHash(4) = &HFFC00B31
                malngHash(5) = &H68581511
                malngHash(6) = &H64F98FA7
                malngHash(7) = &HBEFA4FA4

           Case eSHA256
                ' These words were obtained by taking the first thirty-two
                ' bits (8 chars) of the fractional parts of the SQUARE ROOTS
                ' of the first eight prime numbers.
                ' As per FIPS 180-4, Section 5.3.3
                ' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
                malngHash(0) = &H6A09E667
                malngHash(1) = &HBB67AE85
                malngHash(2) = &H3C6EF372
                malngHash(3) = &HA54FF53A
                malngHash(4) = &H510E527F
                malngHash(5) = &H9B05688C
                malngHash(6) = &H1F83D9AB
                malngHash(7) = &H5BE0CD19
    End Select

    ' These words represent the first thirty-two bits of the fractional parts
    ' of the cube roots of the first sixty-four prime numbers. In other words,
    ' the left half of SHA-512 constants. The same values are used for SHA-224
    ' and SHA-256.
    strData = strData & "428a2f98 71374491 b5c0fbcf e9b5dba5 3956c25b 59f111f1 "
    strData = strData & "923f82a4 ab1c5ed5 d807aa98 12835b01 243185be 550c7dc3 "
    strData = strData & "72be5d74 80deb1fe 9bdc06a7 c19bf174 e49b69c1 efbe4786 "
    strData = strData & "0fc19dc6 240ca1cc 2de92c6f 4a7484aa 5cb0a9dc 76f988da "
    strData = strData & "983e5152 a831c66d b00327c8 bf597fc7 c6e00bf3 d5a79147 "
    strData = strData & "06ca6351 14292967 27b70a85 2e1b2138 4d2c6dfc 53380d13 "
    strData = strData & "650a7354 766a0abb 81c2c92e 92722c85 a2bfe8a1 a81a664b "
    strData = strData & "c24b8b70 c76c51a3 d192e819 d6990624 f40e3585 106aa070 "
    strData = strData & "19a4c116 1e376c08 2748774c 34b0bcb5 391c0cb3 4ed8aa4a "
    strData = strData & "5b9cca4f 682e6ff3 748f82ee 78a5636f 84c87814 8cc70208 "
    strData = strData & "90befffa a4506ceb bef9a3f7 c67178f2 "

    avntData = Split(strData, Chr$(32))

    For lngIndex = 0 To 63
        malngConst(lngIndex) = Val("&H" & avntData(lngIndex))
    Next lngIndex

    strData = vbNullString
    avntData = Empty   ' Always empty variants when not needed

End Sub

' *************************************************************************** 
' Routine:       LoadConstArrays_64 
' 
' Description:   Preload the Work array and Constants array for 32-bit 
'                processing with predetermined values to be used in the 
'                hashing calculations. 
' 
'                Got this explanation from Ask Dr. Math web site on how to
'                calculate The hex representation of the fractional parts of
'                a whole number.
'                http://mathforum.org/dr.math/ 
' 
'                There is a simple process for converting a base 10 decimal
'                fraction to a "decimal" fraction in another base.  In this
'                case, hex will be our new base (16).  We repeatedly multiply
'                our decimal fraction by the new base, picking off the whole
'                number part each time as the next digit of the final output. 
' 
' -----------------------------------------------------------------------------
' Example:   1st Prime number = 2
'            CUBE ROOT of 2 = 1.2599210498948731647672106073 
' 
'            0.2599210498948731647672106073 * 16 =  4.1587367983179706362753697168 
'            0.1587367983179706362753697168 * 16 =  2.5397887730875301804059154688 
'            0.5397887730875301804059154688 * 16 =  8.636620369400482886494647501 
'            0.636620369400482886494647501  * 16 = 10.185925910407726183914360016 
'            0.185925910407726183914360016  * 16 =  2.974814566523618942629760256 
'            0.974814566523618942629760256  * 16 = 15.597033064377903082076164096 
'            0.597033064377903082076164096  * 16 =  9.552529030046449313218625536 
'            0.552529030046449313218625536  * 16 =  8.840464480743189011498008576 
'            0.840464480743189011498008576  * 16 = 13.447431691891024183968137216 
'            0.447431691891024183968137216  * 16 =  7.158907070256386943490195456 
'            0.158907070256386943490195456  * 16 =  2.542513124102191095843127296 
'            0.542513124102191095843127296  * 16 =  8.680209985635057533490036736 
'            0.680209985635057533490036736  * 16 = 10.883359770160920535840587776 
'            0.883359770160920535840587776  * 16 = 14.133756322574728573449404416 
'            0.133756322574728573449404416  * 16 =  2.140101161195657175190470656 
'            0.140101161195657175190470656  * 16 =  2.241618579130514803047530496 
'                                                   |
'               Whole numbers are converted to hex__|
' 
'            The hex representation of the fractional parts of the
'            CUBE ROOTS of 2 is:   428a2f98d728ae22 
' 
' -----------------------------------------------------------------------------
' Example:   First prime number = 2
'            SQUARE ROOT of 2 = 1.4142135623730950488016887242
'
'            0.4142135623730950488016887242 * 16 =  6.6274169979695207808270195872
'            0.6274169979695207808270195872 * 16 = 10.038671967512332493232313395
'            0.038671967512332493232313395  * 16 =  0.61875148019731989171701432
'            0.61875148019731989171701432   * 16 =  9.90002368315711826747222912
'            0.90002368315711826747222912   * 16 = 14.40037893051389227955566592
'            0.40037893051389227955566592   * 16 =  6.40606288822227647289065472
'            0.40606288822227647289065472   * 16 =  6.49700621155642356625047552
'            0.49700621155642356625047552   * 16 =  7.95209938490277706000760832
'            0.95209938490277706000760832   * 16 = 15.23359015844443296012173312
'            0.23359015844443296012173312   * 16 =  3.73744253511092736194772992
'            0.73744253511092736194772992   * 16 = 11.79908056177483779116367872
'            0.79908056177483779116367872   * 16 = 12.78528898839740465861885952
'            0.78528898839740465861885952   * 16 = 12.56462381435847453790175232
'            0.56462381435847453790175232   * 16 =  9.03398102973559260642803712
'            0.03398102973559260642803712   * 16 =  0.54369647576948170284859392
'            0.54369647576948170284859392   * 16 =  8.69914361231170724557750272
'                                                   |
'               Whole numbers are converted to hex__|
'
'            The hex representation of the fractional parts of the
'            SQUARE ROOTS of 2 is:   6a09e667f3bcc908
' 
' -----------------------------------------------------------------------------
'
' For computing SHA-512/t (ex: SHA-512/256).  I have submitted this simplified
' explanation to NIST as a possible enhancement to Section 5.3.6, page 16.
'
' 1.  Denote H01 to be the initial hash value of SHA-512 as specified above.
'
'     Original IV for SHA-512  (See Section 5.3.5)
'         H01(0) = "6A09E667F3BCC908"
'              - - -
'         H01(7) = "5BE0CD19137E2179"
'
' 2.  For i = 0 to 7
'         {
'         H02(i) = H01(i) Xor "A5A5A5A5A5A5A5A5"
'         }
'
'     Results after Xor
'         H02(0) = "CFAC43C256196CAD"
'         H02(1) = "1EC20B20216F029E"
'         H02(2) = "99CB56D75B315D8E"
'         H02(3) = "00EA509FFAB89354"
'         H02(4) = "F4ABF7DA08432774"
'         H02(5) = "3EA0CD298E9BC9BA"
'         H02(6) = "BA267C0E5EE418CE"
'         H02(7) = "FE4568BCB6DB84DC"
'
' 3.  Replace original IV for SHA-512 (in step 1.) with results in H02()
'
' 4.  Now hash data string "SHA-512/t" with modified SHA-512.  Where /t
'     represents the new truncation value.  (ex: "SHA-512/256")
'
'     Results of hashed data string "SHA-512/256"  (See Section 5.3.6.2)
'         H0(0) = "22312194FC2BF72C"
'             - - -
'         H0(7) = "0EB72DDC81C52CA2"
'
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 19-Jan-2007  Kenneth Ives  kenaso@tx.rr.com 
'              Wrote routine 
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' 02-May-2012  Kenneth Ives  kenaso@tx.rr.com 
'              Updated explanation of computing values for SHA-512/t
' *************************************************************************** 
Private Sub LoadConstArrays_64()

    ' Called by Initialize()

    Dim lngIndex As Long
    Dim strData  As String
    Dim avntData As Variant

    ReDim mastrHash(8)    ' Size arrays
    ReDim mastrConst(80)

    strData = vbNullString
    avntData = Empty   ' Start with empty variants

    Select Case mlngHashMethod

           Case eSHA384
                ' These words were obtained by taking the first sixty-four
                ' bits of the fractional parts of the SQUARE ROOTS of the
                ' 9th through 16th prime numbers.  As per FIPS 180-4, Section 5.3.4
                ' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
                mastrHash(0) = "cbbb9d5dc1059ed8"
                mastrHash(1) = "629a292a367cd507"
                mastrHash(2) = "9159015a3070dd17"
                mastrHash(3) = "152fecd8f70e5939"
                mastrHash(4) = "67332667ffc00b31"
                mastrHash(5) = "8eb44a8768581511"
                mastrHash(6) = "db0c2e0d64f98fa7"
                mastrHash(7) = "47b5481dbefa4fa4"

           Case eSHA512
                ' These words were obtained by taking the first sixty-four
                ' bits of the fractional parts of the SQUARE ROOTS of the
                ' first 8 prime numbers.  As per FIPS 180-4, Section 5.3.5
                ' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
                mastrHash(0) = "6a09e667f3bcc908"
                mastrHash(1) = "bb67ae8584caa73b"
                mastrHash(2) = "3c6ef372fe94f82b"
                mastrHash(3) = "a54ff53a5f1d36f1"
                mastrHash(4) = "510e527fade682d1"
                mastrHash(5) = "9b05688c2b3e6c1f"
                mastrHash(6) = "1f83d9abfb41bd6b"
                mastrHash(7) = "5be0cd19137e2179"

                ' For computing SHA-512/t
                ' See flowerbox above for more details
'                mastrHash(0) = "CFAC43C256196CAD"
'                mastrHash(1) = "1EC20B20216F029E"
'                mastrHash(2) = "99CB56D75B315D8E"
'                mastrHash(3) = "00EA509FFAB89354"
'                mastrHash(4) = "F4ABF7DA08432774"
'                mastrHash(5) = "3EA0CD298E9BC9BA"
'                mastrHash(6) = "BA267C0E5EE418CE"
'                mastrHash(7) = "FE4568BCB6DB84DC"

           Case eSHA512_224
                ' See explanation in flowerbox above as to how these words
                ' were obtained.  As per FIPS 180-4, Section 5.3.6.1
                ' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
                mastrHash(0) = "8C3D37C819544DA2"
                mastrHash(1) = "73E1996689DCD4D6"
                mastrHash(2) = "1DFAB7AE32FF9C82"
                mastrHash(3) = "679DD514582F9FCF"
                mastrHash(4) = "0F6D2B697BD44DA8"
                mastrHash(5) = "77E36F7304C48942"
                mastrHash(6) = "3F9D85A86A1D36C8"
                mastrHash(7) = "1112E6AD91D692A1"

           Case eSHA512_256
                ' See explanation in flowerbox above as to how these words
                ' were obtained.  As per FIPS 180-4, Section 5.3.6.2
                ' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
                mastrHash(0) = "22312194FC2BF72C"
                mastrHash(1) = "9F555FA3C84C64C2"
                mastrHash(2) = "2393B86B6F53B151"
                mastrHash(3) = "963877195940EABD"
                mastrHash(4) = "96283EE2A88EFFE3"
                mastrHash(5) = "BE5E1E2553863992"
                mastrHash(6) = "2B0199FC2C85B8AA"
                mastrHash(7) = "0EB72DDC81C52CA2"

           Case eSHA512_320
                ' See explanation in flowerbox above as to how these words
                ' were obtained.   See declarations section of this module
                ' for sample outputs.
                ' Created as per instructions FIPS 180-4, Section 5.3.6
                ' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
                mastrHash(0) = "5EBB5F84338831EA"
                mastrHash(1) = "40A1CD264905B9BA"
                mastrHash(2) = "996D9B4887A7CB2E"
                mastrHash(3) = "1FE0790FCEC6FB5A"
                mastrHash(4) = "DF14A9A3FAE0F4C6"
                mastrHash(5) = "206B146C8AEB54DC"
                mastrHash(6) = "B74FF50097BD4889"
                mastrHash(7) = "0C802F99E9E3209D"
    End Select

    ' Data for constants array. These are the same values
    ' used for SHA-384, SHA-512, SHA-512/224 and SHA-512/256.
    ' These words represent the first sixty-four bits of the
    ' fractional parts of the CUBE ROOTS of the first eighty
    ' prime numbers.
    strData = strData & "428a2f98d728ae22 7137449123ef65cd b5c0fbcfec4d3b2f e9b5dba58189dbbc "
    strData = strData & "3956c25bf348b538 59f111f1b605d019 923f82a4af194f9b ab1c5ed5da6d8118 "
    strData = strData & "d807aa98a3030242 12835b0145706fbe 243185be4ee4b28c 550c7dc3d5ffb4e2 "
    strData = strData & "72be5d74f27b896f 80deb1fe3b1696b1 9bdc06a725c71235 c19bf174cf692694 "
    strData = strData & "e49b69c19ef14ad2 efbe4786384f25e3 0fc19dc68b8cd5b5 240ca1cc77ac9c65 "
    strData = strData & "2de92c6f592b0275 4a7484aa6ea6e483 5cb0a9dcbd41fbd4 76f988da831153b5 "
    strData = strData & "983e5152ee66dfab a831c66d2db43210 b00327c898fb213f bf597fc7beef0ee4 "
    strData = strData & "c6e00bf33da88fc2 d5a79147930aa725 06ca6351e003826f 142929670a0e6e70 "
    strData = strData & "27b70a8546d22ffc 2e1b21385c26c926 4d2c6dfc5ac42aed 53380d139d95b3df "
    strData = strData & "650a73548baf63de 766a0abb3c77b2a8 81c2c92e47edaee6 92722c851482353b "
    strData = strData & "a2bfe8a14cf10364 a81a664bbc423001 c24b8b70d0f89791 c76c51a30654be30 "
    strData = strData & "d192e819d6ef5218 d69906245565a910 f40e35855771202a 106aa07032bbd1b8 "
    strData = strData & "19a4c116b8d2d0c8 1e376c085141ab53 2748774cdf8eeb99 34b0bcb5e19b48a8 "
    strData = strData & "391c0cb3c5c95a63 4ed8aa4ae3418acb 5b9cca4f7763e373 682e6ff3d6b2b8a3 "
    strData = strData & "748f82ee5defb2fc 78a5636f43172f60 84c87814a1f0ab72 8cc702081a6439ec "
    strData = strData & "90befffa23631e28 a4506cebde82bde9 bef9a3f7b2c67915 c67178f2e372532b "
    strData = strData & "ca273eceea26619c d186b8c721c0c207 eada7dd6cde0eb1e f57d4f7fee6ed178 "
    strData = strData & "06f067aa72176fba 0a637dc5a2c898a6 113f9804bef90dae 1b710b35131c471b "
    strData = strData & "28db77f523047d84 32caab7b40c72493 3c9ebe0a15c9bebc 431d67c49c100d4c "
    strData = strData & "4cc5d4becb3e42b6 597f299cfc657e2a 5fcb6fab3ad6faec 6c44198c4a475817 "

    avntData = Split(strData, Chr$(32))

    For lngIndex = 0 To 79
        mastrConst(lngIndex) = CStr(avntData(lngIndex))
    Next lngIndex

    strData = vbNullString
    avntData = Empty   ' Always empty variants when not needed

End Sub

Private Sub ResetVariables()

    Erase malngConst()
    Erase malngHash()
    Erase mastrConst()
    Erase mastrHash()

    ' Free class objects from memory
    Set mobjMath32 = Nothing
    Set mobjMath64 = Nothing

    mstrBitsInHex = vbNullString
    mcurAccrued = 0@
    mcurBitCount = 0@
    mcurCurrBitCnt = 0@

End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objSHA2 = new cSHA2

    ResetVariables

    ' Preset property values
    ReturnLowercase = False   ' Default return of hashed data
    HashRounds = 1            ' Set default number of passes
    HashMethod = eSHA256      ' Set default hash algorithm

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objSHA2 = Nothing

    ResetVariables

End Sub


