Attribute VB_Name = "modTestData"
' ***************************************************************************
' Some of the test data is the same found in FIPS 180-3 publication.
' The hashed results are accurate.
'
' REFERENCE:
'
' NIST (National Institute of Standards and Technology) Publications
' (FIPS, Special Publications)
' http://csrc.nist.gov/publications/PubsFIPS.html
'
' FIPS 180-2 (Federal Information Processing Standards Publication)
' dated 1-Aug-2002, with Change Notice 1, dated 25-Feb-2004
' http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf
'
' FIPS 180-3 (Federal Information Processing Standards Publication)
' dated Oct-2008 (supercedes FIPS 180-2)
' http://csrc.nist.gov/publications/fips/fips180-3/fips180-3_final.pdf
'
' FIPS 180-4 (Federal Information Processing Standards Publication)
' dated Mar-2012 (supercedes FIPS 180-3)
' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
'
' Examples of hash outputs:
' http://csrc.nist.gov/groups/ST/toolkit/examples.html
'
' Additional SHA2 information and test vectors by Aaron Gifford
'     SHA2 Information - http://www.adg.us/computers/sha.html
'     Test vectors    - http://www.adg.us/computers/sha2-1.0.zip
'
' NIST Test vectors are at http://csrc.nist.gov/cryptval/shs.htm
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com 
'              Added reference to SHA-512/224 and SHA-512/256 as per
'              FIPS 180-4 dtd March-2012 (Supercedes FIPS 180-3)
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Public Const TEST_FILE1 As String = "Vector004.dat"  ' Excert from A. Lincoln speech
  Public Const TEST_FILE2 As String = "Vector013.dat"  ' Binary test file
  Public Const TEST_FILE3 As String = "Vector017.dat"  ' test for off-by-one
  Public Const TEST_FILE4 As String = "OneMil_a.dat"   ' One million letter 'a'
  Public Const TEST_FILE5 As String = "OneMil_0.dat"   ' One million zeroes
  
  Private Const MODULE_NAME As String = "modTestData"
  
' ***************************************************************************
' Determine the algorithm used and return the pertinent information
' ***************************************************************************
Public Sub SelectResults(ByVal lngAlgorithm As Long, _
                         ByVal lngExpectedResults As Long, _
                         ByRef strTestData As String, _
                         ByRef strDataLength As String, _
                         ByRef strOutput As String)
    
    Const ROUTINE_NAME As String = "SelectResults"
    
    Select Case lngExpectedResults
           Case 0
                strTestData = "abc"
                strDataLength = "3"
           Case 1
                strTestData = "The quick brown fox jumps over the lazy dog"
                strDataLength = "43"
           Case 2
                strTestData = "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
                strDataLength = "56"
           Case 3
                strTestData = "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmn" & _
                              "oijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu"
                strDataLength = "112"
           Case 4
                strTestData = "One thousand letter 'A'"
                strDataLength = "1000"
           Case 5
                strTestData = "Excert from President Abraham Lincoln in a file named " & TEST_FILE1
                strDataLength = "1515"
           Case 6
                strTestData = "A binary file that is one byte short of 17 times size " & _
                              "of the SHA-384 and SHA-512 block lengths named " & TEST_FILE2
                strDataLength = "2175"
           Case 7
                strTestData = "The length of this binary data set is designed to test for " & _
                              "off-by-one in a file named " & TEST_FILE3
                strDataLength = "12271"
           Case 8
                strTestData = "1,000,000 letter 'a'"
                strDataLength = "1000000"
           Case 9
                strTestData = "1,000,000 binary zeroes"
                strDataLength = "1000000"
           Case Else
                InfoMsg "Cannot identify test case." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Sub
    End Select
    
    Select Case lngAlgorithm
           Case 0   ' SHA-1
                Select Case lngExpectedResults
                       Case 0: strOutput = "A9993E364706816ABA3E25717850C26C9CD0D89D"
                       Case 1: strOutput = "2FD4E1C67A2D28FCED849EE1BB76E7391B93EB12"
                       Case 2: strOutput = "84983E441C3BD26EBAAE4AA1F95129E5E54670F1"
                       Case 3: strOutput = "A49B2446A02C645BF419F995B67091253A04A259"
                       Case 4: strOutput = "3AE3644D6777A1F56A1DEFEABC74AF9C4B313E49"
                       Case 5: strOutput = "3728B3FD827FE2BFD0900E0586A03FFD3394E647"
                       Case 6: strOutput = "A04FDD79DDD249C71F687674329E026C57BCC378"
                       Case 7: strOutput = "3CC047623D26128CC61DFDEF8AF7CA473814063A"
                       Case 8: strOutput = "34AA973CD4C4DAA4F61EEB2BDBAD27316534016F"
                       Case 9: strOutput = "BEF3595266A65A2FF36B700A75E8ED95C68210B6"
                End Select
           Case 1   ' SHA-224
                Select Case lngExpectedResults
                       Case 0: strOutput = "23097D223405D8228642A477BDA255B32AADBCE4BDA0B3F7E36C9DA7"
                       Case 1: strOutput = "730E109BD7A8A32B1CB9D9A09AA2325D2430587DDBC0C38BAD911525"
                       Case 2: strOutput = "75388B16512776CC5DBA5DA1FD890150B0C6455CB4F58B1952522525"
                       Case 3: strOutput = "C97CA9A559850CE97A04A96DEF6D99A9E0E0E2AB14E6B8DF265FC0B3"
                       Case 4: strOutput = "A8D0C66B5C6FDFD836EB3C6D04D32DFE66C3B1F168B488BF4C9C66CE"
                       Case 5: strOutput = "62A41AB0961BCDD22DB70B896DB3955C1D04096AF6DE47F5AAAD1226"
                       Case 6: strOutput = "91E452CFC8F22F9C69E637EC9DCF80D5798607A52234686FCF8880AD"
                       Case 7: strOutput = "A1B0964A6D8188EB2980E126FEFC70EB79D0745A91CC2F629AF34ECE"
                       Case 8: strOutput = "20794655980C91D8BBB4C1EA97618A4BF03F42581948B2EE4EE7AD67"
                       Case 9: strOutput = "3A5D74B68F14F3A4B2BE9289B8D370672D0B3D2F53BC303C59032DF3"
                End Select
           Case 2   ' SHA-256
                Select Case lngExpectedResults
                       Case 0: strOutput = "BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD"
                       Case 1: strOutput = "D7A8FBB307D7809469CA9ABCB0082E4F8D5651E46D3CDB762D02D0BF37C9E592"
                       Case 2: strOutput = "248D6A61D20638B8E5C026930C3E6039A33CE45964FF2167F6ECEDD419DB06C1"
                       Case 3: strOutput = "CF5B16A778AF8380036CE59E7B0492370B249B11E8F07A51AFAC45037AFEE9D1"
                       Case 4: strOutput = "C2E686823489CED2017F6059B8B239318B6364F6DCD835D0A519105A1EADD6E4"
                       Case 5: strOutput = "4D25FCCF8752CE470A58CD21D90939B7EB25F3FA418DD2DA4C38288EA561E600"
                       Case 6: strOutput = "8FF59C6D33C5A991088BC44DD38F037EB5AD5630C91071A221AD6943E872AC29"
                       Case 7: strOutput = "88EE6ADA861083094F4C64B373657E178D88EF0A4674FCE6E4E1D84E3B176AFB"
                       Case 8: strOutput = "CDC76E5C9914FB9281A1C7E284D73E67F1809A48A497200E046D39CCC7112CD0"
                       Case 9: strOutput = "D29751F2649B32FF572B5E0A9F541EA660A50F94FF0BEEDFB0B692B924CC8025"
                End Select
           Case 3   ' SHA-384
                Select Case lngExpectedResults
                       Case 0: strOutput = "CB00753F45A35E8BB5A03D699AC65007272C32AB0EDED1631A8B605A43FF5BED8086072BA1E7CC2358BAECA134C825A7"
                       Case 1: strOutput = "CA737F1014A48F4C0B6DD43CB177B0AFD9E5169367544C494011E3317DBF9A509CB1E5DC1E85A941BBEE3D7F2AFBC9B1"
                       Case 2: strOutput = "3391FDDDFC8DC7393707A65B1B4709397CF8B1D162AF05ABFE8F450DE5F36BC6B0455A8520BC4E6F5FE95B1FE3C8452B"
                       Case 3: strOutput = "09330C33F71147E83D192FC782CD1B4753111B173B3B05D22FA08086E3B0F712FCC7C71A557E2DB966C3E9FA91746039"
                       Case 4: strOutput = "7DF01148677B7F18617EEE3A23104F0EED6BB8C90A6046F715C9445FF43C30D69E9E7082DE39C3452FD1D3AFD9BA0689"
                       Case 5: strOutput = "69CC75B95280BDD9E154E743903E37B1205AA382E92E051B1F48A6DB9D0203F8A17C1762D46887037275606932D3381E"
                       Case 6: strOutput = "92DCA5655229B3C34796A227FF1809E273499ADC2830149481224E0F54FF4483BD49834D4865E508EF53D4CD22B703CE"
                       Case 7: strOutput = "78CC6402A29EB984B8F8F888AB0102CABE7C06F0B9570E3D8D744C969DB14397F58ECD14E70F324BF12D8DD4CD1AD3B2"
                       Case 8: strOutput = "9D0E1809716474CB086E834E310A4A1CED149E9C00F248527972CEC5704C2A5B07B8B3DC38ECC4EBAE97DDD87F3D8985"
                       Case 9: strOutput = "8A1979F9049B3FFF15EA3A43A4CF84C634FD14ACAD1C333FECB72C588B68868B66A994386DC0CD1687B9EE2E34983B81"
                End Select
           Case 4   ' SHA-512
                Select Case lngExpectedResults
                       Case 0: strOutput = "DDAF35A193617ABACC417349AE20413112E6FA4E89A97EA20A9EEEE64B55D39A2192992A274FC1A836BA3C23A3FEEBBD454D4423643CE80E2A9AC94FA54CA49F"
                       Case 1: strOutput = "07E547D9586F6A73F73FBAC0435ED76951218FB7D0C8D788A309D785436BBB642E93A252A954F23912547D1E8A3B5ED6E1BFD7097821233FA0538F3DB854FEE6"
                       Case 2: strOutput = "204A8FC6DDA82F0A0CED7BEB8E08A41657C16EF468B228A8279BE331A703C33596FD15C13B1B07F9AA1D3BEA57789CA031AD85C7A71DD70354EC631238CA3445"
                       Case 3: strOutput = "8E959B75DAE313DA8CF4F72814FC143F8F7779C6EB9F7FA17299AEADB6889018501D289E4900F7E4331B99DEC4B5433AC7D329EEB6DD26545E96E55B874BE909"
                       Case 4: strOutput = "329C52AC62D1FE731151F2B895A00475445EF74F50B979C6F7BB7CAE349328C1D4CB4F7261A0AB43F936A24B000651D4A824FCDD577F211AEF8F806B16AFE8AF"
                       Case 5: strOutput = "23450737795D2F6A13AA61ADCCA0DF5EEF6DF8D8DB2B42CD2CA8F783734217A73E9CABC3C9B8A8602F8AEAEB34562B6B1286846060F9809B90286B3555751F09"
                       Case 6: strOutput = "0E928DB6207282BFB498EE871202F2337F4074F3A1F5055A24F08E912AC118F8101832CDB9C2F702976E629183DB9BACFDD7B086C800687C3599F15DE7F7B9DD"
                       Case 7: strOutput = "211BEC83FBCA249C53668802B857A9889428DC5120F34B3EAC1603F13D1B47965C387B39EF6AF15B3A44C5E7B6BBB6C1096A677DC98FC8F472737540A332F378"
                       Case 8: strOutput = "E718483D0CE769644E2E42C7BC15B4638E1F98B13B2044285632A803AFA973EBDE0FF244877EA60A4CB0432CE577C31BEB009C5C2C49AA2E4EADB217AD8CC09B"
                       Case 9: strOutput = "CE044BC9FD43269D5BBC946CBEBC3BB711341115CC4ABDF2EDBC3FF2C57AD4B15DEB699BDA257FEA5AEF9C6E55FCF4CF9DC25A8C3CE25F2EFE90908379BFF7ED"
                End Select
           Case 5   ' SHA-512/224
                Select Case lngExpectedResults
                       Case 0: strOutput = "4634270F707B6A54DAAE7530460842E20E37ED265CEEE9A43E8924AA"
                       Case 1: strOutput = "944CD2847FB54558D4775DB0485A50003111C8E5DAA63FE722C6AA37"
                       Case 2: strOutput = "E5302D6D54BB242275D1E7622D68DF6EB02DEDD13F564C13DBDA2174"
                       Case 3: strOutput = "23FEC5BB94D60B23308192640B0C453335D664734FE40E7268674AF9"
                       Case 4: strOutput = "3000C31A7AB8E9C760257073C4D3BE370FAB6D1D28EB027C6D874F29"
                       Case 5: strOutput = "E267F1D3F97B2BB2495720F0A6207552012DB33AE43502D0A221DFD7"
                       Case 6: strOutput = "6D85BBD10A382C256ADD592836D42F223DBB503FDC1623FEDA9F3D10"
                       Case 7: strOutput = "75452CE87F5F10840B8EE92419AD9B64B05A81881496A5890F8D0466"
                       Case 8: strOutput = "37AB331D76F0D36DE422BD0EDEB22A28ACCD487B7A8453AE965DD287"
                       Case 9: strOutput = "7576F5B118A2DDC31AB05C641F04027FED5F1CBB65894D17EC664466"
                End Select
           Case 6   ' SHA-512/256
                Select Case lngExpectedResults
                       Case 0: strOutput = "53048E2681941EF99B2E29B76B4C7DABE4C2D0C634FC6D46E0E2F13107E7AF23"
                       Case 1: strOutput = "DD9D67B371519C339ED8DBD25AF90E976A1EEEFD4AD3D889005E532FC5BEF04D"
                       Case 2: strOutput = "BDE8E1F9F19BB9FD3406C90EC6BC47BD36D8ADA9F11880DBC8A22A7078B6A461"
                       Case 3: strOutput = "3928E184FB8690F840DA3988121D31BE65CB9D3EF83EE6146FEAC861E19B563A"
                       Case 4: strOutput = "6AD592C8991FA0FC0FC78B6C2E73F3B55DB74AFEB1027A5AEACB787FB531E64A"
                       Case 5: strOutput = "E7A5FFFD2F07EBEA8E868519AFFC397D8592B7567ACD2368E299C9A2F6055DDF"
                       Case 6: strOutput = "799127629DA5399B236FFEE8AA54ECD6EFDC8B78CB630DF5B409A4C163C231FE"
                       Case 7: strOutput = "EC40BA525E09F73EBC6BE17A6CD4277DC108CC4C8B4436F990D394EE4484ACA5"
                       Case 8: strOutput = "9A59A052930187A97038CAE692F30708AA6491923EF5194394DC68D56C74FB21"
                       Case 9: strOutput = "8B620FF17FD0414C7C3567704F9E275A5C37801720C75DC05CF81558E4A0F965"
                End Select
           Case 7   ' SHA-512/320
                Select Case lngExpectedResults
                       Case 0: strOutput = "0F7567ED5B9C77C089BE0D0F74EBD5DDA19FCC52DB4018D03036E6AE7FBFCEDA2567E1FEE10E37DC"
                       Case 1: strOutput = "7A3A7017B742C90AAE8DBF1296112E66610B33DE57C081AF472857BB1E7F5582D69378B4EEF60DA8"
                       Case 2: strOutput = "30F2DF31D6AA8ACDA222464BF24312A2C0CB3280D6D0B277ADFEF33522A59CF49561E522BE5D63F1"
                       Case 3: strOutput = "56862213733C0C0D89A27E5A9CC266F0E147DEAAF48038BF9C5484DFDB68DB93602D0A4589463C75"
                       Case 4: strOutput = "B2B24C3AD8DC11D82A979E8532D3DCA964C19F8941B42AEBAD57DC43B5C33201F1434B4C346016EB"
                       Case 5: strOutput = "93E1C523995B1150D2E35CF83792A0CB1658AA17453CC623ECC7FA932A7FA1F1C431E496DB5EE6A3"
                       Case 6: strOutput = "9843167660749E74C5460B22A6E4BA3F65D1B93B0C160A9B5F4BBC1FE018173FECDB10B10B87EB49"
                       Case 7: strOutput = "08E9952B925636EC2C6B136E9BB58596E4ADCF92928709C33D0604D90540B9D00E8A20A2EB3CA247"
                       Case 8: strOutput = "5A19AABC267F08806F4FDD070423F9BCFF606481D7E0834A7D2C289DBEA840FF836258EC71E23035"
                       Case 9: strOutput = "68DBEAF6CC9337DF8FACA863DABEAC6388CC60C267B30B41A4019398429D04558C377C761965CDC0"
                End Select
           Case Else
                InfoMsg "Unknown hash algorithm selected." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End Select
    
End Sub
