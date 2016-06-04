; Auto-Parser for XML / HTML by SKAN
/*
StrX( H, BS,BO,BT, ES,EO,ET, NextOffset )

    Parameters

    1 ) H = HayStack. The "Source Text"

    2 ) BS = BeginStr. Pass a String that will result at the left extreme of Resultant String
    3 ) BO = BeginOffset.
        Number of Characters to omit from the left extreme of "Source Text" while searching for BeginStr
        Pass a 0 to search in reverse ( from right-to-left ) in "Source Text"
        If you intend to call StrX() from a Loop, pass the same variable used as 8th Parameter, which will simplify the parsing process.
    4 ) BT = BeginTrim.
        Number of characters to trim on the left extreme of Resultant String
        Pass the String length of BeginStr if you want to omit it from Resultant String
        Pass a Negative value if you want to expand the left extreme of Resultant String
    5 ) ES = EndStr. Pass a String that will result at the right extreme of Resultant String
    6 ) EO = EndOffset.
        Can be only True or False.
        If False, EndStr will be searched from the end of Source Text.
        If True, search will be conducted from the search result offset of BeginStr or from offset 1 whichever is applicable.
    7 ) ET = EndTrim.
        Number of characters to trim on the right extreme of Resultant String
        Pass the String length of EndStr if you want to omit it from Resultant String
        Pass a Negative value if you want to expand the right extreme of Resultant String
    8 ) NextOffset : A name of ByRef Variable that will be updated by StrX() with the current offset, You may pass the same variable as Parameter 3, to simplify data parsing in a loop
 */
 
StrX( H,  BS="",BO=0,BT=1,   ES="",EO=0,ET=1,  ByRef N="" ) { ;    | by Skan | 19-Nov-2009
Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
 <0)?(1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
 +(0-ET)):(X+P)):(X)))-P) ; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
}