
Module modImport
    Private unicodeMapArray(256) As Long   'Used to reference IPA characters to Unicode characters
    Private keymanMapArray(256) As String 'Used for referencing IPA characters to keyman strings
    Private invalidSequences As New ArrayList

    Private Sub InitKeymanMapArray()
        keymanMapArray(97) = "a"    'Lowercase A
        keymanMapArray(140) = "a>"  'Turned A
        keymanMapArray(65) = "a="  'Cursive A
        keymanMapArray(129) = "o="  'Turned Cursive A
        keymanMapArray(81) = "a<"  'Ash Digraph
        keymanMapArray(98) = "b"  'Lowercase B
        keymanMapArray(186) = "b>"  'Hooktop B
        keymanMapArray(245) = "B"  'Small Capital B
        keymanMapArray(66) = "b="  'Beta
        keymanMapArray(99) = "c"  'Lowercase C
        keymanMapArray(141) = "o<"  'Open O
        keymanMapArray(254) = "c<"  'Curly-tail C
        keymanMapArray(67) = "c="  'C Cedilla
        keymanMapArray(100) = "d"  'Lowercase D
        keymanMapArray(235) = "d>"  'Hooktop D
        keymanMapArray(234) = "d<"  'Right-tail D
        keymanMapArray(68) = "d="  'Eth
        keymanMapArray(101) = "e"  'Lowercase E
        keymanMapArray(171) = "e="  'Schwa
        keymanMapArray(130) = "E"  'Reversed E
        keymanMapArray(69) = "e<"  'Epsilon
        keymanMapArray(206) = "e>"  'Reversed Epsilon
        keymanMapArray(207) = "O<"  'Closed Reversed Epsilon
        keymanMapArray(102) = "f"  'Lowercase F
        keymanMapArray(103) = "g"  'Lowercase G
        keymanMapArray(169) = "g>"  ' Hooktop G
        keymanMapArray(71) = "G"  ' Small Capital G
        keymanMapArray(253) = "G>"  ' Hooktop Small Capital G
        keymanMapArray(104) = "h"  ' Lowercase H
        keymanMapArray(72) = "h^"  ' Superscript H
        keymanMapArray(250) = "h<"  ' Hooktop H
        keymanMapArray(238) = "H>"  ' Hooked Heng
        keymanMapArray(240) = "h>"  ' Crossed H
        keymanMapArray(231) = "h="  ' Turned H
        keymanMapArray(75) = "H"  ' Small Capital H
        keymanMapArray(105) = "i"  ' Lowercase I
        keymanMapArray(34) = """"  ' Undotted I
        keymanMapArray(246) = "I="  ' Barred I
        keymanMapArray(174) = ""  ' Undotted Barred I "Empty"
        keymanMapArray(73) = "i="  ' Small Capital I
        keymanMapArray(106) = "j"  ' Lowercase J
        keymanMapArray(190) = ""  ' Dotless J "Empty"
        keymanMapArray(74) = "j^"  ' Superscript J "Conflict" j^ or j 
        keymanMapArray(239) = "j="  ' Barred Dotless J
        keymanMapArray(198) = "j<"  ' Curly-tail J
        keymanMapArray(215) = "j>"  ' Hooktop Barred Dotless J
        keymanMapArray(107) = "k"  ' Lowercase K
        keymanMapArray(108) = "l"  ' Lowercase L
        keymanMapArray(58) = "l^"  ' Superscript L
        keymanMapArray(241) = "l<"  ' Right-tail L
        keymanMapArray(194) = "l="  ' Belted L
        keymanMapArray(76) = "l>"  ' L-Yogh Digraph
        keymanMapArray(59) = "L"  ' Small Capital L
        keymanMapArray(109) = "m"  ' Lowercase M
        keymanMapArray(201) = "m^"  ' Superior m
        keymanMapArray(77) = "m>"  ' Left-tail M (at right)
        keymanMapArray(181) = "u="  ' Turned M
        keymanMapArray(229) = "w>"  ' Turned M, Right Leg
        keymanMapArray(110) = "n"  ' Lowercase N
        keymanMapArray(60) = "n^"  ' Superscript N
        keymanMapArray(78) = "n>"  ' Eng
        keymanMapArray(212) = "n>^"  ' Superscript Eng
        keymanMapArray(247) = "n<"  ' Right-tail N
        keymanMapArray(248) = "n="  ' Left-tail N (at left)
        keymanMapArray(203) = "n=^"  ' Superscript left-tail N (at left)
        keymanMapArray(178) = "N"  ' Small Capital N
        keymanMapArray(111) = "o"  ' Lowercase O
        keymanMapArray(79) = "o>"  ' Slashed O
        keymanMapArray(80) = "O"  ' Barred O
        keymanMapArray(184) = "f="  ' Phi
        keymanMapArray(84) = "t="  ' Theta
        keymanMapArray(191) = "E<"  ' O-E Digraph
        keymanMapArray(175) = "E>"  ' Small Capital O-E Digraph
        keymanMapArray(135) = "p="  ' Bull's Eye
        keymanMapArray(112) = "p"  ' Lowercase P
        keymanMapArray(113) = "q"  ' Lowercase Q
        keymanMapArray(114) = "r"  ' Lowercase R
        keymanMapArray(168) = "r="  ' Turned R
        keymanMapArray(228) = "L>"  ' Turned Long-leg R
        keymanMapArray(125) = "r<"  ' Right-tail R
        keymanMapArray(82) = "r>"  ' Fish-hook R
        keymanMapArray(211) = "R<"  ' Turned R, Right Tail
        keymanMapArray(123) = "R="  ' Small Capital R
        keymanMapArray(210) = "R>"  ' Inverted Small Capital R
        keymanMapArray(115) = "s"  ' Lowercase S
        keymanMapArray(167) = "s<"  ' Right-Tail S (at left)
        keymanMapArray(83) = "s="  ' Esh
        keymanMapArray(116) = "t"  ' Lowercase T
        keymanMapArray(255) = "t<"  ' Right-tail T
        keymanMapArray(117) = "u"  ' Lowercase U
        keymanMapArray(172) = "U"  ' Barred U
        keymanMapArray(86) = "v="  ' Cursive V
        keymanMapArray(85) = "u<"  ' Upsilon
        keymanMapArray(118) = "v"  ' Lowercase V
        keymanMapArray(195) = "u>"  ' Turned V
        keymanMapArray(196) = "g="  ' Gamma 
        keymanMapArray(236) = "g=^"  ' Superscript Gamma
        keymanMapArray(70) = "O>"  ' Ram's Horns "Empty"
        keymanMapArray(119) = "w"  ' Lowercase W
        keymanMapArray(87) = "w^"  ' Superscript W
        keymanMapArray(227) = "w="  ' Turned W
        keymanMapArray(120) = "x"  ' Lowercase X
        keymanMapArray(88) = "x="  ' Chi
        keymanMapArray(121) = "y"  ' Lowercase Y
        keymanMapArray(180) = "L<"  ' Turned Y
        keymanMapArray(89) = "y="  ' Small Capital Y
        keymanMapArray(122) = "z"  ' Lowercase Z
        keymanMapArray(252) = "z>"  ' Curly-tail Z
        keymanMapArray(189) = "z<"  ' Right-tail Z
        keymanMapArray(90) = "z="  ' Yogh
        keymanMapArray(63) = "?="  ' Glottal Stop
        keymanMapArray(251) = "Q"  ' Barred Glottal Stop
        keymanMapArray(192) = "?<"  ' Reversed Glottal Stop
        keymanMapArray(179) = "?<^"  ' Superscript Reversed Glottal Stop
        keymanMapArray(185) = "Q<"  ' Barred Reversed Glottal Stop
        keymanMapArray(151) = "!"  ' Exclamation Point
        keymanMapArray(44) = ","  ' Comma
        keymanMapArray(46) = "."  ' Period
        keymanMapArray(39) = "))"  ' Apostrophe
        keymanMapArray(91) = "("  ' Left Square Bracket
        keymanMapArray(93) = ")"  ' Right Square Bracket
        keymanMapArray(62) = "::"  ' Half-length Mark
        keymanMapArray(249) = ":"  ' Length Mark
        keymanMapArray(131) = "@&"  ' Top Tie Bar  ' Also could be #&, but WS 2.5
        keymanMapArray(237) = "#=="  ' Bottom Tie Bar
        keymanMapArray(150) = ".<"  ' Vertical Line
        keymanMapArray(124) = ")))"  ' Corner
        keymanMapArray(61) = "-"  ' Under-bar (o-width)
        keymanMapArray(173) = "-"  ' Under-bar (i-width)
        keymanMapArray(35) = "@2"  ' Macron (o-width)
        keymanMapArray(220) = "@2"  ' Macron (i-width)
        keymanMapArray(147) = "@22"  ' Macron (high o-width)
        keymanMapArray(148) = "@2"  ' Macron (high i-width)
        keymanMapArray(48) = "~~~"  ' Subscript Tilde (o-width)
        keymanMapArray(188) = "$$$"  ' Subscript Tilde (i-width)
        keymanMapArray(242) = "~~"  ' Superimposed Tilde
        keymanMapArray(41) = "~"  ' Superscript Tilde (owidth)
        keymanMapArray(226) = "~"  ' Superscript Tilde (i-width)
        keymanMapArray(64) = "@3"  ' Acute Accent (o-width)
        keymanMapArray(219) = "@3"  ' Acute Accent (i-width)
        keymanMapArray(143) = "@33"  ' Acute Accent (high owidth)
        keymanMapArray(144) = "@3"  ' Acute Accent (high iwidth)
        keymanMapArray(33) = "@4"  ' Double Acute Accent (owidth)
        keymanMapArray(218) = "@4"  ' Double Acute Accent (iwidth)
        keymanMapArray(136) = "@44"  ' Double Acute Accent (high o-width)
        keymanMapArray(137) = "@4"  ' Double Acute Accent
        keymanMapArray(36) = "@1"  ' Grave Accent (o-width)
        keymanMapArray(221) = "@1"  ' Grave Accent (i-width)
        keymanMapArray(152) = "@11"  ' Grave Accent (high owidth)
        keymanMapArray(153) = "@1"  ' Grave Accent (high iwidth)
        keymanMapArray(37) = "@0"  ' Double Grave Accent (owidth)
        keymanMapArray(222) = "@0"  ' Double Grave Accent (iwidth)
        keymanMapArray(157) = "@00"  ' Double Grave Accent (high o-width)
        keymanMapArray(158) = "@0"  ' Double Grave Accent (high i-width)
        keymanMapArray(94) = "@40"  ' Circumflex (o-width)
        keymanMapArray(223) = "@31"  ' Circumflex (i-width)
        keymanMapArray(233) = "@40"  ' Circumflex (high o-width)
        keymanMapArray(230) = "@40"  ' Circumflex (high i-width)
        keymanMapArray(38) = "@13"  ' Wedge (o-width)
        keymanMapArray(224) = "@13"  ' Wedge (i-width)
        keymanMapArray(244) = "@04"  ' Wedge (high o-width)
        keymanMapArray(243) = "@04"  ' Wedge (high i-width)
        keymanMapArray(45) = "%%%"  ' Subscript Umlaut (o-width) "Conflict" with Minute Space
        keymanMapArray(208) = "%%%"  ' Subscript Umlaut (i-width)
        keymanMapArray(95) = """"  ' Umlaut "Conflict" with Breve, Chose IPA93 keying. In mdb shown as """"
        keymanMapArray(164) = "%%"  ' Subscript Wedge
        keymanMapArray(40) = """"""""""    ' Breve (o-width)
        keymanMapArray(225) = """"""""""""""""""    ' Breve (i-width)
        keymanMapArray(57) = "^"  ' Subscript Arch (o-width)  ' Changed this from "$$", so that we can handle on- and off-glides in WS 2.5 format
        keymanMapArray(187) = "$$"  ' Subscript Arch (i-width) 
        keymanMapArray(209) = "{{{{"  ' Subscript Seagull
        keymanMapArray(56) = "%"  ' Under-ring (o-width)
        keymanMapArray(165) = "%"  ' Under Ring (i-width)
        keymanMapArray(42) = """"""""""""""""""    ' Over-ring (o-width) "Conflict" with Breve.
        keymanMapArray(161) = "@"  ' Over Ring (i-width)
        keymanMapArray(96) = "$"  ' Syllabicity Mark
        keymanMapArray(43) = "+"  ' Subscript Plus (o-width)
        keymanMapArray(177) = "+"  ' Subscript Plus (i-width)
        keymanMapArray(126) = """"""""""""""  ' Over-cross
        keymanMapArray(213) = "(("  ' Right Hook
        keymanMapArray(53) = "{"  ' Subscript Bridge
        keymanMapArray(176) = "{{"  ' Inverted Subscript Bridge
        keymanMapArray(54) = "{{{"  ' Subscript Square
        keymanMapArray(52) = "--"  ' Lowering Sign (o-width)
        keymanMapArray(162) = "--"  ' Lowering Sign (i-width)
        keymanMapArray(51) = "++"  ' Raising Sign (o-width)
        keymanMapArray(163) = "++"  ' Raising Sign (i-width)
        keymanMapArray(49) = "+++"  ' Advancing Sign (o-width)
        keymanMapArray(193) = "+++"  ' Advancing Sign (i-width)
        keymanMapArray(50) = "---"  ' Retracting Sign (o-width)
        keymanMapArray(170) = "---"  ' Retracting Sign (i-width)
        keymanMapArray(55) = "----"  ' Subscript Left Half-ring
        keymanMapArray(166) = "++++"    ' Subscript Right Half-ring
        keymanMapArray(138) = "#4"  ' Extra-high Tone Bar
        keymanMapArray(145) = "#3"  ' High Tone Bar
        keymanMapArray(149) = "#2"  ' Mid Tone Bar
        keymanMapArray(154) = "#1"  ' Low Tone Bar
        keymanMapArray(159) = "#0"  ' Extra-low Tone Bar
        keymanMapArray(232) = "#04"  ' Right Bar 15
        keymanMapArray(134) = "#40"  ' Right Bar 51
        keymanMapArray(216) = "#24"  ' Right Bar 35
        keymanMapArray(128) = "#02"  ' Right Bar 13
        keymanMapArray(133) = "#42"  ' Right Bar 53
        keymanMapArray(217) = "#20"  ' Right Bar 31
        keymanMapArray(155) = "#<"  ' Down Arrow
        keymanMapArray(139) = "#>"  ' Up Arrow
        keymanMapArray(205) = "#<<"  ' Downward Diagonal Arrow
        keymanMapArray(204) = "#>>"  ' Upward Diagonal Arrow
        keymanMapArray(199) = "}}"  ' Vertical Stroke (Inferior)
        keymanMapArray(200) = "}"  ' Vertical Stroke (Superior)
        keymanMapArray(142) = "!|"  ' Pipe
        keymanMapArray(156) = "!="  ' Double-barred Pipe
        keymanMapArray(132) = ".="  ' Double Vertical Line
        keymanMapArray(146) = "!>"  ' Double Pipe
        keymanMapArray(202) = "%%%"  ' Minute Space "Conflict" with Umlaut
        keymanMapArray(92) = "\"  ' Backward Slash
        keymanMapArray(47) = "/"  ' Forward Slash
        keymanMapArray(214) = "-"  ' Hyphen Dash
        keymanMapArray(32) = " "  ' Space

        keymanMapArray(13) = vbCrLf  ' Newline
        keymanMapArray(23) = "#"   ' Pound
        keymanMapArray(22) = "?"  ' Want to produce and match ?= for glottal stop, so this is a dummy character. 
        keymanMapArray(21) = "I"  ' Roman capital I on its own is valid in WS25 but is not representable in PalmSurv or WS 4. 

        ' These map to an unprinting character in order to allow the larger valued double-quote functions to map properly.
        keymanMapArray(7) = """"""   'two double quotes
        keymanMapArray(8) = """"""""  ' three double quotes"
        keymanMapArray(9) = """"""""""""  ' five double quotes"
        keymanMapArray(10) = """"""""""""""""  ' seven double quotes"
        keymanMapArray(11) = "R"  ' Capital R
        keymanMapArray(12) = "#="  ' #= -- not used for a character
    End Sub


    ' Takes a Keyman string and returns a string of the associated IPA characters.
    Public Function MapKeymanStrToIPAStr(ByVal KeymanStr As String) As String
        InitKeymanMapArray()

        Dim substr As String = ""
        Dim final As String = ""
        Dim testIPAchar As Integer = 0
        Dim validIPAchar As Integer = 0
        Dim initialIndex As Integer = 0
        Dim sequenceLength As Integer = 1
        Dim badParse As Boolean = False

        While (initialIndex < KeymanStr.Length)
            If (initialIndex + sequenceLength <= KeymanStr.Length) Then
                testIPAchar = MapKeymanStrToIPA(KeymanStr.Substring(initialIndex, sequenceLength))
            Else
                testIPAchar = 0
            End If
            If (testIPAchar <> 0) AndAlso (sequenceLength < (KeymanStr.Length - initialIndex + 1)) Then
                validIPAchar = testIPAchar
                sequenceLength += 1
            Else
                If (sequenceLength = 1) Then      ' This statement is for error checking if the string contains invalid keyman sequences
                    If Not invalidSequences.Contains(KeymanStr.Substring(initialIndex, sequenceLength)) Then
                        invalidSequences.Add(KeymanStr.Substring(initialIndex, sequenceLength))
                    End If
                    final += "[x" & returnXs(invalidSequences.IndexOf(KeymanStr.Substring(initialIndex, sequenceLength))) & "]"
                    initialIndex += 1
                    sequenceLength = 1
                Else
                    final += Chr(validIPAchar)
                    initialIndex += sequenceLength - 1
                    sequenceLength = 1
                End If
            End If
        End While
        Return final
    End Function
    Private Function returnXs(ByVal count As Integer) As String
        Dim str As String = ""
        For i As Integer = 0 To count - 1
            str += "x"
        Next
        Return str
    End Function
    Public Function MapKeymanStrToIPA(ByVal KeymanStr As String) As Integer
        Dim i As Integer = 0

        For i = 0 To 255  ' Searches through the entire keymanMapArray of strings to see if the sent
            If KeymanStr.Equals(keymanMapArray(i)) Then
                Return i
            End If
        Next

        Return 0
    End Function

    Private Sub InitUnicodeMapArray()
        Dim i As Integer
        For i = 0 To &HD - 1
            unicodeMapArray(i) = &H0     ' Setting every value in between values to zero to ensure no search will
            ' run against random values.
        Next
        unicodeMapArray(&HD) = &HD     ' Newline
        For i = &HE To &H15 - 1
            unicodeMapArray(i) = &H0     ' Setting every value in between values to zero to ensure no search will
            ' run against random values.
        Next

        unicodeMapArray(&H15) = &H49   ' Roman capital I on its own is valid in WS25 but is not representable in PalmSurv or WS 4. 
        unicodeMapArray(&H16) = &H3F  ' Want to produce and match ?= for glottal stop, so this is a dummy character. 
        unicodeMapArray(&H17) = &H17  ' Pound "#"
        unicodeMapArray(&H20) = &H20
        unicodeMapArray(&H21) = &H30B
        '	unicodeMapArray(&H22) = &H0069 'bactxt="byte-dia'uactxt="dotless  WEIRD XML FUNCTION
        unicodeMapArray(&H22) = &H131     ' Dotless "i"
        unicodeMapArray(&H23) = &H304
        unicodeMapArray(&H24) = &H300
        unicodeMapArray(&H25) = &H30F
        unicodeMapArray(&H26) = &H30C
        unicodeMapArray(&H27) = &H2BC
        unicodeMapArray(&H28) = &H306
        unicodeMapArray(&H29) = &H303
        unicodeMapArray(&H2A) = &H30A
        unicodeMapArray(&H2B) = &H31F
        unicodeMapArray(&H2C) = &H2C
        unicodeMapArray(&H2D) = &H324
        unicodeMapArray(&H2E) = &H2E
        unicodeMapArray(&H2F) = &H2F
        unicodeMapArray(&H30) = &H330
        unicodeMapArray(&H31) = &H318
        unicodeMapArray(&H32) = &H319
        unicodeMapArray(&H33) = &H31D
        unicodeMapArray(&H34) = &H31E
        unicodeMapArray(&H35) = &H32A
        unicodeMapArray(&H36) = &H33B
        unicodeMapArray(&H37) = &H31C
        unicodeMapArray(&H38) = &H325
        unicodeMapArray(&H39) = &H32F
        unicodeMapArray(&H3A) = &H2E1
        unicodeMapArray(&H3B) = &H29F
        unicodeMapArray(&H3C) = &H207F
        unicodeMapArray(&H3D) = &H320
        unicodeMapArray(&H3E) = &H2D1
        unicodeMapArray(&H3F) = &H294
        unicodeMapArray(&H40) = &H301
        unicodeMapArray(&H41) = &H251
        unicodeMapArray(&H42) = &H3B2
        unicodeMapArray(&H43) = &HE7   ' Combo, was combo of c and cedilla, now just one character
        unicodeMapArray(&H44) = &HF0
        unicodeMapArray(&H45) = &H25B     ' LATIN SMALL LETTER EPSILON
        unicodeMapArray(&H46) = &H264
        unicodeMapArray(&H47) = &H262
        unicodeMapArray(&H48) = &H2B0
        unicodeMapArray(&H49) = &H26A
        unicodeMapArray(&H4A) = &H2B2
        unicodeMapArray(&H4B) = &H29C
        unicodeMapArray(&H4C) = &H26E
        unicodeMapArray(&H4D) = &H271
        unicodeMapArray(&H4E) = &H14B
        unicodeMapArray(&H4F) = &HF8
        unicodeMapArray(&H50) = &H275
        unicodeMapArray(&H51) = &HE6
        unicodeMapArray(&H52) = &H27E
        unicodeMapArray(&H53) = &H283
        unicodeMapArray(&H54) = &H3B8
        unicodeMapArray(&H55) = &H28A
        unicodeMapArray(&H56) = &H28B
        unicodeMapArray(&H57) = &H2B7
        unicodeMapArray(&H58) = &H3C7
        unicodeMapArray(&H59) = &H28F
        unicodeMapArray(&H5A) = &H292
        unicodeMapArray(&H5B) = &H5B
        unicodeMapArray(&H5C) = &H5C
        unicodeMapArray(&H5D) = &H5D
        unicodeMapArray(&H5E) = &H302
        unicodeMapArray(&H5F) = &H308
        unicodeMapArray(&H60) = &H329
        unicodeMapArray(&H61) = &H61   ' lowercase a
        unicodeMapArray(&H62) = &H62
        unicodeMapArray(&H63) = &H63
        unicodeMapArray(&H64) = &H64
        unicodeMapArray(&H65) = &H65
        unicodeMapArray(&H66) = &H66
        unicodeMapArray(&H67) = &H67   ' lowercase g
        unicodeMapArray(&H68) = &H68
        unicodeMapArray(&H69) = &H69
        unicodeMapArray(&H6A) = &H6A
        unicodeMapArray(&H6B) = &H6B
        unicodeMapArray(&H6C) = &H6C
        unicodeMapArray(&H6D) = &H6D
        unicodeMapArray(&H6E) = &H6E
        unicodeMapArray(&H6F) = &H6F
        unicodeMapArray(&H70) = &H70
        unicodeMapArray(&H71) = &H71
        unicodeMapArray(&H72) = &H72
        unicodeMapArray(&H73) = &H73
        unicodeMapArray(&H74) = &H74
        unicodeMapArray(&H75) = &H75
        unicodeMapArray(&H76) = &H76
        unicodeMapArray(&H77) = &H77
        unicodeMapArray(&H78) = &H78
        unicodeMapArray(&H79) = &H79
        unicodeMapArray(&H7A) = &H7A   ' lowercase z
        unicodeMapArray(&H7B) = &H280
        unicodeMapArray(&H7C) = &H31A
        unicodeMapArray(&H7D) = &H27D
        unicodeMapArray(&H7E) = &H33D
        unicodeMapArray(&H7F) = &H7F   ' NOT RECOGNIZED IN SIL IPA FONT SET
        '	unicodeMapArray(&H80) = &H02E902E7  ' Problem Missing
        unicodeMapArray(&H80) = &H0    ' TEMPORARY FIX
        unicodeMapArray(&H81) = &H252
        unicodeMapArray(&H82) = &H258
        unicodeMapArray(&H83) = &H361
        unicodeMapArray(&H84) = &H2016
        '	unicodeMapArray(&H85) = &H02E502E7  ' Problem Missing
        unicodeMapArray(&H85) = &H0    ' TEMPORARY FIX
        '	unicodeMapArray(&H86) = &H02E502E9  ' Problem Missing
        unicodeMapArray(&H86) = &H0    ' TEMPORARY FIX
        unicodeMapArray(&H87) = &H298
        unicodeMapArray(&H88) = &H30B    'ubctxt="udia 
        unicodeMapArray(&H89) = &H30B    'ubctxt="i-udia  
        unicodeMapArray(&H8A) = &H2E5
        unicodeMapArray(&H8B) = &H2191
        unicodeMapArray(&H8C) = &H250
        unicodeMapArray(&H8D) = &H254
        unicodeMapArray(&H8E) = &H1C0
        unicodeMapArray(&H8F) = &H301    'ubctxt="udia 
        unicodeMapArray(&H90) = &H301    'ubctxt="i-udia  
        unicodeMapArray(&H91) = &H2E6
        unicodeMapArray(&H92) = &H1C1
        unicodeMapArray(&H93) = &H304    'ubctxt="udia 
        unicodeMapArray(&H94) = &H304    'ubctxt="i-udia 
        unicodeMapArray(&H95) = &H2E7
        unicodeMapArray(&H96) = &H7C
        unicodeMapArray(&H97) = &H1C3
        unicodeMapArray(&H98) = &H300    'ubctxt="udia  
        unicodeMapArray(&H99) = &H300    'ubctxt="i-udia  
        unicodeMapArray(&H9A) = &H2E8
        unicodeMapArray(&H9B) = &H2193
        unicodeMapArray(&H9C) = &H1C2
        unicodeMapArray(&H9D) = &H30F    'ubctxt="udia  
        unicodeMapArray(&H9E) = &H30F    'ubctxt="i-udia  
        unicodeMapArray(&H9F) = &H2E9
        unicodeMapArray(&HA1) = &H30A    'ubctxt="iwidth  
        unicodeMapArray(&HA2) = &H31E    'ubctxt="iwidth  
        unicodeMapArray(&HA3) = &H31D    'ubctxt="iwidth  
        unicodeMapArray(&HA4) = &H32C
        unicodeMapArray(&HA5) = &H325    'ubctxt="iwidth  
        unicodeMapArray(&HA6) = &H339
        unicodeMapArray(&HA7) = &H282
        unicodeMapArray(&HA8) = &H279
        unicodeMapArray(&HA9) = &H260
        unicodeMapArray(&HAA) = &H319    'ubctxt="iwidth  
        unicodeMapArray(&HAB) = &H259
        unicodeMapArray(&HAC) = &H289
        unicodeMapArray(&HAD) = &H320    'ubctxt="iwidth  
        unicodeMapArray(&HAE) = &H268    'uactxt="dotless  
        unicodeMapArray(&HAF) = &H276
        unicodeMapArray(&HB0) = &H33A
        unicodeMapArray(&HB1) = &H31F    'ubctxt="iwidth  
        unicodeMapArray(&HB2) = &H274
        unicodeMapArray(&HB3) = &H2E4
        unicodeMapArray(&HB4) = &H28E
        unicodeMapArray(&HB5) = &H26F
        unicodeMapArray(&HB8) = &H278
        unicodeMapArray(&HB9) = &H2A2
        unicodeMapArray(&HBA) = &H253
        unicodeMapArray(&HBB) = &H32F    'ubctxt="iwidth  
        unicodeMapArray(&HBC) = &H330    'ubctxt="iwidth  
        unicodeMapArray(&HBD) = &H290
        unicodeMapArray(&HBE) = &H6A    'uactxt="dotless  
        unicodeMapArray(&HBF) = &H153
        unicodeMapArray(&HC0) = &H295
        unicodeMapArray(&HC1) = &H318    'ubctxt="iwidth  
        unicodeMapArray(&HC2) = &H26C
        unicodeMapArray(&HC3) = &H28C
        unicodeMapArray(&HC4) = &H263
        unicodeMapArray(&HC6) = &H29D
        unicodeMapArray(&HC7) = &H2CC
        unicodeMapArray(&HC8) = &H2C8
        unicodeMapArray(&HC9) = &HF180   ' private use: Superscript m
        unicodeMapArray(&HCA) = &H200A
        unicodeMapArray(&HCB) = &HF181    ' private use: Superscript nya
        unicodeMapArray(&HCC) = &H2197
        unicodeMapArray(&HCD) = &H2198
        unicodeMapArray(&HCE) = &H25C
        unicodeMapArray(&HCF) = &H25E
        unicodeMapArray(&HD0) = &H324     'ubctxt="iwidth  
        unicodeMapArray(&HD1) = &H33C
        unicodeMapArray(&HD2) = &H281
        unicodeMapArray(&HD3) = &H27B
        unicodeMapArray(&HD4) = &HF182    ' private use: Superscript eng
        unicodeMapArray(&HD5) = &H2DE
        unicodeMapArray(&HD6) = &H2D
        unicodeMapArray(&HD7) = &H284
        'unicodeMapArray(&HD8) = &H02E7 02E5 ' Problem missing Right bar #35 
        unicodeMapArray(&HD8) = &H0    'TEMPORARY FIX
        '	unicodeMapArray(&HD9) = &H02E7 02E9 ' Problem missing Right bar #31
        unicodeMapArray(&HD9) = &H0    'TEMPORARY FIX
        unicodeMapArray(&HDA) = &H30B    'ubctxt="iwidth  
        unicodeMapArray(&HDB) = &H301    'ubctxt="iwidth  
        unicodeMapArray(&HDC) = &H304    'ubctxt="iwidth  
        unicodeMapArray(&HDD) = &H300    'ubctxt="iwidth  
        unicodeMapArray(&HDE) = &H30F    'ubctxt="iwidth  
        unicodeMapArray(&HDF) = &H302    'ubctxt="iwidth  
        unicodeMapArray(&HE0) = &H30C    'ubctxt="iwidth  
        unicodeMapArray(&HE1) = &H306    'ubctxt="iwidth  
        unicodeMapArray(&HE2) = &H303    'ubctxt="iwidth  
        unicodeMapArray(&HE3) = &H28D
        unicodeMapArray(&HE4) = &H27A
        unicodeMapArray(&HE5) = &H270
        unicodeMapArray(&HE6) = &H302    'ubctxt="i-udia  
        unicodeMapArray(&HE7) = &H265
        '	unicodeMapArray(&HE8) = &H02E9 02E5  ' Problem Missing
        unicodeMapArray(&HE8) = &H0    ' TEMPORARY FIX
        unicodeMapArray(&HE9) = &H302     'ubctxt="udia  
        unicodeMapArray(&HEA) = &H256
        unicodeMapArray(&HEB) = &H257
        unicodeMapArray(&HEC) = &H2E0
        unicodeMapArray(&HED) = &H203F    ' Under-bar only in SIL Unicode Beta (private character)
        unicodeMapArray(&HEE) = &H267
        unicodeMapArray(&HEF) = &H25F
        unicodeMapArray(&HF0) = &H127
        unicodeMapArray(&HF1) = &H26D
        ' also gave this map, unsure about it unicodeMapArray(&H6C F2) = &H026B
        unicodeMapArray(&HF2) = &H334
        unicodeMapArray(&HF3) = &H30C    'ubctxt="i-udia  
        unicodeMapArray(&HF4) = &H30C    'ubctxt="udia  
        unicodeMapArray(&HF5) = &H299
        unicodeMapArray(&HF6) = &H268
        unicodeMapArray(&HF7) = &H273
        unicodeMapArray(&HF8) = &H272
        unicodeMapArray(&HF9) = &H2D0
        unicodeMapArray(&HFA) = &H266
        unicodeMapArray(&HFB) = &H2A1
        unicodeMapArray(&HFC) = &H291
        unicodeMapArray(&HFD) = &H29B
        unicodeMapArray(&HFE) = &H255
        unicodeMapArray(&HFF) = &H288
    End Sub

    Public Function MapIPAtoUnicode(ByVal IPAchar As Integer) As Long

        If (IPAchar < 256) AndAlso (IPAchar > 0) Then
            Return unicodeMapArray(IPAchar)
        End If

        Return &HFFFF

    End Function
    Public Function MapIPAstrtoUnicodeStr(ByVal IPAstr As String) As String
        Dim Unicodestr As String = ""
        Dim i, c As Integer

        InitUnicodeMapArray()

        For i = 0 To IPAstr.Length - 1
            c = Asc(IPAstr.Substring(i, 1))

            Select Case c
                Case &HE8
                    Unicodestr += ChrW(&H2E9)
                    Unicodestr += ChrW(&H2E5)
                Case &H80
                    Unicodestr += ChrW(&H2E9)
                    Unicodestr += ChrW(&H2E7)
                Case &HD9
                    Unicodestr += ChrW(&H2E7)
                    Unicodestr += ChrW(&H2E9)
                Case &HD8
                    Unicodestr += ChrW(&H2E7)
                    Unicodestr += ChrW(&H2E5)
                Case &H86
                    Unicodestr += ChrW(&H2E5)
                    Unicodestr += ChrW(&H2E9)
                Case &H85
                    Unicodestr += ChrW(&H2E5)
                    Unicodestr += ChrW(&H2E7)
                Case Else
                    If MapIPAtoUnicode(c) = &H0 Then
                        Unicodestr += "[x]"        'Error Checking, if the character is 'invalid, the program returns a [x]
                    Else
                        Unicodestr += Convert.ToChar(MapIPAtoUnicode(c))
                    End If
            End Select
        Next

        Return Unicodestr
    End Function
    Public Function MapKeymanStrToUnicodeStr(ByVal keystr As String) As String

        Dim unistr As String
        InitUnicodeMapArray()
        InitKeymanMapArray()

        unistr = MapKeymanStrToIPAStr(keystr)    'maps the keyman string to an IPA string
        unistr = MapIPAstrtoUnicodeStr(unistr)    'maps the IPA string created in the previous function to its unicode equivalent
        unistr = replaceXs(unistr) 'Replaces multiple x's with a single x with a number denoting how many there were, for error changing

        Return unistr

    End Function

    Private Function replaceXs(ByVal ipaStr As String) As String
        Try
            If Not ipaStr Is Nothing Then
                'Debug.WriteLine("ipaStr = " & ipaStr)
                Dim i, a As Integer
                Dim returnStr As String = ""
                For i = 0 To ipaStr.Length - 1
                    'Debug.WriteLine("i = " & i & " returnStr = " & returnStr & " ipaStr.chars(i) = " & ipaStr.Chars(i))
                    If ipaStr.Chars(i) = "[" Then
                        a = i
                        Try 'AJW***
                            While ipaStr.Chars(i) <> "]" And i < ipaStr.Length - 1
                                i += 1
                            End While
                        Catch
                            Debug.Print("i = " & i.ToString)
                            Debug.Print(ipaStr)
                        End Try
                        returnStr += "[x" & i - a - 2 & "]"
                    Else
                        returnStr += ipaStr.Chars(i)
                    End If
                Next
                Return returnStr
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.StackTrace)
        End Try
        Return ""
    End Function
End Module
