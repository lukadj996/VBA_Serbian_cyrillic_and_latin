Attribute VB_Name = "prevod_cyr_lat"
Function lat2cyr(in_string As String)

Dim A As String * 1
Dim B As String * 1
Dim i As Integer

latchars = ChrW(97) + ChrW(98) + ChrW(99) + ChrW(269) + ChrW(263) + ChrW(100) + ChrW(273) + ChrW(101) + ChrW(102) + ChrW(103) + ChrW(104) + ChrW(105) + ChrW(106) + ChrW(107) + ChrW(108) + ChrW(109) + ChrW(110) + ChrW(111) + ChrW(112) + ChrW(114) + ChrW(115) + ChrW(353) + ChrW(116) + ChrW(117) + ChrW(118) + ChrW(122) + ChrW(382) + ChrW(65) + ChrW(66) + ChrW(67) + ChrW(268) + ChrW(262) + ChrW(68) + ChrW(272) + ChrW(69) + ChrW(70) + ChrW(71) + ChrW(72) + ChrW(73) + ChrW(74) + ChrW(75) + ChrW(76) + ChrW(77) + ChrW(78) + ChrW(79) + ChrW(80) + ChrW(82) + ChrW(83) + ChrW(352) + ChrW(84) + ChrW(85) + ChrW(86) + ChrW(90) + ChrW(381)
cyrchars = ChrW(1072) + ChrW(1073) + ChrW(1094) + ChrW(1095) + ChrW(1115) + ChrW(1076) + ChrW(1106) + ChrW(1077) + ChrW(1092) + ChrW(1075) + ChrW(1093) + ChrW(1080) + ChrW(1112) + ChrW(1082) + ChrW(1083) + ChrW(1084) + ChrW(1085) + ChrW(1086) + ChrW(1087) + ChrW(1088) + ChrW(1089) + ChrW(1096) + ChrW(1090) + ChrW(1091) + ChrW(1074) + ChrW(1079) + ChrW(1078) + ChrW(1040) + ChrW(1041) + ChrW(1062) + ChrW(1063) + ChrW(1035) + ChrW(1044) + ChrW(1026) + ChrW(1045) + ChrW(1060) + ChrW(1043) + ChrW(1061) + ChrW(1048) + ChrW(1032) + ChrW(1050) + ChrW(1051) + ChrW(1052) + ChrW(1053) + ChrW(1054) + ChrW(1055) + ChrW(1056) + ChrW(1057) + ChrW(1064) + ChrW(1058) + ChrW(1059) + ChrW(1042) + ChrW(1047) + ChrW(1046)

'6 specijalnih karaktera
in_string = Replace(in_string, ChrW(100) + ChrW(382), ChrW(1119)) 'dž
in_string = Replace(in_string, ChrW(108) + ChrW(106), ChrW(1113)) 'lj
in_string = Replace(in_string, ChrW(110) + ChrW(106), ChrW(1114)) 'nj
in_string = Replace(in_string, ChrW(68) + ChrW(382), ChrW(1039)) 'Dž
in_string = Replace(in_string, ChrW(76) + ChrW(106), ChrW(1033)) 'Lj
in_string = Replace(in_string, ChrW(78) + ChrW(106), ChrW(1034)) 'Nj

For i = 1 To Len(latchars) '54 slova plus 6 specijalnih = 60 = 30 malih + 30 velikih
A = Mid(latchars, i, 1)
B = Mid(cyrchars, i, 1)
in_string = Replace(in_string, A, B)
Next
lat2cyr = in_string
End Function
Function cyr2lat(in_string As String)

Dim A As String * 1
Dim B As String * 1
Dim i As Integer

latchars = ChrW(97) + ChrW(98) + ChrW(99) + ChrW(269) + ChrW(263) + ChrW(100) + ChrW(273) + ChrW(101) + ChrW(102) + ChrW(103) + ChrW(104) + ChrW(105) + ChrW(106) + ChrW(107) + ChrW(108) + ChrW(109) + ChrW(110) + ChrW(111) + ChrW(112) + ChrW(114) + ChrW(115) + ChrW(353) + ChrW(116) + ChrW(117) + ChrW(118) + ChrW(122) + ChrW(382) + ChrW(65) + ChrW(66) + ChrW(67) + ChrW(268) + ChrW(262) + ChrW(68) + ChrW(272) + ChrW(69) + ChrW(70) + ChrW(71) + ChrW(72) + ChrW(73) + ChrW(74) + ChrW(75) + ChrW(76) + ChrW(77) + ChrW(78) + ChrW(79) + ChrW(80) + ChrW(82) + ChrW(83) + ChrW(352) + ChrW(84) + ChrW(85) + ChrW(86) + ChrW(90) + ChrW(381)
cyrchars = ChrW(1072) + ChrW(1073) + ChrW(1094) + ChrW(1095) + ChrW(1115) + ChrW(1076) + ChrW(1106) + ChrW(1077) + ChrW(1092) + ChrW(1075) + ChrW(1093) + ChrW(1080) + ChrW(1112) + ChrW(1082) + ChrW(1083) + ChrW(1084) + ChrW(1085) + ChrW(1086) + ChrW(1087) + ChrW(1088) + ChrW(1089) + ChrW(1096) + ChrW(1090) + ChrW(1091) + ChrW(1074) + ChrW(1079) + ChrW(1078) + ChrW(1040) + ChrW(1041) + ChrW(1062) + ChrW(1063) + ChrW(1035) + ChrW(1044) + ChrW(1026) + ChrW(1045) + ChrW(1060) + ChrW(1043) + ChrW(1061) + ChrW(1048) + ChrW(1032) + ChrW(1050) + ChrW(1051) + ChrW(1052) + ChrW(1053) + ChrW(1054) + ChrW(1055) + ChrW(1056) + ChrW(1057) + ChrW(1064) + ChrW(1058) + ChrW(1059) + ChrW(1042) + ChrW(1047) + ChrW(1046)

'6 specijalnih karaktera
in_string = Replace(in_string, ChrW(1119), ChrW(100) + ChrW(382)) 'dž
in_string = Replace(in_string, ChrW(1113), ChrW(108) + ChrW(106)) 'lj
in_string = Replace(in_string, ChrW(1114), ChrW(110) + ChrW(106)) 'nj
in_string = Replace(in_string, ChrW(1039), ChrW(68) + ChrW(382)) 'Dž
in_string = Replace(in_string, ChrW(1033), ChrW(76) + ChrW(106)) 'Lj
in_string = Replace(in_string, ChrW(1034), ChrW(78) + ChrW(106)) 'Nj

For i = 1 To Len(latchars) '54 slova plus 6 specijalnih = 60 = 30 malih + 30 velikih
A = Mid(latchars, i, 1)
B = Mid(cyrchars, i, 1)
in_string = Replace(in_string, B, A)
Next
cyr2lat = in_string
End Function
