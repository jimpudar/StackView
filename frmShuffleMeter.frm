VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmShuffleMeter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joyal ShuffleMeter"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "frmShuffleMeter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   9075
   Begin TabDlg.SSTab ShuffleMeterForm 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   1710
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Spread"
      TabPicture(0)   =   "frmShuffleMeter.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Shape1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Shape1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Shape1(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Shape1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Shape1(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Shape1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Shape1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Shape1(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Shape1(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Shape1(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Shape1(10)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Shape1(11)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Shape1(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Shape1(13)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Shape1(14)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Shape1(15)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Shape1(16)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Shape1(17)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Shape1(18)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Shape1(19)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Shape1(20)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Shape1(21)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Shape1(22)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Shape1(23)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Shape1(24)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Shape1(25)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Shape1(26)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Shape1(27)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Shape1(28)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Shape1(29)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Shape1(30)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Shape1(31)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Shape1(32)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Shape1(33)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Shape1(34)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Shape1(35)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Shape1(36)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Shape1(37)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Shape1(38)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Shape1(39)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Shape1(40)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Shape1(41)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Shape1(42)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Shape1(43)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Shape1(44)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Shape1(45)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Shape1(46)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Shape1(47)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Shape1(48)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Shape1(49)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Shape1(50)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Shape1(51)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Shape2(0)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Shape2(1)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Shape2(2)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Shape2(3)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Shape2(4)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Shape2(5)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Shape2(6)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Shape2(7)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Shape2(8)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Shape2(9)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "Shape2(10)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "Shape2(11)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Shape2(12)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "Shape2(13)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "Shape2(14)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "Shape2(15)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "Shape2(16)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "Shape2(17)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "Shape2(18)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "Shape2(19)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "Shape2(20)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "Shape2(21)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "Shape2(22)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "Shape2(23)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "Shape2(24)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "Shape2(25)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "Shape2(26)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "Shape2(27)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Shape2(28)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "Shape2(29)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "Shape2(30)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "Shape2(31)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "Shape2(32)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "Shape2(33)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "Shape2(34)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "Shape2(35)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "Shape2(36)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "Shape2(37)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "Shape2(38)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "Shape2(39)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "Shape2(40)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "Shape2(41)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "Shape2(42)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "Shape2(43)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "Shape2(44)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "Shape2(45)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "Shape2(46)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "Shape2(47)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "Shape2(48)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "Shape2(49)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "Shape2(50)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "Shape2(51)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "Shape3(0)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "Shape3(1)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "Shape3(2)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "Shape3(3)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "Shape3(4)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "Shape3(5)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "Shape3(6)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "Shape3(7)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "Shape3(8)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "Shape3(9)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "Shape3(10)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "Shape3(11)"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "Shape3(12)"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "Shape3(13)"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "Shape3(14)"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "Shape3(15)"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "Shape3(16)"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "Shape3(17)"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "Shape3(18)"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "Shape3(19)"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "Shape3(20)"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "Shape3(21)"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "Shape3(22)"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "Shape3(23)"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "Shape3(24)"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "Shape3(25)"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "Shape3(26)"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "Shape3(27)"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "Shape3(28)"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "Shape3(29)"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "Shape3(30)"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "Shape3(31)"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "Shape3(32)"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "Shape3(33)"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "Shape3(34)"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "Shape3(35)"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "Shape3(36)"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "Shape3(37)"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "Shape3(38)"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "Shape3(39)"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "Shape3(40)"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).Control(151)=   "Shape3(41)"
      Tab(0).Control(151).Enabled=   0   'False
      Tab(0).Control(152)=   "Shape3(42)"
      Tab(0).Control(152).Enabled=   0   'False
      Tab(0).Control(153)=   "Shape3(43)"
      Tab(0).Control(153).Enabled=   0   'False
      Tab(0).Control(154)=   "Shape3(44)"
      Tab(0).Control(154).Enabled=   0   'False
      Tab(0).Control(155)=   "Shape3(45)"
      Tab(0).Control(155).Enabled=   0   'False
      Tab(0).Control(156)=   "Shape3(46)"
      Tab(0).Control(156).Enabled=   0   'False
      Tab(0).Control(157)=   "Shape3(47)"
      Tab(0).Control(157).Enabled=   0   'False
      Tab(0).Control(158)=   "Shape3(48)"
      Tab(0).Control(158).Enabled=   0   'False
      Tab(0).Control(159)=   "Shape3(49)"
      Tab(0).Control(159).Enabled=   0   'False
      Tab(0).Control(160)=   "Shape3(50)"
      Tab(0).Control(160).Enabled=   0   'False
      Tab(0).Control(161)=   "Shape3(51)"
      Tab(0).Control(161).Enabled=   0   'False
      Tab(0).Control(162)=   "Label1"
      Tab(0).Control(162).Enabled=   0   'False
      Tab(0).Control(163)=   "Label2"
      Tab(0).Control(163).Enabled=   0   'False
      Tab(0).Control(164)=   "Label3"
      Tab(0).Control(164).Enabled=   0   'False
      Tab(0).Control(165)=   "Label4(0)"
      Tab(0).Control(165).Enabled=   0   'False
      Tab(0).Control(166)=   "Label7"
      Tab(0).Control(166).Enabled=   0   'False
      Tab(0).Control(167)=   "Label8"
      Tab(0).Control(167).Enabled=   0   'False
      Tab(0).Control(168)=   "Label4(1)"
      Tab(0).Control(168).Enabled=   0   'False
      Tab(0).Control(169)=   "Label4(2)"
      Tab(0).Control(169).Enabled=   0   'False
      Tab(0).Control(170)=   "Label4(3)"
      Tab(0).Control(170).Enabled=   0   'False
      Tab(0).ControlCount=   171
      TabCaption(1)   =   "Permutation"
      TabPicture(1)   =   "frmShuffleMeter.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5(58)"
      Tab(1).Control(1)=   "Label5(57)"
      Tab(1).Control(2)=   "Label5(56)"
      Tab(1).Control(3)=   "Label5(55)"
      Tab(1).Control(4)=   "Label5(54)"
      Tab(1).Control(5)=   "Label5(53)"
      Tab(1).Control(6)=   "Label5(52)"
      Tab(1).Control(7)=   "Label5(51)"
      Tab(1).Control(8)=   "Label5(50)"
      Tab(1).Control(9)=   "Label5(49)"
      Tab(1).Control(10)=   "Label5(48)"
      Tab(1).Control(11)=   "Label5(47)"
      Tab(1).Control(12)=   "Label5(46)"
      Tab(1).Control(13)=   "Label5(45)"
      Tab(1).Control(14)=   "Label5(44)"
      Tab(1).Control(15)=   "Label5(43)"
      Tab(1).Control(16)=   "Label5(42)"
      Tab(1).Control(17)=   "Label5(41)"
      Tab(1).Control(18)=   "Label5(40)"
      Tab(1).Control(19)=   "Label5(39)"
      Tab(1).Control(20)=   "Label5(38)"
      Tab(1).Control(21)=   "Label5(37)"
      Tab(1).Control(22)=   "Label5(36)"
      Tab(1).Control(23)=   "Label5(35)"
      Tab(1).Control(24)=   "Label5(34)"
      Tab(1).Control(25)=   "Label5(33)"
      Tab(1).Control(26)=   "Label5(32)"
      Tab(1).Control(27)=   "Label5(31)"
      Tab(1).Control(28)=   "Label5(30)"
      Tab(1).Control(29)=   "Label5(29)"
      Tab(1).Control(30)=   "Label5(28)"
      Tab(1).Control(31)=   "Label5(27)"
      Tab(1).Control(32)=   "Label5(26)"
      Tab(1).Control(33)=   "Label5(25)"
      Tab(1).Control(34)=   "Label5(24)"
      Tab(1).Control(35)=   "Label5(23)"
      Tab(1).Control(36)=   "Label5(22)"
      Tab(1).Control(37)=   "Label5(21)"
      Tab(1).Control(38)=   "Label5(20)"
      Tab(1).Control(39)=   "Label5(19)"
      Tab(1).Control(40)=   "Line12"
      Tab(1).Control(41)=   "Line11"
      Tab(1).Control(42)=   "Line10"
      Tab(1).Control(43)=   "Label5(18)"
      Tab(1).Control(44)=   "Label5(17)"
      Tab(1).Control(45)=   "Label5(16)"
      Tab(1).Control(46)=   "Label5(15)"
      Tab(1).Control(47)=   "Label5(14)"
      Tab(1).Control(48)=   "Label5(13)"
      Tab(1).Control(49)=   "Label5(12)"
      Tab(1).Control(50)=   "Label5(11)"
      Tab(1).Control(51)=   "Label5(10)"
      Tab(1).Control(52)=   "Label5(9)"
      Tab(1).Control(53)=   "Label5(8)"
      Tab(1).Control(54)=   "Label5(7)"
      Tab(1).Control(55)=   "Label5(6)"
      Tab(1).Control(56)=   "Label5(5)"
      Tab(1).Control(57)=   "Label5(4)"
      Tab(1).Control(58)=   "Label5(3)"
      Tab(1).Control(59)=   "Label5(2)"
      Tab(1).Control(60)=   "Label5(1)"
      Tab(1).Control(61)=   "Label5(0)"
      Tab(1).Control(62)=   "Line9(0)"
      Tab(1).ControlCount=   63
      TabCaption(2)   =   "Distribution"
      TabPicture(2)   =   "frmShuffleMeter.frx":1D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5(59)"
      Tab(2).Control(1)=   "Label5(60)"
      Tab(2).Control(2)=   "Line9(1)"
      Tab(2).Control(3)=   "Label5(61)"
      Tab(2).Control(4)=   "Label5(62)"
      Tab(2).Control(5)=   "Label5(63)"
      Tab(2).Control(6)=   "Line9(2)"
      Tab(2).Control(7)=   "Label5(64)"
      Tab(2).Control(8)=   "Label5(65)"
      Tab(2).Control(9)=   "Label5(66)"
      Tab(2).Control(10)=   "Label5(67)"
      Tab(2).Control(11)=   "Label5(68)"
      Tab(2).Control(12)=   "Label5(69)"
      Tab(2).Control(13)=   "Label5(70)"
      Tab(2).Control(14)=   "Label5(71)"
      Tab(2).Control(15)=   "Label5(72)"
      Tab(2).Control(16)=   "Label5(73)"
      Tab(2).Control(17)=   "Label5(74)"
      Tab(2).Control(18)=   "Label5(75)"
      Tab(2).Control(19)=   "Label5(76)"
      Tab(2).Control(20)=   "Label5(77)"
      Tab(2).Control(21)=   "Label5(78)"
      Tab(2).Control(22)=   "Label5(79)"
      Tab(2).Control(23)=   "Label5(80)"
      Tab(2).Control(24)=   "Label5(81)"
      Tab(2).Control(25)=   "Label5(82)"
      Tab(2).Control(26)=   "Label5(83)"
      Tab(2).Control(27)=   "Label5(84)"
      Tab(2).Control(28)=   "Label5(85)"
      Tab(2).Control(29)=   "Label5(86)"
      Tab(2).Control(30)=   "Label5(87)"
      Tab(2).Control(31)=   "Line9(3)"
      Tab(2).Control(32)=   "Label5(88)"
      Tab(2).Control(33)=   "Label5(89)"
      Tab(2).Control(34)=   "Label5(90)"
      Tab(2).Control(35)=   "Label5(91)"
      Tab(2).Control(36)=   "Line9(4)"
      Tab(2).Control(37)=   "Label5(92)"
      Tab(2).Control(38)=   "Label5(93)"
      Tab(2).Control(39)=   "Label5(94)"
      Tab(2).Control(40)=   "Label5(95)"
      Tab(2).Control(41)=   "Label5(96)"
      Tab(2).Control(42)=   "Label5(97)"
      Tab(2).Control(43)=   "Label5(98)"
      Tab(2).Control(44)=   "Label5(99)"
      Tab(2).Control(45)=   "Label5(100)"
      Tab(2).Control(46)=   "Label5(101)"
      Tab(2).Control(47)=   "Label5(102)"
      Tab(2).Control(48)=   "Label5(103)"
      Tab(2).Control(49)=   "Label5(104)"
      Tab(2).Control(50)=   "Label5(105)"
      Tab(2).Control(51)=   "Label5(106)"
      Tab(2).Control(52)=   "Label5(107)"
      Tab(2).Control(53)=   "Label5(108)"
      Tab(2).Control(54)=   "Label5(109)"
      Tab(2).Control(55)=   "Label5(110)"
      Tab(2).Control(56)=   "Label5(111)"
      Tab(2).Control(57)=   "Label5(112)"
      Tab(2).Control(58)=   "Label5(113)"
      Tab(2).Control(59)=   "Label5(114)"
      Tab(2).Control(60)=   "Label5(115)"
      Tab(2).Control(61)=   "Label5(116)"
      Tab(2).Control(62)=   "Label5(117)"
      Tab(2).Control(63)=   "Label5(118)"
      Tab(2).Control(64)=   "Label5(119)"
      Tab(2).Control(65)=   "Line9(5)"
      Tab(2).Control(66)=   "Label5(120)"
      Tab(2).Control(67)=   "Label5(121)"
      Tab(2).Control(68)=   "Label5(122)"
      Tab(2).Control(69)=   "Label5(123)"
      Tab(2).Control(70)=   "Label5(124)"
      Tab(2).Control(71)=   "Label5(125)"
      Tab(2).Control(72)=   "Label5(126)"
      Tab(2).Control(73)=   "Label5(127)"
      Tab(2).Control(74)=   "Label5(128)"
      Tab(2).Control(75)=   "Label5(129)"
      Tab(2).Control(76)=   "Label5(130)"
      Tab(2).Control(77)=   "Label5(131)"
      Tab(2).Control(78)=   "Label5(132)"
      Tab(2).Control(79)=   "Label5(133)"
      Tab(2).Control(80)=   "Label5(134)"
      Tab(2).Control(81)=   "Label5(135)"
      Tab(2).Control(82)=   "Label5(136)"
      Tab(2).Control(83)=   "Label5(137)"
      Tab(2).Control(84)=   "Label5(138)"
      Tab(2).Control(85)=   "Label5(139)"
      Tab(2).Control(86)=   "Label5(140)"
      Tab(2).Control(87)=   "Label5(141)"
      Tab(2).Control(88)=   "Label5(142)"
      Tab(2).Control(89)=   "Label5(143)"
      Tab(2).Control(90)=   "Line9(6)"
      Tab(2).Control(91)=   "Label5(144)"
      Tab(2).Control(92)=   "Label5(145)"
      Tab(2).Control(93)=   "Label5(146)"
      Tab(2).Control(94)=   "Label5(147)"
      Tab(2).Control(95)=   "Line9(7)"
      Tab(2).Control(96)=   "Label5(148)"
      Tab(2).Control(97)=   "Label5(149)"
      Tab(2).Control(98)=   "Label5(150)"
      Tab(2).Control(99)=   "Label5(151)"
      Tab(2).Control(100)=   "Label5(152)"
      Tab(2).Control(101)=   "Label5(153)"
      Tab(2).Control(102)=   "Label5(154)"
      Tab(2).Control(103)=   "Label5(155)"
      Tab(2).Control(104)=   "Label5(156)"
      Tab(2).Control(105)=   "Label5(157)"
      Tab(2).Control(106)=   "Label5(158)"
      Tab(2).Control(107)=   "Label5(159)"
      Tab(2).Control(108)=   "Label5(160)"
      Tab(2).Control(109)=   "Label5(161)"
      Tab(2).Control(110)=   "Label5(162)"
      Tab(2).Control(111)=   "Label5(163)"
      Tab(2).Control(112)=   "Label5(164)"
      Tab(2).Control(113)=   "Label5(165)"
      Tab(2).Control(114)=   "Label5(166)"
      Tab(2).Control(115)=   "Label5(167)"
      Tab(2).Control(116)=   "Label5(168)"
      Tab(2).Control(117)=   "Label5(169)"
      Tab(2).Control(118)=   "Label5(170)"
      Tab(2).Control(119)=   "Label5(171)"
      Tab(2).Control(120)=   "Label5(172)"
      Tab(2).Control(121)=   "Label5(173)"
      Tab(2).ControlCount=   122
      TabCaption(3)   =   "Group"
      TabPicture(3)   =   "frmShuffleMeter.frx":1D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label5(174)"
      Tab(3).Control(1)=   "Label5(175)"
      Tab(3).Control(2)=   "Label5(176)"
      Tab(3).Control(3)=   "Label5(177)"
      Tab(3).Control(4)=   "Label5(178)"
      Tab(3).Control(5)=   "Label5(179)"
      Tab(3).Control(6)=   "Label5(180)"
      Tab(3).Control(7)=   "Label5(181)"
      Tab(3).Control(8)=   "Label5(182)"
      Tab(3).Control(9)=   "Label5(183)"
      Tab(3).Control(10)=   "Label5(184)"
      Tab(3).Control(11)=   "Label5(185)"
      Tab(3).Control(12)=   "Label5(186)"
      Tab(3).Control(13)=   "Label5(187)"
      Tab(3).Control(14)=   "Line9(8)"
      Tab(3).Control(15)=   "Label5(188)"
      Tab(3).Control(16)=   "Label5(189)"
      Tab(3).Control(17)=   "Label5(190)"
      Tab(3).Control(18)=   "Label5(191)"
      Tab(3).Control(19)=   "Line9(9)"
      Tab(3).Control(20)=   "Label5(192)"
      Tab(3).Control(21)=   "Label5(193)"
      Tab(3).Control(22)=   "Label5(194)"
      Tab(3).Control(23)=   "Label5(195)"
      Tab(3).Control(24)=   "Label5(196)"
      Tab(3).Control(25)=   "Label5(197)"
      Tab(3).Control(26)=   "Label5(198)"
      Tab(3).Control(27)=   "Label5(199)"
      Tab(3).Control(28)=   "Label5(200)"
      Tab(3).Control(29)=   "Label5(201)"
      Tab(3).Control(30)=   "Label5(202)"
      Tab(3).Control(31)=   "Label5(203)"
      Tab(3).Control(32)=   "Label5(204)"
      Tab(3).Control(33)=   "Label5(205)"
      Tab(3).Control(34)=   "Label5(206)"
      Tab(3).Control(35)=   "Label5(207)"
      Tab(3).Control(36)=   "Label5(208)"
      Tab(3).Control(37)=   "Label5(209)"
      Tab(3).Control(38)=   "Label5(210)"
      Tab(3).Control(39)=   "Label5(211)"
      Tab(3).Control(40)=   "Line9(10)"
      Tab(3).Control(41)=   "Label5(212)"
      Tab(3).Control(42)=   "Label5(213)"
      Tab(3).Control(43)=   "Label5(214)"
      Tab(3).Control(44)=   "Label5(215)"
      Tab(3).Control(45)=   "Label5(216)"
      Tab(3).Control(46)=   "Label5(217)"
      Tab(3).Control(47)=   "Label5(218)"
      Tab(3).Control(48)=   "Label5(219)"
      Tab(3).Control(49)=   "Label5(220)"
      Tab(3).Control(50)=   "Label5(221)"
      Tab(3).Control(51)=   "Label5(222)"
      Tab(3).Control(52)=   "Label5(223)"
      Tab(3).Control(53)=   "Label5(224)"
      Tab(3).Control(54)=   "Label5(225)"
      Tab(3).Control(55)=   "Label5(226)"
      Tab(3).Control(56)=   "Label5(227)"
      Tab(3).Control(57)=   "Label5(228)"
      Tab(3).Control(58)=   "Label5(229)"
      Tab(3).Control(59)=   "Label5(230)"
      Tab(3).Control(60)=   "Label5(231)"
      Tab(3).Control(61)=   "Label5(232)"
      Tab(3).Control(62)=   "Label5(233)"
      Tab(3).Control(63)=   "Label5(234)"
      Tab(3).Control(64)=   "Label5(235)"
      Tab(3).Control(65)=   "Label5(236)"
      Tab(3).Control(66)=   "Label5(237)"
      Tab(3).Control(67)=   "Label5(238)"
      Tab(3).Control(68)=   "Label5(239)"
      Tab(3).Control(69)=   "Line9(11)"
      Tab(3).Control(70)=   "Label5(240)"
      Tab(3).Control(71)=   "Label5(241)"
      Tab(3).Control(72)=   "Label5(242)"
      Tab(3).Control(73)=   "Line9(12)"
      Tab(3).Control(74)=   "Label5(243)"
      Tab(3).Control(75)=   "Label5(244)"
      Tab(3).ControlCount=   76
      TabCaption(4)   =   "Break"
      TabPicture(4)   =   "frmShuffleMeter.frx":1D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label9"
      Tab(4).Control(1)=   "Line7"
      Tab(4).Control(2)=   "Label10"
      Tab(4).Control(3)=   "Label11"
      Tab(4).Control(4)=   "Label12"
      Tab(4).Control(5)=   "Label13"
      Tab(4).Control(6)=   "Label14"
      Tab(4).Control(7)=   "Label15"
      Tab(4).Control(8)=   "Label16"
      Tab(4).Control(9)=   "Label17"
      Tab(4).Control(10)=   "Label18"
      Tab(4).Control(11)=   "Line8"
      Tab(4).Control(12)=   "Label19"
      Tab(4).Control(13)=   "SMB_Actual_C"
      Tab(4).Control(14)=   "SMB_Actual_S"
      Tab(4).Control(15)=   "SMB_Actual_V"
      Tab(4).Control(16)=   "SMB_Diff_C"
      Tab(4).Control(17)=   "SMB_Diff_S"
      Tab(4).Control(18)=   "SMB_Diff_V"
      Tab(4).Control(19)=   "SMB_TotalDiff"
      Tab(4).ControlCount=   20
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "xx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3060
         TabIndex        =   274
         Top             =   1170
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "xx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3060
         TabIndex        =   273
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "xx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3060
         TabIndex        =   272
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Analysis of Groups in the Stack"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   244
         Left            =   -74760
         TabIndex        =   271
         Top             =   645
         Width           =   6840
      End
      Begin VB.Label Label5 
         Caption         =   "Colors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   243
         Left            =   -74760
         TabIndex        =   270
         Top             =   1110
         Width           =   1485
      End
      Begin VB.Line Line9 
         Index           =   12
         X1              =   -74760
         X2              =   -66870
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   242
         Left            =   -73290
         TabIndex        =   269
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   241
         Left            =   -72540
         TabIndex        =   268
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   240
         Left            =   -71730
         TabIndex        =   267
         Top             =   1110
         Width           =   435
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   -74760
         X2              =   -71265
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   239
         Left            =   -74760
         TabIndex        =   266
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   238
         Left            =   -74760
         TabIndex        =   265
         Top             =   1830
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   237
         Left            =   -74760
         TabIndex        =   264
         Top             =   2130
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   236
         Left            =   -74760
         TabIndex        =   263
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   235
         Left            =   -74760
         TabIndex        =   262
         Top             =   2730
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   234
         Left            =   -74760
         TabIndex        =   261
         Top             =   3030
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   233
         Left            =   -73290
         TabIndex        =   260
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   232
         Left            =   -72510
         TabIndex        =   259
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   231
         Left            =   -71730
         TabIndex        =   258
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   230
         Left            =   -73290
         TabIndex        =   257
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   229
         Left            =   -72510
         TabIndex        =   256
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   228
         Left            =   -71730
         TabIndex        =   255
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   227
         Left            =   -73290
         TabIndex        =   254
         Top             =   2130
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   226
         Left            =   -72510
         TabIndex        =   253
         Top             =   2130
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   225
         Left            =   -71730
         TabIndex        =   252
         Top             =   2130
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   224
         Left            =   -73290
         TabIndex        =   251
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   223
         Left            =   -72510
         TabIndex        =   250
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   222
         Left            =   -71730
         TabIndex        =   249
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   221
         Left            =   -73290
         TabIndex        =   248
         Top             =   2730
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   220
         Left            =   -72510
         TabIndex        =   247
         Top             =   2730
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   219
         Left            =   -71730
         TabIndex        =   246
         Top             =   2730
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   218
         Left            =   -73290
         TabIndex        =   245
         Top             =   3030
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   217
         Left            =   -72510
         TabIndex        =   244
         Top             =   3030
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   216
         Left            =   -71730
         TabIndex        =   243
         Top             =   3030
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Suits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   215
         Left            =   -70470
         TabIndex        =   242
         Top             =   1110
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   214
         Left            =   -69000
         TabIndex        =   241
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   213
         Left            =   -68250
         TabIndex        =   240
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   212
         Left            =   -67440
         TabIndex        =   239
         Top             =   1110
         Width           =   435
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   -70470
         X2              =   -66975
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   211
         Left            =   -70470
         TabIndex        =   238
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   210
         Left            =   -70470
         TabIndex        =   237
         Top             =   1830
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   209
         Left            =   -70470
         TabIndex        =   236
         Top             =   2130
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   208
         Left            =   -70470
         TabIndex        =   235
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "7.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   207
         Left            =   -69000
         TabIndex        =   234
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   206
         Left            =   -68220
         TabIndex        =   233
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   205
         Left            =   -67440
         TabIndex        =   232
         Top             =   1530
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   204
         Left            =   -69000
         TabIndex        =   231
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   203
         Left            =   -68220
         TabIndex        =   230
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   202
         Left            =   -67440
         TabIndex        =   229
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   201
         Left            =   -69000
         TabIndex        =   228
         Top             =   2130
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   200
         Left            =   -68220
         TabIndex        =   227
         Top             =   2130
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   199
         Left            =   -67440
         TabIndex        =   226
         Top             =   2130
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   198
         Left            =   -69000
         TabIndex        =   225
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   197
         Left            =   -68220
         TabIndex        =   224
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   196
         Left            =   -67440
         TabIndex        =   223
         Top             =   2430
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   195
         Left            =   -74760
         TabIndex        =   222
         Top             =   3330
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   194
         Left            =   -73290
         TabIndex        =   221
         Top             =   3330
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   193
         Left            =   -72510
         TabIndex        =   220
         Top             =   3330
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   192
         Left            =   -71730
         TabIndex        =   219
         Top             =   3330
         Width           =   435
      End
      Begin VB.Line Line9 
         Index           =   9
         X1              =   -74760
         X2              =   -66870
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Label Label5 
         Caption         =   "Values"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   191
         Left            =   -70470
         TabIndex        =   218
         Top             =   3480
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   190
         Left            =   -69000
         TabIndex        =   217
         Top             =   3480
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   189
         Left            =   -68250
         TabIndex        =   216
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   188
         Left            =   -67440
         TabIndex        =   215
         Top             =   3480
         Width           =   435
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   -70470
         X2              =   -66975
         Y1              =   3795
         Y2              =   3795
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   187
         Left            =   -70470
         TabIndex        =   214
         Top             =   3900
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   186
         Left            =   -70470
         TabIndex        =   213
         Top             =   4200
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Group of 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   185
         Left            =   -70470
         TabIndex        =   212
         Top             =   4500
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "2.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   184
         Left            =   -69000
         TabIndex        =   211
         Top             =   3900
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   183
         Left            =   -68220
         TabIndex        =   210
         Top             =   3900
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   182
         Left            =   -67440
         TabIndex        =   209
         Top             =   3900
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   181
         Left            =   -69000
         TabIndex        =   208
         Top             =   4200
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   180
         Left            =   -68220
         TabIndex        =   207
         Top             =   4200
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   179
         Left            =   -67440
         TabIndex        =   206
         Top             =   4200
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   178
         Left            =   -69000
         TabIndex        =   205
         Top             =   4500
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   177
         Left            =   -68220
         TabIndex        =   204
         Top             =   4500
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   176
         Left            =   -67440
         TabIndex        =   203
         Top             =   4500
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "TOTAL OF DIFFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   175
         Left            =   -74760
         TabIndex        =   202
         Top             =   4995
         Width           =   3345
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   174
         Left            =   -67650
         TabIndex        =   201
         Top             =   4995
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   173
         Left            =   -67590
         TabIndex        =   200
         Top             =   5760
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "TOTAL OF DIFFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   172
         Left            =   -74700
         TabIndex        =   199
         Top             =   5760
         Width           =   3345
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   171
         Left            =   -67380
         TabIndex        =   198
         Top             =   5310
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   170
         Left            =   -68160
         TabIndex        =   197
         Top             =   5310
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   169
         Left            =   -68940
         TabIndex        =   196
         Top             =   5310
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   168
         Left            =   -67380
         TabIndex        =   195
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   167
         Left            =   -68160
         TabIndex        =   194
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   166
         Left            =   -68940
         TabIndex        =   193
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   165
         Left            =   -67380
         TabIndex        =   192
         Top             =   4710
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   164
         Left            =   -68160
         TabIndex        =   191
         Top             =   4710
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   163
         Left            =   -68940
         TabIndex        =   190
         Top             =   4710
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   162
         Left            =   -67380
         TabIndex        =   189
         Top             =   4410
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   161
         Left            =   -68160
         TabIndex        =   188
         Top             =   4410
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   160
         Left            =   -68940
         TabIndex        =   187
         Top             =   4410
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   159
         Left            =   -67380
         TabIndex        =   186
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   158
         Left            =   -68160
         TabIndex        =   185
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   157
         Left            =   -68940
         TabIndex        =   184
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   156
         Left            =   -67380
         TabIndex        =   183
         Top             =   3810
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   155
         Left            =   -68160
         TabIndex        =   182
         Top             =   3810
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   154
         Left            =   -68940
         TabIndex        =   181
         Top             =   3810
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Diamonds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   153
         Left            =   -70410
         TabIndex        =   180
         Top             =   5310
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Spades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   152
         Left            =   -70410
         TabIndex        =   179
         Top             =   5010
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Hearts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   151
         Left            =   -70410
         TabIndex        =   178
         Top             =   4710
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Clubs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   150
         Left            =   -70410
         TabIndex        =   177
         Top             =   4410
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   149
         Left            =   -70410
         TabIndex        =   176
         Top             =   4110
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   148
         Left            =   -70410
         TabIndex        =   175
         Top             =   3810
         Width           =   1035
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   -70410
         X2              =   -66915
         Y1              =   3705
         Y2              =   3705
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   147
         Left            =   -67380
         TabIndex        =   174
         Top             =   3390
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   146
         Left            =   -68190
         TabIndex        =   173
         Top             =   3390
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   145
         Left            =   -68940
         TabIndex        =   172
         Top             =   3390
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "4th Quarter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   144
         Left            =   -70410
         TabIndex        =   171
         Top             =   3390
         Width           =   1170
      End
      Begin VB.Line Line9 
         Index           =   6
         X1              =   -74700
         X2              =   -66810
         Y1              =   5655
         Y2              =   5655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   143
         Left            =   -71670
         TabIndex        =   170
         Top             =   5310
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   142
         Left            =   -72450
         TabIndex        =   169
         Top             =   5310
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   141
         Left            =   -73230
         TabIndex        =   168
         Top             =   5310
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   140
         Left            =   -71670
         TabIndex        =   167
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   139
         Left            =   -72450
         TabIndex        =   166
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   138
         Left            =   -73230
         TabIndex        =   165
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   137
         Left            =   -71670
         TabIndex        =   164
         Top             =   4710
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   136
         Left            =   -72450
         TabIndex        =   163
         Top             =   4710
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   135
         Left            =   -73230
         TabIndex        =   162
         Top             =   4710
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   134
         Left            =   -71670
         TabIndex        =   161
         Top             =   4410
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   133
         Left            =   -72450
         TabIndex        =   160
         Top             =   4410
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   132
         Left            =   -73230
         TabIndex        =   159
         Top             =   4410
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   131
         Left            =   -71670
         TabIndex        =   158
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   130
         Left            =   -72450
         TabIndex        =   157
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   129
         Left            =   -73230
         TabIndex        =   156
         Top             =   4110
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   128
         Left            =   -71670
         TabIndex        =   155
         Top             =   3810
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   127
         Left            =   -72450
         TabIndex        =   154
         Top             =   3810
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   126
         Left            =   -73230
         TabIndex        =   153
         Top             =   3810
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Diamonds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   125
         Left            =   -74700
         TabIndex        =   152
         Top             =   5310
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Spades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   124
         Left            =   -74700
         TabIndex        =   151
         Top             =   5010
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Hearts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   123
         Left            =   -74700
         TabIndex        =   150
         Top             =   4710
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Clubs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   122
         Left            =   -74700
         TabIndex        =   149
         Top             =   4410
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   121
         Left            =   -74700
         TabIndex        =   148
         Top             =   4110
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   120
         Left            =   -74700
         TabIndex        =   147
         Top             =   3810
         Width           =   1035
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   -74700
         X2              =   -71205
         Y1              =   3705
         Y2              =   3705
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   119
         Left            =   -71670
         TabIndex        =   146
         Top             =   3390
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   118
         Left            =   -72480
         TabIndex        =   145
         Top             =   3390
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   117
         Left            =   -73230
         TabIndex        =   144
         Top             =   3390
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "3rd Quarter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   116
         Left            =   -74700
         TabIndex        =   143
         Top             =   3390
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   115
         Left            =   -67380
         TabIndex        =   142
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   114
         Left            =   -68160
         TabIndex        =   141
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   113
         Left            =   -68940
         TabIndex        =   140
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   112
         Left            =   -67380
         TabIndex        =   139
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   111
         Left            =   -68160
         TabIndex        =   138
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   110
         Left            =   -68940
         TabIndex        =   137
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   109
         Left            =   -67380
         TabIndex        =   136
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   108
         Left            =   -68160
         TabIndex        =   135
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   107
         Left            =   -68940
         TabIndex        =   134
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   106
         Left            =   -67380
         TabIndex        =   133
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   105
         Left            =   -68160
         TabIndex        =   132
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   104
         Left            =   -68940
         TabIndex        =   131
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   103
         Left            =   -67380
         TabIndex        =   130
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   102
         Left            =   -68160
         TabIndex        =   129
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   101
         Left            =   -68940
         TabIndex        =   128
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   100
         Left            =   -67380
         TabIndex        =   127
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   99
         Left            =   -68160
         TabIndex        =   126
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   98
         Left            =   -68940
         TabIndex        =   125
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Diamonds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   97
         Left            =   -70410
         TabIndex        =   124
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Spades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   96
         Left            =   -70410
         TabIndex        =   123
         Top             =   2640
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Hearts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   95
         Left            =   -70410
         TabIndex        =   122
         Top             =   2340
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Clubs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   94
         Left            =   -70410
         TabIndex        =   121
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   93
         Left            =   -70410
         TabIndex        =   120
         Top             =   1740
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   92
         Left            =   -70410
         TabIndex        =   119
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   -70410
         X2              =   -66915
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   91
         Left            =   -67380
         TabIndex        =   118
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   90
         Left            =   -68190
         TabIndex        =   117
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   89
         Left            =   -68940
         TabIndex        =   116
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "2nd Quarter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   88
         Left            =   -70410
         TabIndex        =   115
         Top             =   1020
         Width           =   1170
      End
      Begin VB.Line Line9 
         Index           =   3
         X1              =   -74700
         X2              =   -66810
         Y1              =   3285
         Y2              =   3285
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   87
         Left            =   -71670
         TabIndex        =   114
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   86
         Left            =   -72450
         TabIndex        =   113
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   85
         Left            =   -73230
         TabIndex        =   112
         Top             =   2940
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   84
         Left            =   -71670
         TabIndex        =   111
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   83
         Left            =   -72450
         TabIndex        =   110
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   82
         Left            =   -73230
         TabIndex        =   109
         Top             =   2640
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   81
         Left            =   -71670
         TabIndex        =   108
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   80
         Left            =   -72450
         TabIndex        =   107
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   79
         Left            =   -73230
         TabIndex        =   106
         Top             =   2340
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   78
         Left            =   -71670
         TabIndex        =   105
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   77
         Left            =   -72450
         TabIndex        =   104
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   76
         Left            =   -73230
         TabIndex        =   103
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   75
         Left            =   -71670
         TabIndex        =   102
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   74
         Left            =   -72450
         TabIndex        =   101
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   73
         Left            =   -73230
         TabIndex        =   100
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   72
         Left            =   -71670
         TabIndex        =   99
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   71
         Left            =   -72450
         TabIndex        =   98
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   70
         Left            =   -73230
         TabIndex        =   97
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Diamonds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   69
         Left            =   -74700
         TabIndex        =   96
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Spades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   68
         Left            =   -74700
         TabIndex        =   95
         Top             =   2640
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Hearts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   67
         Left            =   -74700
         TabIndex        =   94
         Top             =   2340
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Clubs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   66
         Left            =   -74700
         TabIndex        =   93
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   65
         Left            =   -74700
         TabIndex        =   92
         Top             =   1740
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   64
         Left            =   -74700
         TabIndex        =   91
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   -74700
         X2              =   -71205
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   63
         Left            =   -71670
         TabIndex        =   90
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   62
         Left            =   -72480
         TabIndex        =   89
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   61
         Left            =   -73230
         TabIndex        =   88
         Top             =   1020
         Width           =   435
      End
      Begin VB.Line Line9 
         Index           =   1
         X1              =   -74700
         X2              =   -66810
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label5 
         Caption         =   "1st Quarter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   60
         Left            =   -74700
         TabIndex        =   87
         Top             =   1020
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "Analysis of Distribution in Thirteen-Card Packets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   59
         Left            =   -74700
         TabIndex        =   86
         Top             =   555
         Width           =   6840
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   58
         Left            =   -67530
         TabIndex        =   85
         Top             =   5655
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   57
         Left            =   -67530
         TabIndex        =   84
         Top             =   5220
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   56
         Left            =   -69255
         TabIndex        =   83
         Top             =   5220
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "7.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   55
         Left            =   -71280
         TabIndex        =   82
         Top             =   5220
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   54
         Left            =   -67530
         TabIndex        =   81
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   53
         Left            =   -69255
         TabIndex        =   80
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "4.7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   52
         Left            =   -71280
         TabIndex        =   79
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   51
         Left            =   -67530
         TabIndex        =   78
         Top             =   4620
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   50
         Left            =   -69255
         TabIndex        =   77
         Top             =   4620
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   49
         Left            =   -71280
         TabIndex        =   76
         Top             =   4620
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   48
         Left            =   -67530
         TabIndex        =   75
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   47
         Left            =   -69255
         TabIndex        =   74
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   46
         Left            =   -71280
         TabIndex        =   73
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   45
         Left            =   -67530
         TabIndex        =   72
         Top             =   4020
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   44
         Left            =   -69255
         TabIndex        =   71
         Top             =   4020
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   43
         Left            =   -71280
         TabIndex        =   70
         Top             =   4020
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   42
         Left            =   -67530
         TabIndex        =   69
         Top             =   3630
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   41
         Left            =   -69255
         TabIndex        =   68
         Top             =   3630
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   40
         Left            =   -71280
         TabIndex        =   67
         Top             =   3630
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   39
         Left            =   -67530
         TabIndex        =   66
         Top             =   3330
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   38
         Left            =   -69255
         TabIndex        =   65
         Top             =   3330
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "7.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   37
         Left            =   -71280
         TabIndex        =   64
         Top             =   3330
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   36
         Left            =   -67530
         TabIndex        =   63
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   35
         Left            =   -69255
         TabIndex        =   62
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   34
         Left            =   -71280
         TabIndex        =   61
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   33
         Left            =   -67530
         TabIndex        =   60
         Top             =   2730
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   32
         Left            =   -69255
         TabIndex        =   59
         Top             =   2730
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "2.4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   31
         Left            =   -71280
         TabIndex        =   58
         Top             =   2730
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   30
         Left            =   -67530
         TabIndex        =   57
         Top             =   2430
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   29
         Left            =   -69255
         TabIndex        =   56
         Top             =   2430
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0.2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   28
         Left            =   -71280
         TabIndex        =   55
         Top             =   2430
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   27
         Left            =   -67530
         TabIndex        =   54
         Top             =   2025
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   26
         Left            =   -69255
         TabIndex        =   53
         Top             =   2025
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "4.9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   25
         Left            =   -71280
         TabIndex        =   52
         Top             =   2025
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   -67530
         TabIndex        =   51
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   -69255
         TabIndex        =   50
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   -71280
         TabIndex        =   49
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   -67530
         TabIndex        =   48
         Top             =   1425
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   -69255
         TabIndex        =   47
         Top             =   1425
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1.6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   -71280
         TabIndex        =   46
         Top             =   1425
         Width           =   615
      End
      Begin VB.Line Line12 
         X1              =   -74700
         X2              =   -66810
         Y1              =   5550
         Y2              =   5550
      End
      Begin VB.Line Line11 
         X1              =   -74700
         X2              =   -66810
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line10 
         X1              =   -74700
         X2              =   -66810
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Label Label5 
         Caption         =   "Packets With..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   -74700
         TabIndex        =   45
         Top             =   1125
         Width           =   2025
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   -71543
         TabIndex        =   44
         Top             =   1125
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   -69345
         TabIndex        =   43
         Top             =   1125
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Difference"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   -67830
         TabIndex        =   42
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "2 Pairs of Same Suit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -74700
         TabIndex        =   41
         Top             =   3030
         Width           =   2190
      End
      Begin VB.Label Label5 
         Caption         =   "3 Cards of Same Suit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   -74700
         TabIndex        =   40
         Top             =   2730
         Width           =   2220
      End
      Begin VB.Label Label5 
         Caption         =   "1 Pair of Same Suit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -74700
         TabIndex        =   39
         Top             =   3330
         Width           =   2040
      End
      Begin VB.Label Label5 
         Caption         =   "4 Cards of Same Suit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -74700
         TabIndex        =   38
         Top             =   2430
         Width           =   2265
      End
      Begin VB.Label Label5 
         Caption         =   "4 Different Suits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -74700
         TabIndex        =   37
         Top             =   3630
         Width           =   1740
      End
      Begin VB.Label Label5 
         Caption         =   "4 Cards of Same Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -74700
         TabIndex        =   36
         Top             =   4020
         Width           =   2460
      End
      Begin VB.Label Label5 
         Caption         =   "3 Cards of Same Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -74700
         TabIndex        =   35
         Top             =   4320
         Width           =   2430
      End
      Begin VB.Label Label5 
         Caption         =   "2 Pairs of Same Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -74700
         TabIndex        =   34
         Top             =   4620
         Width           =   2355
      End
      Begin VB.Label Label5 
         Caption         =   "1 Pair of Same Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -74700
         TabIndex        =   33
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "2 Pairs of Same Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -74700
         TabIndex        =   32
         Top             =   2025
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "3 Cards of Same Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -74700
         TabIndex        =   31
         Top             =   1725
         Width           =   2490
      End
      Begin VB.Label Label5 
         Caption         =   "4 Cards of Same Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -74700
         TabIndex        =   30
         Top             =   1425
         Width           =   2550
      End
      Begin VB.Label Label5 
         Caption         =   "4 Different Values"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -74700
         TabIndex        =   29
         Top             =   5220
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "TOTAL OF DIFFERENCES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -74700
         TabIndex        =   28
         Top             =   5655
         Width           =   3180
      End
      Begin VB.Label Label5 
         Caption         =   "Analysis of Permutations in Four-Card Packets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -74700
         TabIndex        =   27
         Top             =   675
         Width           =   6360
      End
      Begin VB.Line Line9 
         Index           =   0
         X1              =   -74700
         X2              =   -66810
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label SMB_TotalDiff 
         Alignment       =   2  'Center
         Caption         =   "49.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68355
         TabIndex        =   26
         Top             =   2970
         Width           =   900
      End
      Begin VB.Label SMB_Diff_V 
         Alignment       =   2  'Center
         Caption         =   "49.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68280
         TabIndex        =   25
         Top             =   2310
         Width           =   735
      End
      Begin VB.Label SMB_Diff_S 
         Alignment       =   2  'Center
         Caption         =   "39.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68280
         TabIndex        =   24
         Top             =   1905
         Width           =   735
      End
      Begin VB.Label SMB_Diff_C 
         Alignment       =   2  'Center
         Caption         =   "26.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68280
         TabIndex        =   23
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label SMB_Actual_V 
         Alignment       =   2  'Center
         Caption         =   "49.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70170
         TabIndex        =   22
         Top             =   2310
         Width           =   645
      End
      Begin VB.Label SMB_Actual_S 
         Alignment       =   2  'Center
         Caption         =   "39.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70170
         TabIndex        =   21
         Top             =   1905
         Width           =   645
      End
      Begin VB.Label SMB_Actual_C 
         Alignment       =   2  'Center
         Caption         =   "26.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70170
         TabIndex        =   20
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label Label19 
         Caption         =   "TOTAL OF DIFFERENCES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74760
         TabIndex        =   19
         Top             =   2970
         Width           =   3270
      End
      Begin VB.Line Line8 
         X1              =   -74775
         X2              =   -66945
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "49.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72060
         TabIndex        =   18
         Top             =   2310
         Width           =   645
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "39.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72060
         TabIndex        =   17
         Top             =   1905
         Width           =   645
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "26.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72060
         TabIndex        =   16
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Difference"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68520
         TabIndex        =   15
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70320
         TabIndex        =   14
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72218
         TabIndex        =   13
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label12 
         Caption         =   "Value Breaks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74760
         TabIndex        =   12
         Top             =   2310
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Suit Breaks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74760
         TabIndex        =   11
         Top             =   1905
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Color Breaks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74760
         TabIndex        =   10
         Top             =   1500
         Width           =   1935
      End
      Begin VB.Line Line7 
         X1              =   -74775
         X2              =   -66885
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label9 
         Caption         =   "Analysis of Breaks in the Stack"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   3585
      End
      Begin VB.Label Label8 
         Caption         =   "Number of Value Cycles:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   1170
         Width           =   2565
      End
      Begin VB.Label Label7 
         Caption         =   "Number of Suit Cycles:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   5
         Top             =   840
         Width           =   2490
      End
      Begin VB.Label Label4 
         Caption         =   "Number of Color Cycles:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   510
         Width           =   2565
      End
      Begin VB.Label Label3 
         Caption         =   "Value Spread"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   3
         Top             =   6210
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Suit Spread"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   3975
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Color Spread"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   2475
         Width           =   1995
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   51
         Left            =   7950
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   50
         Left            =   7800
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   49
         Left            =   7650
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   48
         Left            =   7500
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   47
         Left            =   7350
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   46
         Left            =   7200
         Top             =   5265
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   45
         Left            =   7050
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   44
         Left            =   6900
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   43
         Left            =   6750
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   42
         Left            =   6600
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   41
         Left            =   6450
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   40
         Left            =   6300
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   39
         Left            =   6150
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   38
         Left            =   6000
         Top             =   5340
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   37
         Left            =   5850
         Top             =   5355
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   36
         Left            =   5700
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   35
         Left            =   5550
         Top             =   5340
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   34
         Left            =   5400
         Top             =   5355
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   33
         Left            =   5250
         Top             =   5355
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   32
         Left            =   5100
         Top             =   5340
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   31
         Left            =   4950
         Top             =   5355
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   30
         Left            =   4800
         Top             =   5370
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   29
         Left            =   4650
         Top             =   5340
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   28
         Left            =   4500
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   27
         Left            =   4350
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   26
         Left            =   4200
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   25
         Left            =   4050
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   24
         Left            =   3900
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   23
         Left            =   3750
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   22
         Left            =   3600
         Top             =   5340
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   21
         Left            =   3450
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   20
         Left            =   3300
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   19
         Left            =   3150
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   18
         Left            =   3000
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   17
         Left            =   2850
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   16
         Left            =   2700
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   15
         Left            =   2550
         Top             =   5340
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   14
         Left            =   2400
         Top             =   5325
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   13
         Left            =   2250
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000834&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   12
         Left            =   2100
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   11
         Left            =   1950
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   10
         Left            =   1800
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   9
         Left            =   1650
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   8
         Left            =   1500
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   7
         Left            =   1350
         Top             =   5310
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   6
         Left            =   1200
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   5
         Left            =   1050
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   4
         Left            =   900
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   3
         Left            =   750
         Top             =   5295
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   2
         Left            =   600
         Top             =   5280
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   450
         Top             =   5265
         Width           =   105
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   0
         Left            =   300
         Top             =   5265
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   51
         Left            =   7950
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   50
         Left            =   7800
         Top             =   3585
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   49
         Left            =   7650
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   48
         Left            =   7500
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   47
         Left            =   7350
         Top             =   3570
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   46
         Left            =   7200
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   45
         Left            =   7050
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   44
         Left            =   6900
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   43
         Left            =   6750
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   42
         Left            =   6600
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   41
         Left            =   6450
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   40
         Left            =   6300
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   39
         Left            =   6150
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   38
         Left            =   6000
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   37
         Left            =   5850
         Top             =   3630
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   36
         Left            =   5700
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   35
         Left            =   5550
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   34
         Left            =   5400
         Top             =   3570
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   33
         Left            =   5250
         Top             =   3585
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   32
         Left            =   5100
         Top             =   3585
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   31
         Left            =   4950
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   30
         Left            =   4800
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   29
         Left            =   4650
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   28
         Left            =   4500
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   27
         Left            =   4350
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   26
         Left            =   4200
         Top             =   3585
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   25
         Left            =   4050
         Top             =   3585
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   24
         Left            =   3900
         Top             =   3570
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   23
         Left            =   3750
         Top             =   3585
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   22
         Left            =   3600
         Top             =   3570
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   21
         Left            =   3450
         Top             =   3585
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   20
         Left            =   3300
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   19
         Left            =   3150
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   18
         Left            =   3000
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   17
         Left            =   2850
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   16
         Left            =   2700
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   15
         Left            =   2550
         Top             =   3630
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   14
         Left            =   2400
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   13
         Left            =   2250
         Top             =   3630
         Width           =   105
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000834&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   12
         Left            =   2100
         Top             =   3630
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   11
         Left            =   1950
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   10
         Left            =   1800
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   9
         Left            =   1650
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   8
         Left            =   1500
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   7
         Left            =   1350
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   6
         Left            =   1200
         Top             =   3630
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   5
         Left            =   1050
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   4
         Left            =   900
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   3
         Left            =   750
         Top             =   3615
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   2
         Left            =   600
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   450
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   0
         Left            =   300
         Top             =   3600
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   51
         Left            =   7950
         Top             =   2070
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   50
         Left            =   7800
         Top             =   2085
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   49
         Left            =   7650
         Top             =   2085
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   48
         Left            =   7500
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   47
         Left            =   7350
         Top             =   2085
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   46
         Left            =   7200
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   45
         Left            =   7050
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   44
         Left            =   6900
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   43
         Left            =   6750
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   42
         Left            =   6600
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   41
         Left            =   6450
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   40
         Left            =   6300
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   39
         Left            =   6150
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   38
         Left            =   6000
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   37
         Left            =   5850
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   36
         Left            =   5700
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   35
         Left            =   5550
         Top             =   2160
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   34
         Left            =   5400
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   33
         Left            =   5250
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   32
         Left            =   5100
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   31
         Left            =   4950
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   30
         Left            =   4800
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   29
         Left            =   4650
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   28
         Left            =   4500
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   27
         Left            =   4350
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   26
         Left            =   4200
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   25
         Left            =   4050
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   24
         Left            =   3900
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   23
         Left            =   3750
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   22
         Left            =   3600
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   21
         Left            =   3450
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   20
         Left            =   3300
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   19
         Left            =   3150
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   18
         Left            =   3000
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   17
         Left            =   2850
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   16
         Left            =   2700
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   15
         Left            =   2550
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   14
         Left            =   2400
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   13
         Left            =   2250
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000834&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   12
         Left            =   2100
         Top             =   2115
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   11
         Left            =   1950
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   10
         Left            =   1800
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   9
         Left            =   1650
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   8
         Left            =   1500
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   7
         Left            =   1350
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   6
         Left            =   1200
         Top             =   2160
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   5
         Left            =   1050
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   4
         Left            =   900
         Top             =   2145
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   3
         Left            =   750
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   2
         Left            =   600
         Top             =   2130
         Width           =   105
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   450
         Top             =   2100
         Width           =   105
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   0
         Left            =   300
         Top             =   2100
         Width           =   105
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         X1              =   300
         X2              =   8190
         Y1              =   6180
         Y2              =   6180
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   300
         X2              =   8190
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   300
         X2              =   8085
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   300
         X2              =   8190
         Y1              =   5745
         Y2              =   5745
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   300
         X2              =   8190
         Y1              =   5325
         Y2              =   5325
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   300
         X2              =   8190
         Y1              =   4905
         Y2              =   4905
      End
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8400
      TabIndex        =   276
      Top             =   1275
      Width           =   150
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6030
      TabIndex        =   275
      Top             =   1275
      Width           =   150
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1125
      Left            =   255
      Picture         =   "frmShuffleMeter.frx":1D56
      Top             =   240
      Width           =   3795
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Shuffle Index"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6150
      TabIndex        =   8
      Top             =   150
      Width           =   2355
   End
   Begin VB.Label ShuffleIndexLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6105
      TabIndex        =   7
      Top             =   690
      Width           =   405
   End
   Begin VB.Image ShuffleIndexArrow 
      Height          =   270
      Left            =   6165
      Picture         =   "frmShuffleMeter.frx":FC42
      Top             =   960
      Width           =   255
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1050
      Left            =   5910
      Picture         =   "frmShuffleMeter.frx":1002E
      Top             =   510
      Width           =   2775
   End
End
Attribute VB_Name = "frmShuffleMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SetShuffleMeterDeck()
'transpose the current deck to the ShuffleMeterDeck code
Dim pSuit As Variant
Dim pValue As Variant
For i% = 1 To 52
    pSuit = Right(Deck(2, i%), 1)
    pValue = Left(Deck(2, i%), Len(Deck(2, i%)) - 1)
    'set color
    If pSuit = "H" Or pSuit = "D" Then
        ShuffleMeterDeck(1, i%) = 0
    Else
        ShuffleMeterDeck(1, i%) = 1
    End If
    'set suit
    If pSuit = "C" Then
        ShuffleMeterDeck(2, i%) = 1
    ElseIf pSuit = "H" Then
        ShuffleMeterDeck(2, i%) = 2
    ElseIf pSuit = "S" Then
        ShuffleMeterDeck(2, i%) = 3
    ElseIf pSuit = "D" Then
        ShuffleMeterDeck(2, i%) = 4
    End If
    'set value
    If pValue = "A" Then
        ShuffleMeterDeck(3, i%) = 1
    ElseIf pValue = "2" Then
        ShuffleMeterDeck(3, i%) = 2
    ElseIf pValue = "3" Then
        ShuffleMeterDeck(3, i%) = 3
    ElseIf pValue = "4" Then
        ShuffleMeterDeck(3, i%) = 4
    ElseIf pValue = "5" Then
        ShuffleMeterDeck(3, i%) = 5
    ElseIf pValue = "6" Then
        ShuffleMeterDeck(3, i%) = 6
    ElseIf pValue = "7" Then
        ShuffleMeterDeck(3, i%) = 7
    ElseIf pValue = "8" Then
        ShuffleMeterDeck(3, i%) = 8
    ElseIf pValue = "9" Then
        ShuffleMeterDeck(3, i%) = 9
    ElseIf pValue = "10" Then
        ShuffleMeterDeck(3, i%) = 10
    ElseIf pValue = "J" Then
        ShuffleMeterDeck(3, i%) = 11
    ElseIf pValue = "Q" Then
        ShuffleMeterDeck(3, i%) = 12
    ElseIf pValue = "K" Then
        ShuffleMeterDeck(3, i%) = 13
    End If
Next i%
End Sub

Public Sub SetSpread()
'turn off the bars
For i% = 0 To 51
    Shape1(i%).Visible = False
    Shape2(i%).Visible = False
    Shape3(i%).Visible = False
Next i%
'set the color spread bar charts
For i% = 1 To 52
    If ShuffleMeterDeck(1, i%) = 0 Then
        'Shape1(i% - 1).Left = 300 + (i% - 1) * 150
        Shape1(i% - 1).Top = 2100
        'Shape1(i% - 1).Width = 105
        Shape1(i% - 1).Height = 300
        Shape1(i% - 1).BorderColor = &HFF&
        Shape1(i% - 1).FillColor = &HFF&
        'Shape1(i% - 1).FillStyle = 0
    Else
        'Shape1(i% - 1).Left = 300 + (i% - 1) * 150
        Shape1(i% - 1).Top = 1800
        'Shape1(i% - 1).Width = 105
        Shape1(i% - 1).Height = 600
        Shape1(i% - 1).BorderColor = &H0&
        Shape1(i% - 1).FillColor = &H0&
        'Shape1(i% - 1).FillStyle = 0
    End If
Next i%
'set the suit spread bar charts
For i% = 1 To 52
    If ShuffleMeterDeck(2, i%) = 1 Then
        'Shape2(i% - 1).Left = 300 + (i% - 1) * 150
        Shape2(i% - 1).Top = 3600 + 150
        'Shape2(i% - 1).Width = 105
        Shape2(i% - 1).Height = 150
        Shape2(i% - 1).BorderColor = &H0&
        Shape2(i% - 1).FillColor = &H0&
        'Shape2(i% - 1).FillStyle = 0
    ElseIf ShuffleMeterDeck(2, i%) = 2 Then
        'Shape2(i% - 1).Left = 300 + (i% - 1) * 150
        Shape2(i% - 1).Top = 3600
        'Shape2(i% - 1).Width = 105
        Shape2(i% - 1).Height = 300
        Shape2(i% - 1).BorderColor = &HFF&
        Shape2(i% - 1).FillColor = &HFF&
        'Shape2(i% - 1).FillStyle = 0
    ElseIf ShuffleMeterDeck(2, i%) = 3 Then
        'Shape2(i% - 1).Left = 300 + (i% - 1) * 150
        Shape2(i% - 1).Top = 3600 - 150
        'Shape2(i% - 1).Width = 105
        Shape2(i% - 1).Height = 450
        Shape2(i% - 1).BorderColor = &H0&
        Shape2(i% - 1).FillColor = &H0&
        'Shape2(i% - 1).FillStyle = 0
    ElseIf ShuffleMeterDeck(2, i%) = 4 Then
        'Shape2(i% - 1).Left = 300 + (i% - 1) * 150
        Shape2(i% - 1).Top = 3600 - 300
        'Shape2(i% - 1).Width = 105
        Shape2(i% - 1).Height = 600
        Shape2(i% - 1).BorderColor = &HFF&
        Shape2(i% - 1).FillColor = &HFF&
        'Shape2(i% - 1).FillStyle = 0
    End If
Next i%
'set the value spread bar charts
For i% = 1 To 52
    'Shape3(i% - 1).Left = 300 + (i% - 1) * 150
    Shape3(i% - 1).Top = 4800 + 105 * (13 - ShuffleMeterDeck(3, i%))
    'Shape3(i% - 1).Width = 105
    Shape3(i% - 1).Height = 105 * ShuffleMeterDeck(3, i%)
    Shape3(i% - 1).BorderColor = &H0&
    Shape3(i% - 1).FillColor = &H0&
    'Shape3(i% - 1).FillStyle = 0
Next i%
'turn on the bars
For i% = 0 To 51
    Shape1(i%).Visible = True
    Shape2(i%).Visible = True
    Shape3(i%).Visible = True
Next i%
End Sub

Public Sub SetBreak()
Dim pColorCount As Integer
Dim pSuitCount As Integer
Dim pValueCount As Integer
Dim pLastColor As Integer
Dim pLastSuit As Integer
Dim pLastValue As Integer
Dim pColorDiff As Double
Dim pSuitDiff As Double
Dim pValueDiff As Double
Dim pTotalDiff As Double
Dim pShuffleIndex As Integer
'set the initial values of pLastxxx to the 52nd card
pLastColor = ShuffleMeterDeck(1, 52)
pLastSuit = ShuffleMeterDeck(2, 52)
pLastValue = ShuffleMeterDeck(3, 52)
'set the initial values of pxxxCount to zeros
pColorCount = 0
pSuitCount = 0
pValueCount = 0
'start counting
For i% = 1 To 52
    'color
    If pLastColor <> ShuffleMeterDeck(1, i%) Then
        pLastColor = ShuffleMeterDeck(1, i%)
        pColorCount = pColorCount + 1
    End If
    'suit
    If pLastSuit <> ShuffleMeterDeck(2, i%) Then
        pLastSuit = ShuffleMeterDeck(2, i%)
        pSuitCount = pSuitCount + 1
    End If
    'value
    If pLastValue <> ShuffleMeterDeck(3, i%) Then
        pLastValue = ShuffleMeterDeck(3, i%)
        pValueCount = pValueCount + 1
    End If
Next i%
SMB_Actual_C.Caption = pColorCount
SMB_Actual_S.Caption = pSuitCount
SMB_Actual_V.Caption = pValueCount
pColorDiff = Round(Abs(26.3 - pColorCount), 1)
pSuitDiff = Round(Abs(39 - pSuitCount), 1)
pValueDiff = Round(Abs(49 - pValueCount), 1)
SMB_Diff_C.Caption = pColorDiff
SMB_Diff_S.Caption = pSuitDiff
SMB_Diff_V.Caption = pValueDiff
pTotalDiff = pColorDiff + pSuitDiff + pValueDiff
pShuffleIndex = Round((2 * pColorDiff + 4 * pSuitDiff + 13 * pValueDiff), 0)
SMB_TotalDiff.Caption = pTotalDiff
ShuffleIndexLabel.Caption = pShuffleIndex
If pShuffleIndex > 250 Then
    ShuffleIndexLabel.Left = 6105 + 8 * 260
    ShuffleIndexArrow.Left = 6165 + 8 * 260
Else
    ShuffleIndexLabel.Left = 6105 + 8 * pShuffleIndex
    ShuffleIndexArrow.Left = 6165 + 8 * pShuffleIndex
End If
End Sub

Public Sub SetPermutations()
'colors
Dim SMP4CSC As Double
Dim SMP4CSCd As Double
Dim SMP3CSC As Double
Dim SMP3CSCd As Double
Dim SMP2PSC As Double
Dim SMP2PSCd As Double
'suits
Dim SMP4CSS As Double
Dim SMP4CSSd As Double
Dim SMP3CSS As Double
Dim SMP3CSSd As Double
Dim SMP2PSS As Double
Dim SMP2PSSd As Double
Dim SMP1PSS As Double
Dim SMP1PSSd As Double
Dim SMP4CDS As Double
Dim SMP4CDSd As Double
'values
Dim SMP4CSV As Double
Dim SMP4CSVd As Double
Dim SMP3CSV As Double
Dim SMP3CSVd As Double
Dim SMP2PSV As Double
Dim SMP2PSVd As Double
Dim SMP1PSV As Double
Dim SMP1PSVd As Double
Dim SMP4CDV As Double
Dim SMP4CDVd As Double
Dim SMPTD As Double
'counts
Dim pRedCount As Integer
Dim pBlackCount As Integer
Dim pClubCount As Integer
Dim pHeartCount As Integer
Dim pSpadeCount As Integer
Dim pDiamondCount As Integer
Dim pAceCount As Integer
Dim pTwoCount As Integer
Dim pThreeCount As Integer
Dim pFourCount As Integer
Dim pFiveCount As Integer
Dim pSixCount As Integer
Dim pSevenCount As Integer
Dim pEightCount As Integer
Dim pNineCount As Integer
Dim pTenCount As Integer
Dim pJackCount As Integer
Dim pQueenCount As Integer
Dim pKingCount As Integer
'set parameters to zero
SMP4CSC = 0
SMP4CSCd = 0
SMP3CSC = 0
SMP3CSCd = 0
SMP2PSC = 0
SMP2PSCd = 0
SMP4CSS = 0
SMP4CSSd = 0
SMP3CSS = 0
SMP3CSSd = 0
SMP2PSS = 0
SMP2PSSd = 0
SMP1PSS = 0
SMP1PSSd = 0
SMP4CDS = 0
SMP4CDSd = 0
SMP4CSV = 0
SMP4CSVd = 0
SMP3CSV = 0
SMP3CSVd = 0
SMP2PSV = 0
SMP2PSVd = 0
SMP1PSV = 0
SMP1PSVd = 0
SMP4CDV = 0
SMP4CDVd = 0
'run counters for permutations
For p% = 1 To 13
    'set counters to zero
    pRedCount = 0
    pBlackCount = 0
    pClubCount = 0
    pHeartCount = 0
    pSpadeCount = 0
    pDiamondCount = 0
    pAceCount = 0
    pTwoCount = 0
    pThreeCount = 0
    pFourCount = 0
    pFiveCount = 0
    pSixCount = 0
    pSevenCount = 0
    pEightCount = 0
    pNineCount = 0
    pTenCount = 0
    pJackCount = 0
    pQueenCount = 0
    pKingCount = 0
    For q% = 1 To 4
        'colors
        If ShuffleMeterDeck(1, (p% - 1) * 4 + q%) = 0 Then
            pRedCount = pRedCount + 1
        Else
            pBlackCount = pBlackCount + 1
        End If
        'suits
        If ShuffleMeterDeck(2, (p% - 1) * 4 + q%) = 1 Then
            pClubCount = pClubCount + 1
        ElseIf ShuffleMeterDeck(2, (p% - 1) * 4 + q%) = 2 Then
            pHeartCount = pHeartCount + 1
        ElseIf ShuffleMeterDeck(2, (p% - 1) * 4 + q%) = 3 Then
            pSpadeCount = pSpadeCount + 1
        ElseIf ShuffleMeterDeck(2, (p% - 1) * 4 + q%) = 4 Then
            pDiamondCount = pDiamondCount + 1
        End If
        'values
        If ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 1 Then
            pAceCount = pAceCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 2 Then
            pTwoCount = pTwoCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 3 Then
            pThreeCount = pThreeCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 4 Then
            pFourCount = pFourCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 5 Then
            pFiveCount = pFiveCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 6 Then
            pSixCount = pSixCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 7 Then
            pSevenCount = pSevenCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 8 Then
            pEightCount = pEightCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 9 Then
            pNineCount = pNineCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 10 Then
            pTenCount = pTenCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 11 Then
            pJackCount = pJackCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 12 Then
            pQueenCount = pQueenCount + 1
        ElseIf ShuffleMeterDeck(3, (p% - 1) * 4 + q%) = 13 Then
            pKingCount = pKingCount + 1
        End If
    Next q%
    'process colors
    If pRedCount = 4 Or pBlackCount = 4 Then
        SMP4CSC = SMP4CSC + 1
    ElseIf pRedCount = 3 Or pBlackCount = 3 Then
        SMP3CSC = SMP3CSC + 1
    ElseIf pRedCount = 2 And pBlackCount = 2 Then
        SMP2PSC = SMP2PSC + 1
    End If
    'process suits
    If pClubCount = 4 Or pHeartCount = 4 Or _
        pSpadeCount = 4 Or pDiamondCount = 4 Then
            SMP4CSS = SMP4CSS + 1
    End If
    If pClubCount * pHeartCount * pSpadeCount * pDiamondCount = 1 Then
        SMP4CDS = SMP4CDS + 1
    End If
    'plug zeros for easier combination checks
    If pClubCount = 0 Then
        pClubCount = 1
    End If
    If pHeartCount = 0 Then
        pHeartCount = 1
    End If
    If pSpadeCount = 0 Then
        pSpadeCount = 1
    End If
    If pDiamondCount = 0 Then
        pDiamondCount = 1
    End If
    'plug a 1 for any 4's
    'this makes the check for 2 pairs correct
    If pClubCount = 4 Then
        pClubCount = 1
    End If
    If pHeartCount = 4 Then
        pHeartCount = 1
    End If
    If pSpadeCount = 4 Then
        pSpadeCount = 1
    End If
    If pDiamondCount = 4 Then
        pDiamondCount = 1
    End If
    'continue processing suits
    If pClubCount * pHeartCount * pSpadeCount * pDiamondCount = 4 Then
        SMP2PSS = SMP2PSS + 1
    End If
    If pClubCount * pHeartCount * pSpadeCount * pDiamondCount = 3 Then
        SMP3CSS = SMP3CSS + 1
    End If
    If pClubCount * pHeartCount * pSpadeCount * pDiamondCount = 2 Then
        SMP1PSS = SMP1PSS + 1
    End If
    'process values
    If pAceCount = 4 Or pTwoCount = 4 Or pThreeCount = 4 Or pFourCount = 4 _
        Or pFiveCount = 4 Or pSixCount = 4 Or pSevenCount = 4 Or pEightCount = 4 _
        Or pNineCount = 4 Or pTenCount = 4 Or pJackCount = 4 Or pQueenCount = 4 _
        Or pKingCount = 4 Then
            SMP4CSV = SMP4CSV + 1
    End If
    'plug zeros for easier combination checks
    If pAceCount = 0 Then
        pAceCount = 1
    End If
    If pTwoCount = 0 Then
        pTwoCount = 1
    End If
    If pThreeCount = 0 Then
        pThreeCount = 1
    End If
    If pFourCount = 0 Then
        pFourCount = 1
    End If
    If pFiveCount = 0 Then
        pFiveCount = 1
    End If
    If pSixCount = 0 Then
        pSixCount = 1
    End If
    If pSevenCount = 0 Then
        pSevenCount = 1
    End If
    If pEightCount = 0 Then
        pEightCount = 1
    End If
    If pNineCount = 0 Then
        pNineCount = 1
    End If
    If pTenCount = 0 Then
        pTenCount = 1
    End If
    If pJackCount = 0 Then
        pJackCount = 1
    End If
    If pQueenCount = 0 Then
        pQueenCount = 1
    End If
    If pKingCount = 0 Then
        pKingCount = 1
    End If
    'plug a 1 for any 4's
    'this makes the check for 2 pairs correct
    If pAceCount = 4 Then
        pAceCount = 1
    End If
    If pTwoCount = 4 Then
        pTwoCount = 1
    End If
    If pThreeCount = 4 Then
        pThreeCount = 1
    End If
    If pFourCount = 4 Then
        pFourCount = 1
    End If
    If pFiveCount = 4 Then
        pFiveCount = 1
    End If
    If pSixCount = 4 Then
        pSixCount = 1
    End If
    If pSevenCount = 4 Then
        pSevenCount = 1
    End If
    If pEightCount = 4 Then
        pEightCount = 1
    End If
    If pNineCount = 4 Then
        pNineCount = 1
    End If
    If pTenCount = 4 Then
        pTenCount = 1
    End If
    If pJackCount = 4 Then
        pJackCount = 1
    End If
    If pQueenCount = 4 Then
        pQueenCount = 1
    End If
    If pKingCount = 4 Then
        pKingCount = 1
    End If
    'continue processing values
    If pAceCount * pTwoCount * pThreeCount * pFourCount _
        * pFiveCount * pSixCount * pSevenCount * pEightCount _
        * pNineCount * pTenCount * pJackCount * pQueenCount _
        * pKingCount = 4 Then
            SMP2PSV = SMP2PSV + 1
    End If
    If pAceCount * pTwoCount * pThreeCount * pFourCount _
        * pFiveCount * pSixCount * pSevenCount * pEightCount _
        * pNineCount * pTenCount * pJackCount * pQueenCount _
        * pKingCount = 3 Then
            SMP3CSV = SMP3CSV + 1
    End If
    If pAceCount * pTwoCount * pThreeCount * pFourCount _
        * pFiveCount * pSixCount * pSevenCount * pEightCount _
        * pNineCount * pTenCount * pJackCount * pQueenCount _
        * pKingCount = 2 Then
            SMP1PSV = SMP1PSV + 1
    End If
    If pAceCount * pTwoCount * pThreeCount * pFourCount _
        * pFiveCount * pSixCount * pSevenCount * pEightCount _
        * pNineCount * pTenCount * pJackCount * pQueenCount _
        * pKingCount = 1 Then
            SMP4CDV = SMP4CDV + 1
    End If
Next p%
'calculate differences
SMP4CSCd = Abs(1.6 - SMP4CSC)
SMP3CSCd = Abs(6.5 - SMP3CSC)
SMP2PSCd = Abs(4.9 - SMP2PSC)
SMP4CSSd = Abs(0.2 - SMP4CSS)
SMP3CSSd = Abs(2.4 - SMP3CSS)
SMP2PSSd = Abs(1.8 - SMP2PSS)
SMP1PSSd = Abs(7.3 - SMP1PSS)
SMP4CDSd = Abs(1.2 - SMP4CDS)
SMP4CSVd = Abs(0 - SMP4CSV)
SMP3CSVd = Abs(0.3 - SMP3CSV)
SMP2PSVd = Abs(0.2 - SMP2PSV)
SMP1PSVd = Abs(4.7 - SMP1PSV)
SMP4CDVd = Abs(7.8 - SMP4CDV)
SMPTD = SMP4CSCd + SMP3CSCd + SMP2PSCd + _
    SMP4CSSd + SMP3CSSd + SMP2PSSd + SMP1PSSd + SMP4CDSd + _
    SMP4CSVd + SMP3CSVd + SMP2PSVd + SMP1PSVd + SMP4CDVd
'set labels
Label5(20).Caption = SMP4CSC
Label5(21).Caption = SMP4CSCd
Label5(23).Caption = SMP3CSC
Label5(24).Caption = SMP3CSCd
Label5(26).Caption = SMP2PSC
Label5(27).Caption = SMP2PSCd
Label5(29).Caption = SMP4CSS
Label5(30).Caption = SMP4CSSd
Label5(32).Caption = SMP3CSS
Label5(33).Caption = SMP3CSSd
Label5(35).Caption = SMP2PSS
Label5(36).Caption = SMP2PSSd
Label5(38).Caption = SMP1PSS
Label5(39).Caption = SMP1PSSd
Label5(41).Caption = SMP4CDS
Label5(42).Caption = SMP4CDSd
Label5(44).Caption = SMP4CSV
Label5(45).Caption = SMP4CSVd
Label5(47).Caption = SMP3CSV
Label5(48).Caption = SMP3CSVd
Label5(50).Caption = SMP2PSV
Label5(51).Caption = SMP2PSVd
Label5(53).Caption = SMP1PSV
Label5(54).Caption = SMP1PSVd
Label5(56).Caption = SMP4CDV
Label5(57).Caption = SMP4CDVd
Label5(58).Caption = SMPTD
End Sub

Public Sub SetDistribution()
'establish parameters
'quarter 1
Dim SMDQ1B As Double
Dim SMDQ1R As Double
Dim SMDQ1C As Double
Dim SMDQ1H As Double
Dim SMDQ1S As Double
Dim SMDQ1D As Double
Dim SMDQ1Bd As Double
Dim SMDQ1Rd As Double
Dim SMDQ1Cd As Double
Dim SMDQ1Hd As Double
Dim SMDQ1Sd As Double
Dim SMDQ1Dd As Double
'quarter 2
Dim SMDQ2B As Double
Dim SMDQ2R As Double
Dim SMDQ2C As Double
Dim SMDQ2H As Double
Dim SMDQ2S As Double
Dim SMDQ2D As Double
Dim SMDQ2Bd As Double
Dim SMDQ2Rd As Double
Dim SMDQ2Cd As Double
Dim SMDQ2Hd As Double
Dim SMDQ2Sd As Double
Dim SMDQ2Dd As Double
'quarter 3
Dim SMDQ3B As Double
Dim SMDQ3R As Double
Dim SMDQ3C As Double
Dim SMDQ3H As Double
Dim SMDQ3S As Double
Dim SMDQ3D As Double
Dim SMDQ3Bd As Double
Dim SMDQ3Rd As Double
Dim SMDQ3Cd As Double
Dim SMDQ3Hd As Double
Dim SMDQ3Sd As Double
Dim SMDQ3Dd As Double
'quarter 4
Dim SMDQ4B As Double
Dim SMDQ4R As Double
Dim SMDQ4C As Double
Dim SMDQ4H As Double
Dim SMDQ4S As Double
Dim SMDQ4D As Double
Dim SMDQ4Bd As Double
Dim SMDQ4Rd As Double
Dim SMDQ4Cd As Double
Dim SMDQ4Hd As Double
Dim SMDQ4Sd As Double
Dim SMDQ4Dd As Double
'total
Dim SMDTD As Double
'counts
Dim pRedCount As Integer
Dim pBlackCount As Integer
Dim pClubCount As Integer
Dim pHeartCount As Integer
Dim pSpadeCount As Integer
Dim pDiamondCount As Integer
'set parameters to zero
'quarter 1
SMDQ1B = 0
SMDQ1R = 0
SMDQ1C = 0
SMDQ1H = 0
SMDQ1S = 0
SMDQ1D = 0
SMDQ1Bd = 0
SMDQ1Rd = 0
SMDQ1Cd = 0
SMDQ1Hd = 0
SMDQ1Sd = 0
SMDQ1Dd = 0
'quarter 2
SMDQ2B = 0
SMDQ2R = 0
SMDQ2C = 0
SMDQ2H = 0
SMDQ2S = 0
SMDQ2D = 0
SMDQ2Bd = 0
SMDQ2Rd = 0
SMDQ2Cd = 0
SMDQ2Hd = 0
SMDQ2Sd = 0
SMDQ2Dd = 0
'quarter 3
SMDQ3B = 0
SMDQ3R = 0
SMDQ3C = 0
SMDQ3H = 0
SMDQ3S = 0
SMDQ3D = 0
SMDQ3Bd = 0
SMDQ3Rd = 0
SMDQ3Cd = 0
SMDQ3Hd = 0
SMDQ3Sd = 0
SMDQ3Dd = 0
'quarter 4
SMDQ4B = 0
SMDQ4R = 0
SMDQ4C = 0
SMDQ4H = 0
SMDQ4S = 0
SMDQ4D = 0
SMDQ4Bd = 0
SMDQ4Rd = 0
SMDQ4Cd = 0
SMDQ4Hd = 0
SMDQ4Sd = 0
SMDQ4Dd = 0
'run counters for distributions
For p% = 1 To 4
    'set counters to zero
    pRedCount = 0
    pBlackCount = 0
    pClubCount = 0
    pHeartCount = 0
    pSpadeCount = 0
    pDiamondCount = 0
    For q% = 1 To 13
        'colors
        If ShuffleMeterDeck(1, (p% - 1) * 13 + q%) = 0 Then
            pRedCount = pRedCount + 1
        Else
            pBlackCount = pBlackCount + 1
        End If
        'suits
        If ShuffleMeterDeck(2, (p% - 1) * 13 + q%) = 1 Then
            pClubCount = pClubCount + 1
        ElseIf ShuffleMeterDeck(2, (p% - 1) * 13 + q%) = 2 Then
            pHeartCount = pHeartCount + 1
        ElseIf ShuffleMeterDeck(2, (p% - 1) * 13 + q%) = 3 Then
            pSpadeCount = pSpadeCount + 1
        ElseIf ShuffleMeterDeck(2, (p% - 1) * 13 + q%) = 4 Then
            pDiamondCount = pDiamondCount + 1
        End If
    Next q%
    'process data
    If p% = 1 Then
        SMDQ1B = pBlackCount
        SMDQ1R = pRedCount
        SMDQ1C = pClubCount
        SMDQ1H = pHeartCount
        SMDQ1S = pSpadeCount
        SMDQ1D = pDiamondCount
    ElseIf p% = 2 Then
        SMDQ2B = pBlackCount
        SMDQ2R = pRedCount
        SMDQ2C = pClubCount
        SMDQ2H = pHeartCount
        SMDQ2S = pSpadeCount
        SMDQ2D = pDiamondCount
    ElseIf p% = 3 Then
        SMDQ3B = pBlackCount
        SMDQ3R = pRedCount
        SMDQ3C = pClubCount
        SMDQ3H = pHeartCount
        SMDQ3S = pSpadeCount
        SMDQ3D = pDiamondCount
    ElseIf p% = 4 Then
        SMDQ4B = pBlackCount
        SMDQ4R = pRedCount
        SMDQ4C = pClubCount
        SMDQ4H = pHeartCount
        SMDQ4S = pSpadeCount
        SMDQ4D = pDiamondCount
    End If
Next p%
'calculate differences
    'quarter 1
    SMDQ1Bd = Abs(6.5 - SMDQ1B)
    SMDQ1Rd = Abs(6.5 - SMDQ1R)
    SMDQ1Cd = Abs(3.3 - SMDQ1C)
    SMDQ1Hd = Abs(3.3 - SMDQ1H)
    SMDQ1Sd = Abs(3.3 - SMDQ1S)
    SMDQ1Dd = Abs(3.3 - SMDQ1D)
    'quarter 2
    SMDQ2Bd = Abs(6.5 - SMDQ2B)
    SMDQ2Rd = Abs(6.5 - SMDQ2R)
    SMDQ2Cd = Abs(3.3 - SMDQ2C)
    SMDQ2Hd = Abs(3.3 - SMDQ2H)
    SMDQ2Sd = Abs(3.3 - SMDQ2S)
    SMDQ2Dd = Abs(3.3 - SMDQ2D)
    'quarter 3
    SMDQ3Bd = Abs(6.5 - SMDQ3B)
    SMDQ3Rd = Abs(6.5 - SMDQ3R)
    SMDQ3Cd = Abs(3.3 - SMDQ3C)
    SMDQ3Hd = Abs(3.3 - SMDQ3H)
    SMDQ3Sd = Abs(3.3 - SMDQ3S)
    SMDQ3Dd = Abs(3.3 - SMDQ3D)
    'quarter 4
    SMDQ4Bd = Abs(6.5 - SMDQ4B)
    SMDQ4Rd = Abs(6.5 - SMDQ4R)
    SMDQ4Cd = Abs(3.3 - SMDQ4C)
    SMDQ4Hd = Abs(3.3 - SMDQ4H)
    SMDQ4Sd = Abs(3.3 - SMDQ4S)
    SMDQ4Dd = Abs(3.3 - SMDQ4D)
    'total
    SMDTD = SMDQ1Bd + SMDQ1Rd + SMDQ1Cd + SMDQ1Hd + SMDQ1Sd + SMDQ1Dd + _
        SMDQ2Bd + SMDQ2Rd + SMDQ2Cd + SMDQ2Hd + SMDQ2Sd + SMDQ2Dd + _
        SMDQ3Bd + SMDQ3Rd + SMDQ3Cd + SMDQ3Hd + SMDQ3Sd + SMDQ3Dd + _
        SMDQ4Bd + SMDQ4Rd + SMDQ4Cd + SMDQ4Hd + SMDQ4Sd + SMDQ4Dd
'set labels
    'quarter 1
    Label5(71).Caption = SMDQ1B
    Label5(74).Caption = SMDQ1R
    Label5(77).Caption = SMDQ1C
    Label5(80).Caption = SMDQ1H
    Label5(83).Caption = SMDQ1S
    Label5(86).Caption = SMDQ1D
    Label5(72).Caption = SMDQ1Bd
    Label5(75).Caption = SMDQ1Rd
    Label5(78).Caption = SMDQ1Cd
    Label5(81).Caption = SMDQ1Hd
    Label5(84).Caption = SMDQ1Sd
    Label5(87).Caption = SMDQ1Dd
    'quarter 2
    Label5(99).Caption = SMDQ2B
    Label5(102).Caption = SMDQ2R
    Label5(105).Caption = SMDQ2C
    Label5(108).Caption = SMDQ2H
    Label5(111).Caption = SMDQ2S
    Label5(114).Caption = SMDQ2D
    Label5(100).Caption = SMDQ2Bd
    Label5(103).Caption = SMDQ2Rd
    Label5(106).Caption = SMDQ2Cd
    Label5(109).Caption = SMDQ2Hd
    Label5(112).Caption = SMDQ2Sd
    Label5(115).Caption = SMDQ2Dd
    'quarter 3
    Label5(127).Caption = SMDQ3B
    Label5(130).Caption = SMDQ3R
    Label5(133).Caption = SMDQ3C
    Label5(136).Caption = SMDQ3H
    Label5(139).Caption = SMDQ3S
    Label5(142).Caption = SMDQ3D
    Label5(128).Caption = SMDQ3Bd
    Label5(131).Caption = SMDQ3Rd
    Label5(134).Caption = SMDQ3Cd
    Label5(137).Caption = SMDQ3Hd
    Label5(140).Caption = SMDQ3Sd
    Label5(143).Caption = SMDQ3Dd
    'quarter 4
    Label5(155).Caption = SMDQ4B
    Label5(158).Caption = SMDQ4R
    Label5(161).Caption = SMDQ4C
    Label5(164).Caption = SMDQ4H
    Label5(167).Caption = SMDQ4S
    Label5(170).Caption = SMDQ4D
    Label5(156).Caption = SMDQ4Bd
    Label5(159).Caption = SMDQ4Rd
    Label5(162).Caption = SMDQ4Cd
    Label5(165).Caption = SMDQ4Hd
    Label5(168).Caption = SMDQ4Sd
    Label5(171).Caption = SMDQ4Dd
    'totals
    Label5(173).Caption = SMDTD
End Sub

Public Sub SetGroup()
'establish parameters
'colors
Dim SMG2C As Double
Dim SMG3C As Double
Dim SMG4C As Double
Dim SMG5C As Double
Dim SMG6C As Double
Dim SMG7C As Double
Dim SMG8C As Double
Dim SMG2Cd As Double
Dim SMG3Cd As Double
Dim SMG4Cd As Double
Dim SMG5Cd As Double
Dim SMG6Cd As Double
Dim SMG7Cd As Double
Dim SMG8Cd As Double
'suit
Dim SMG2S As Double
Dim SMG3S As Double
Dim SMG4S As Double
Dim SMG5S As Double
Dim SMG2Sd As Double
Dim SMG3Sd As Double
Dim SMG4Sd As Double
Dim SMG5Sd As Double
'value
Dim SMG2V As Double
Dim SMG3V As Double
Dim SMG4V As Double
Dim SMG2Vd As Double
Dim SMG3Vd As Double
Dim SMG4Vd As Double
'total
Dim SMGTD As Double
'counters
Dim pColorCounter As Integer
Dim pSuitCounter As Integer
Dim pValueCounter As Integer
'Dim pC2 As Integer
'Dim pC3 As Integer
'Dim pC4 As Integer
'Dim pC5 As Integer
'Dim pC6 As Integer
'Dim pC7 As Integer
'Dim pC8 As Integer
'Dim pS2 As Integer
'Dim pS3 As Integer
'Dim pS4 As Integer
'Dim pS5 As Integer
'Dim pV2 As Integer
'Dim pV3 As Integer
'Dim pV4 As Integer
Dim pCurrentColor As Integer
Dim pCurrentSuit As Integer
Dim pCurrentValue As Integer
'set parameters to zero
SMG2C = 0
SMG3C = 0
SMG4C = 0
SMG5C = 0
SMG6C = 0
SMG7C = 0
SMG8C = 0
SMG2Cd = 0
SMG3Cd = 0
SMG4Cd = 0
SMG5Cd = 0
SMG6Cd = 0
SMG7Cd = 0
SMG8Cd = 0
SMG2S = 0
SMG3S = 0
SMG4S = 0
SMG5S = 0
SMG2Sd = 0
SMG3Sd = 0
SMG4Sd = 0
SMG5Sd = 0
SMG2V = 0
SMG3V = 0
SMG4V = 0
SMG2Vd = 0
SMG3Vd = 0
SMG4Vd = 0
SMGTD = 0
pColorCounter = 0
pSuitCounter = 0
pValueCounter = 0
'pC2 = 0
'pC3 = 0
'pC4 = 0
'pC5 = 0
'pC6 = 0
'pC7 = 0
'pC8 = 0
'pS2 = 0
'pS3 = 0
'pS4 = 0
'pS5 = 0
'pV2 = 0
'pV3 = 0
'pV4 = 0
'initialize current values
'run groups for colors
pCurrentColor = ShuffleMeterDeck(1, 1)
For p% = 1 To 52
    If ShuffleMeterDeck(1, p%) = pCurrentColor Then
        pColorCounter = pColorCounter + 1
        If pColorCounter = 8 Then
            SMG8C = SMG8C + 1
            pColorCounter = 0
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p% + 1)
            End If
        End If
        'need to check the final group size
        If p% = 52 Then
            If pColorCounter = 7 Then
                SMG7C = SMG7C + 1
            ElseIf pColorCounter = 6 Then
                SMG6C = SMG6C + 1
            ElseIf pColorCounter = 5 Then
                SMG5C = SMG5C + 1
            ElseIf pColorCounter = 4 Then
                SMG4C = SMG4C + 1
            ElseIf pColorCounter = 3 Then
                SMG3C = SMG3C + 1
            ElseIf pColorCounter = 2 Then
                SMG2C = SMG2C + 1
            End If
        End If
    Else
        If pColorCounter = 7 Then
            SMG7C = SMG7C + 1
            pColorCounter = 1
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p%)
            End If
        ElseIf pColorCounter = 6 Then
            SMG6C = SMG6C + 1
            pColorCounter = 1
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p%)
            End If
        ElseIf pColorCounter = 5 Then
            SMG5C = SMG5C + 1
            pColorCounter = 1
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p%)
            End If
        ElseIf pColorCounter = 4 Then
            SMG4C = SMG4C + 1
            pColorCounter = 1
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p%)
            End If
        ElseIf pColorCounter = 3 Then
            SMG3C = SMG3C + 1
            pColorCounter = 1
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p%)
            End If
        ElseIf pColorCounter = 2 Then
            SMG2C = SMG2C + 1
            pColorCounter = 1
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p%)
            End If
        ElseIf pColorCounter = 1 Then
            pColorCounter = 1
            If p% < 52 Then
                pCurrentColor = ShuffleMeterDeck(1, p%)
            End If
        End If
    End If
Next p%
'run groups for suits
pCurrentSuit = ShuffleMeterDeck(2, 1)
For p% = 1 To 52
    If ShuffleMeterDeck(2, p%) = pCurrentSuit Then
        pSuitCounter = pSuitCounter + 1
        If pSuitCounter = 5 Then
            SMG5S = SMG5S + 1
            pSuitCounter = 0
            If p% < 52 Then
                pCurrentSuit = ShuffleMeterDeck(2, p% + 1)
            End If
        End If
        'need to check the final group size
        If p% = 52 Then
            If pSuitCounter = 4 Then
                SMG4S = SMG4S + 1
            ElseIf pSuitCounter = 3 Then
                SMG3S = SMG3S + 1
            ElseIf pSuitCounter = 2 Then
                SMG2S = SMG2S + 1
            End If
        End If
    Else
        If pSuitCounter = 4 Then
            SMG4S = SMG4S + 1
            pSuitCounter = 1
            If p% < 52 Then
                pCurrentSuit = ShuffleMeterDeck(2, p%)
            End If
        ElseIf pSuitCounter = 3 Then
            SMG3S = SMG3S + 1
            pSuitCounter = 1
            If p% < 52 Then
                pCurrentSuit = ShuffleMeterDeck(2, p%)
            End If
        ElseIf pSuitCounter = 2 Then
            SMG2S = SMG2S + 1
            pSuitCounter = 1
            If p% < 52 Then
                pCurrentSuit = ShuffleMeterDeck(2, p%)
            End If
        ElseIf pSuitCounter = 1 Then
            pSuitCounter = 1
            If p% < 52 Then
                pCurrentSuit = ShuffleMeterDeck(2, p%)
            End If
        End If
    End If
Next p%
'run groups for values
pCurrentValue = ShuffleMeterDeck(3, 1)
For p% = 1 To 52
    If ShuffleMeterDeck(3, p%) = pCurrentValue Then
        pValueCounter = pValueCounter + 1
        If pValueCounter = 4 Then
            SMG4V = SMG4V + 1
            pValueCounter = 0
            If p% < 52 Then
                pCurrentValue = ShuffleMeterDeck(3, p% + 1)
            End If
        End If
        'need to check the final group size
        If p% = 52 Then
            If pValueCounter = 3 Then
                SMG3V = SMG3V + 1
            ElseIf pValueCounter = 2 Then
                SMG2V = SMG2V + 1
            End If
        End If
    Else
        If pValueCounter = 3 Then
            SMG3V = SMG3V + 1
            pValueCounter = 1
            If p% < 52 Then
                pCurrentValue = ShuffleMeterDeck(3, p%)
            End If
        ElseIf pValueCounter = 2 Then
            SMG2V = SMG2V + 1
            pValueCounter = 1
            If p% < 52 Then
                pCurrentValue = ShuffleMeterDeck(3, p%)
            End If
        ElseIf pValueCounter = 1 Then
            pValueCounter = 1
            If p% < 52 Then
                pCurrentValue = ShuffleMeterDeck(3, p%)
            End If
        End If
    End If
Next p%
'calculate differences
    'colors
    SMG2Cd = Abs(6.6 - SMG2C)
    SMG3Cd = Abs(3.4 - SMG3C)
    SMG4Cd = Abs(1.6 - SMG4C)
    SMG5Cd = Abs(0.8 - SMG5C)
    SMG6Cd = Abs(0.4 - SMG6C)
    SMG7Cd = Abs(0.2 - SMG7C)
    SMG8Cd = Abs(0.1 - SMG8C)
    'suits
    SMG2Sd = Abs(7.1 - SMG2S)
    SMG3Sd = Abs(1.9 - SMG3S)
    SMG4Sd = Abs(0.5 - SMG4S)
    SMG5Sd = Abs(0.1 - SMG5S)
    'values
    SMG2Vd = Abs(2.8 - SMG2V)
    SMG3Vd = Abs(0.1 - SMG3V)
    SMG4Vd = Abs(0 - SMG4V)
    'total differences
    SMGTD = SMG2Cd + SMG3Cd + SMG4Cd + _
        SMG5Cd + SMG6Cd + SMG7Cd + SMG8Cd + _
        SMG2Sd + SMG3Sd + SMG4Sd + SMG5Sd + _
        SMG2Vd + SMG3Vd + SMG4Vd
'set labels
    'colors
    Label5(232).Caption = SMG2C
    Label5(229).Caption = SMG3C
    Label5(226).Caption = SMG4C
    Label5(223).Caption = SMG5C
    Label5(220).Caption = SMG6C
    Label5(217).Caption = SMG7C
    Label5(193).Caption = SMG8C
    Label5(231).Caption = SMG2Cd
    Label5(228).Caption = SMG3Cd
    Label5(225).Caption = SMG4Cd
    Label5(222).Caption = SMG5Cd
    Label5(219).Caption = SMG6Cd
    Label5(216).Caption = SMG7Cd
    Label5(192).Caption = SMG8Cd
    'suits
    Label5(206).Caption = SMG2S
    Label5(203).Caption = SMG3S
    Label5(200).Caption = SMG4S
    Label5(197).Caption = SMG5S
    Label5(205).Caption = SMG2Sd
    Label5(202).Caption = SMG3Sd
    Label5(199).Caption = SMG4Sd
    Label5(196).Caption = SMG5Sd
    'values
    Label5(183).Caption = SMG2V
    Label5(180).Caption = SMG3V
    Label5(177).Caption = SMG4V
    Label5(182).Caption = SMG2Vd
    Label5(179).Caption = SMG3Vd
    Label5(176).Caption = SMG4Vd
    'total differences
    Label5(174).Caption = SMGTD
End Sub

Public Sub SetCycles()
Dim pCycle(26) As Integer
Dim pCycleExists As Boolean
Dim pColorCycles As Integer
Dim pSuitCycles As Integer
Dim pValueCycles As Integer
'CHECK COLOR CYCLES
    'load pCycle(x)
    For i% = 1 To 26
        pCycle(i%) = ShuffleMeterDeck(1, i%)
    Next i%
    'start with default of 1 for pColorCycles
        'this assumes a single full deck with just one unique cycle
    pColorCycles = 1
    'check 26 card cycles
    'assume cycle exists until proven otherwise
    pCycleExists = True
    For p% = 1 To 2
        For q% = 1 To 26
            If pCycle(q%) <> ShuffleMeterDeck(1, (p% - 1) * 26 + q%) Then
                pCycleExists = False
            End If
        Next q%
    Next p%
    'if 26 card cycle exists, set the pColorCycles value to 2
    If pCycleExists Then
        pColorCycles = 2
    End If
    'check 4 card cycles
    'assume cycle exists until proven otherwise
    pCycleExists = True
    For p% = 1 To 13
        For q% = 1 To 4
            If pCycle(q%) <> ShuffleMeterDeck(1, (p% - 1) * 4 + q%) Then
                pCycleExists = False
            End If
        Next q%
    Next p%
    'if 4 card cycle exists, set the pColorCycles value to 13
    If pCycleExists Then
        pColorCycles = 13
    End If
    'check 2 card cycles
    'assume cycle exists until proven otherwise
    pCycleExists = True
    For p% = 1 To 26
        For q% = 1 To 2
            If pCycle(q%) <> ShuffleMeterDeck(1, (p% - 1) * 2 + q%) Then
                pCycleExists = False
            End If
        Next q%
    Next p%
    'if 2 card cycle exists, set the pColorCycles value to 26
    If pCycleExists Then
        pColorCycles = 26
    End If
'CHECK SUIT CYCLES
    'load pCycle(x)
    For i% = 1 To 26
        pCycle(i%) = ShuffleMeterDeck(2, i%)
    Next i%
    'start with default of 1 for pSuitCycles
        'this assumes a single full deck with just one unique cycle
    pSuitCycles = 1
    'check 26 card cycles
    'assume cycle exists until proven otherwise
    pCycleExists = True
    For p% = 1 To 2
        For q% = 1 To 26
            If pCycle(q%) <> ShuffleMeterDeck(2, (p% - 1) * 26 + q%) Then
                pCycleExists = False
            End If
        Next q%
    Next p%
    'if 26 card cycle exists, set the pColorCycles value to 2
    If pCycleExists Then
        pSuitCycles = 2
    End If
    'check 4 card cycles
    'assume cycle exists until proven otherwise
    pCycleExists = True
    For p% = 1 To 13
        For q% = 1 To 4
            If pCycle(q%) <> ShuffleMeterDeck(2, (p% - 1) * 4 + q%) Then
                pCycleExists = False
            End If
        Next q%
    Next p%
    'if 4 card cycle exists, set the pSuitCycles value to 13
    If pCycleExists Then
        pSuitCycles = 13
    End If
'CHECK VALUE CYCLES
    'load pCycle(x)
    For i% = 1 To 26
        pCycle(i%) = ShuffleMeterDeck(3, i%)
    Next i%
    'start with default of 1 for pValueCycles
        'this assumes a single full deck with just one unique cycle
    pValueCycles = 1
    'check 26 card cycles
    'assume cycle exists until proven otherwise
    pCycleExists = True
    For p% = 1 To 2
        For q% = 1 To 26
            If pCycle(q%) <> ShuffleMeterDeck(3, (p% - 1) * 26 + q%) Then
                pCycleExists = False
            End If
        Next q%
    Next p%
    'if 26 card cycle exists, set the pColorCycles value to 2
    If pCycleExists Then
        pValueCycles = 2
    End If
    'check 13 card cycles
    'assume cycle exists until proven otherwise
    pCycleExists = True
    For p% = 1 To 4
        For q% = 1 To 13
            If pCycle(q%) <> ShuffleMeterDeck(3, (p% - 1) * 13 + q%) Then
                pCycleExists = False
            End If
        Next q%
    Next p%
    'if 13 card cycle exists, set the pValueCycles value to 4
    If pCycleExists Then
        pValueCycles = 4
    End If
'set Labels
Label4(1).Caption = pColorCycles
Label4(2).Caption = pSuitCycles
Label4(3).Caption = pValueCycles
End Sub

Private Sub Form_Activate()
SetShuffleMeterParameters
End Sub

Private Sub Form_Load()
SetShuffleMeterParameters
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.mnuShuffleMeter.Checked = False
End Sub

Public Sub SetShuffleMeterParameters()
SetShuffleMeterDeck
SetPermutations
SetDistribution
SetGroup
SetBreak
SetSpread
SetCycles
End Sub

