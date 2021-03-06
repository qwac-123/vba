Attribute VB_Name = "poker"
Dim hand(7, 2) '0-7六个玩家
Dim playerMoney(7)
Dim deck(52)    '牌,如sT,d8
Dim shuffleDeck(52) '洗好的牌,是牌的编号，51=sA,50=sK,0=d2
Dim zongpai(7) '手牌+公共牌,编号
Dim communityCards(5)  'cards on table
Dim pattern(7) '每个人最后的牌型，0是high,用于内部比大小
Dim patternDeck(7) '每个人最后的牌型，用于显示给人看
Dim maxPat(7) '最大的牌型
Dim maxPatDeck(7)
Dim maxPlayer(7) '最大的几个玩家
Dim playerFold(7) 'default false
Dim playerQuit(7) 'def false

Dim playerPosition(7)

Dim rngPosition

Dim rngHand  '手牌格子
Dim rngMoney '手头的钱所在的格子
Dim rngPot
Dim rngFlop As Range
Dim rngTurn As Range
Dim rngRiver As Range

Dim callChips(7) '跟注花费


Dim finalSeven As String
Dim deckCard As Integer    'dealed cards发出去的牌的张数
Dim player As Integer
Dim prevPlayer As Integer
Dim gamePlayed As Integer '玩了几次
Dim stage As Integer 'preflop=0,flop=1,etc
Dim SB As Integer
Dim BB As Integer
Dim betTurn As Integer
Dim pot 'dichi
Function TexasHoldem()
'v0.3 171120可以比大小了
'v0.2 171116 现在可以判定同花
'v0.1 洗牌发牌理顺最后牌组

'Range("a1:ak9").Clear
'Range("a12:ak15").Clear

If gamePlayed < 1 Then


    Call preparecells
    Call preparePlayers
    Call prepareDeck
    BB = 10
    SB = 5
    betTurn = 0
    pot = 0
    gamePlayed = 1
End If
Range(rngPosition(gamePlayed - 1)) = "D"
Range(rngPosition(gamePlayed)) = "SB"
Range(rngPosition(gamePlayed + 1)) = "BB"
For player = 0 To 7
    If playerMoney(player) <= 0 Then
        playerQuit(player) = True
    End If
Next player

Call shuffle    '洗牌
Call deal       '发牌
stage = 0 'preflop
    Call raiseCallFold
Call flop
stage = 1 'flop
    Call raiseCallFold
Call turn
stage = 2 'turn
    Call raiseCallFold
Call river 'river
stage = 3
    Call raiseCallFold
Call compareCards
gamePlayed = gamePlayed + 1
End Function
Function preparecells()
'准备表
    Cells.ColumnWidth = 3.13
    Cells.RowHeight = 14.25
    Set rngFlop = Range("T12:V12")
    Set rngTurn = Range("w12")
    Set rngRiver = Range("x12")
    'N=14,V=22,AD=30,CODROW1=4,R2=12,R3=20
    Range("t10:u10").Merge
    Set rngPot = Range("t10:u10")
    rngPot.Borders.LineStyle = xlContinuous
    Range("r10:s10").Merge
    Range("r10:s10").Value = "pot:"
    
    Range("T12:X12").NumberFormatLocal = ";;;"
    Range("M12:N12,N4:O4,V4:W4,AD4:AE4,AE12:AF12,AD20:AE20,V20:W20,N20:O20,T12:X12" _
        ).NumberFormatLocal = ";;;"
    Range("M12:N12,N4:O4,V4:W4,AD4:AE4,AE12:AF12,AD20:AE20,V20:W20,N20:O20,T12:X12" _
        ).Value = ""
    i = 0
    For Each Rng In Range("T21,L21,K13,L5,T5,AB5,AC13,AB21")
        Rng.Value = "P" & i
        i = i + 1
    Next Rng
    rngPosition = Array("s21", "k21", "j13", "k5", "s5", "aa5", "Ab13", "Aa21")
    Range("L6:M6,K14:L14,L22:M22,T22:U22,AB22:AC22,AC14:AD14,AB6:AC6,T6:U6").Merge
    Range("L6:M6,K14:L14,L22:M22,T22:U22,AB22:AC22,AC14:AD14,AB6:AC6,T6:U6") = "chips:"
    
    Range("N6:O6,M14:N14,N22:O22,V22:W22,AD22:AE22,AE14:AF14,AD6:AE6,V6:W6").Merge 'rngMoney
    With Range("N5:O5,V5:W5,AD5:AE5,AE13:AF13,AD21:AE21,V21:W21,N21:O21,M13:N13,T13:X13").Borders '给人看的部分
        .LineStyle = xlContinuous
    End With
End Function
Function preparePlayers()
rngHand = [{"v20","w20";"n20","o20";"m12","n12";"n4","o4";"v4","w4";"ad4","ae4";"ae12","af12";"ad20","ae20"}]
rngMoney = Array("V22", "N22", "M14", "N6", "V6", "AD6", "AE14", "AD22")
For player = 0 To 7
    playerMoney(player) = 1000
    Range(rngMoney(player)) = playerMoney(player)
    playerQuit(player) = False
    playerFold(player) = False
Next player
End Function
Function prepareDeck()
Dim suit
Dim rank
suit = Array("d", "c", "h", "s")
rank = Array(2, 3, 4, 5, 6, 7, 8, 9, "T", "J", "Q", "K", "A")
k = 0
For i = 0 To 3
    For j = 0 To 12
        deck(k) = suit(i) & rank(j)
        k = k + 1
    Next j
Next i
End Function
Function shuffle()
Dim wushier(52)
Dim k
'Dim a
'For a = 0 To 5
For k = 0 To 51
    wushier(k) = k   '一个00-51的数组
    shuffleDeck(k) = 0
Next k
uBond = 52
lBond = 0
For k = 0 To 51  '循环产生52个不重复随机数，做成一个数组shuffleDeck
    If uBond > 0 Then
        paixu = Int(Rnd(Timer) * uBond)   '[0,uBond-1] 之间随机整数
        shuffleDeck(k) = wushier(paixu) '随机到的值一个个放到新数组里
        wushier(paixu) = wushier(uBond - 1) '把最后一个值挪到随机到的位置
        uBond = uBond - 1                  '舍弃掉最后一个数
    End If
Next k
'Next a
End Function
Function raiseCallFold() '施工现场

Do
For i = 0 To 7
    player = (i + gamePlayed) Mod 8
    prevPlayer = (i + gamePlayed - 1) Mod 8
    If stage = 0 And betTurn = 0 Then 'preflop,start
        Select Case player
        Case 1 '小盲注位
            callChips(player) = SB
        Case 2 '大盲注位
            callChips(player) = BB
        Case Else
            If player = 0 Then '是人
                Call youBet
            Else
                Call aiBet
            End If
        End Select
    Else
        If player = 0 Then 'you
            Call youBet
        Else
            Call aiBet
        End If
    End If
        pot = pot + callChips(player)
        playerMoney(player) = playerMoney(player) - callChips(player)
        Range(rngMoney(player)).Value = playerMoney(player)
Next i
rngPot.Value = pot
betTurn = betTurn + 1
Loop While callChips(prevPlayer) < callChips(player)
For i = 0 To 7
    callChips(i) = 0
Next i
End Function
Function aiBet()
'v0 傻根
callChips(player) = callChips(prevPlayer)
End Function
Function youBet()
yourOPT:
yourOption = InputBox("输入你的选项序号：" & vbCrLf & "0.过牌" & vbCrLf & "1.跟注 " & callChips(player) & vbCrLf & "2.加注" & vbCrLf & "3.弃牌", "过牌，跟牌，加注或者弃牌")
    Select Case yourOption
        Case 0 'check
            If callChips(prevPlayer) > callChips(player) Then
                MsgBox "不能过牌，你的注小"
                GoTo yourOPT
            End If
        Case 1 'call
            callChips(player) = callChips(prevPlayer)
            playerMoney(player) = playerMoney(player) - callChips(player)
            Range(rngMoney(player)).Value = playerMoney(player)
        Case 2 'raise
            raiseChips = callChips(prevPlayer) + InputBox("raise?", "加注", 2 * callChips(player))
            callChips(player) = raiseChips
        Case 3 'fold
            playerFold(0) = True
    End Select
    
End Function
Function deal()
deckCard = 0
For handCard = 0 To 1   'preflop发牌
    For player = 0 To 7 '目前还是支持7个人
        If playerQuit(player) Then
            Exit For
        End If
        hand(player, handCard) = shuffleDeck(deckCard)
        Range(rngHand(player + 1, handCard + 1)) = deck(shuffleDeck(deckCard))
        deckCard = deckCard + 1
    Next player
Next handCard
End Function
Function flop()
i = 0
For Each Rng In rngFlop
    Rng.Value = deck(shuffleDeck(deckCard))
    communityCards(i) = shuffleDeck(deckCard)
    deckCard = deckCard + 1
    i = i + 1
Next
End Function
Function turn()
rngTurn = deck(shuffleDeck(deckCard))
communityCards(3) = shuffleDeck(deckCard)
deckCard = deckCard + 1
End Function
Function river()
rngRiver = deck(shuffleDeck(deckCard))
communityCards(4) = shuffleDeck(deckCard)
End Function
Function compareCards()
countSameMax = 0
For i = 0 To 7
    maxPat(i) = ""
Next i
For player = 0 To 7
    If playerFold(player) = "fold" Then
        Exit For
    End If
    i = 0
    For handCard = 0 To 1
        zongpai(i) = hand(player, handCard)
        i = i + 1
    Next handCard
    For tableCard = 0 To 4
        zongpai(i) = communityCards(tableCard)
        i = i + 1
    Next tableCard
    '测试部分'ceshi'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'zongpai(0) = 38
'zongpai(1) = 49
'zongpai(2) = 48
'zongpai(3) = 51
'zongpai(4) = 46
'zongpai(5) = 30
'zongpai(6) = 45

    'end测试部分'ceshi'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call insertSort
    Call rules '比量牌组
    Debug.Print patternDeck(player) & " " & deck(hand(player, 0)) & " " & deck(hand(player, 1))
    Debug.Print "------"
    If pattern(player) > maxPat(countSameMax) Then '选出最大的牌型和玩家
        countSameMax = 0
        maxPat(countSameMax) = pattern(player)
        maxPatDeck(countSameMax) = patternDeck(player)
        maxPlayer(countSameMax) = player
    ElseIf pattern(player) = maxPat(countSameMax) Then '如果有相同大小牌型，则增加分底池的玩家
        countSameMax = countSameMax + 1
        maxPat(countSameMax) = pattern(player)
        maxPatDeck(countSameMax) = patternDeck(player)
        maxPlayer(countSameMax) = player
    End If
'    Exit For '测试部分，待删'ceshi'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next player

earnChips = pot / (countSameMax + 1)
For i = 0 To countSameMax
    player = maxPlayer(i)
    a = MsgBox("winner is :   " & maxPlayer(i) & vbCrLf & _
                "pattern is :   " & maxPatDeck(i) & vbCrLf & _
                player & " win  $" & earnChips, vbOKOnly, "Winner")
    
    playerMoney(player) = playerMoney(player) + earnChips
    rngMoney(player) = playerMoney(player)
    Debug.Print "第 " & gamePlayed & " 轮"
    Debug.Print "得奖人数:  " & countSameMax + 1
    Debug.Print "最大牌型:  " & maxPatDeck(i)
    Debug.Print "公共牌:   " & deck(communityCards(0)) & deck(communityCards(1)) & deck(communityCards(2)) & _
    deck(communityCards(3)) & deck(communityCards(4))
    Debug.Print "赢家:  " & player, deck(hand(player, 0)), deck(hand(player, 1))
    Debug.Print "------------------------------------"
'    Exit For 'ceshi'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next i
pot = 0
rngPot.Value = ""
earnChips = 0
End Function
Function zctex()
For i = 0 To 3
Call TexasHoldem
Next i
End Function
Function insertSort()
'v1.0   把牌按照点数从大到小，从黑桃到方片排序
zok = ""
zokd = ""
For i = 0 To 5
    For j = i + 1 To 6
        rankZj = zongpai(j) Mod 13
        suitZj = Int(zongpai(j) / 13)
        rankZi = zongpai(i) Mod 13
        suitZi = Int(zongpai(i) / 13)
        If rankZj > rankZi Then '先比牌的点数
            temp = zongpai(j)
            zongpai(j) = zongpai(i)
            zongpai(i) = temp
        ElseIf rankZj = rankZi And suitZj > suitZi Then '相同就比花色
            temp = zongpai(j)
            zongpai(j) = zongpai(i)
            zongpai(i) = temp
        End If
    Next j

Next i
'ceshi'''''''''''''''''''''''''''''''''''''''
For i = 0 To 6
    zok = zok & " " & zongpai(i)
    zokd = zokd & " " & deck(zongpai(i))
    Next i
Debug.Print "player: " & player
Debug.Print zok
Debug.Print zokd 'ceshi'''''''''''''''''''''''''''''''''''''''
End Function

Function rules()
'v0.2   在判断牌点数rank时，保证每个牌是两位数，用字符串表示，00是2,12是A
'       重新改写了if树，现在能够判断同花，顺，985
'v0.1 判断同花顺（失败
'chr(48)="0"

Dim namePat 'name
Dim countSuits(3) '统计这7张牌里面各个花色的数量
Dim flushCards(3) '四个花色，把牌连起来
Dim sameCards(6) '统计这7张牌里点数相同的牌数
Dim sameDeck(6)
Dim rankZ(6)
Dim rankDeck(6)
Dim suitZ(6)
Dim suitDeck(6)
Dim typePat(9) '0-9十个统计能组成的牌型，内容是rankZ
Dim deckPat(9) '用牌id组成的牌型,zongpai
'patternDeck(player)
'pattern(player)
Dim rnkSev As String
Dim sutSev As String
rnkSev = ""
sutSev = ""

rulHe:
isFlush = False
strFlush = ""
strHigh = "0"
namePat = Array("high card", "one pair", "two pairs", "3 of a kind", "straight", "flush", "full house", "4 of a kind", "straight flush", "royal flush")
For i = 0 To 9 '开头定义区
    typePat(i) = "" '牌型号开头，跟着5张牌，11位字符，除了9和8之外的牌型和牌，在最后选最大的牌型
    If i < 7 Then
        rankZ(i) = "" & zongpai(i) Mod 13 '牌组中第i张牌的点数,00是2,，08是T，12是A
        rankDeck(i) = deck(zongpai(i)) '按点数排序的牌的牌面，deck
        If rankZ(i) < 10 Then '，确保两位数
            rankZ(i) = "0" & rankZ(i)
        End If
        suitZ(i) = Int(zongpai(i) / 13) '牌组中第i张牌的花色，0是diamonds，3是spades
'        Cells(19, i + 3) = rankZ(i)
'        Cells(20, i + 3) = suitZ(i)
'        rnkSev = rnkSev & rankZ(i)
'        sutSev = sutSev & suitZ(i)
        countSuits(suitZ(i)) = countSuits(suitZ(i)) + 1 '统计这七张牌各种花色(0-3数组),每个组里是有几张牌int
         '后面的还是得要，因为后面可能成小的同花顺 '后面的不要了，因为牌是按点数排好序的，后面的小牌可以不管
        flushCards(suitZ(i)) = flushCards(suitZ(i)) & rankZ(i) '整理同花的牌
        suitDeck(suitZ(i)) = suitDeck(suitZ(i)) & deck(zongpai(i)) '整理同花的牌的牌面deck
        If Len(flushCards(suitZ(i))) = 10 Then '如果同花牌有5张了
            typePat(5) = "5" & flushCards(suitZ(i)) '只是同花
            deckPat(5) = suitDeck(suitZ(i)) & Space(5) & namePat(5)
        End If
        sameCards(i) = rankZ(i) '相同牌组，用于最后组牌型
        sameDeck(i) = deck(zongpai(i))
    End If
    If i < 5 Then
        strHigh = strHigh & rankZ(i) '0high
        deckHigh = deckHigh & deck(zongpai(i))
    End If
    If i < 4 Then
        flushCards(i) = ""  '准备同花牌牌组
        countSuits(i) = 0 '统计这七张牌各种花色(0-3数组),每个组里是有几张牌int
    End If
Next i
typePat(0) = strHigh '& "    high"
deckPat(0) = deckHigh '
sameNum = 0 '相同牌里第一张的位置
endSameCard = 0 '用于跳过相同的牌
straightSuits = 0 '用于判断顺子的花色
straightFlush = 0

countStraightFlush = 0 '计数同花顺的个数
countStraight = 0 '计算顺子的个数
straightCards = "" '收集顺子牌，如kqjt9,1110090807,10位字符
straightDecks = "" '顺子以牌面收集，如 sTd9h8c7d6
pattern(player) = "" '最后的牌型，牌型号开头，跟着5张牌，11位字符，如皇家同花顺就是91211100908


For i = 0 To 5 '组牌型
    j = i + 1
    If i < endSameCard - 1 Then
        GoTo nexti
    End If
    If rankZ(i) - rankZ(j) = 1 Then   '判断是否顺子而且在范围内
        countStraight = countStraight + 1 'i牌的顺子次数
        straightCards = straightCards & rankZ(i) '收集顺子的牌
        ifSame = i
        If i = endSameCard - 1 Then
            ifSame = sameNum
        End If
        straightDecks = straightDecks & rankDeck(ifSame)
        straightSuits = straightSuits + Abs(suitZ(i) - suitZ(j)) '统计花色，如果都是同花色，=0
        If countStraight = 4 And straightSuits = 0 And Left(straightCards, 2) = 12 Then   '4rankZi+1rankZj 5张连接张，是A的同花顺，9royal
            pattern(player) = 9 & straightCards & rankZ(j) '& "    royal flush"
            patternDeck(player) = straightDecks & rankDeck(j) & Space(5) & namePat(9)
            Exit Function
        ElseIf countStraight = 4 And straightSuits = 0 Then '普通同花顺,8straight flush
            pattern(player) = 8 & straightCards & rankZ(j) '& "    straight flush"
            patternDeck(player) = straightDecks & rankDeck(j) & Space(5) & namePat(8)
            Exit Function
        ElseIf countStraight = 4 Then
            pattern(player) = 4 & straightCards & rankZ(j) '& "    straight" '只是顺子,或者3 of a kind"
            patternDeck(player) = straightDecks & rankDeck(j) & Space(5) & namePat(4)
            Exit Function
        End If
    ElseIf rankZ(i) - rankZ(j) = 0 Then '判断是否相同点数牌
        sameNum = i
        For j = j To 6
            If rankZ(i) - rankZ(j) = 0 Then
                sameCards(i) = sameCards(i) & rankZ(j)
                sameDeck(i) = sameDeck(i) & rankDeck(j)
                sameCards(j) = ""
                sameDeck(j) = ""
            Else
                Exit For
            End If
        Next j
        If j > 6 Then
            j = 6
        End If
        endSameCard = j
        suitZ(j - 1) = suitZ(i) 'j是连续相同牌后面第一张不同的牌，最后一张相同牌是j-1
    Else '有间断点
        countStraight = 0
        straightCards = ""
    End If
nexti:
Next i

For i = 0 To 6
    For j = i + 1 To 6
        If Len(sameCards(j)) > Len(sameCards(i)) Then
            temp = sameCards(i)
            sameCards(i) = sameCards(j)
            sameCards(j) = temp
            
            tempD = sameDeck(i)
            sameDeck(i) = sameDeck(j)
            sameDeck(j) = tempD
        ElseIf Len(sameCards(j)) = Len(sameCards(i)) And sameCards(j) > sameCards(i) Then
            temp = sameCards(i)
            sameCards(i) = sameCards(j)
            sameCards(j) = temp
            tempD = sameDeck(i)
            sameDeck(i) = sameDeck(j)
            sameDeck(j) = tempD
        End If
    Next j
Next i

Select Case Len(sameCards(0))
    Case 8
        typePat(7) = 7 & sameCards(0) & Left(sameCards(1), 2) ' & "    4 of a kind" '74 of a kind:       AAAAx
        deckPat(7) = sameDeck(0) & Left(sameDeck(1), 2) & Space(5) & namePat(7)
        pattern(player) = typePat(7)
        patternDeck(player) = deckPat(7)
        Exit Function
    Case 6
        If Len(sameCards(1)) > 3 Then
            typePat(6) = 6 & sameCards(0) & Left(sameCards(1), 4) ' & "    full house" ':        AAAKK
            deckPat(6) = sameDeck(0) & Left(sameDeck(1), 4) & Space(5) & namePat(6)
            pattern(player) = typePat(6)
            patternDeck(player) = deckPat(6)
            Exit Function
        ElseIf Len(typePat(5)) > 10 Then
            pattern(player) = typePat(5) '只是同花，前面写了
            'deckPat(5) = suitDeck(suitZ(i)) & Space(5) & namePat(5)前面已经写了
            patternDeck(player) = deckPat(5)
            Exit Function
        Else
            typePat(3) = 3 & sameCards(0) & sameCards(1) & sameCards(2) '& "    3 of a kind" ':       AAAxx
            deckPat(3) = sameDeck(0) & sameDeck(1) & sameDeck(2) & Space(5) & namePat(3)
            pattern(player) = typePat(3)
            patternDeck(player) = deckPat(3)
            Exit Function
        End If
    Case 4
        If Len(sameCards(1)) > 3 Then
            typePat(2) = 2 & sameCards(0) & sameCards(1) & Left(sameCards(2), 2) '& "    two pairs" '2          AAKKx
            deckPat(2) = sameDeck(0) & sameDeck(1) & Left(sameDeck(2), 2) & Space(5) & namePat(2)
            pattern(player) = typePat(2)
            patternDeck(player) = deckPat(2)
            Exit Function
        Else
            typePat(1) = 1 & sameCards(0) & sameCards(1) & sameCards(2) & sameCards(3) '& "    one pair" '1           AAxxx
            deckPat(1) = sameDeck(0) & sameDeck(1) & sameDeck(2) & sameDeck(3) & Space(5) & namePat(1)
            pattern(player) = typePat(1)
            patternDeck(player) = deckPat(1)
            Exit Function
        End If
End Select
pattern(player) = typePat(0)
patternDeck(player) = deckPat(0) & Space(5) & namePat(0)
    
'pattern(player) = Application.WorksheetFunction.Max(typePat)
'final:
'9royal flush:       TJQKAs
'8straight flush:    56789s
'74 of a kind:       AAAAx
'6full house:        AAAKK
'5flush:             479TKs
'4straight:          56789o  'bicycle:A2345  'broadway:TJQKA
'33 of a kind:       AAAxx
'2two pairs          AAKKx
'1one pair           AAxxx
'0high card          xxxxx

End Function
Function ceee()
Dim beu(5)
bii = 51
i = 1
For i = i To 5
'beu(i) = i
Debug.Print Rnd
Next
'beu(4) = "0" & beu(3)
'Debug.Print beu(4), 4
'Debug.Print beu(3), 3
'Application.WorksheetFunction.Max(beu)
End Function
Function helpme()
Application.ScreenUpdating = 1
Application.DisplayAlerts = 1

'玩家位置
'1. Button--庄家位置，也被称作按钮位
'线上游戏中第一局庄家位置由系统随机指定，线下游戏时可以大家抽牌决定，抽到最大牌的人的做第一局的庄家，以后每局庄家位置按照顺时针方向下移一位。
'2. Big Blind--大盲注，简称BB
'庄家左手数起第二位即为大盲注，牌局开始前需固定下注的位置，一般下注额为当前牌桌底注。
'3. Small Blind--小盲注，简称SB
'庄家左手数起第一位即为小盲注，也是牌局开始前需固定下注的位置，一般下注额为大盲注的一半。
'4. Under the Gun--枪口位，简称UTG
'大盲注左手数起第一位即为枪口位，枪口位的位置相对来说比较被动，往往会被迫弃牌。
'5. Cutoff--关煞位，庄家右边的位置。
'牌局操作
'Action --叫注?说话
'德州扑克里共有七种操作:
'1. Bet--押注：押上筹码。
'2. Call--跟进 / 跟注：跟随众人押上同等的注额。
'3. Fold--弃牌 / 不跟：放弃继续牌局的机会。
'4. Check--让牌 / 看牌：在无需跟进的情况下选择把决定“让”给下一位。
'5. Raise--加注：把现有的注金抬高。
'6. Re-raise--再加注：在别人加注以后回过来再加注。
'7. All-in--全下：一次过把手上的筹码全押上。
'四轮下注
'1. Pre-flop--翻牌前
'公共牌出现以前的第一轮叫注?
'2. Flop--翻牌，首三张公用牌。Flop round--翻牌圈：首三张公共牌出现以后的押注圈。
'3. Turn--转牌，第四张公共牌。Turn round--转牌圈： 第四张公共牌出现以后的押注圈 。
'4. River--河牌，第五张公共牌。River round--河牌圈：第五张公共牌出现以后 , 也即是摊牌以前的押注圈 。
'四种花色
'1. H(Heart)--红桃：在扑克牌里是爱情的象征
'2. S(Spade)--黑桃：在扑克牌里是权力的象征
'3. D(Diamond)--方块：在扑克牌里是财富的象征
'4. C(Club)--草花：在扑克牌里是幸运的象征
'各种牌型
'各种极品美
'五张牌组合由大至小依次为:
'
'suited --同一花色: 比如AKs 表示AK同一花色
'off suit - -不同花色: 比如AKo 表示AK不同花色
'Set--暗三条：比如你3-3 翻牌 A-3-4 你就是一个Set
'Bicycle --最小的顺子: a -2 - 3 - 4 - 5
'Broadway--10到A的顺子
'Connectors--连牌：比如 9-10、10-J这样的起手牌
'Draw hand - -听牌: 多为凑同花和凑顺子的牌 比如黑桃10 - J这样的起手牌
'Open-ended Straight--两端开口顺子：比如你手牌Q-K，台面是10-J-3
'Pocket pair--口袋对子：比如2-2、3-3、4-4这样的起手牌
'American Airlines - -AA: 一对A的起手牌
'Cowboys --KK: 一对K的起手牌
'Rainbow -彩虹面: 指的是翻牌三张不同花色的情况
'Nuts--坚果：比如你手牌A-A，台面 A-A-6-J-8，你的四条最大，就叫nuts
'其他
'Pot--底池：每一个牌局里众人已押上的筹码总额，也即该局的奖金数目。
'Outs--出路： 一个玩家在某个阶段所可能获胜的几种方法。比如一个拥有一对口袋9的玩家需要多一张9来取胜,他的就有两条“出路”(剩下的两个花色的9)。
'Bluff--诈唬：在没有什么胜算的情况下押上很多筹码，虚张声势。
'Slowplay--慢玩：比如坚果不下注，钓鱼的意思。
'Heads-up--单挑，缩写HU
'Showdown--摊牌比大小：双方都不肯弃牌，只好比大小。
'free card--免费牌：指无人下注，免费看一张牌。
'Fish--鱼：一般高水平的玩家对那些输不起，牌品差的玩家的贬意称呼。
'Shark --鲨鱼: 一般指能够赢钱的高手?
End Function


