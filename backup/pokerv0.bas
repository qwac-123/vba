Attribute VB_Name = "poker"
Dim hand(5, 2)
Dim deck(52)
Dim suit
Dim rank
Dim shuffleDeck(52) '洗好的牌
Dim deckCard    'dealed cards
Dim communityCards(5)  'cards on table
Dim rngMoney As Object
Dim rngFlops As Object
Dim rngTurn As Object
Dim rngRiver As Object
Function TexasHoldem()

Cells(16, 3) = "chips:"
Set rngMoney = Cells(16, 4)
Set rngFlops = Range("c9:e9")
Set rngTurn = Cells(9, 6)
Set rngRiver = Cells(9, 7)
Call prepareDeck
rngMoney = "$1000"
Call shuffle    '洗牌
Call deal       '发牌
'preflop

Call flop

Call turn

Call river

Call compareCards
End Function
Function prepareDeck()
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
For k = 0 To 51
    wushier(k) = k '一个0-51的数组
Next k
uBond = 52
lBond = 0
For k = 0 To 51  '循环产生52个不重复随机数，做成一个数组shuffleDeck
    If uBond > 0 Then
        paixu = Int(Rnd * uBond)     '[0,uBond-1] 之间随机整数
        shuffleDeck(k) = wushier(paixu) '随机到的值一个个放到新数组里
        wushier(paixu) = wushier(uBond - 1) '把最后一个值挪到随机到的位置
        uBond = uBond - 1                  '舍弃掉最后一个数
    End If
Next k
End Function
Function deal()
deckCard = 0
For player = 0 To 4 '目前还是支持五个人
    For handCard = 0 To 1   'preflop发牌
        hand(player, handCard) = shuffleDeck(deckCard)
        deckCard = deckCard + 1
    Next handCard
Next player
Range("d15") = deck(hand(0, 0))
Range("e15") = deck(hand(0, 1))
End Function
Function flop()
For i = 0 To 2
    Cells(9, 3 + i) = deck(shuffleDeck(deckCard))
    communityCards(i) = shuffleDeck(deckCard)
    deckCard = deckCard + 1
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
Dim zongpai(7)
For player = 0 To 0
    i = 0
    For handCard = 0 To 1
        zongpai(i) = hand(player, handCard)
        i = i + 1
    Next handCard
    For i = 2 To 6
        zongpai(i) = communityCards(i - 2)
    Next i
    For i = 0 To 6
    Debug.Print zongpai(i); "compare"
    Next
Next
End Function
Sub c()
Call TexasHoldem
End Sub
Function rules()

'final:
'royal flush:       sT,sJ,sQ,sK,sA
'straight flush:    56789s
'4 of a kind:       AAAAx
'full hose:         AAAKK
'flush:             479TKs
'straight:          56789o  'bicycle:A2345  'broadway:TJQKA
'3 of a kind:       AAAxx
'two pairs          AAKKx
'one pair           AAxxx
'high card          xxxxx

'hand:
'suited
'off suited
'set:pocket pair+hit flop
'connectors:45,9T,TJ
'draw hand:wait for flush or straight
'pocket pair:22,33,44 in hand
'American Airlines:AA pocket pair
'cowboys:KK
'Nuts:you are the strongest, AA hit AA6J8


'public:
'rainbow:flop off suited

'
'Pot:chips on table
'Outs:
'Bluff
'Slowplay
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
Function cccee()
Dim b(5)
a = Array(1, 2, 1, 4, 5)
u = 5
l = 0
For i = 0 To 4
    ci = Int(Rnd * u + l)
    b(i) = a(ci)
    a(ci) = a(l)
    l = l + 1
    u = u - 1
Next i
'For i = 0 To 4

End Function
