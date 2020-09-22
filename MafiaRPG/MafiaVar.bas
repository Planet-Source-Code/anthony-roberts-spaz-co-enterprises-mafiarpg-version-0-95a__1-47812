Attribute VB_Name = "modVar"
Public strName As String, intHP As Integer, intEnHP As Integer
Public intAtt As Integer, intEnAtt As Integer, intWeap As Integer
Public intEnWeap As Integer, intHit As Integer, intLevel As Integer
Public intSpec As Integer, strWeap As String, strEnWeap As String
Public intEnLevel As Integer, intExp As Double, intMaxHP As Integer
Public intMaxEnHP As Integer, strMessage As String, strEnemy As String
Public strMessage2 As String, intEnemy As Integer, intHeal As Integer
Public intSet As Integer, intHPIncrese As Integer, intTrain As Integer
Public intSelect As Integer

'Level Variables
Public intL2 As Double, intL3 As Double, intL4 As Double
Public intL5 As Double, intL6 As Double, intL7 As Double
Public intL8 As Double, intL9 As Double, intL10 As Double
Public intCurLevel As Double

'Character Things
Public intMoney As Double, bolAIDS As Boolean, intDrugUse As Integer
Public bolDrugLine As Boolean, intTime As Integer, strMonth As String
Public intLoanTotalTime As Integer, intMonth As Integer, intAmmo As Integer
Public StartLoan As Double, CurrentLoan As Double, intStealing As Integer
Public intBank As Double, bolBank As Boolean, bolCondom As Boolean
Public CurrentMonthTime As Integer, CityMap As Boolean

'"Chance"/Random Variable
Public intChance As Double, intBorder As Integer, bolLoanShark As Boolean
Public bolCalfAvailable As Boolean

'Bar Variables
Public bolBreast As Boolean, bolPussy As Boolean, bolSlap As Boolean
Public bolWeap As Boolean, bolDrink As Boolean, intBangCost As Integer
Public intGetMax As Integer, intCurrentGet As Integer, intHPIncreseChick As Integer
Public bolBar As Boolean

'Company
Public strCompany As String

'Drugs
Public intSpoon As Integer, intLazy As Integer, intKoocie As Integer
Public intLucky As Integer, intHair As Integer, intCoka As Integer
Public intCalf As Integer

'Weapons
Public bolBrass As Boolean, bolGold As Boolean, bolChain As Boolean
Public bolSwitch As Boolean, bolPistol As Boolean, bolUzi As Boolean
Public bolAssault As Boolean, bolPunisher As Boolean
    'Weapon Warrents
    Public bolW1 As Boolean, bolW2 As Boolean, bolW3 As Boolean, bolW4 As Boolean

'Casino
Public intBet As Double, intNum1 As Integer, intNum2 As Integer
Public intNum3 As Integer

'Stocks
Public intSCE As Integer, intSHE As Integer, intBI As Integer, intLADC As Integer
Public intTWM As Integer, intCNS As Integer, intBA As Integer, intHC As Integer
Public priSCE As Integer, priSHE As Integer, priBI As Integer, priLADC As Integer
Public priTWM As Integer, priCNS As Integer, priBA As Integer, priHC As Integer
Public totSCE As Integer, totSHE As Integer, totBI As Integer, totLADC As Integer
Public totTWM As Integer, totCNS As Integer, totBA As Integer, totHC As Integer
Public bolSCE As Boolean, bolSHE As Boolean, bolBI As Boolean, bolLADC As Boolean
Public bolTWM As Boolean, bolCNS As Boolean, bolBA As Boolean, bolHC As Boolean

'Save
Public strPath As String, strFileName As String, strExtension As String, strFileNameAndPath As String
Public intLoadTemp As Integer

'Lackys
Public intProsEarn As Integer, intChemEarn As Integer, intBounceEarn As Integer
Public intPros As Integer, intChem As Integer, intBounce As Integer

'Arena
Public strArenaName As String, strArenaWeapon As String, intArenaWeapon As Integer
Public intArenaMaxHP As Integer, intArenaHP As Integer, intArenaOdds As Integer
Public intArenaBet As Double, intArenaEarn As Double

'Cars
Public bolTaxi As Boolean, bolFirefly As Boolean, bolMustang As Boolean
Public bolViper As Boolean, bolGTO As Boolean, intSpeed As Integer

'Company Earnings
Public dbl1 As Double, dbl2 As Double, dbl3 As Double, dbl4 As Double
Public dbl5 As Double, dbl6 As Double, dbl7 As Double, dbl8 As Double

'Speed of Game
Public strSpeed As String

'Button


'Players Position
Public X As Integer, Y As Integer
Public MapX(10) As Boolean, MapY(10) As Boolean

'Enemy Present
Public EnemyThere As Boolean
