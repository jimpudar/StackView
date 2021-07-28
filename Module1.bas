Attribute VB_Name = "Module1"
Option Base 1
Public StackedDeck(6, 52) As Variant
Public Deck(6, 52) As Variant
'(1,x) is the stack value (1-52)
'(2,x) is the card value
'(3,x) says "CardXX" (such as CardAC for the Ace of Clubs)
'(4,x) says "Selection" if the card is selected
'(5,x) says "PositionXX" (such as Position12 for where the card currently
    'is in the deck)
'(6,x) is logical (True/False) for whether card is Reversed in deck
Public TempCard(6, 2) As Variant
Public ShuffleMeterDeck(3, 52) As Integer
'(1,x) is color - 0 for red, 1 for black
'(2,x) is suit - 1=club, 2=heart, 3=spade, 4=diamond
'(3,x) is value - 1=ace, 2=two,... 11=jack, 12=queen, 13=king
'where x= card position
Public CustomDeck(6, 52) As Variant
'(1,x) is the position slot order (1-52)
'(2,x) is the card value
'(3,x) is the original position values of the card indexes for Transfer: Retain
'(4,x) is the retained Selection info from the Deck
'(6,x) is the retained Reverse info from the Deck
Public CustomDeckAsImported(6, 52)
'retains all info of original cards for "Transfer: Retain Positions"
Public TestOriginalDeck(2, 52)
Public DeckImages(52) As Object
Public ChangedDeck(6, 52) As Variant
Public UnwindDeck(6, 52)
Public DisplayDeck(6, 52) As Variant
Public StartDeck(2, 52) As Variant
'used for Search Sequence
Public StartDeckName
Public StartDeckInitial(2, 52) As Variant
'used for Search - initial deck load
Public TargetDeck(2, 52) As Variant
'used for Search Sequence
Public TargetDeckName
Public MatchFound As Integer
'either a 0 or 1 to determine whether a match occurs in the Search Sequence procedure
' results (1 for match, 0 for no match)
Public NoMatchFound As Integer
'either a 0 or 1 to determine whether a "no match" conclusion occurs
'after a Search Sequence procedure
'results (1 for verified "no match", 0 for "no conclusion reached")
Public PartialMatchFound As Integer
'either a 0 or 1 to determine whether a partial match occurs in the
'Search Sequence procedure results (1 for match, 0 for no match)
Public PartialMatchCounter As Integer
'used in SearchTransactionCheckPartialMatch()
Public SearchStartDeckSet As Integer
'when form loads, it is set to 0
'when the start deck is correctly loaded, it is set to 1
Public SearchTargetDeckSet As Integer
'when form loads, it is set to 0
'when the target deck is correctly loaded, it is set to 1
Public SearchCurrentLevel As Integer
'keeps track of how many levels have been tried in the search
Public SearchCurrentLevelRestart As Integer
'need to save and restore the variable after a stop
Public SearchProcessing As Boolean
'true/false variable to interupt search
Public FirstPass As Boolean
'true/false variable indicating whether or not the progress bar is in its first pass
Public SearchFileLoading As Boolean
'true/false variable to control automatic processing when checkbox values are set
Public SearchProgressCounter As Integer
'counter for search progress bar segments across levels
'when level count is less than the required interval
Public Manipulations() As Variant
Public TopCut() As Variant
Public BottomCut() As Variant
Public CutDepth As Integer
Public SelectedCutDepth As Integer
Public RemainingCut As Integer
Public DeckCount As Integer
Public DragX
Public DragY
Public OriginalTop As Integer
Public OriginalLeft As Integer
Public ProtectedBlock As Integer
Public InteriorPosition As Integer
Public InteriorCard As Integer
Public MeshedBlockPosition As Integer
Public MeshedBlock As Integer
Public ShiftBlock As Integer
Public ShiftDepth As Integer
Public RifflePortion As Integer
Public CutError As Integer
Public Hands As Integer
Public SelectedCards(5, 20) As Variant
Public NumberOfSelectedCards As Integer
Public DeckProperties As Integer
    'DeckProperties represents the first index in the Deck(x,52)
    'and ChangedDeck(x,52) stack manipulation.  This will allow
    'for later addition of deck properties.
Public FirstCutStart As Integer
Public FirstCutDepth As Integer
Public SecondCutStart As Integer
Public SecondCutDepth As Integer
Public ThirdCutStart As Integer
Public ThirdCutDepth As Integer
Public FourthCutDepth As Integer
Public FourthCutStart As Integer
Public FifthCutDepth As Integer
Public SelectedCard As Integer
Public ReturnPosition As Integer
Public fMainForm As frmMain
Public SuitOrder(4)
Public NextCardIncrements(4)
Public NextCardIncrementsReport(4)
Public StanyonStartingCard
Public StanyonVariationDeck(2, 52)
Public StanyonParameterError As Integer
Public StanyonSuitError As Integer
Public StanyonCycleError As Integer
Public SuitPointer As Integer
Public SuitPointerOffset As Integer
Public IncrementPointer As Integer
Public CardValue
Public CardText
Public FFS
'FFS stands for "First Four Suits" in the Stanyon Variation error checking
Public SessionAlreadyOpen As Integer
'this variable sets to nonzero after initial loading.
'this allows opening and closing of Control Form without reseting the deck.
Public SessionCommand
'this text variable passes the command name to the SessionListBox
Public SessionPresent As Integer
'this binary variable keeps track if there are
'session entries in the SessionListBox
Public SessionSaved As Integer
'this binary variable keeps track if SessionListBox has not been saved
'a 0 means that it has not been saved.  A 1 means it has been saved, or there
'is nothing to save.
Public MnemonicSaved As Integer
'this binary variable keeps track if Mnemonic Table changes have not been saved
'a 0 means that it has not been saved.  A 1 means it has been saved, or there
'is nothing to save.
Public MnemonicCompare As Variant
Public SearchSaved As Integer
'this binary variable keeps track if the search parameters have been saved
'a 0 means that it has not been saved.  A 1 means it has been saved, or there
'is nothing to save.  It is used in the search module before a new search
'is initiated in the middle of a paused search.
Public SearchSessionTransferred As Integer
'binary variable 0=not transferred; 1=transferred
Public SessionRecordMode As Boolean
'logical value of False=not recording, True=recording
Public TestingMode As Boolean
'logical value of False=not testing, True=testing
Public SearchingMode As Boolean
'logical value of False=not searching, True=searching
Public SearchCounter As Integer
'used as an array index for Manuipluations(x,y)
Public SearchContinueReady As Integer
'binary variable indicating if search can continue
'if search parameters are changed, it will be set to zero
Public SearchCounterMax As Integer
'the value used to dimension Manipulations (2,SearchCounterMax)
Public SessionParseError As Boolean
'logical value for error traps in SessionParse and event plays
Public SearchParseError As Boolean
'logical value for error traps in SearchParse and event plays
Public SessionAllowableParameters()
'contains an array of the allowable event parameters for SessionParse
Public SessionNumParameters
'contains the number of allowable text-based session parameters
Public StartOrder(52)
'contains the card index values for building the TestOrder
Public TestOrder(52)
'contains the actual test order
Public TestCounter
'keeps track of how many cards have been shown
Public DeckRangeStart
'for the Test module Partial Deck Range starting value
Public DeckRangeFinish
'for the Test module Partial Deck Range finish value
Public DeckRangeCount
'for the Test module will equal to DeckRangeFinish-DeckRangeStart + 1
Public TestRandomValue
'either a 0 or 1 to determine whether the test is for card(1) or position(0)
Public TestCardMode
'either a 0, 1, or 2 to determine Current (0), Next (1), or Previous (2)
Public TestProgressIntervals
Public ShowProgressIntervals
'these two are used to set the progress bars with the timers
Public ShowingMode As Integer
'set to 1 if the answer is being showed.
'this helps set the "Show" button as a toggle for behaving like "Next"
Public CummTimeIntervals
'the cummulative testing/showing time in Timer Intervals
Public CummTimeSeconds
'the cummulative testing/showing time in seconds (calculated at end)
Public StartTime
Public SearchStartTime
'used to calculate Elapsed Time
Public ElapsedTime
Public SearchElapsedTime As Double
Public SpeedMod As Long
Public ImportedCustomDeck As Integer
'used to check that a deck has actually been imported
'0 if not imported (reset), and 1 if imported
Public CreatedStanyonDeck As Integer
'used to check that a Stanyon deck has actually been created
'0 if not created (reset), and 1 if created
Public PokerCardsDealt As Integer
'either a 0 or 1 to determine whether the cards are in poker
' deal position (1 for yes, 0 for no)
'ShowDeal sets the variable to 1
'ShowCards sets the variable to 0
Public SearchLevelCounter(26) As Integer
'these variables are used to track the Search moves
'each value can be between 1 and SearchCounterMax
'  Manipulations(1, SearchLevelCounter(4)) indicates the 4th manipulation in the
'  latest search run.

'these next variables are used to determine which special manipulations are used
' in a search for a deck match
Public SearchSpecialName
'this variable contains the checkboxname used for special search parameters
Public SearchCDPMin As Integer
Public SearchCDPMax As Integer
Public SearchCDP As Integer
Public SearchRSCMin As Integer
Public SearchRSCMax As Integer
Public SearchRSC As Integer
Public SearchRSCRMin As Integer
Public SearchRSCRMax As Integer
Public SearchRSCR As Integer
Public SearchMC1Min As Integer
Public SearchMC1Max As Integer
Public SearchMC2Min As Integer
Public SearchMC2Max As Integer
Public SearchMC As Integer
Public SearchSTB1Min As Integer
Public SearchSTB1Max As Integer
Public SearchSTB2Min As Integer
Public SearchSTB2Max As Integer
Public SearchSTB As Integer
Public SearchSTBR1Min As Integer
Public SearchSTBR1Max As Integer
Public SearchSTBR2Min As Integer
Public SearchSTBR2Max As Integer
Public SearchSTBR As Integer
Public SearchOFST1Min As Integer
Public SearchOFST1Max As Integer
Public SearchOFST2Min As Integer
Public SearchOFST2Max As Integer
Public SearchOFST As Integer
Public SearchOFSTR1Min As Integer
Public SearchOFSTR1Max As Integer
Public SearchOFSTR2Min As Integer
Public SearchOFSTR2Max As Integer
Public SearchOFSTR As Integer
Public SearchOFSB1Min As Integer
Public SearchOFSB1Max As Integer
Public SearchOFSB2Min As Integer
Public SearchOFSB2Max As Integer
Public SearchOFSB As Integer
Public SearchOFSBR1Min As Integer
Public SearchOFSBR1Max As Integer
Public SearchOFSBR2Min As Integer
Public SearchOFSBR2Max As Integer
Public SearchOFSBR As Integer
Public SearchIFST1Min As Integer
Public SearchIFST1Max As Integer
Public SearchIFST2Min As Integer
Public SearchIFST2Max As Integer
Public SearchIFST As Integer
Public SearchIFSTR1Min As Integer
Public SearchIFSTR1Max As Integer
Public SearchIFSTR2Min As Integer
Public SearchIFSTR2Max As Integer
Public SearchIFSTR As Integer
Public SearchIFSB1Min As Integer
Public SearchIFSB1Max As Integer
Public SearchIFSB2Min As Integer
Public SearchIFSB2Max As Integer
Public SearchIFSB As Integer
Public SearchIFSBR1Min As Integer
Public SearchIFSBR1Max As Integer
Public SearchIFSBR2Min As Integer
Public SearchIFSBR2Max As Integer
Public SearchIFSBR As Integer
Public SearchMatchStartCard As Integer
Public SearchMatchEndCard As Integer
Public SearchSpecialCancel As Boolean
'used to trap a "cancal" button press after a SearchSpecial dialog box
'this is to prevent the manipulation count from incorrectly increasing
Public ThresholdMatchCards As Integer
Public TrapThreshold As Boolean
Public SearchTotalPossibleTime As Double
'this value is in seconds
Public SearchRemainingPossibleTime As Double
'this value is in seconds
Public SearchTotalManipulations As Double
Public ManipulationsPerSecond As Double
'this variable keeps track of how many deck manipulations can be performed by
'the computer each second.  It is used to adjust the estimated possible time
'display.  At the time of original programming, this value was 3500 for my computer
Public WholeDeckMatchSet As Boolean
'when true, the WholeDeckMatch parameters have been set and saved, and are current
'when false, the PartialDeckMatch parameters have been set and saved and current
'when entering one of the two dialog boxes, this logical variable will help
'exiting with the "cancel" button in an elegant manner.
Public TrapFileWhole
'file name for trap log file
Public TrapFileWholeFinal
'file name for trap log file when the ok button is pressed
Public TrapPathWhole
'path name for trap log file
Public TrapPathWholeFinal
'path name for trap log file when the ok button is pressed
Public TrapFilePartial
'file name for trap log file
Public TrapFilePartialFinal
'file name for trap log file when the ok button is pressed
Public TrapPathPartial
'path name for trap log file
Public TrapPathPartialFinal
'path name for trap log file when the ok button is pressed
Public TrapFileFinal
'file name used in match check routine
Public TrapPathFinal
'path name used in match check routine
Public SuspendTrapWhole
'used to control the Option buttons on frmWholeDeckMatch
Public SuspendTrapPartial
'used to control the Option buttons on frmPartialDeckMatch
Public SuspendTrapWholeFinal
'used to control the Option buttons on frmWholeDeckMatch when the 'ok button is pressed
Public SuspendTrapPartialFinal
'used to control the Option buttons on frmPartialDeckMatch when the 'ok button is pressed
Public SuspendTrapFinal
'used for match check
Public AdvTestCounter
'keeps track of how many cards have been shown for the Advanced Test
Public AdvTestingMode As Boolean
'logical value of False=not testing, True=testing
Public AdvShowingMode As Integer
'set to 1 if the answer is being showed.
'this helps set the "Show" button as a toggle for behaving like "Next"
Public AdvTestRange
'the number of Advanced Test cards to show
Public AdvStartTime
Public AdvTestProgressIntervals
Public AdvShowProgressIntervals
Public DesiredCardSequence(30) As Variant
Public DesiredPositionSequence(30) As Variant
Public AdvElapsedTime
Public AdvDesiredCardText
Public AnswerInputKeyEntry
Public AdvDesiredCard
Public AdvDesiredCardShift
Public AdvDesiredPosition
Public CardsToCut
Public NewTopCardShift
Public NewTopCard
Public NewBottomCardShift
Public NewBottomCard
Public OriginalDeckOrder(6, 52) As Variant
Public AdvDeckOriginal(6, 52) As Variant
Public AdvDeckCurrent(6, 52) As Variant
Public CorrectAnswerCount
Public BlockSize
Public PileTable(8, 2) As Variant
'the first element is the pile number
'the second element is the start/stop position values
'   (x, 1) = start of pile card position for pile x
'   (x, 2) = stop of pile card position for pile x
Public cPileTable(8, 2) As Variant
'this is the changedPileTable during the CutPiles sequences
Public sPileTable(8, 2) As Variant
'this is the swappedPileTable during the SwapPiles sequences
Public ChangedPileTable(8, 2) As Variant
'the first element is the pile number
'the second element is the start/stop position values
'   (x, 1) = start of pile card position for pile x
'   (x, 2) = stop of pile card position for pile x
Public PileLocations(8, 3) As Variant
'the first element is the pile number
'the second element is the top/left location values
'   (x, 1) = top location of first card position for pile x
'   (x, 2) = left location of first card position for pile x
'   (x, 3) = right location of right edge of pile x
Public PileDeck(6, 52) As Variant
Public ChangedPileDeck(6, 52) As Variant
Public PilesShown As Integer
'either a 0 or 1 to determine whether the cards are in pile
' dealt position (1 for yes, 0 for no)
'ShowDeal sets the variable to 0
'ShowCards sets the variable to 0
Public NumPiles As Integer
'the actual number of piles showing
Public NumPilesPlan As Integer
'the number of piles that want to be created
Public Row1 As Integer
Public Row2 As Integer
Public Row3 As Integer
Public RowWidth As Integer
Public MaxWidth As Integer
Public MaxHeight As Integer
Public PileError As Boolean
'logical value of False=no error, True=error
Public PileCode(8) As String
Public PileInputData(2, 8) As Variant
'PileInputData(1,x) = pile type (RP=RandomPure, RA=Random Approx, S=Specified Exact)
'PileInputData(2,x) = raw input date before randoms are calculated
Public PileOutputData(2, 8) As Variant
'PileOutputData(1,x) = pile type (RP=RandomPure, RA=Random Approx, SE=Specified Exact)
'PileOutputData(2,x) = calculated output data
Public PileParseError As Boolean
Public RandomPiles(8, 52) As Integer
'this variable is used for random pile genereation to distribute the cards
'before they are reassembled into a full deck for pile display
'the First parameter is the pile number
'the Second parameter is the stack value of the card at that position in the pile
Public PileMatrixRow As Integer
Public PileMatrixColumn As Integer
Public MnemonicCards(52) As String
Public MnemonicPositions(52) As String
Public MnemonicCardIndex As Integer
Public MnemonicPositionIndex As Integer
Public TemporaryMnemonicHint As Boolean
Public BackDesignSelected As String
'this is for the back design that is selected in the dialog box
Public BackDesignSelected2 As String
'this is for the back design that is selected in the dialog box, but has the yellow overlay
Public BackDesignCurrent As String
Public pSEP(11) As Variant
'this is used by the session parsing procedure
'I had to make it a public variable because the procedure got too long
Public SessionRecursionLevel As Integer
'this variable is used to keep track of how deep recursive sessions are
Public SessionRecursionLimit As Integer
'this variable is used to limit how deep recursive sessions are
Public SessionRecursing As Boolean
'this variable is used in the parse section
Public InsertMacroError As Boolean
'keeps track of error for macro inserts in sessions
Public GilbreathOffset As Integer
'keeps track of screen placement of offset cards
Public GilbreathStatus(52) As Boolean
'identifies which cards are offset in the temporary deck
Public GilbreathDeck(52) As Boolean
'identifies which cards are offset in the final deck
Public GilbreathPileNum As Integer
'identifies the pile with the offset cards
Public GilbreathActive As Boolean
'if True, the Gilbreath view can be created
Public GilbreathShown As Boolean
'if True, a Gilbreath view is showing





Sub Main()
    frmSplash.Show
    frmSplash.Refresh
'    frmTest.Visible = False
'    frmCustomDeck.Visible = False
'    frmShuffleMeter.Visible = False
'    frmTestAdvanced.Visible = False
'    frmPiles.Visible = False
'    frmMnemonic.Visible = False
'    frmDeck.Visible = False
'    frmStackView.Visible = False
'    Unload frmSplash
    SessionAlreadyOpen = 0
    frmMain.Show
    Unload frmSplash
End Sub


