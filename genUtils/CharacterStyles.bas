Attribute VB_Name = "CharacterStyles"
'Created by Erica Warren -- erica.warren@macmillan.com
'Split off from MacmillanCleanupMacro: https://github.com/macmillanpublishers/Word-template/blob/master/macmillan/CleanupMacro.bas

' ======== PURPOSE ============================
' Applies Macmillan character styles to direct-styled text in current document

' ======== DEPENDENCIES =======================
' 1. Requires ProgressBar userform module

' Note: have already used all numerals and capital letters for tagging, starting with lowercase letters. through a.

Option Explicit
Option Base 1

Dim activeRng As Range

Sub MacmillanCharStyles()
    
    Dim CharacterProgress As ProgressBar
    Set CharacterProgress = New ProgressBar
    
    CharacterProgress.Title = "Macmillan Character Styles Macro"
    
    Call ActualCharStyles(oProgressChar:=CharacterProgress, StartPercent:=0, TotalPercent:=1)

End Sub

Sub ActualCharStyles(oProgressChar As ProgressBar, StartPercent As Single, TotalPercent As Single)
    ' Have to pass the ProgressBar so this can be run from within another macro
    ' StartPercent is the percentage the progress bar is at when this sub starts
    ' TotalPercent is the total percent of the progress bar that this sub will cover
    
    
    '------------------Time Start-----------------
    'Dim StartTime As Double
    'Dim SecondsElapsed As Double
    
    'Remember time when macro starts
    'StartTime = Timer
    
    '------------check for endnotes and footnotes--------------------------
    Dim stStories() As Variant
    stStories = StoryArray
    
    ' ======= Run startup checks ========
    ' True means a check failed (e.g., doc protection on)
    If StartupSettings(StoriesUsed:=stStories) = True Then
        Call Cleanup
        Exit Sub
    End If
    
    '--------Progress Bar------------------------------
    'Percent complete and status for progress bar (PC) and status bar (Mac)
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
    
    'First status shown will be randomly pulled from array, for funzies
    Dim funArray() As String
    ReDim funArray(1 To 10)      'Declare bounds of array here
    
    funArray(1) = "* Mixing metaphors..."
    funArray(2) = "* Arguing about the serial comma..."
    funArray(3) = "* Un-mixing metaphors..."
    funArray(4) = "* Avoiding the passive voice..."
    funArray(5) = "* Ending sentences in prepositions..."
    funArray(6) = "* Splitting infinitives..."
    funArray(7) = "* Ooh, what an interesting manuscript..."
    funArray(8) = "* Un-dangling modifiers..."
    funArray(9) = "* Jazzing up author bio..."
    funArray(10) = "* Filling in plot holes..."
    
    Dim X As Integer
    
    'Rnd returns random number between (0,1], rest of expression is to return an integer (1,10)
    Randomize           'Sets seed for Rnd below to value of system timer
    X = Int(UBound(funArray()) * Rnd()) + 1
    
    'Debug.Print x
    

'    strTitle = "Macmillan Character Styles Macro"

'   first number is percent of THIS macro completed
    sglPercentComplete = (0.09 * TotalPercent) + StartPercent
    strStatus = funArray(X)

    
    ' This is now passed from calling sub
'    Dim oProgressChar As ProgressBar
'    Set oProgressChar = New ProgressBar ' Triggers Initialize event, which calls Show method for PC

'    oProgressChar.Title = strTitle
    
    ' Calls ProgressBar.Increment mathod and waits for it to complete
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    
    '-----------Delete hidden text ------------------------------------------------
    Dim S As Long
    
    ' Note, if you don't delete hidden text, this macro turns it into reg. text.
    For S = 1 To UBound(stStories())
        If HiddenTextSucks(StoryType:=(stStories(S))) = True Then
            ' Notify user maybe?
        End If
    Next S
    
    Call zz_clearFind
    
    
    ' -------------- Clear formatting from paragraph marks -------------------
    ' can cause errors
    
    For S = 1 To UBound(stStories())
        Call ClearPilcrowFormat(StoryType:=(stStories(S)))
    Next S
    '===================== Replace Local Styles Start ========================

    '-----------------------Tag space break styles----------------------------
    Call zz_clearFind                          'Clear find object
    
    sglPercentComplete = (0.18 * TotalPercent) + StartPercent
    strStatus = "* Preserving styled whitespace..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    
    For S = 1 To UBound(stStories())
        Call PreserveWhiteSpaceinBrkStylesA(StoryType:=(stStories(S)))     'Part A tags styled blank paragraphs so they don't get deleted
    Next S
    Call zz_clearFind
    
    '----------------------------Fix hyperlinks---------------------------------------
    sglPercentComplete = (0.28 * TotalPercent) + StartPercent
    strStatus = "* Applying styles to hyperlinks..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call StyleAllHyperlinks(StoriesInUse:=stStories)
    
    Call zz_clearFind

    
    '--------------------------Remove unstyled space breaks---------------------------
    sglPercentComplete = (0.39 * TotalPercent) + StartPercent
    strStatus = "* Removing unstyled breaks..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    For S = 1 To UBound(stStories())
        Call RemoveBreaks(StoryType:=(stStories(S)))  ''new sub v. 3.7, removed manual page breaks and multiple paragraph returns
    Next S
    Call zz_clearFind
    
    '--------------------------Tag existing character styles------------------------
    sglPercentComplete = (0.52 * TotalPercent) + StartPercent
    strStatus = "* Tagging character styles..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    For S = 1 To UBound(stStories())
        Call TagExistingCharStyles(StoryType:=(stStories(S)))            'tag existing styled items
    Next S
    Call zz_clearFind
    
    '-------------------------Tag direct formatting----------------------------------
    sglPercentComplete = (0.65 * TotalPercent) + StartPercent
    strStatus = "* Tagging direct formatting..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    ' allBkmkrStyles is a jagged array (array of arrays) to hold in-use Bookmaker styles.
    ' i.e., one array for each story. Must be Variant.
    Dim allBkmkrStyles() As Variant
    For S = 1 To UBound(stStories())
    'tag local styling, reset local styling, remove text highlights
        Call LocalStyleTag(StoryType:=(stStories(S)))
        
        ReDim Preserve allBkmkrStyles(1 To S)
        allBkmkrStyles(S) = TagBkmkrCharStyles(StoryType:=stStories(S))
    Next S
    Call zz_clearFind


    '----------------------------Apply Macmillan character styles to tagged text--------
    sglPercentComplete = (0.81 * TotalPercent) + StartPercent
    strStatus = "* Applying Macmillan character styles..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    For S = 1 To UBound(stStories())
        Call LocalStyleReplace(StoryType:=(stStories(S)), BkmkrStyles:=allBkmkrStyles(S))            'reapply local styling through char styles
    Next S
    Call zz_clearFind
    
    '---------------------------Remove tags from styled space breaks---------------------
    sglPercentComplete = (0.95 * TotalPercent) + StartPercent
    strStatus = "* Cleaning up styled whitespace..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    For S = 1 To UBound(stStories())
        Call PreserveWhiteSpaceinBrkStylesB(StoryType:=(stStories(S)))     'Part B removes the tags and reapplies the styles
    Next S
    Call zz_clearFind
    
    
    ' -------------------------- Tag un-styled paragraphs as TX / TX1 / COTX1 -----------
    ' NOTE: must be done AFTER character styles, because if whole para has direct format
    ' it will be removed when apply style (but style won't be removed)
    ' This is total progress bar that will be covered in TagUnstyledText
    Dim sglTotalForText As Single
    sglTotalForText = TotalPercent - sglPercentComplete

    Call TagUnstyledText(objTagProgress:=oProgressChar, StartingPercent:=sglPercentComplete, _
        TotalPercent:=sglTotalForText, Status:=strStatus)

    ' Only tagging through main text story, because Endnotes story and Footnotes story should
    ' already be tagged at Endnote Text and Footnote Text by dafault when created

    '---------------------------Return settings to original------------------------------
    sglPercentComplete = TotalPercent + StartPercent
    strStatus = "* Finishing up..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressChar, Status:=strStatus, Percent:=sglPercentComplete)
    
    ' If this is the whole macro, close out; otherwise calling macro will close it all down
    If TotalPercent = 1 Then
        Call Cleanup
        Unload oProgressChar
        MsgBox "Macmillan character styles have been applied throughout your manuscript."
    End If
    
    
    '----------------------Timer End-------------------------------------------
    'Determine how many seconds code took to run
    '  SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Notify user in seconds
    '  Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"

End Sub


Private Sub PreserveWhiteSpaceinBrkStylesA(StoryType As WdStoryType)
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
    Dim tagArray(13) As String                                   ' number of items in array should be declared here
    Dim StylePreserveArray(13) As String              ' number of items in array should be declared here
    Dim E As Long
    
    StylePreserveArray(1) = "Space Break (#)"
    StylePreserveArray(2) = "Space Break with Ornament (orn)"
    StylePreserveArray(3) = "Space Break with ALT Ornament (orn2)"
    StylePreserveArray(4) = "Section Break (sbr)"
    StylePreserveArray(5) = "Part Start (pts)"
    StylePreserveArray(6) = "Part End (pte)"
    StylePreserveArray(7) = "Page Break (pb)"
    StylePreserveArray(8) = "Space Break - 1-Line (ls1)"
    StylePreserveArray(9) = "Space Break - 2-Line (ls2)"
    StylePreserveArray(10) = "Space Break - 3-Line (ls3)"
    StylePreserveArray(11) = "Column Break (cbr)"
    StylePreserveArray(12) = "Design Note (dn)"
    StylePreserveArray(13) = "Bookmaker Page Break (br)"
    
    tagArray(1) = "`1`^&`1``"                                       'v. 3.1 patch  added extra backtick on trailing tag for all of these.
    tagArray(2) = "`2`^&`2``"
    tagArray(3) = "`3`^&`3``"
    tagArray(4) = "`4`^&`4``"
    tagArray(5) = "`5`^&`5``"
    tagArray(6) = "`6`^&`6``"
    tagArray(7) = "`7`^&`7``"
    tagArray(8) = "`8`^&`8``"
    tagArray(9) = "`9`^&`9``"
    tagArray(10) = "`0`^&`0``"
    tagArray(11) = "`L`^&`L``"
    tagArray(12) = "`R`^&`R``"
    tagArray(13) = "`N`^&`N``"
    
    On Error GoTo BreaksStyleError:
    
    For E = 1 To UBound(StylePreserveArray())
        With activeRng.Find
          .ClearFormatting
          .Replacement.ClearFormatting
          .Text = "^13"
          .Replacement.Text = tagArray(E)
          .Wrap = wdFindContinue
          .Format = True
          .Style = StylePreserveArray(E)
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = True
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute Replace:=wdReplaceAll
        End With
NextLoop:
    Next
    
    On Error GoTo 0
    Exit Sub
    
BreaksStyleError:
    ' skips tagging that style if it's missing from doc; if missing, obv nothing has that style
    'Debug.Print StylePreserveArray(e)
    '5834 "Item with specified name does not exist" i.e. style not present in doc
    '5941 item not available in collection
    If Err.Number = 5834 Or Err.Number = 5941 Then
        Resume NextLoop:
    End If

End Sub

Private Sub RemoveBreaks(StoryType As WdStoryType)
    'Created v. 3.7
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
    Dim wsFindArray(4) As String              'number of items in array should be declared here
    Dim wsReplaceArray(4) As String       'and here
    Dim Q As Long
    
    wsFindArray(1) = "^m^13"              'manual page breaks
    wsFindArray(2) = "^13{2,}"          '2 or more paragraphs
    wsFindArray(3) = "(`[0-9]``)^13"    'remove para following a preserved break style                     v. 3.1 patch
    wsFindArray(4) = "(^m`7`^13`7``)`7`^13`7``"  'remove blank para following page break
                                                    ' even if styled.
    
    wsReplaceArray(1) = "^p"
    wsReplaceArray(2) = "^p"
    wsReplaceArray(3) = "\1"
    wsReplaceArray(4) = "\1"
    
    
    For Q = 1 To UBound(wsFindArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = wsFindArray(Q)
            .Replacement.Text = wsReplaceArray(Q)
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next
    
    ''' the bit below to remove the first or last paragraph if it's blank
    Dim myRange As Range
    Set myRange = ActiveDocument.Paragraphs(1).Range
        If myRange.Text = Chr(13) Then myRange.Delete
    
    Set myRange = ActiveDocument.Paragraphs.Last.Range
        If myRange.Text = Chr(13) Then myRange.Delete

End Sub

Private Sub PreserveWhiteSpaceinBrkStylesB(StoryType As WdStoryType)
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
    Dim tagArrayB(13) As String                                   ' number of items in array should be declared here
    Dim F As Long
    
    tagArrayB(1) = "`1`(^13)`1``"                             'v. 3.1 patch  added extra backtick on trailing tag for all of these.
    tagArrayB(2) = "`2`(^13)`2``"
    tagArrayB(3) = "`3`(^13)`3``"
    tagArrayB(4) = "`4`(^13)`4``"
    tagArrayB(5) = "`5`(^13)`5``"
    tagArrayB(6) = "`6`(^13)`6``"
    tagArrayB(7) = "`7`(^13)`7``"
    tagArrayB(8) = "`8`(^13)`8``"
    tagArrayB(9) = "`9`(^13)`9``"
    tagArrayB(10) = "`0`(^13)`0``"
    tagArrayB(11) = "`L`(^13)`L``"              ' for new column break, added v. 3.4.1
    tagArrayB(12) = "`R`(^13)`R``"
    tagArrayB(13) = "`N`(^13)`N``"
    
    For F = 1 To UBound(tagArrayB())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = tagArrayB(F)
            .Replacement.Text = "\1"
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next

End Sub

Private Sub TagExistingCharStyles(StoryType As WdStoryType)
    Set activeRng = ActiveDocument.StoryRanges(StoryType)                        'this whole sub (except last stanza) is basically a v. 3.1 patch.  correspondingly updated sub name, call in main, and replacements go along with bold and common replacements
    
    Dim tagCharStylesArray(12) As String                                   ' number of items in array should be declared here
    Dim CharStylePreserveArray(12) As String              ' number of items in array should be declared here
    Dim D As Long
    
    CharStylePreserveArray(1) = "span hyperlink (url)"
    CharStylePreserveArray(2) = "span symbols (sym)"
    CharStylePreserveArray(3) = "span accent characters (acc)"
    CharStylePreserveArray(4) = "span cross-reference (xref)"
    CharStylePreserveArray(5) = "span material to come (tk)"
    CharStylePreserveArray(6) = "span carry query (cq)"
    CharStylePreserveArray(7) = "span key phrase (kp)"
    CharStylePreserveArray(8) = "span preserve characters (pre)"  'added v. 3.2

    CharStylePreserveArray(9) = "span ISBN (isbn)"  'added v. 3.7
    CharStylePreserveArray(10) = "span symbols ital (symi)"     'added v. 3.8
    CharStylePreserveArray(11) = "span symbols bold (symb)"
    CharStylePreserveArray(12) = "span run-in computer type (comp)"
    
    
    tagCharStylesArray(1) = "`H|^&|H`"
    tagCharStylesArray(2) = "`Z|^&|Z`"
    tagCharStylesArray(3) = "`Y|^&|Y`"
    tagCharStylesArray(4) = "`X|^&|X`"
    tagCharStylesArray(5) = "`W|^&|W`"
    tagCharStylesArray(6) = "`V|^&|V`"
    tagCharStylesArray(7) = "`T|^&|T`"
    tagCharStylesArray(8) = "`F|^&|F`"
    tagCharStylesArray(9) = "`Q|^&|Q`"
    tagCharStylesArray(10) = "`E|^&|E`"
    tagCharStylesArray(11) = "`G|^&|G`"
    tagCharStylesArray(12) = "`J|^&|J`"
    
    On Error GoTo CharStyleError
    
    For D = 1 To UBound(CharStylePreserveArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Replacement.Text = tagCharStylesArray(D)
            .Wrap = wdFindContinue
            .Format = True
            .Style = CharStylePreserveArray(D)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
NextLoop:
    Next
    
    On Error GoTo 0
    Exit Sub
    
CharStyleError:
        ' skips tagging that style if it's missing from doc; if missing, obv nothing has that style
        'Debug.Print CharStylePreserveArray(d)
        '5834 "Item with specified name does not exist" i.e. style not present in doc
        '5941 item is not present in collection
        If Err.Number = 5834 Or Err.Number = 5941 Then
            Resume NextLoop
        End If

End Sub

Private Sub LocalStyleTag(StoryType As WdStoryType)
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
    '------------tag key styles-------------------------------
    Dim tagStyleFindArray(11) As Boolean              ' number of items in array should be declared here
    Dim tagStyleReplaceArray(11) As String         'and here
    Dim G As Long
    
    tagStyleFindArray(1) = False        'Bold
    tagStyleFindArray(2) = False        'Italic
    tagStyleFindArray(3) = False        'Underline
    tagStyleFindArray(4) = False        'Smallcaps
    tagStyleFindArray(5) = False        'Subscript
    tagStyleFindArray(6) = False        'Superscript
    tagStyleFindArray(7) = False        'Highlights
    ' note 8 - 10 are below
    tagStyleFindArray(11) = False       'Strikethrough
    
    tagStyleReplaceArray(1) = "`B|^&|B`"
    tagStyleReplaceArray(2) = "`I|^&|I`"
    tagStyleReplaceArray(3) = "`U|^&|U`"
    tagStyleReplaceArray(4) = "`M|^&|M`"
    tagStyleReplaceArray(5) = "`S|^&|S`"
    tagStyleReplaceArray(6) = "`P|^&|P`"
    tagStyleReplaceArray(8) = "`A|^&|A`"
    tagStyleReplaceArray(9) = "`C|^&|C`"
    tagStyleReplaceArray(10) = "`D|^&|D`"
    tagStyleReplaceArray(11) = "`a|^&|a`"
    
    For G = 1 To UBound(tagStyleFindArray())
    
    tagStyleFindArray(G) = True
        
        If tagStyleFindArray(8) = True Then tagStyleFindArray(1) = True: tagStyleFindArray(2) = True     'bold and italic                        v. 3.1 update
        If tagStyleFindArray(9) = True Then tagStyleFindArray(1) = True: tagStyleFindArray(4) = True: tagStyleFindArray(2) = False  'bold and smallcaps                 v. 3.1 update
        If tagStyleFindArray(10) = True Then tagStyleFindArray(2) = True: tagStyleFindArray(4) = True: tagStyleFindArray(1) = False 'smallcaps and italic               v. 3.1 update
        If tagStyleFindArray(11) = True Then tagStyleFindArray(2) = False: tagStyleFindArray(4) = False ' reset tags for strikethrough
        
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Replacement.Text = tagStyleReplaceArray(G)
            .Wrap = wdFindContinue
            .Format = True
            .Font.Bold = tagStyleFindArray(1)
            .Font.Italic = tagStyleFindArray(2)
            .Font.Underline = tagStyleFindArray(3)
            .Font.SmallCaps = tagStyleFindArray(4)
            .Font.Subscript = tagStyleFindArray(5)
            .Font.Superscript = tagStyleFindArray(6)
            .Highlight = tagStyleFindArray(7)
            .Font.StrikeThrough = tagStyleFindArray(11)
            .Replacement.Highlight = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    
    tagStyleFindArray(G) = False
    
    Next
    


End Sub

Private Sub LocalStyleReplace(StoryType As WdStoryType, BkmkrStyles As Variant)
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
    ' Determine if we need to do the bookmaker styles thing
    ' BkmkrStyles is an array of bookmaker character styles in use. If it's empty,
    ' there are none in use so we don't have to check
    
    Dim blnCheckBkmkr As Boolean
    
    If IsArrayEmpty(BkmkrStyles) = False Then
        blnCheckBkmkr = True
    Else
        blnCheckBkmkr = False
    End If
    
    '-------------apply styles to tags
    'number of items in array should = styles in LocalStyleTag + styles in TagExistingCharStyles
    Dim tagFindArray(22) As String              ' number of items in array should be declared here
    Dim tagReplaceArray(22) As String         'and here
    Dim H As Long
    
    tagFindArray(1) = "`B|(*)|B`"
    tagFindArray(2) = "`I|(*)|I`"
    tagFindArray(3) = "`U|(*)|U`"
    tagFindArray(4) = "`M|(*)|M`"
    tagFindArray(5) = "`H|(*)|H`"
    tagFindArray(6) = "`S|(*)|S`"
    tagFindArray(7) = "`P|(*)|P`"
    tagFindArray(8) = "`Z|(*)|Z`"
    tagFindArray(9) = "`Y|(*)|Y`"
    tagFindArray(10) = "`X|(*)|X`"
    tagFindArray(11) = "`W|(*)|W`"
    tagFindArray(12) = "`V|(*)|V`"
    tagFindArray(13) = "`T|(*)|T`"
    tagFindArray(14) = "`A|(*)|A`"                'v. 3.1 patch
    tagFindArray(15) = "`C|(*)|C`"                 'v. 3.1 patch
    tagFindArray(16) = "`D|(*)|D`"                       'v. 3.1 patch
    tagFindArray(17) = "`F|(*)|F`"
    tagFindArray(18) = "`Q|(*)|Q`"          'v. 3.7 added
    tagFindArray(19) = "`E|(*)|E`"
    tagFindArray(20) = "`G|(*)|G`"          'v. 3.8 added
    tagFindArray(21) = "`J|(*)|J`"
    tagFindArray(22) = "`a|(*)|a`"

    tagReplaceArray(1) = "span boldface characters (bf)"
    tagReplaceArray(2) = "span italic characters (ital)"
    tagReplaceArray(3) = "span underscore characters (us)"
    tagReplaceArray(4) = "span small caps characters (sc)"
    tagReplaceArray(5) = "span hyperlink (url)"
    tagReplaceArray(6) = "span subscript characters (sub)"
    tagReplaceArray(7) = "span superscript characters (sup)"
    tagReplaceArray(8) = "span symbols (sym)"
    tagReplaceArray(9) = "span accent characters (acc)"
    tagReplaceArray(10) = "span cross-reference (xref)"
    tagReplaceArray(11) = "span material to come (tk)"
    tagReplaceArray(12) = "span carry query (cq)"
    tagReplaceArray(13) = "span key phrase (kp)"
    tagReplaceArray(14) = "span bold ital (bem)"
    tagReplaceArray(15) = "span smcap bold (scbold)"
    tagReplaceArray(16) = "span smcap ital (scital)"
    tagReplaceArray(17) = "span preserve characters (pre)"
    tagReplaceArray(18) = "span ISBN (isbn)"                        'v. 3.7 added
    tagReplaceArray(19) = "span symbols ital (symi)"                ' v. 3.8 added
    tagReplaceArray(20) = "span symbols bold (symb)"                ' v. 3.8 added
    tagReplaceArray(21) = "span run-in computer type (comp)"
    tagReplaceArray(22) = "span strikethrough characters (str)"
    
    For H = LBound(tagFindArray()) To UBound(tagFindArray())
    
    ' ----------- bookmaker char styles ----------------------
        ' tag bookmaker line-ending character styles and
        ' adjust name if have additional styles applied
        ' because if you append "tighten" or "loosen" to
        ' regular style name, Bookmaker does that.
        If blnCheckBkmkr = True Then
        
On Error GoTo BkmkrError

            Dim Q As Long
            Dim qCount As Long
            Dim strAction As String
            Dim strNewName As String
            Dim strTag As String
            
            ' deal with bookmaker styles
            For Q = LBound(BkmkrStyles) To UBound(BkmkrStyles)
                ' replace bookmaker-tagged text with bookmaker styles
                strTag = "bk" & Format(Q, "0000")
                With activeRng.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = "`" & strTag & "|(*)|" & strTag & "`"
                    .Replacement.Text = "\1"
                    .Wrap = wdFindContinue
                    .Format = True
                    .Replacement.Style = BkmkrStyles(Q)
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=wdReplaceAll
                End With

                'Move selection to start of document
                Selection.HomeKey Unit:=wdStory
            

                qCount = 0
                With Selection.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = tagFindArray(H)
                    .Replacement.Text = "\1"
                    .Wrap = wdFindStop
                    .Forward = True
                    .Style = BkmkrStyles(Q)
                    .Replacement.Style = tagReplaceArray(H)
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False

                    Do While .Execute = True And qCount < 200
                        qCount = qCount + 1
                        .Execute Replace:=wdReplaceOne
                        ' pull just action to add to style name
                        ' always starts w/ "bookmaker ", but we want to include the space,
                        ' hence start at 10
                        strAction = Mid(BkmkrStyles(Q), 10, InStr(BkmkrStyles(Q), "(") - 11)
                        strNewName = tagReplaceArray(H) & strAction
                        Debug.Print "current style is: " & Selection.Style
                        Debug.Print "new style is: " & strNewName
                        
                        ' Note these hybrid styles aren't in std template, so if they
                        ' haven't been created in this doc yet, will error.
                        Selection.Style = strNewName
                    Loop
                End With
            Next Q
            
            
        End If
        
On Error GoTo ErrorHandler
    ' tag the rest of the character styles
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = tagFindArray(H)
            .Replacement.Text = "\1"
            .Wrap = wdFindContinue
            .Format = True
            .Replacement.Style = tagReplaceArray(H)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    
NextLoop:
    Next
    
    On Error GoTo 0
    
    Exit Sub

ErrorHandler:
'    Debug.Print Err.Number & ": " & Err.Description
'    Debug.Print "Error style: " & tagReplaceArray(h)
    
    Dim myStyle As Style
    
    If Err.Number = 5834 Or Err.Number = 5941 Then
        Select Case tagReplaceArray(H)
            
            'If style from LocalStyleTag is not present, add style
            Case "span boldface characters (bf)":
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .Bold = True
                End With
                Resume
            
            Case "span italic characters (ital)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .Italic = True
                End With
                Resume
                
            Case "span underscore characters (us)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .Underline = wdUnderlineSingle
                End With
                Resume
            
            Case "span small caps characters (sc)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .SmallCaps = False
                    .AllCaps = True
                    .Size = 9
                End With
                Resume
            
            Case "span subscript characters (sub)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .Subscript = True
                End With
                Resume
                
            Case "span superscript characters (sup)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .Superscript = True
                End With
                Resume

            Case "span bold ital (bem)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .Bold = True
                    .Italic = True
                End With
                Resume
                
            Case "span smcap bold (scbold)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .SmallCaps = False
                    .AllCaps = True
                    .Size = 9
                    .Bold = True
                End With
                Resume

            Case "span smcap ital (scital)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), _
                    Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .SmallCaps = False
                    .AllCaps = True
                    .Size = 9
                    .Italic = True
                End With
                Resume
                
            Case "span strikethrough characters (str)"
                Set myStyle = ActiveDocument.Styles.Add(Name:=tagReplaceArray(H), Type:=wdStyleTypeCharacter)
                With myStyle.Font
                    .Shading.BackgroundPatternColor = wdColorLightTurquoise
                    .StrikeThrough = True
                End With
                Resume
            
            'Else just skip if not from direct formatting
            Case Else
                Resume NextLoop:
        
        End Select
    End If
    
    Exit Sub

BkmkrError:
'    Debug.Print Err.Number & ": " & Err.Description
'    Debug.Print "New name: " & strNewName
'    Debug.Print "Old name: " & tagReplaceArray(h)
    
    Dim myStyle2 As Style
    
    If Err.Number = 5834 Or Err.Number = 5941 Then

        Set myStyle2 = ActiveDocument.Styles.Add(Name:=strNewName, _
            Type:=wdStyleTypeCharacter)
            
On Error GoTo ErrorHandler
        ' If the original style did not exist yet, will error here
        ' but ErrorHandler will add the style
        myStyle2.BaseStyle = tagReplaceArray(H)
        ' Then go back to BkmkrError so further errors will route
        ' correctly
On Error GoTo BkmkrError
        Resume
    Else
        ' something else was the error
        MsgBox Err.Number & ": " & Err.Description
        Resume Next
    End If
    
End Sub




'Private Sub TagBkmkrCharStyles()
Private Function TagBkmkrCharStyles(StoryType As Variant) As Variant
'    Set activeRng = ActiveDocument.Range
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
' Will need to loop through stories as well
' And be a function that returns an array

    Dim objStyle As Style
    Dim strBkmkrNames() As String
    Dim Z As Long
    
    ' Loop through all styles to get array of bkmkr styles in document
    ' NOTE! The .InUse property does NOT mean "in use in the document"; it means
    ' "any custom style or any modified built-in style". Ugh. Anyway, now we
    ' have to loop through all styles to see if bookmaker styles are present,
    ' then search for each of those styles to see if they are in use.
    
    For Each objStyle In ActiveDocument.Styles
        ' If char style with "bookmaker" in name is in use...
'        Debug.Print objStyle.NameLocal & " InUse: " & objStyle.InUse
        ' binary compare is default, but adding here to be clear that we are doing
        ' a CASE SENSITIVE search, because "Bookmaker" is only for Paragraph styles,
        ' which we don't want to mess with.
        If InStr(1, objStyle.NameLocal, "bookmaker", vbBinaryCompare) <> 0 And _
            objStyle.Type = wdStyleTypeCharacter Then
            Debug.Print StoryType & ": " & objStyle.NameLocal
            Selection.HomeKey Unit:=wdStory
            ' Now see if it's being used ...
            With Selection.Find
                .ClearFormatting
                .Text = ""
                .Style = objStyle.NameLocal
                .Wrap = wdFindContinue
                .Format = True
                .Forward = True
                .Execute
            End With
            
            If Selection.Find.Found = True Then
                '... add it to an array
                Z = Z + 1
                ReDim Preserve strBkmkrNames(1 To Z)
                strBkmkrNames(Z) = objStyle.NameLocal
            End If
        End If
    Next objStyle

    If IsArrayEmpty(strBkmkrNames) = True Then
'        Debug.Print "No bookmaker character styles in use."
        TagBkmkrCharStyles = strBkmkrNames
        Exit Function
    End If

    ' Tag in-use bkmkr styles
    ' Make sure if text also has formatting,
    ' the tags do not have it...
    Dim X As Long
    Dim strTag As String
    Dim strAction As String
    Dim lngCount As Long

    For X = LBound(strBkmkrNames) To UBound(strBkmkrNames)
        strTag = "bk" & Format(X, "0000")
'        Debug.Print strTag
        
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Replacement.Text = "`" & strTag & "|^&|" & strTag & "`"
            .Wrap = wdFindContinue
            .Format = True
            .Style = strBkmkrNames(X)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With

    Next

    '-------------Reset everything -- clears all direct formatting!
    activeRng.Font.Reset
    
    ' return array of in-use bookmaker styles so we can tag later
    TagBkmkrCharStyles = strBkmkrNames
    

End Function

Private Sub TagUnstyledText(objTagProgress As ProgressBar, StartingPercent As Single, _
    TotalPercent As Single, Status As String)
    ' Make sure we're always working with the right document
    Dim thisDoc As Document
    Set thisDoc = ActiveDocument

    ' Rename built-in style that has parens
    thisDoc.Styles("Normal (Web)").NameLocal = "_"

    Dim lngParaCount As Long
    Dim A As Long
    Dim strCurrentStyle As String
    Dim strTX As String
    Dim strTX1 As String
    Dim strNewStyle As String
    Dim strParaStatus As String
    Dim sglStartingPercent As Single
    Dim sglTotalPercent As Single
    Dim strNextStyle As String
    Dim strNextNextStyle As String
    Dim strCOTX1 As String
    Dim sglPercentComplete As Single

    ' Making these variables so we don't get any input errors with the style names t/o
    strTX = "Text - Standard (tx)"
    strTX1 = "Text - Std No-Indent (tx1)"
    strCOTX1 = "Chap Opening Text No-Indent (cotx1)"

    lngParaCount = thisDoc.Paragraphs.Count

    Dim myStyle As Style ' For error handlers

    ' Loop through all paras, tag any w/o close parens as TX or TX1
    ' (or COTX1 if following chap opener)
    For A = 1 To lngParaCount

        If A Mod 100 = 0 Then
            ' Increment progress bar
            sglPercentComplete = (((A / lngParaCount) * TotalPercent) + _
                StartingPercent)
            strParaStatus = "* Tagging non-Macmillan paragraphs with Text " _
                & "- Standard (tx): " & A & " of " & lngParaCount & vbNewLine & Status
            Call UpdateBarAndWait(Bar:=objTagProgress, Status:=strParaStatus, _
                Percent:=sglPercentComplete)
        End If

        strCurrentStyle = thisDoc.Paragraphs(A).Style
        'Debug.Print a & ": " & strCurrentStyle

On Error GoTo ErrorHandler1     ' adds this style if it is not in the document
        ' tag all non-Macmillan-style paragraphs with standard Macmillan styles
        ' Macmillan styles all end in close parens
        If Right(strCurrentStyle, 1) <> ")" Then    ' it's not a Macmillan style
            ' If flush left, make No-Indent
            If thisDoc.Paragraphs(A).FirstLineIndent = 0 Then
                strNewStyle = strTX1
            Else
                strNewStyle = strTX
            End If

            ' Change the style of the paragraph in question
            ' This is where it will error if no style present
            thisDoc.Paragraphs(A).Style = strNewStyle

        ElseIf A < lngParaCount Then ' it is already a Macmillan style
        ' but can't check next para if it's the last para
        
On Error GoTo ErrorHandler2
            ' is it a chap head?
            If InStr(strCurrentStyle, "(cn)") > 0 Or _
                InStr(strCurrentStyle, "(ct)") > 0 Or _
                InStr(strCurrentStyle, "(ctnp)") > 0 Then

                strNextStyle = thisDoc.Paragraphs(A + 1).Style

                ' is the next para non-Macmillan (and thus should be COTX1)
                If Right(strNextStyle, 1) <> ")" Then     ' it's not a Macmillan style
                    ' so it should be COTX1
                    ' Will error if this style not present in doc
                    thisDoc.Paragraphs(A + 1).Style = strCOTX1
                Else ' it IS a Macmillan style too
                    ' it IT a chap opener? (can have CN followed by CT)
                    If InStr(strNextStyle, "(cn)") > 0 Or _
                        InStr(strNextStyle, "(ct)") > 0 Or _
                        InStr(strNextStyle, "(ctnp)") > 0 Then

                        strNextNextStyle = thisDoc.Paragraphs(A + 2).Style

                        If Right(strNextNextStyle, 1) <> ")" Then ' it's not Macmillan
                            ' so it should be COTX1
                            thisDoc.Paragraphs(A + 2).Style = strCOTX1
                        End If
                    End If
                End If
            Else
                ' It's a styled para but NOT a chap head, just move on
            End If
        End If
    Next A

On Error GoTo 0

    ' Change Normal (Web) back
    thisDoc.Styles("Normal (Web),_").NameLocal = "Normal (Web)"

    Exit Sub

ErrorHandler1:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        ' so we'll add it
        Set myStyle = thisDoc.Styles.Add(Name:=strNewStyle, Type:=wdStyleTypeParagraphOnly)
        ' wdStyleType is enumerated here <https://msdn.microsoft.com/en-us/library/office/ff196870(v=office.15).aspx>
        ' but mysteriously does NOT include wdStyleTypeParagraphOnly. Using wdStyleTypeParagraph
        ' can (always?) create a linked style, wherein the paragraph style is linked to a
        ' character style, and the paragraph style now has that character formatting, and you can't
        ' unlink the styles even from the Modify menu. It's a mess -- just use wdStyleTypeParagraph,
        ' which creates a new paragraph style based on Normal w/ no additional formatting.
        
        With myStyle
            '.QuickStyle = True ' not available for Mac

            ' If we set the style to "(no style)", the new style picks up the original
            ' direct formatting of the paragraph in it's definition, which could be anything.
            ' If we keep BaseStyle = Normal, it could also technically be anything. I'm
            ' going to assume that Normal is more likely to have reasonable formatting
            ' and stay with that as a base style. If this proves to not work out, we'll
            ' have to define every possible formatting option when creating the new style.
            
            ' If it all works out, we'll just have to define the most important formatting,
            ' and hope that Normal doesn't have crazy things like borders or red text.
            
            ' ALSO! If you set a format to something that is the same as Normal, it WON'T
            ' be added to the style definition of the new style. So we may have to define
            ' everything at some point anyway...
            With .Font
                .Name = "Times New Roman"
                .Size = 12
            End With
            
            With .ParagraphFormat
                .LineSpacingRule = wdLineSpaceDouble
                .SpaceAfter = 0
                .SpaceBefore = 0
                .Borders(wdBorderLeft).Color = RGB(102, 204, 255)
'                .Borders.DistanceFromLeft = 4 ' in points
                
                ' Different settings for each style
                Select Case strNewStyle
                    Case strTX
                        .FirstLineIndent = 36 ' default unit is points, 36pt = 0.5in
                        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                        .Borders(wdBorderLeft).LineWidth = wdLineWidth600pt

                    Case strTX1
                        .FirstLineIndent = 0
                        .Borders(wdBorderLeft).LineStyle = wdLineStyleDouble
                        .Borders(wdBorderLeft).LineWidth = wdLineWidth225pt

                End Select
            End With
        End With

        ' Now go back and try to assign that style again
        Resume

    Else
        Debug.Print "ErrorHandler1: " & Err.Number & " " & Err.Description
        On Error GoTo 0
        Call Cleanup
        Exit Sub
    End If

    Exit Sub

ErrorHandler2:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        Debug.Print Err.Number & ": " & Err.Description
        Set myStyle = thisDoc.Styles.Add(Name:=strCOTX1, Type:=wdStyleTypeParagraphOnly)
        With myStyle
            '.QuickStyle = True ' not available for Mac
            ' Set BaseStyle to nothing before adding features to clear
            ' other formatting from orig style.
            .BaseStyle = vbNullString
            ' will error if no TX1 in doc, so go back to Error1 and create it
On Error GoTo ErrorHandler1
            strNewStyle = strTX1
            .BaseStyle = strNewStyle
            With .ParagraphFormat
                .SpaceBefore = 144      ' in points
                With .Borders(wdBorderLeft)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = wdLineWidth600pt
                    .Color = RGB(0, 255, 0)
                End With
            End With
        End With
        ' Reset On Erro to ErrorHandler2, or else it will continue to go
        ' to ErrorHandler1 for any future errors
On Error GoTo ErrorHandler2
        ' Now go back and try to assign that style again
        Resume
    Else
        Debug.Print "ErrorHandler2: " & Err.Number & " " & Err.Description
        On Error GoTo 0
        Call Cleanup
        Exit Sub
    End If
    
    Exit Sub
    
End Sub

Sub test()
End Sub
