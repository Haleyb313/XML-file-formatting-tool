VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MsgBoxWarning()
    
    Dim ws As Worksheet
    Dim LastRow As Integer
    Dim rng As Range
    
    Set ws = ActiveSheet
    
    LastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    LastRow = LastRow
    
    Set rng = ws.Range("A2:J" & LastRow)
    
    Result = MsgBox("Updating the table will refresh all old data except for columns Disc., Result, Notes, and Upload to ProjNET. This action cannot be undone. Are you sure you want to continue?", vbYesNo + vbQuestion)

    If Result = vbYes Then
        
        'copy the notes + upload status into the Holding tab
        Call copyNotes
        
        'clear out old data
        rng.Select
        Selection.ClearContents
        Selection.Validation.Delete
    
        'fill in table with new data
        Call updateTable
        
    Else: ''
    End If

    ws.Range("A1").Select

End Sub
    
    
Sub updateTable():

    Application.ScreenUpdating = False

    'set up worksheets
    Dim wsXML As Worksheet
    Set wsXML = ThisWorkbook.Sheets("Copy_XML_table")
    
    Dim wsReview As Worksheet
    Set wsReview = ThisWorkbook.Sheets("Review_Table")

    'get the last row of the XML table
    Dim last_row As Long
    last_row = wsXML.Range("F2", wsXML.Range("F2").End(xlDown)).Rows.count + 1 'add one to skip header

' --------- XML TABLE DATA SETUP --------- '

    'if someone is doing mass uploads of comments, enter their name here to later replace with the actual Evaluator
    Dim uploader As String
    uploader = "Haley Baron"

    'setting up our XML Table columns - these will likely need updating each pull
    Dim xmlID As String
    Dim xmlSpec As String
    Dim xmlSheet As String
    Dim xmlDetail As String
    Dim xmlCommentText As String
    Dim xmlAttachment As String
    Dim xmlDocRef As String
    Dim xmlCreatedBy As String
    Dim xmlCreatedOn As String
    Dim xmlStatusOverall As String
    Dim xmlDiscipline As String
    Dim xmlDocType As String
    Dim xmlEvaluation As String
    Dim xmlEvaluationText1 As String
    Dim xmlEvalAttachment1 As String
    Dim xmlEvalCreatedBy1 As String
    Dim xmlEvalCreatedOn1 As String
    Dim xmlBackChecks As String


    xmlID = "F"
    xmlSpec = "G"
    xmlSheet = "H"
    xmlDetail = "I"
    xmlCommentText = "K"
    xmlAttachment = "L"
    xmlDocRef = "M"
    xmlCreatedBy = "N"
    xmlCreatedOn = "O"
    xmlStatusOverall = "P"
    xmlDiscipline = "Q"
    xmlDocType = "R"
    xmlEvaluation = "U"
    xmlEvaluationText1 = "Y"
    xmlEvalAttachment1 = "Z"
    xmlEvalCreatedBy1 = "AA"
    xmlEvalCreatedOn1 = "AB"
    xmlBackChecks = "AC"


' --------- REVIEW TABLE DATA SETUP --------- '
    
    'set up the Review Table columns we'll be inputing data on
    Dim reviewRow As Integer
    reviewRow = 2 'the review table starts at row 2
    
    'setting up our Review Table columns
    Dim colRow As String
    colRow = "A"
    
    Dim colEvaluation As String
    colEvaluation = "B"
       
    Dim colStatus As String
    colStatus = "C"
    
    Dim colID As String
    colID = "D"
       
    Dim colComment As String 'this is the big comment block
    colComment = "E"
       
    Dim colCommentator As String
    colCommentator = "F"
    
    Dim colEvaluator As String
    colEvaluator = "G"
       
    Dim colDiscipline As String
    colDiscipline = "H"
    
    Dim colResult As String
    colResult = "I"
       
    Dim colNotes As String
    colNotes = "J"
    
    Dim colUploadStatus As String
    colUploadStatus = "K"
       
    'Get a handle to all the stuff in the big text block under the Comment / Evaluation / Backcheck section
    
    'ID : Created by XXX on YYY
    Dim id As String
    Dim submittedByID As String
    Dim submittedOnID As String
    
    'Document Type: XXX
    Dim docType As String
    
    'COMMENT by XXX on YYY
    Dim commentText As String
    
    'EVALUATION by XXX on YYY
    Dim submittedByEv1 As String
    Dim submittedOnEv1 As String
    Dim evaluationText1 As String

    'EVALUATION 2 by XXX on YYY
    Dim evaluationTitle2 As String
    Dim submittedByEv2 As String
    Dim submittedOnEv2 As String
    Dim evaluationText2 As String
    
    'BACKCHECK COMMENT THREAD
    Dim submittedByBC As String
    Dim submittedOnBC As String
    Dim backcheckText As String
    
    'setting up our Review table range that we'll be pasting data into
    Dim rngID As Range
    Set rngID = wsReview.Range("D2:D" & last_row) 'using ID because it will never be blank
    
    Dim row As Range
       
' --------- BUILD THE REVIEW TABLE ---------

    For Each row In rngID:
        
        'get a handle to the XML table row we're copying from
        copyRow = row.row
        
        'fill in the simple stuff
        wsReview.Range(colRow & reviewRow).Value = copyRow
        wsReview.Range(colEvaluation & reviewRow).Value = wsXML.Range(xmlEvaluation & copyRow)
        wsReview.Range(colStatus & reviewRow).Value = wsXML.Range(xmlStatusOverall & copyRow)
        wsReview.Range(colID & reviewRow).Value = wsXML.Range(xmlID & copyRow)
        'colComment taken care of below
        wsReview.Range(colCommentator & reviewRow).Value = wsXML.Range(xmlCreatedBy & copyRow)
        
        'for the Evaluator, make sure the name is correct and not just the uploader for all the comments
            'if name = uploader, instead go into the comment and pull the Evaluators name from the first character until ":"
            'example:  "John Smith: Electrical | MM/DD/YYYY | Text Text Text..."  <-- the below will result in "John Smith"
            If wsXML.Range(xmlEvalCreatedBy1 & copyRow).Value = uploader Then
                Dim evalText As String
                evalText = wsXML.Range(xmlEvaluationText1 & copyRow).Value
                Dim position As Integer
                position = InStr(evalText, ":")
                wsReview.Range(colEvaluator & reviewRow).Value = Left(evalText, (position - 1))
            Else 'If not the uploader, use the name as provided
                wsReview.Range(colEvaluator & reviewRow).Value = wsXML.Range(xmlEvalCreatedBy1 & copyRow)
            End If
            
'        'build the dropdown menu for Result (first remove any old validation)
        wsReview.Range(colResult & reviewRow).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Concur,Non-Concur,Check and Resolve, For Information Only"
            
            
        'the rest of the columns can have input from evaluators, so it was copied and stored that data in the tab "Holding" using the function copyNotes earlier
        'set up the Holding columns to grab data from
        Dim holdID As String
        holdID = "Holding!A:A"
        
        Dim holdDisc As String
        holdDisc = "Holding!B:B"
        
        Dim holdResult As String
        holdResult = "Holding!C:C"
        
        Dim holdNotes As String
        holdNotes = "Holding!D:D"
        
        Dim holdUploadStatus As String
        holdUploadStatus = "Holding!E:E"
        
        Dim checkID As String 'using this just to make reading the forumlas easier
        checkID = "Review_Table!" & colID
        
        'build the INDEX MATCH formulas
        wsReview.Range(colDiscipline & reviewRow).Value = "=IFERROR(INDEX(" & holdDisc & "," & "MATCH(" & checkID & reviewRow & "," & holdID & ", 0)), ""error"")"
        wsReview.Range(colResult & reviewRow).Value = "=IFERROR(INDEX(" & holdResult & "," & "MATCH(" & checkID & reviewRow & "," & holdID & ", 0)), ""error"")"
        wsReview.Range(colNotes & reviewRow).Value = "=IFERROR(IF(" & colUploadStatus & reviewRow & "<>""Yes"", INDEX(" & holdNotes & "," & "MATCH(" & checkID & reviewRow & "," & holdID & ", 0)), """"), """")"
        wsReview.Range(colUploadStatus & reviewRow).Value = "=IFERROR(INDEX(" & holdUploadStatus & "," & "MATCH(" & checkID & reviewRow & "," & holdID & ", 0)), ""No"")"
             
        
        
        'now it's time to fill in the big comment block!
        
        'get a handle to all the big block information
        id = wsXML.Range(xmlID & copyRow).Value
        
        docType = wsXML.Range(xmlDocType & copyRow).Value
        docRef = wsXML.Range(xmlDocRef & copyRow).Value
            'if blank, show n/a instead
            If docRef = "" Then
                docRef = "n/a"
            End If
        docSpec = wsXML.Range(xmlSpec & copyRow).Value
            'if blank, show n/a instead
            If docSpec = "" Then
                docSpec = "n/a"
            End If
        docSheet = wsXML.Range(xmlSheet & copyRow).Value
            'if blank, show n/a instead
            If docSheet = "" Then
                docSheet = "n/a"
            End If
        docDetail = wsXML.Range(xmlDetail & copyRow).Value
            'if blank, show n/a instead
            If docDetail = "" Then
                docDetail = "n/a"
            End If
        
        submittedByID = wsXML.Range(xmlCreatedBy & copyRow).Value
        submittedOnID = wsXML.Range(xmlCreatedOn & copyRow).Value
        commentText = wsXML.Range(xmlCommentText & copyRow).Value
        attachment = wsXML.Range(xmlAttachment & copyRow).Value
                  
        submittedByEv1 = wsXML.Range(xmlEvalCreatedBy1 & copyRow).Value
        submittedOnEv1 = wsXML.Range(xmlEvalCreatedOn1 & copyRow).Value
        evaluationText1 = wsXML.Range(xmlEvaluationText1 & copyRow).Value
        evaluationAttachment1 = wsXML.Range(xmlEvalAttachment1 & copyRow).Value
        
''-------- ADD IN FOLLOW UP EVAL STUFF LATER WHEN YOU HAVE IT
'        submittedByEv2 = wsXML.Range(xmlEvalCreatedBy2 & copyRow).Value
'        submittedOnEv2 = wsXML.Range(xmlEvalCreatedOn2 & copyRow).Value
'        evaluationText2 = wsXML.Range(xmlEvaluationText2 & copyRow).Value
'        evaluationAttachment2 = wsXML.Range(xmlEvalAttachment2 & copyRow).Value

''-------- ADD IN BACKCHECK STUFF LATER WHEN YOU HAVE IT
'        submittedByBC1 = wsXML.Range(xmlBackCheckCreatedBy1 & copyRow).Value
'        submittedOnBC1 = wsXML.Range(xmlBackCheckCreatedOn1 & copyRow).Value
'        backcheckStatus1 = wsXML.Range(xmlBackCheckStatus1 & copyRow).Value
'        backcheckText1 = wsXML.Range(xmlBackCheckText1 & copyRow).Value
'        backcheckAttachment1 = wsXML.Range(xmlBackCheckAttachment1 & copyRow).Value
'
'        submittedByBC2 = wsXML.Range(xmlBackCheckCreatedBy2 & copyRow).Value
'        submittedOnBC2 = wsXML.Range(xmlBackCheckCreatedOn2 & copyRow).Value
'        backcheckStatus2 = wsXML.Range(xmlBackCheckStatus2 & copyRow).Value
'        backcheckText2 = wsXML.Range(xmlBackCheckText2 & copyRow).Value
'        backcheckAttachment2 = wsXML.Range(xmlBackCheckAttachment2 & copyRow).Value
       
       
'--------- Build the Big Comment Block!  looks a bit like below ---------
        
        '        "ID: " & id
        '        "Document Type: " & docType & " | Document Reference: " & docRef
        '        "Spec: " & docSpec & " | Sheet: " & docSheet & " | Detail: " & docDetail
        '        "Attachment: " & attachment  <--- only shown if exists
        
        '        "COMMENT by: " & submittedByID
        '        "Submitted by: " & submittedByID & " | Submitted on: " & submittedOnID
        '        commentText
        
        '        "EVALUATION by: " & submittedByEv1
        '        "Submitted by: " & submittedByEv1 & " | Submitted on: " & submittedOnEv1
        '        evaluationStatus1
        '        evaluationText1
        '        "Attachment: " & evalAttachment1 <--- only shown if exists
        
        '        (2nd evaluation text only shown if exists)
        '        "Submitted by: " & submittedByEv2 & " | Submitted on: " & submittedOnEv2
        '        evaluationStatus2
        '        evaluationText2
        '        "Attachment: " & evalAttachment2 <--- only shown if exists
        
        '        "BACKCHECK COMMENT THREAD"
        '        "Submitted by: " & submittedByBC1 & " | Submitted on: " & submittedOnBC1
        '        backcheckText1
        '        "Attachment: " & backCheckAttachment1 <--- only shown if exists
        
        '        (2nd backcheck only shown if exists)
        '        "Submitted by: " & submittedByBC2 & " | Submitted on: " & submittedOnBC2
        '        backcheckText2
        '        "Attachment: " & backCheckAttachment2
        
        
        Dim block As String 'set up a holding area
               
        'using  Chr(10) for adding a new line within the cell
            'pretty much every line will start a new paragraph in the cell
        
        Dim rngB 'we will be looping through all characters to make some bold and/or underlined
        Set rngB = wsReview.Range(colComment & reviewRow)
        
        'setting up all the sections of text (bold+underlined, regular, bold, regular, underlined, regular, etc.)
        'B indicates bold, U indicates underlined
        
        Dim s1B As String
        s1B = "ID: " & id 'bold + underlined
        
        Dim s2 As String
        s2 = Chr(10) & "Document Type: " & docType & " | Document Reference: " & docRef _
            & Chr(10) & "Spec: " & docSpec & " | Sheet: " & docSheet & " | Detail: " & docDetail
        'if there is an attachment, add a line in
        If Len(attachment) > 0 Then
            s2 = s2 & Chr(10) & "Attachment: " & attachment
        End If
        
        Dim s3B As String
        s3B = Chr(10) & "COMMENT" 'bold
        
        Dim s4U As String
        s4U = Chr(10) & "Submitted by: " & submittedByID & " | Submitted on: " & submittedOnID 'underlined
            

        Dim s5 As String
        s5 = Chr(10) & commentText

        Dim s6B As String
        s6B = Chr(10) & "EVALUATION" 'bold

        Dim s7U As String
        s7U = Chr(10) & "Submitted by: " & submittedByEv1 & " | Submitted on: " & submittedOnEv1 'underlined
        
        Dim s8 As String
        s8 = Chr(10) & evaluationText1
        'if there is an attachment, add a line in
        If Len(evaluationAttachment1) > 0 Then
            s8 = s8 & Chr(10) & "Attachment: " & evaluationAttachment1
        End If
        
''-------- ADD IN FOLLOW UP EVAL STUFF LATER WHEN YOU HAVE IT
'        Dim s9U As String
'        Dim s10 As String
'        'if follow up eval blank, leave blank
'        If Len(submittedByEv2) > 0 Then
'            s9U = s9U & Chr(10) & "Submitted by: " & submittedByEv2 & " | Submitted on: " & submittedOnEv2
'
'            s10 = Chr(10) & evaluationStatus2 & " | " & evaluationText2
'            'if there is an attachment, add a line in
'            If Len(evaluationAttachment2) > 0 Then
'                s10 = s10 & Chr(10) & "Attachment: " & evaluationAttachment2
'            End If
'        Else
'            s9U = ""
'            s10 = ""
'        End If

       
''-------- ADD IN FOLLOW UP BACKCHECK STUFF LATER WHEN YOU HAVE IT
'        Dim s11B As String
'        s11B = Chr(10) & "BACKCHECK COMMENT THREAD"  'bold
'
'        Dim s12U As String
'        Dim s13 As String
'
'        'if backcheck blank, leave blank
'        If Len(submittedByBC1) > 0 Then
'            s12U = Chr(10) & "Submitted by: " & submittedByBC1 & " | Submitted on: " & submittedOnBC1
'
'            s13 = Chr(10) & backcheckStatus1 & " | " & backcheckText1
'            'if there is an attachment, add a line in
'            If Len(backcheckAttachment1) > 0 Then
'                s13 = s13 & Chr(10) & "Attachment: " & backcheckAttachment1
'            End If
'        Else
'            s12U = ""
'            s13 = ""
'        End If
'
'        Dim s14U As String
'        Dim s15 As String
'        'if follow up backcheck blank, leave blank
'        If Len(submittedByBC2) > 0 Then
'            s14U = Chr(10) & "Submitted by: " & submittedByBC2 & " | Submitted on: " & submittedOnBC2
'
'            s15 = Chr(10) & backcheckStatus2 & " | " & backcheckText2
'            'if there is an attachment, add a line in
'            If Len(backcheckAttachment2) > 0 Then
'                s15 = s15 & Chr(10) & "Attachment: " & backcheckAttachment2
'            End If
'        Else
'            s14U = ""
'            s15 = ""
'        End If


        '--------- add the formating! ---------
            'to add bold formatting to specific words, use the below link as reference
            'https://stackoverflow.com/questions/53877637/bolding-text-with-vba
       
        Dim k1 As Integer
        Dim k2 As Integer
        Dim k3 As Integer
        Dim k4 As Integer
        Dim k5 As Integer
'        Dim k6 As Integer 'for future follow up evaluation
'        Dim k7 As Integer
'        Dim k8 As Integer 'for future backcheck
'        Dim k9 As Integer
       
        'build block, no formating yet
        block = s1B & s2 & s3B & s4U & s5 & s6B & s7U & s8 ' & s9U & s10 & s11B & s12U &s13 & s14U & s15
        block = block & Chr(10) 'add one last Chr(10) line to seperate the rows nicely
        
        'have to clean up HTML weird text in block, i.e. remove the <br/> and &quot
        block = WorksheetFunction.Substitute(block, "<br />", Chr(10))
        block = WorksheetFunction.Substitute(block, "&quot;", Chr(34))
        block = WorksheetFunction.Substitute(block, "&amp;", Chr(38))
        
        
        'paste block into the Comment/Eval/BackCheck cell
        wsReview.Range(colComment & reviewRow).Value = block
        
        'add formating to the cell by creating character ranges
        k1 = InStr(block, s1B)  'bold and underlined ID
        k2 = InStr(block, s3B)  'bold COMMENT title
        k3 = InStr(block, s4U)  'underlined Submitted By
        k4 = InStr(block, s6B)  'bold EVALUATION
        k5 = InStr(block, s7U)  'underlined eval Submitted By
'        k6 = InStr(block, s9U)  'underlined follow up eval Submitted By
'        k7 = InStr(block, s11B) 'bold BACKCHECK
'        k8 = InStr(block, s12U) 'underlined backcheck Submitted By
'        k9 = InStr(block, s14U) 'underlined follow up backcheck Submitted By

        With rngB
            rngB.Font.Bold = False
            rngB.Font.Underline = False
            rngB.Characters(k1, Len(s1B)).Font.Bold = True
            rngB.Characters(k1, Len(s1B)).Font.Underline = True
            rngB.Characters(k2, Len(s3B)).Font.Bold = True
            rngB.Characters(k3, Len(s4U)).Font.Underline = True
            rngB.Characters(k4, Len(s6B)).Font.Bold = True
            rngB.Characters(k5, Len(s7U)).Font.Underline = True
'            rngB.Characters(k6, Len(s9U)).Font.Underline = True
'            rngB.Characters(k7, Len(s11B)).Font.Bold = True
'            rngB.Characters(k8, Len(s12U)).Font.Underline = True
'            rngB.Characters(k9, Len(s14U)).Font.Underline = True



        End With
        
        'add a border around this new entry
        Dim rngBorder As String
        rngBorder = (colRow & reviewRow) & ":" & (colUploadStatus & reviewRow)
        
        wsReview.Range(rngBorder).BorderAround xlContinuous, xlThin

       
        reviewRow = reviewRow + 1 'move down to the next row to restart the process to enter data
        
        Next row
       
       
       'finally, copy and paste the formulas as values only
       Dim formulaRange As Range
       Set formulaRange = wsReview.Range(colDiscipline & ":" & colUploadStatus)
       
       formulaRange.Select
       Selection.Copy
       Selection.PasteSpecial Paste:=xlPasteValues
       
       Application.CutCopyMode = False
              
    Application.ScreenUpdating = True

End Sub


Sub copyNotes():

    Application.ScreenUpdating = False

'   For each row, loop through Review Table and add ID + Disc. + Result + Notes + Upload status to tab "Holding"

    Dim wsRT As Worksheet
    Set wsRT = ThisWorkbook.Sheets("Review_Table")
    
    Dim wsH As Worksheet
    Set wsH = ThisWorkbook.Sheets("Holding")
    
'    ' getting the range of Holding list
'    Dim holdingCount As Long
'    holdingCount = wsH.Range("A3", wsH.Range("A3").End(xlDown)).Rows.count + 2 'add two to skip warning and header
'
'    Dim holdingList As Range
'    Set holdingList = wsH.Range("A3:D" & holdingCount)

    Dim pasteRow As String
    pasteRow = "3" ' Holding will always start on row 3

    ' getting the range of the Review Table
    Dim revCount As Long
    revCount = wsRT.Range("D2", wsRT.Range("D2").End(xlDown)).Rows.count + 1 'add one to skip header

    Dim revList As Range
    Set revList = wsRT.Range("D2:D" & revCount) 'using column D - "ID"

    Dim copyID As Range
        
    Dim copyDiscipline As Range
        
    Dim copyResult As Range
        
    Dim copyNotes As Range
    
    Dim copyUploadStatus As Range
    
    Dim row As Range
    
    For Each row In revList
      
        copyRow = row.row
        
        Set copyID = wsRT.Range("D" & copyRow) 'ID is in column D
        Set copyDiscipline = wsRT.Range("H" & copyRow) 'Upload Status is in column H
        Set copyResult = wsRT.Range("I" & copyRow) 'Result is in column I
        Set copyNotes = wsRT.Range("J" & copyRow) ' Notes is in column J
        Set copyUploadStatus = wsRT.Range("K" & copyRow) 'Upload Status is in column K
        
        
        Union(copyID, copyDiscipline, copyResult, copyNotes, copyUploadStatus).Copy

        wsH.Select
        wsH.Range("A" & pasteRow).Select
        ActiveSheet.Paste

        wsRT.Select
        
        pasteRow = pasteRow + 1
                  
    Next row
    

    Application.ScreenUpdating = True

End Sub



