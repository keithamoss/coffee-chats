Public DebugOn As Boolean
Public Worksheet

'-- @TODO Implement a decent algorithm for a round-robin tournament that
'-- randomises selections and allows people to join after the competition has begun

Sub Main()
    DebugOn = False
    Set Worksheet = Sheets("Sheet1")
    PrintDebug (Chr(10) & "#################")
    
    '-- First, work out how many possible coffee partners each
    '-- person has.
    Dim potentialCoffeePartners As Collection
    Dim matchesPerPerson As Collection
    Dim matchesForPerson As Collection
    Set potentialCoffeePartners = GetPotentialCoffeePartners()
    Set matchesPerPerson = potentialCoffeePartners("matchesPerPerson")
    Set matchesForPerson = potentialCoffeePartners("matchesForPerson")
    
    '-- Second, sort by the number of possible matches each person has
    '-- so that we'll process the person with the fewest matches first.
    '-- This is a naive approach to making sure that we maximise the
    '-- chances of someone getting a coffee date.
    Dim sortedMatchesPerPerson As Collection
    Set sortedMatchesPerPerson = SortCollection(matchesPerPerson)
    PrintDebug ("matchesPerPerson: " & CollectionToString(matchesPerPerson, ", "))
    PrintDebug ("sortedMatchesPerPerson: " & CollectionToString(sortedMatchesPerPerson, ", "))
    PrintDebug ("---")
    
    '-- Third, go through and assign coffee date partners to everyone.
    Dim coffeePartners As Collection
    Set coffeePartners = AssignCoffeePartners(sortedMatchesPerPerson, matchesForPerson)
    PrintDebug ("---")
    
    '-- Lastly, create a report on all of the matches and copy it
    '-- to our clipboard.
    PrintDebug ("Report")
    nextRowNum = GetNextAvailableRow()
    foo = WriteCoffeePartnersRow(coffeePartners, nextRowNum)
    foo = CopyCoffeePartnersReport(coffeePartners, nextRowNum)
End Sub

Function GetPotentialCoffeePartners() As Collection
  Dim response As New Collection
  Dim matchesPerPerson As New Collection
  Dim matchesForPerson As New Collection
  
  For Each Item In GetColumns()
      Dim personName As String
      personName = Worksheet.Cells(1, Item.Column)
      PrintDebug (personName)
      
      Dim possiblePartners As Collection
      Set possiblePartners = GetPossibleCoffeePartners(personName, Item.Column)
      PrintDebug ("possiblePartners: " & CollectionToString(possiblePartners, ", "))
      
      matchesPerPerson.Add possiblePartners.Count & "#" & personName, personName
      matchesForPerson.Add possiblePartners, personName
      
      PrintDebug ("---")
  Next

  response.Add matchesForPerson, "matchesForPerson"
  response.Add matchesPerPerson, "matchesPerPerson"
  Set GetPotentialCoffeePartners = response
End Function

Function GetPossibleCoffeePartners(personName As String, colNumber As Integer) As Collection
    Dim partners As New Collection
    
    Dim spokenTo As Collection
    Set spokenTo = HasAlreadySpokenTo(personName, colNumber)
    
    For Each rw In GetColumns()
        cell = Worksheet.Cells(1, rw.Column)
        If cell <> personName And Contains(spokenTo, cell) = False Then
            partners.Add cell, cell
        End If
    Next
    
    Set GetPossibleCoffeePartners = partners
End Function

Function HasAlreadySpokenTo(personName As String, colNumber As Integer) As Collection
    Dim spokenTo As New Collection
    
    For Each rw In GetRows()
        cell = Worksheet.Cells(rw.Row, colNumber)
        If cell <> "" Then
            '-- PrintDebug "cell: " & cell
            spokenTo.Add cell, cell
        End If
    Next

    Set HasAlreadySpokenTo = spokenTo
End Function

Function AssignCoffeePartners(sortedMatchesPerPerson As Collection, matchesForPerson As Collection) As Collection
  Dim coffeePartners As New Collection
  PrintDebug ("AssignCoffeePartners")
    
  For Each Item In sortedMatchesPerPerson
      Dim personName As String
      Dim possiblePartnersActual As Collection
      personName = Split(Item, "#")(1)
      
      If Contains(coffeePartners, personName) = False Then
          Set possiblePartnersActual = GetAvailableMatchesThisRound(matchesForPerson(personName), coffeePartners)
          
          If possiblePartnersActual.Count > 0 Then
              Dim coffeePartner As String
              coffeePartner = GetRandomItemFromCollection(possiblePartnersActual)
              PrintDebug (personName & " (" & possiblePartnersActual.Count & "): " & coffeePartner)
              
              coffeePartners.Add coffeePartner, personName
              coffeePartners.Add personName, coffeePartner
          Else
              PrintDebug (personName & " has no possible partners")
          End If
      Else
          '-- PrintDebug (personName & " already chosen")
      End If
  Next

  Set AssignCoffeePartners = coffeePartners
End Function

Function WriteCoffeePartnersRow(coffeePartners As Collection, nextRowNum)
    Worksheet.Cells(nextRowNum, 1).NumberFormat = "mmm-yy"
    Worksheet.Cells(nextRowNum, 1) = Date
    
    For Each rw In GetColumns()
        personName = Worksheet.Cells(1, rw.Column)
        
        If Contains(coffeePartners, personName) = True Then
            coffeePartner = coffeePartners(personName)
            PrintDebug (personName & ": " & coffeePartner)
            Worksheet.Cells(nextRowNum, rw.Column) = coffeePartner
        End If
    Next
End Function

Function CopyCoffeePartnersReport(coffeePartners As Collection, nextRowNum)
    Dim coffeePartnersReport As New Collection
    Dim coffeePartnersWithoutPartnersReport As New Collection
    
    For Each rw In GetColumns()
        personName = Worksheet.Cells(1, rw.Column)
        
        If Contains(coffeePartners, personName) = True And Contains(coffeePartnersReport, personName) = False Then
            coffeePartner = coffeePartners(personName)
            coffeePartnersReport.Add personName & ": " & coffeePartner, coffeePartner
        ElseIf Worksheet.Cells(nextRowNum, rw.Column) = "" Then
            '-- PrintDebug ("Add " & personName & " to withoutPartners")
            coffeePartnersWithoutPartnersReport.Add personName
        End If
    Next
    
    report = CollectionToString(coffeePartnersReport, Chr(10))
    If coffeePartnersWithoutPartnersReport.Count > 0 Then
        report = report & Chr(10) & Chr(10) & "No partners found for: " & Chr(10) & CollectionToString(coffeePartnersWithoutPartnersReport, Chr(10))
    End If
    
    Worksheet.Range("A1") = report
    Worksheet.Range("A1").Copy
    MsgBox "Coffee partners generated. Please paste the list before closing this dialog."
    Worksheet.Range("A1") = ""
End Function

Function GetAvailableMatchesThisRound(matches As Collection, alreadyMatched As Collection) As Collection
    Dim actualMatches As New Collection
    
    For Each Match In matches
        If Contains(alreadyMatched, Match) = False Then
            actualMatches.Add Match
        End If
    Next
    
    Set GetAvailableMatchesThisRound = actualMatches
End Function

Function GetRandomItemFromCollection(col As Collection) As Variant
    GetRandomItemFromCollection = col(Int(col.Count * Rnd + 0.999999))
End Function

Function GetLastRow() As Integer
    GetLastRow = Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Row
End Function

Function GetLastColumn() As Integer
    Dim lastRow As Long, lastCol As Long
    Set sh = Worksheet
    lastRow = sh.Cells.Find("*", sh.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lastCol = sh.Cells.Find("*", sh.Cells(1, 1), xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    GetLastColumn = sh.Cells(lastRow, lastCol).Column
End Function

Function GetColumns() As Variant
    Set GetColumns = Worksheet.Range(Columns(2), Columns(GetLastColumn()))
End Function

Function GetRows() As Variant
    Set GetRows = Worksheet.Range("A2:A" & GetLastRow())
End Function

Function GetNextAvailableRow() As Integer
    GetNextAvailableRow = Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Row + 1
End Function

Sub PrintDebug(str As String)
    If DebugOn = True Then
        Debug.Print str
    End If
End Sub





