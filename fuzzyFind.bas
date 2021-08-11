'''''''''''''''''''''''''''''''''''''''''''''''
' Fuzzy Find                                  '
'''''''''''''''''''''''''''''''''''''''''''''''
' *** Requires Enum Type "CaseSensitivity" ***
' *** Requires Function "originalMetric" ***
' *** Requires Function "damerau" ***
'       *** Requires Reference "Microsoft Scripting Library" ***
' *** Requires Function "hamming" ***
' *** Requires Function "levenshtein" ***
' *** Requires Function "sorensenDice" ***
'       *** Requires Function "ngrams" ***
' *** Requires Function "tversky" ***
'       *** Requires Function "uniqueArrayElements" ***
' *** Requires Function "jaccard" ***
' *** Requires Function "jaroWinkler" ***
' *** Requires Function "simpleMatching" ***
' *** Requires Function "min" ***
' *** Requires Function "max" ***

'recieves input of
'       query (ex. "foo") as string
'       searchRange (ex. Range(A1:B5)) range
'       searchSheet (ex. wb.Sheets("Sheet 1")) as worksheet

'       optional caseSensitive (ex. False) as caseSensitivity: default = 2

'       optional weights (ex. Array(1, .2, 3, 4, 5, .06, 7, 8, .009)) as Variant: default = Array(1, 1, 1, 1, 1, 1, 1, 1, 1)
'       *** Array must contain exactly nine elements or it will be reverted to default.
'       *** The Lbound base is not relevant.

'       optional tverskySymmetry ex. True) as boolean: default = false
'       *** Determines the symetry of the tversky algorithm.

'       optional tverskyWeights (ex. Array(1, 2)) as variant: default = Array(1, 1)
'       *** Controls the weights of each side of the tversky algorithim.
'       *** Array must contain exactly two elements or it will be reverted to default.
'       *** LBound base is not relevant.


'outputs the closest match to the query text as a string


''' From the Author '''
'@Description: Configurable fuzzy find algorithim for string matching.
'@Author: Justin Icenhour
'@Version: 1.0.0
'@License: GPL-3.0

Function fuzzyFind(query As String, searchRange As range, searchSheet As Worksheet, caseSensitive As CaseSensitivity, Optional weights As Variant, Optional tverskySymmetry As Boolean = False, Optional tverskyWeights As Variant) As String

        'dimension variables
        Dim ws As Worksheet: Set ws = searchSheet
        Dim lookupRange As range: Set lookupRange = searchRange
        Dim cell As range
        Dim i As Integer
        Dim topScore As Double
        Dim currentScore As Double
        Dim topScoringCell As range
        
        Dim originalMetric As Double: originalMetric = 0
        Dim damerau As Double: damerau = 0
        Dim hamming As Double: hamming = 0
        Dim levenshtein As Double: levenshtein = 0
        Dim sorensenDice As Double: sorensenDice = 0
        Dim tversky As Double: tversky = 0
        Dim jaccard As Double: jaccard = 0
        Dim jaroWinkler As Double: jaroWinkler = 0
        Dim simpleMatching As Double: simpleMatching = 0
        
        Dim originalMetricWeight As Double: originalMetricWeight = 1
        Dim damerauWeight As Double: damerauWeight = 1
        Dim hammingWeight As Double: hammingWeight = 1
        Dim levenshteinWeight As Double: levenshteinWeight = 1
        Dim sorensenDiceWeight As Double: sorensenDiceWeight = 1
        Dim tverskyWeight As Double: tverskyWeight = 1
        Dim jaccardWeight As Double: jaccardWeight = 1
        Dim jaroWinklerWeight As Double: jaroWinklerWeight = 1
        Dim simpleMatchingWeight As Double: simpleMatchingWeight = 1
        
        If IsArray(weights) Then
            If (UBound(weights) - LBound(weights) + 1) = 9 Then
                originalMetricWeight = weights(LBound(weights))
                damerauWeight = weights(LBound(weights) + 1)
                hammingWeight = weights(LBound(weights) + 2)
                levenshteinWeight = weights(LBound(weights) + 3)
                sorensenDiceWeight = weights(LBound(weights) + 4)
                tverskyWeight = weights(LBound(weights) + 5)
                jaccardWeight = weights(LBound(weights) + 6)
                jaroWinklerWeight = weights(LBound(weights) + 7)
                simpleMatchingWeight = weights(LBound(weights) + 8)
            End If
        End If
        
        ReDim originalMetricScores(1 To lookupRange.Count) As Double
        ReDim damerauScores(1 To lookupRange.Count) As Double
        ReDim hammingScores(1 To lookupRange.Count) As Double
        ReDim levenshteinScores(1 To lookupRange.Count) As Double
        ReDim sorensenDiceScores(1 To lookupRange.Count) As Double
        ReDim tverskyScores(1 To lookupRange.Count) As Double
        ReDim jaccardScores(1 To lookupRange.Count) As Double
        ReDim jaroWinklerScores(1 To lookupRange.Count) As Double
        ReDim simpleMatchingScores(1 To lookupRange.Count) As Double
        
        Dim originalMetricMinScore As Double
        Dim originalMetricMaxScore As Double
        Dim damerauMinScore As Double
        Dim damerauMaxScore As Double
        Dim hammingMinScore As Double
        Dim hammingMaxScore As Double
        Dim levenshteinMinScore As Double
        Dim levenshteinMaxScore As Double
        Dim sorensenDiceMinScore As Double
        Dim sorensenDiceMaxScore As Double
        Dim tverskyMinScore As Double
        Dim tverskyMaxScore As Double
        Dim jaccardMinScore As Double
        Dim jaccardMaxScore As Double
        Dim jaroWinklerMinScore As Double
        Dim jaroWinklerMaxScore As Double
        Dim simpleMatchingScore As Double
        
        ReDim cellAddressBook(1 To lookupRange.Count) As String
        
        i = 1
        For Each cell In lookupRange
        
            'fill cell addressbook
            cellAddressBook(i) = cell.Address
            
            'original metric
            If originalMetricWeight = 0 Then
                originalMetricScores(i) = 0
            Else
                originalMetric = common.originalMetric(query, cell.value, caseSensitive)
                originalMetricScores(i) = originalMetric
            End If
            
            'damerau metric
            If damerauWeight = 0 Then
                damerauScores(i) = 0
            Else
                damerau = common.damerau(query, cell.value, caseSensitive)
                If Not damerau = 0 Then damerau = 1 / damerau
                damerauScores(i) = damerau
            End If
            
            'hamming metric
            If hammingWeight = 0 Then
                hammingScores(i) = 0
            Else
                hamming = common.hamming(query, cell.value, caseSensitive)
                If Not hamming = 0 Then hamming = 1 / hamming
                hammingScores(i) = hamming
            End If
            
            'levenshtein metric
            If levenshteinWeight = 0 Then
                levenshteinScores(i) = 0
            Else
                levenshtein = common.levenshtein(query, cell.value, caseSensitive)
                If Not levenshtein = 0 Then levenshtein = 1 / levenshtein
                levenshteinScores(i) = levenshtein
            End If
            
            ' sorensen dice metric
            If sorensenDiceWeight = 0 Then
                sorensenDiceScores(i) = 0
            Else
                sorensenDice = common.sorensenDice(query, cell.value, caseSensitive)
                sorensenDiceScores(i) = sorensenDice
            End If
            
            'tversky metric
            If tverskyWeight = 0 Then
                tverskyScores(i) = 0
            Else
                If IsArray(tverskyWeights) Then
                    If (UBound(tverskyWeights) - LBound(tverskyWeights) + 1) = 2 Then
                        tversky = common.tversky(query, cell.value, caseSensitive, tverskySymmetry, CDbl(tverskyWeights(LBound(tverskyWeights))), CDbl(tverskyWeights(LBound(tverskyWeights)) + 1))
                    Else
                        tversky = common.tversky(query, cell.value, caseSensitive, tverskySymmetry)
                    End If
                Else
                    tversky = common.tversky(query, cell.value, caseSensitive, tverskySymmetry)
                End If
                If Not tversky = 0 Then tversky = 1 / tversky
                tverskyScores(i) = tversky
            End If
            
            'jaccard metric
            If jaccardWeight = 0 Then
                jaccardScores(i) = 0
            Else
                jaccard = common.jaccard(query, cell.value, caseSensitive)
                jaccardScores(i) = jaccard
            End If
            
            'jaroWinkler metric
            If jaroWinklerWeight = 0 Then
                jaroWinklerScores(i) = 0
            Else
                jaroWinkler = common.jaroWinkler(query, cell.value, caseSensitive)
                jaroWinklerScores(i) = jaroWinkler
            End If
            
            'simpleMatching metric
            If simpleMatchingWeight = 0 Then
                simpleMatchingScores(i) = 0
            Else
                simpleMatching = common.simpleMatching(query, cell.value, caseSensitive)
                simpleMatchingScores(i) = simpleMatching
            End If
            
            
            i = i + 1
        Next cell
        
        'determin min / max scores
        originalMetricMinScore = common.min(originalMetricScores)
        originalMetricMaxScore = common.max(originalMetricScores)
        damerauMinScore = common.min(damerauScores)
        damerauMaxScore = common.max(damerauScores)
        hammingMinScore = common.min(hammingScores)
        hammingMaxScore = common.max(hammingScores)
        levenshteinMinScore = common.min(levenshteinScores)
        levenshteinMaxScore = common.max(levenshteinScores)
        sorensenDiceMinScore = common.min(sorensenDiceScores)
        sorensenDiceMaxScore = common.max(sorensenDiceScores)
        tverskyMinScore = common.min(tverskyScores)
        tverskyMaxScore = common.max(tverskyScores)
        jaccardMinScore = common.min(jaccardScores)
        jaccardMaxScore = common.max(jaccardScores)
        jaroWinklerMinScore = common.min(jaroWinklerScores)
        jaroWinklerMaxScore = common.max(jaroWinklerScores)
        simpleMatchingMinScore = common.min(simpleMatching)
        simpleMatchingMaxScore = common.max(simpleMatching)
                
        
        For i = 1 To lookupRange.Count
        
            'normailize original metric
            If Not originalMetricWeight = 0 And originalMetricMinScore <> originalMetricMaxScore Then originalMetricScores(i) = (originalMetricScores(i) - originalMetricMinScore) / (originalMetricMaxScore - originalMetricMinScore)
            
            'normalize damerau metric
            If Not damerauWeight = 0 And damerauMinScore <> damerauMaxScore Then damerauScores(i) = (damerauScores(i) - damerauMinScore) / (damerauMaxScore - damerauMinScore)
            
            'normalize hamming metric
            If Not hammingWeight = 0 And hammingMinScore <> hammingMaxScore Then hammingScores(i) = (hammingScores(i) - hammingMinScore) / (hammingMaxScore - hammingMinScore)
            
            'normailize levenshtein metric
            If Not levenshteinWeight = 0 And levenshteinMinScore <> levenshteinMaxScore Then levenshteinScores(i) = (levenshteinScores(i) - levenshteinMinScore) / (levenshteinMaxScore - levenshteinMinScore)
            
            'normalize sorensen dice metric
            If Not sorensenDiceWeight = 0 And sorensenDiceMinScore <> sorensenDiceMaxScore Then sorensenDiceScores(i) = (sorensenDiceScores(i) - sorensenDiceMinScore) / (sorensenDiceMaxScore - sorensenDiceMinScore)
            
            'normalize tversky metric
            If Not tverskyWeight = 0 And tverskyMinScore <> tverskyMaxScore Then tverskyScores(i) = (tverskyScores(i) - tverskyMinScore) / (tverskyMaxScore - tverskyMinScore)
            
            'normalize jaccard metric
            If Not jaccardWeight = 0 And jaccardMinScore <> jaccardMaxScore Then jaccardScores(i) = (jaccardScores(i) - jaccardMinScore) / (jaccardMaxScore - jaccardMinScore)
            
            'normalize jaroWinkler metric
            If Not jaroWinklerWeight = 0 And jaroWinklerMinScore <> jaroWinklerMaxScore Then jaroWinklerScores(i) = (jaroWinklerScores(i) - jaroWinklerMinScore) / (jaroWinklerMaxScore - jaroWinklerMinScore)
            
            'normalize simpleMatching metric
            If Not simpleMatchingWeight = 0 And simpleMatchingMinScore <> simpleMatchingMaxScore Then simpleMatchingScores(i) = (simpleMatchingScores(i) - simpleMatchingMinScore) / (simpleMatchingMaxScore - simpleMatchingMinScore)
        Next i
        
        For i = 1 To lookupRange.Count
        
            currentScore = (originalMetricScores(i) * originalMetricWeight) + (damerauScores(i) * damerauWeight) + (hammingScores(i) * hammingWeight) + (levenshteinScores(i) * levenshteinWeight) + (sorensenDiceScores(i) * sorensenDiceWeight) + (tverskyScores(i) * tverskyWeight) + (jaccardScores(i) * jaccardWeight) + (jaroWinklerScores(i) * jaroWinklerWeight) + (simpleMatchingScores(i) * simpleMatchingWeight)
            
            'if the current cells is the highest scoring cell then record it as the top score and record its range
            If currentScore > topScore Then
                    Set topScoringCell = ws.range(cellAddressBook(i))
                    topScore = currentScore
            End If
            
            'reset current scores and positions
            currentScore = 0
        
        Next i
        
        If topScore = 0 Then
            Set topScoringCell = ws.range(cellAddressBook(Int((UBound(cellAddressBook) * Rnd) + 1)))
        End If
        
        'return the value of the top scoring cell
        fuzzyFind = topScoringCell.value
End Function