Sub Run_Script()
    'Varible Naming Standards here are terrible, this needs to be chagned to be more clear to read'
    
    '*Header*'
    Dim Choice_Selection As Integer
    Choice_Selection = Sheet1.Range("D2").Value
    
    Dim Max_Group_Sizes As Integer
    Max_Group_Sizes = Sheet1.Range("A5").Value
    
    Dim Num_Participants As Integer
    Num_Participants = Sheet1.Range("A2").Value
    
    Dim Required_Groups As Integer
    Required_Groups = Get_Groups(Num_Participants, Max_Group_Sizes)
    Sheet1.Range("L2").Value = Required_Groups
    
    Dim Start_Time As Date
    Start_Time = Sheet1.Range("P2").Value
    
    
    '*Making Times*'
    Dim Session_Selection As Integer
    Session_Selection = Sheet1.Range("D2").Value
    
    Dim Num_Karting As Integer
    Dim Num_Laser As Integer
    If Session_Selection = 6 Then
        Dim Session As Variant
        Session = Get_Custom_Sessions()
        Num_Karting = Session(0)
        Num_Laser = Session(1)
    Else
        Num_Karting = Get_Session_Num(Session_Selection)(0)
        Num_Laser = Get_Session_Num(Session_Selection)(1)
    End If
    
    
    Dim Max_Num_Sessions As Integer
    Max_Num_Sessions = WorksheetFunction.Max(Num_Karting, Num_Laser) 'Finds the max value between the two sessions'
    
    'Makes Times'
    Dim Times As Variant
    Times = Make_Times(Required_Groups, Max_Num_Sessions, Sheet1.Range("A8").Value)
    
    'Set Finish Time Cell'
    Sheet1.Range("Q2").Value = Sheet1.Range(Times).Value
    
    'Finds Total Time & Set Total Time Cell'
    Total_Time_Seconds = DateDiff("s", Start_Time, Sheet1.Range(Times).Value)
    Sheet1.Range("N2").Value = Total_Time_Seconds / 3600
    
    'Makes Session'
    Sessions = Make_Sessions(Num_Karting, Num_Laser, Required_Groups)
    
    Debug.Print "------------------------------------------------------------------------"
    
    
End Sub
 
'Returns Number of Groups to be used | Take max size of groups & total number of attendies as parameters'
Public Function Get_Groups(Num_Parts As Integer, Max_Group_Size As Integer)
    Dim Num_Groups 'Number of total groups'
    
    'Num Group is quotient of Num_Parts to Max_Group_Size | Cannot divided by 0, must check before returning value'
    If Max_Group_Size > 0 And Max_Group_Size < 12 Then
        Num_Groups = Num_Parts \ Max_Group_Size '\ Int Division Operator'
        Get_Groups = Num_Groups 'returns number of groups required'
        
        If Num_Parts Mod Max_Group_Size <> 0 Then 'Check if division has no remainder | If remainder, +1 groups required | <> is != operator, VBA suck :/'
            Get_Groups = Num_Groups + 1
        Else
            Get_Groups = Num_Groups
        End If
    Else
        Err.Raise vbObjectError + 513, "Module1.Num_Group_Size()", "Unable to pass Max_Group_Size ammount, select amount > 0 & =< 11" 'Throws error if inputted amount incorrect'
        Get_Groups = 0 'Returns 0 as error'
    End If
        
End Function
 
 
Public Function Get_Session_Num(Session_ID As Integer)
    'Session ID corrisponding towards Session Choice | e.g 1Kart 2Laser = 1, 2Kart 2Laser = 2 etc.'
    'There is a more optimised method for this, but I cant be arsed to make it :)'
    If Session_ID > 7 Or Session_ID < 0 Then
        'check if ID passed is within range of ID's, throw error if not'
        Err.Raise vbObjectError + 513, "Module1.Get_Session_Num", "Unable to pass Session_ID. Check case method function to see what is being passed in"
    Else
        Dim Session_Total(1) As Variant 'create array of lenght 2 (indexed begining at 0) to hold how many of each session an ID has'
        
        Select Case Session_ID 'works out how many session of Karting & Laser are needed'
            'Index 0 = number of karting sessions'
            'Index 1 = number of laser sessions'
            
            Case 1 '1Kart 2Laser'
                Session_Total(0) = 1
                Session_Total(1) = 2
            
            Case 2 '2Kart 2Laser'
                Session_Total(0) = 2
                Session_Total(1) = 2
            
            Case 3 '3Kart 2laser'
                Session_Total(0) = 3
                Session_Total(1) = 2
            
            Case 4 '2Laser'
                Session_Total(0) = 0
                Session_Total(1) = 2
            
            Case 5 '3Laser'
                Session_Total(0) = 0
                Session_Total(1) = 3
                
            Case 6 'Custom | This condition shouldn't be met from this loop, and show be caught so that this whole function isn't ran'
                Err.Raise vbObjectError + 513, "Module1.Num_Group_Size()", "Unable to pass Max_Group_Size ammount, select amount > 0 & =< 11" 'Throws error if this condition is somehow met'
                
        End Select
                
            Get_Session_Num = Session_Total 'returns array of session'
    End If
        
End Function
Public Function Get_Custom_Sessions()
    Dim Session_Total(1) As Variant 'create array of lenght 2 (indexed begining at 0) to hold how many of each session an ID has'
    Session_Total(0) = InputBox("How Many Karting Sessions")
    Debug.Print Session_Total(0) & " <- Custom Karting"
    Session_Total(1) = InputBox("How Many Laser Sessions")
    Debug.Print Session_Total(1) & " <- Custom Laser"
    
    Get_Custom_Sessions = Session_Total
End Function
 
Public Function Make_Times(Required_Num_Groups As Integer, Most_Num_Sessions As Integer, Start_Time As Date)
    
    'clear cells to be used later as times, karting, and laser'
    Range("H2:H60").ClearContents
    Range("I2:I60").ClearContents
    Range("J2:J60").ClearContents
    
    'This was made into a parameter'
'    Dim Start_Time As Date
'    Start_Time = Sheet1.Range("A8").Value
    
    'Set first cell in time collum to Start_Time'
    Sheet1.Range("H2").Value = Start_Time
    
    Dim Selected_Time As Date
    Selected_Time = Start_Time 'To hold value of selected time for a later FOR Loop'
    
    Dim End_Count As Integer
    End_Count = (Most_Num_Sessions * Required_Num_Groups) + 2 '+2 as correction amount. I'm not sure why this works as +2 and not +1, but if it aint broke, dont fix it :)'
    
    Dim Counter As Integer
    For Counter = 3 To End_Count
        Selected_Time = DateAdd("n", 15, Selected_Time)
        Cells(Counter, "H").Value = Selected_Time
    Next Counter
    
    Cells(End_Count, "I").Value = "FIN" 'Marks last karting session as Finish Time'
    Cells(End_Count, "J").Value = "FIN" 'Marks last laser session as Finish Time'
    
    Dim End_Cell As String
    End_Cell = "H" & End_Count 'holds location of finishing cell'
    
    
    Sheet1.Range("P2").Value = Sheet1.Range("H2").Value
    Sheet1.Range("Q2").Value = Cells(End_Count, "H").Value
    
    Make_Times = End_Cell 'returns value of last time made'
    
End Function
 
 
Public Function Make_Sessions(Num_Kart As Integer, Num_Laser As Integer, Num_Groups As Integer)
    
    'Array to hold what sessions are to be run for each group'
    Set Venue = CreateObject("System.Collections.ArrayList")
    
    'Appending to Sessions To Array'
    For i = 1 To Num_Kart
        Venue.Add "Karting"
    Next i
        
    For i = 1 To Num_Laser
        Venue.Add "Laser"
    Next i
    
    
    
    'Putting Sessions In Sheet'
    Dim Karting_Pointer As Integer 'Points to latest cell ammended in karting collum'
    Dim Laser_Pointer As Integer 'Points to latest cell ammended in laser collum'
    
    Karting_Pointer = 2 'First cell to be ammneded is allways in position 2'
    Laser_Pointer = 2
    
    Dim G As String
    G = "G" 'to be used to concatinate to group number later'
    
    Set Laser_Group_Order_Arr = CreateObject("System.Collections.ArrayList") 'To be used to amned laser session order to stop overlap with karting'
    If Num_Groups > 4 Then
        For i = 1 To Num_Groups
            GroupSTR = G & i
            Laser_Group_Order_Arr.Add GroupSTR
        Next i
    ElseIf Num_Groups = 4 Then 'Manual Group Orders | Only way to organise with low group numbers'
        Laser_Group_Order_Arr.Add G & "3"
        Laser_Group_Order_Arr.Add G & "4"
        Laser_Group_Order_Arr.Add G & "1"
        Laser_Group_Order_Arr.Add G & "2"
    
    ElseIf Num_Groups = 3 Then 'Only need to test for 4&3 group sizes | Min required groups is 3'
        Laser_Group_Order_Arr.Add G & "3"
        Laser_Group_Order_Arr.Add G & "1"
        Laser_Group_Order_Arr.Add G & "2"
    End If
    
    If Num_Groups > 4 Then 'Last two groups have laser first | After last two groups finish laser game, normal assending order continues | Only works with >3 Groups'
        For i = 2 To 1 Step -1
            Laser_Group_Order_Arr.Insert 0, Laser_Group_Order_Arr(Laser_Group_Order_Arr.Count - i)
        Next i
        
        For i = 1 To 2
            Last_Index = Laser_Group_Order_Arr.Count - 1
            Laser_Group_Order_Arr.RemoveAt Last_Index
        Next i
        
        For i = 0 To Laser_Group_Order_Arr.Count - 1
            Debug.Print Laser_Group_Order_Arr.Item(i)
        Next i
    
    End If
    
    Dim Pointer As Integer
    Pointer = 0
    
    Dim Session As Variant
    If Num_Kart <> 0 Then 'checks for just laser sessions?'
        For Each Session In Venue
            Select Case Session
                Case "Karting"
                    For i = 1 To Num_Groups
                        Cells(Karting_Pointer, "I").Value = G & i 'concatinates "G" to the group number, outputting G1,G2,G3 etc.'
                        Karting_Pointer = Karting_Pointer + 1 'iterates pointer'
                    Next i
                Case "Laser"
'                    For i = Num_Groups To 1 Step -1 'reverse order'
'                        Cells(Laser_Pointer, "J").Value = G & i
'                        Laser_Pointer = Laser_Pointer + 1 'iterates pointer'
'                    Next i
                    For i = 1 To Num_Groups
                        Debug.Print i
                        Debug.Print Pointer & "<- Before"
                        Cells(Laser_Pointer, "J").Value = Laser_Group_Order_Arr(Pointer)
                        Laser_Pointer = Laser_Pointer + 1 'iterates pointer'
                        Pointer = Pointer + 1
                    Next i
                    
                    Pointer = 0
                    
                    
            End Select
            'After For loops processed, we know 2 things:'
            '1) If Karting AND-OR Laser Pointer = 2, then no karting and or laser has been chosen'
            '2) The position of the pointer once complete is the end timing. A check needs to be made for the pointer with the higher value to know this'
        Next
    Else
        'If no karting, then group order does not need to be in reverse for laser'
        'Same code as above karting for loop, but working with laser instead'
        For Each Session In Venue
            For i = 1 To Num_Groups
                Cells(Laser_Pointer, "J").Value = G & i
                Laser_Pointer = Laser_Pointer + 1 'iterates pointer'
            Next i
        Next Session
    End If
    
    
End Function
 
 
Public Function Find_Max_Wait(Fin_Time As Date, Group_Num As Integer)
    
End Function
 
 
'Bad Implimentation Below for making groups, not in use'
'Public Function Make_Groups(Sessions_Per_Group As Integer, Groups_Required As Integer, Last_Session As String)
'    Dim Sessions_To_Groups As Integer 'number of sessions made for all groups'
'    Dim Group_Num As Integer 'current group being made & placed into cell'
'
'    Dim Karting_Target_Cell As Integer
'    Karting_Target_Cell = 2
'
'    Dim G As String
'    G = "G"
'
'    'Making Karting Groups'
'    For Sessions_To_Groups = 1 To Sessions_Per_Group
'        For Group_Num = 1 To Groups_Required
'            Cells(Karting_Target_Cell, "I").Value = G & Group_Num 'concatinates string "G" to the group number, outputting G1,G2,G3 etc.'
'            Karting_Target_Cell = Karting_Target_Cell + 1
'        Next Group_Num
'    Next Sessions_To_Groups
'
'
'    'Making Laser Groups'
'
'    Dim Laser_Target_Cell As String
'    Laser_Target_Cell = Sheet1.Range(Last_Session).Row
'
'    Debug.Print Laser_Target_Cell
'
'    For Sessions_To_Groups = 1 To Sessions_Per_Group
'        For Group_Num = 1 To Groups_Required
'            Cells(Karting_Target_Cell, "J").Value = G & Group_Num 'concatinates string "G" to the group number, outputting G1,G2,G3 etc.'
'            Karting_Target_Cell = Laser_Target_Cell - 1
'        Next Group_Num
'    Next Sessions_To_Groups
'
'End Function
