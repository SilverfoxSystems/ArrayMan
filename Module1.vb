'This example is meant to be used in 64 bit console app

Imports System.Runtime.InteropServices


Module Module1

  <DllImport("Arrayman.dll", EntryPoint:="SortLng")>
    Sub sortLong(ByRef Array1stItem As Long, ByRef Indices1stItem As Integer, ByVal nItemsToSort As Long)
    End Sub
    
<DllImport("Arrayman.dll", EntryPoint:="SortULng")>
    Sub sortULong(ByRef Array1stItem As Long, ByRef Indices1stItem As Integer, ByVal nItemsToSort As Long)
    End Sub

<DllImport("Arrayman.dll", EntryPoint:="SortDbl")>
    Sub sortDouble(ByRef Array1stItem As Double, ByRef Indices1stItem As Integer, ByVal nItemsToSort As Long)
    End Sub
    
    
        Sub Main()
        sortExample()
    End Sub

    Private Sub sortExample()
    Dim u% = 10000

        Dim aD(,,) As Double
        Dim aL&(,,)
        Dim aaL&(u)
        Dim aaD#(u)
        Dim k%(u)
        'arrays "aaL" and "aaD" are used for testing the default sorting procedure
        ReDim aD(3, 2, u)
        ReDim aL(3, 2, u)
        Dim indx%(u)

        ' Dim gc1 As GCHandle = GCHandle.Alloc(aL, GCHandleType.Pinned)
        ' Dim gc3 As GCHandle = GCHandle.Alloc(aD(1, 1, 0), GCHandleType.Pinned)
        'Dim gc2 As GCHandle = GCHandle.Alloc(indx(0), GCHandleType.Pinned)

        Dim t1, t2, t3, t4 As Double 'timers

        Randomize()
        Dim i%, i2%
        
j1:

        'Fill up the arrays with random numbers
        For i% = 0 To u
            aL(1, 1, i) = 2000 - Rnd() * 4321 'We'll sort numbers contained in aD(1, 1, {0 to u})
            aD(1, 1, i) = aL(1, 1, i)
            aaL(i) = aL(1, 1, i)
            aaD(i) = aD(1, 1, i)
        Next
        t1# = DateAndTime.Timer
    
        sortDouble(aD(1, 1, 0), indx(0), indx.Count) ' do the sorting
        'sortLong(aL(1, 1, 0), indx(0), indx.Count)
        
        t2# = DateAndTime.Timer
    ' here, aD(1,1, indx(0)) holds the lowest number
    ' and aD(1,1, indx(u)) contains the highest one
        
        ' test the time taken for the default procedure
        For i = 0 To u
            k(i) = i
        Next
        t3# = DateAndTime.Timer
        Array.Sort(aaD, k)
        t4# = DateAndTime.Timer

    'Confirm the result
        Dim lowest& = aL(1, 1, indx(0))
        Dim i2% = 0
        For i2% = 0 To u
            If lowest > aL(1, 1, indx(i2)) Then
                MsgBox("Array was sorted incorrectly.")
                End
            Else
                lowest = aL(1, 1, indx(i2))
            End If
         Next
        
        If ii < 5 Then
            Console.WriteLine("")
            Console.WriteLine("ArrayMan.dll 'SortLng' completion time: " & t2 - t1 & " sec")
            Console.WriteLine("    Default Array.Sort completion time: " & t4 - t3 & " sec")
            Console.WriteLine("Press any key")
            Console.ReadKey() 'wait for key press
            ii += 1
            GoTo j1 'Loop for 5 times
        End If
    End Sub

End Module
