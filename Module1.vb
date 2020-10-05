'This example is meant to be used in 64 bit console app

Imports System.Runtime.InteropServices


Module Module1

    <DllImport("Arrayman.dll", EntryPoint:="SortLng")>
    Sub sort(ByRef Array1stItem As Long, ByRef Indices1stItem As Integer, ByVal nItemsToSort As Long)
        '  Note: For sorting ULong integers, replace EntryPoint:="SortLng" with EntryPoint:="SortULng"
        '        For sorting Double or Date arrays, use EntryPoint:="SortDbl"
    End Sub

    Sub Main()
        sortLngExample()
    End Sub

    Private Sub sortLngExample()

        Dim u% = 1000
        Dim a(,,), aa(u) As Long
        'array "aa" is used for testing the default sorting procedure
        ReDim a(3, 2, u)
        Dim indx%(u)


        Dim t1, t2, t3, t4 As Double 'timers

        
        
        Randomize()
        Dim i% = 0 'Fill up the arrays with random numbers
        For i% = 0 To u 
            a(1, 1, i) = Rnd() * 43214321  'We'll sort numbers contained in a(1, 1, {0 to u})
            aa(i) = a(1, 1, i)
        Next

        t1# = DateAndTime.Timer
        SortLng(a(1, 1, 0), indx) ' do the sorting
        t2# = DateAndTime.Timer
        ' here, a(1,1, indx(0)) holds the lowest number
        ' and a(1,1, indx(u)) contains the highest one

        Dim k%(u)
        For i = 0 To u
            k(i) = i
        Next
        ' test the time taken for the default procedure
        t3# = DateAndTime.Timer
        Array.Sort(aa, k)
        t4# = DateAndTime.Timer

        ' display the times
        For i = 0 To u
            If k(i) <> indx(i) Then
                Console.WriteLine("Results do NOT match") 'This never happens. It serves only as proof of concept
                Exit For
            ElseIf i = u Then
                'Results match
                Console.WriteLine("")
                Console.WriteLine("ArrayMan.dll 'SortLng' completion time: " & t2 - t1 & " sec")
                Console.WriteLine("    Default Array.Sort completion time: " & t4 - t3 & " sec")
                Console.WriteLine("")
                Console.WriteLine("Sorted " & u + 1 & " Long integers (64bit).")
            End If
        Next

        Console.ReadKey() 'wait for key press


    End Sub

    Private Sub SortLng(ByRef array1stItem As Long, ByRef indices() As Integer)
        'A wrapper procedure for easier usage
        sort(array1stItem, indices(0), indices.Count)
    End Sub


End Module
