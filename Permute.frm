VERSION 5.00
Begin VB.Form fPermute 
   Caption         =   "Permutation, Combination, Partition"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   10680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Page 3"
      Height          =   375
      Index           =   2
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Page 2"
      Height          =   375
      Index           =   1
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Page 1"
      Height          =   375
      Index           =   0
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Integer Factorization, Counting Functions"
      Height          =   6855
      Index           =   2
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   3375
      Begin VB.CommandButton cmdCount 
         Caption         =   "Involution Count"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   65
         Top             =   6120
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Eulerian Number <n k>"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   64
         Top             =   5520
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Bell Number b(n)"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   63
         Top             =   4920
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Stirling Subset {n k}"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   62
         Top             =   4560
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Stirling Cycle [n k]"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   61
         Top             =   4200
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Partition Parts |n k|"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   60
         Top             =   3600
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Partition p(n)"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   59
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Multiset <n k>"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   58
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CommandButton cmdCount 
         Caption         =   "Binomial (n k)"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   2280
         Width           =   2895
      End
      Begin VB.CommandButton cmdFactorAll 
         Caption         =   "All Factorizations"
         Height          =   375
         Left            =   240
         TabIndex        =   56
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdFactorPrime 
         Caption         =   "Prime Factorization"
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtCountK 
         Height          =   285
         Left            =   1320
         TabIndex        =   54
         Text            =   "4"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCountN 
         Height          =   285
         Left            =   1320
         TabIndex        =   53
         Text            =   "12"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "k ="
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "n ="
         Height          =   255
         Left            =   840
         TabIndex        =   51
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tuples, Gray Codes, Partitions"
      Height          =   6855
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   3375
      Begin VB.TextBox txtPartitionMultisetVisit 
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Text            =   "aaabb"
         Top             =   5760
         Width           =   2895
      End
      Begin VB.CommandButton cmdPartitionMultisetVisit 
         Caption         =   "Multiset Partitions"
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   6120
         Width           =   2895
      End
      Begin VB.CommandButton cmdPartitionVisit 
         Caption         =   "Set Partitions of n"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   48
         Top             =   5280
         Width           =   2895
      End
      Begin VB.CommandButton cmdPartitionVisit 
         Caption         =   "Integer Partitions of n into k Parts"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   47
         Top             =   4800
         Width           =   2895
      End
      Begin VB.CommandButton cmdPartitionVisit 
         Caption         =   "Integer Partitions of n"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   4440
         Width           =   2895
      End
      Begin VB.ComboBox cboPartitionVisitK 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cboPartitionVisitN 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton cmdGrayBalancedVisit 
         Caption         =   "Balanced Gray"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CommandButton cmdGrayLongVisit 
         Caption         =   "Long"
         Height          =   375
         Left            =   2520
         TabIndex        =   39
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmdGrayVisit 
         Caption         =   "Loopless Gray"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ComboBox cboGrayCodeVisitN 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdTupleVisit 
         Caption         =   "Loopless Modulo"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton cmdTupleVisit 
         Caption         =   "Loopless Reflected"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   2895
      End
      Begin VB.ComboBox cboTupleVisitM 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cboTupleVisitN 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "k ="
         Height          =   255
         Left            =   2040
         TabIndex        =   43
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "n ="
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Partitions"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   3600
         Width           =   855
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   3240
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label10 
         Caption         =   "n ="
         Height          =   255
         Left            =   1920
         TabIndex        =   36
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Gray codes"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   3240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label8 
         Caption         =   "Tuples"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Permutations, Combinations"
      Height          =   6855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      Begin VB.CommandButton cmdCombinationCountMultiset 
         Caption         =   "Cnt"
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   6120
         Width           =   615
      End
      Begin VB.CommandButton cmdPermutationCountMultiset 
         Caption         =   "Cnt"
         Height          =   375
         Left            =   2520
         TabIndex        =   28
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdCombinationVisitMultiset 
         Caption         =   "All Combinations"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   6120
         Width           =   2295
      End
      Begin VB.CommandButton cmdPermutationVisitMultiset 
         Caption         =   "All Permutations"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   5760
         Width           =   2295
      End
      Begin VB.ComboBox cboMultisetK 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox txtMultiset 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "aaabc"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.ComboBox cboCombinationVisitK 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2760
         Width           =   735
      End
      Begin VB.ComboBox cboCombinationVisitN 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdCombinationVisit 
         Caption         =   "Chase Sequence"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   3960
         Width           =   2895
      End
      Begin VB.CommandButton cmdCombinationVisit 
         Caption         =   "Revolving Door"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   3600
         Width           =   2895
      End
      Begin VB.CommandButton cmdCombinationVisit 
         Caption         =   "Lexicographic"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CommandButton cmdPermutationVisit 
         Caption         =   "Ehrlich Swap"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton cmdPermutationVisitIdx 
         Caption         =   "Idx"
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdPermutationVisit 
         Caption         =   "Plain Changes"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cboPermutationVisitN 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "k ="
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   5220
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Multiset permutations, combinations"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   4800
         Width           =   3015
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3240
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label5 
         Caption         =   "k ="
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   2820
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "n ="
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2820
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Combinations of ""n choose k"" elements"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3240
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label2 
         Caption         =   "n ="
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Permutations of n distinct elements"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "fPermute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'FileName:     Permute.frm
'Author:       John Korejwa <korejwa@tiac.net>
'Description:  Algorithms to count and generate permutations, combinations, tuples, Gray codes, and partitions, and
'              a collection of combinatorics counting functions.  Algorithm numbers refer to "The Art of Computer
'              Programming" fascicles, written by Donald Knuth.
'
'Request:      Please contact the author if you know an efficient way to count multiset partitions.

Private Type RECT
    Left   As Long
    top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const WM_VSCROLL As Long = &H115
Private Const SB_BOTTOM  As Long = 7

Private m_ReserveSpace As RECT

Private m_prime()  As Long
Private m_primes   As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long 'The left and top members are zero. The right and bottom members contain the width and height of the window.
Private Declare Function GetTickCount Lib "kernel32" () As Long


'Add Item to Text1 and scroll to bottom; trim Text1.Text if necessary to comply with TextBox assign limitation
Private Sub AddItemText(ByRef Item As String)
    Dim r As Long, s As Long, t As Long

    With Text1
        s = Len(.Text)
        t = Len(Item)
        If (t > 65533) Then
            r = InStr(t - 65533, Item, vbCrLf)
            If (r = 0) Then
                r = 65533
            Else
                r = t - (r + 1)
            End If
            .Text = Right$(Item, r) & vbCrLf
        ElseIf (s + t > 65533) Then
            r = InStr(s + t - 65533, .Text, vbCrLf)
            If (r = 0) Then
                r = 65533 - t
            Else
                r = s - (r + 1)
            End If
            .Text = Right$(.Text, r) & Item & vbCrLf
        Else
            .Text = .Text & Item & vbCrLf
        End If
        SendMessage .hwnd, WM_VSCROLL, SB_BOTTOM, 0
    End With
End Sub

'Initialize table of prime numbers.  To factor any Long integer, we need all prime numbers <= sqrt(max_long)
'The number of primes <= sqrt(max_long) is 4792, the smallest of which is 2, the largest of which is 46337
Public Sub InitPrimeTable()
    Dim j As Long, n As Long, u() As Long
    Const max_prime As Long = 46337 'floor(sqrt(max_long))

    ReDim m_prime(4791)
    ReDim u(max_prime)
    m_prime(0) = 2
    m_prime(1) = 3
    m_primes = 2
    n = 5
    Do
        If (u(n) = 0) Then 'n is prime
            m_prime(m_primes) = n
            m_primes = m_primes + 1
            If (n = max_prime) Then Exit Do
            For j = n To max_prime Step n
                u(j) = 1
            Next j
        End If
        n = n + 2
        If (u(n) = 0) Then 'n is prime
            m_prime(m_primes) = n
            m_primes = m_primes + 1
            If (n = max_prime) Then Exit Do
            For j = n To max_prime Step n
                u(j) = 1
            Next j
        End If
        n = n + 4
    Loop
End Sub

'factor n into prime product[p(i)^e(i)]  for 0 <= i < m
'returns: m
'InitPrimeTable() must be called before using this function
Public Function FactorInteger(p() As Long, e() As Long, ByVal n As Long) As Long
    Dim i As Long, j As Long, k As Long

    Select Case n
    Case 0              'treat zero as prime, although this is a math error
        p(i) = 0
        e(i) = 0
        i = i + 1
        n = 1
    Case &H80000000     'special case; need to avoid overflow when negating
        p(i) = -1
        e(i) = 1
        i = i + 1
        n = 1
    Case Else
        If (n < 0) Then 'for negative numbers, return -1 as first factor
            p(i) = -1
            e(i) = 1
            i = i + 1
            n = -n
        End If
        j = 0
        k = Int(Sqr(n))
        Do
            If ((n Mod m_prime(j)) = 0) Then
                p(i) = m_prime(j)
                e(i) = 1
                Do
                    n = n \ m_prime(j)
                    If ((n Mod m_prime(j)) <> 0) Then Exit Do
                    e(i) = e(i) + 1
                Loop
                i = i + 1
                k = Int(Sqr(n))
            End If
            j = j + 1
        Loop While (m_prime(j) <= k)
    End Select
    If ((i = 0) Or (n <> 1)) Then
        p(i) = n
        e(i) = 1
        i = i + 1
    End If

    FactorInteger = i
End Function

'Returns: (n k) Binomial Coefficient = n! / (k! * (n-k)!)
'        -1 if result overflows a long
'The number of ways k elements can be chosen from n elements, regardless of order
'InitPrimeTable() must be called before using this function
Public Function BinomialCoefficient(n As Long, ByVal k As Long) As Long
    Const primes_l      As Long = 54 'number of primes in prime table (54 primes_l <= 256)
    Dim g As Long, h As Long, i As Long, j As Long, m As Long, prime_max As Long
    Dim e(primes_l - 1) As Long, f(primes_l - 1) As Long
    prime_max = m_prime(primes_l)

    'check input for error/out of range/trivial solution
    If (n < 0) Or (n >= prime_max) Then GoTo ErrOverflow
    If ((k <= 0) Or (k > n)) Then
        BinomialCoefficient = 0
        Exit Function
    End If

    'calculate directly if overflow is not possible
    If (n <= 12) Then
        i = 1
        If (2 * k > n) Then
            For j = k + 1 To n
                i = i * j
            Next j
            For j = 2 To (n - k)
                i = i \ j
            Next j
        Else
            For j = n - k + 1 To n
                i = i * j
            Next j
            For j = 2 To k
                i = i \ j
            Next j
        End If
        BinomialCoefficient = i
        Exit Function
    End If

    'calculate e(0..m), the exponents of primes representing the result
    m = 0
    For i = 2 To n
        h = i 'factor i, add exponents to f()
        For j = 0 To primes_l - 1
            g = 0
            While ((h Mod m_prime(j)) = 0)
                h = h \ m_prime(j)
                g = g + 1
            Wend
            If (g <> 0) Then
                f(j) = f(j) + g
                If (h = 1) Then Exit For
            End If
        Next j
        If (m < j) Then m = j
        'status: i! = product[prime(0..m)^f(0..m)]
        If (i = k) Then       'divide result by k!
            For j = 0 To m
                e(j) = e(j) - f(j)
            Next j
        End If
        If (i = (n - k)) Then 'subtract exponents denominator divide result by (n-k)!
            For j = 0 To m
                e(j) = e(j) - f(j)
            Next j
        End If
    Next i
    For j = 0 To m            'multiply result by n!
        e(j) = e(j) + f(j)
    Next j

    'result product[prime(0..m)^e(0..m)] can overflow a Long
    On Error GoTo ErrOverflow
    g = 1
    For j = 0 To m
        For i = 1 To e(j)
            g = g * m_prime(j)
        Next i
    Next j
    BinomialCoefficient = g
Exit Function
ErrOverflow:
BinomialCoefficient = -1
End Function

'Returns:  Multinomial Coefficient = n! / (n1! * n2! * ... * nr!)
'        -1 if result overflows a long
'The number of ways a multiset of n elements a(0..n-1), of which n1 are alike, n2 are alike, ... nr are alike,  can be arranged
'Procedure InitPrimes() must be called before using this function
Public Function MultinomialCoefficient(a() As Long, ByVal n As Long) As Long
    Const primes_l      As Long = 54 'number of primes in prime table (54 primes_l <= 256)
    Dim h As Long, i As Long, j As Long, k As Long, q As Long, prime_max As Long
    Dim e(primes_l - 1) As Long, f(primes_l - 1) As Long, v() As Long, m() As Long

    prime_max = m_prime(primes_l)
    ReDim v(prime_max - 2), m(prime_max - 1)

    'check input for error/out of range
    If ((n <= 0) Or (n >= prime_max)) Then GoTo ErrOverflow

    'generate v(), a sorted index of u, so we can ...
    For i = 0 To n - 1
        v(i) = i
    Next i
    For i = 1 To n - 1
        h = v(i)
        k = a(v(i))
        For j = i - 1 To 0 Step -1
            If (a(v(j)) <= k) Then Exit For
            v(j + 1) = v(j)
        Next j
        v(j + 1) = h
    Next i

    '... generate m();  m(i) is number of i-tuplets  (n1, n2, .. nr)
    i = 0
    While (i < n)
        j = 0
        k = a(v(i))
        Do
            i = i + 1
            j = j + 1
            If (i = n) Then Exit Do
        Loop While (a(v(i)) = k)
        m(j) = m(j) + 1
    Wend 'is there a more efficient way to generate m()?

    'calculate directly if overflow is not possible
    If (n <= 12) Then
        j = 1 '       n!
        k = 1 '(n1! * n2! ... nr!)
        For i = 2 To n
            j = j * i
            For h = 1 To m(i)
                k = k * j
            Next h
        Next i
        MultinomialCoefficient = j \ k
        Exit Function
    End If

    'calculate e(0..q), the exponents of primes representing the result
    q = 0
    For i = 2 To n
        h = i 'factor i, add exponents to f()
        For j = 0 To primes_l - 1
            k = 0
            While ((h Mod m_prime(j)) = 0)
                h = h \ m_prime(j)
                k = k + 1
            Wend
            If (k <> 0) Then
                f(j) = f(j) + k
                If (h = 1) Then Exit For
            End If
        Next j
        If (q < j) Then q = j
        'status: i! = product[prime(0..q)^f(0..q)]
        If (m(i) <> 0) Then    'divide result
            k = m(i)
            For j = 0 To q
                e(j) = e(j) - f(j) * k
            Next j
        End If
    Next i
    For j = 0 To q             'multiply result by n!
        e(j) = e(j) + f(j)
    Next j

    'result product[prime(0..q)^e(0..q)] can overflow a Long
    On Error GoTo ErrOverflow
    k = 1
    For j = 0 To q
        For i = 1 To e(j)
            k = k * m_prime(j)
        Next i
    Next j
    MultinomialCoefficient = k
Exit Function
ErrOverflow:
MultinomialCoefficient = -1
End Function

'Returns: p(n)
'The number of integer partitions of n into any number of parts
Public Function PartitionCount(n As Long) As Long
    Dim i As Long, j As Long, k As Long
    Const table_max As Long = 116
    Static p(table_max) As Long

    Select Case n
    Case Is <= 1
        PartitionCount = 1
    Case Is > table_max
        PartitionCount = -1 'overflow
    Case Else
        If (p(0) = 0) Then 'build table
            p(0) = 1
            For i = 1 To table_max
                k = 1
                p(i) = 0
                Do
                    j = i - k * (3 * k - 1) \ 2
                    If (j < 0) Then Exit Do
                    p(i) = p(i) + p(j)
                    j = i - k * (3 * k + 1) \ 2
                    If (j < 0) Then Exit Do
                    p(i) = p(i) + p(j)
                    k = k + 1
                    j = i - k * (3 * k - 1) \ 2
                    If (j < 0) Then Exit Do
                    p(i) = p(i) - p(j)
                    j = i - k * (3 * k + 1) \ 2
                    If (j < 0) Then Exit Do
                    p(i) = p(i) - p(j)
                    k = k + 1
                Loop
            Next i
        End If
        PartitionCount = p(n)
    End Select
End Function

'Returns: |n k|  p(n,k)
'The number of integer partitions of n into k parts
Public Function PartitionCountParts(n As Long, k As Long) As Long
    Dim i As Long, j As Long, u() As Long
On Error GoTo ErrOverflow

    If (n = 0) Then
        If (k = 0) Then
            PartitionCountParts = 1
        Else
            PartitionCountParts = 0
        End If
    ElseIf (k <= 0) Or (k > n) Then
        PartitionCountParts = 0
    ElseIf (n > 1000) Then
        GoTo ErrOverflow
    Else
        ReDim u(n - k, k)
        u(0, 0) = 1
        For i = 1 To (n - k)
            u(i, 0) = 0
        Next i
        For j = 0 To k
            u(0, j) = 1
        Next j
        For i = 1 To n - k
            For j = 1 To k
                If ((i - j) < 0) Then
                    u(i, j) = u(i, j - 1)
                Else
                    u(i, j) = u(i, j - 1) + u(i - j, j)
                End If
            Next j
        Next i
        PartitionCountParts = u(i - 1, j - 1)
    End If
Exit Function
ErrOverflow:
    PartitionCountParts = -1
End Function

'Returns: [n k] "n cycle k" Stirling Cycle Number (the "first kind")
'The number of ways to arrange n objects into k cycles
Public Function StirlingCycleNumber(n As Long, k As Long) As Long
    Dim i As Long, j As Long, u() As Long
On Error GoTo ErrOverflow

    If ((k <= 0) Or (k > n)) Then
        If ((k = 0) And (n = 0)) Then
            StirlingCycleNumber = 1
        Else
            StirlingCycleNumber = 0
        End If
    ElseIf (k = 1) Then 'return (n-1)!
        j = 1
        For i = 2 To n - 1
            j = j * i
        Next i
        StirlingCycleNumber = j
    ElseIf (k = n) Then
        StirlingCycleNumber = 1
    ElseIf (k = n - 1) Then 'return (n 2)
        StirlingCycleNumber = BinomialCoefficient(n, 2)
    Else
        ReDim u(k) 'recurrence: [n k] = (n - 1) * [n-1 k] + [n-1 k-1]  integer n>0
        u(0) = 0 'if n>0
        For j = 1 To k
            u(j) = 1
        Next j
        For i = 1 To (n - k)
            For j = 1 To k 'construct u(0..k) = first k elements of diagonal row i
                u(j) = u(j) * (i + j - 1) + u(j - 1)
            Next j
        Next i
        StirlingCycleNumber = u(k)
    End If
Exit Function
ErrOverflow:
    StirlingCycleNumber = -1
End Function

'Returns: {n k} "n subset k" Stirling Subset Number (the "second kind")
'The number of ways to partition n elements into k nonempty subsets
Public Function StirlingSubsetNumber(n As Long, k As Long) As Long
    Dim i As Long, j As Long, u() As Long
On Error GoTo ErrOverflow

    If ((k <= 0) Or (k > n)) Then
        If ((k = 0) And (n = 0)) Then
            StirlingSubsetNumber = 1
        Else
            StirlingSubsetNumber = 0
        End If
    ElseIf ((k = 1) Or (k = n)) Then
        StirlingSubsetNumber = 1
    Else
        ReDim u(k) 'recurrence: {n k} = k*{n-1 k} + {n-1 k-1}  integer n>0
        u(0) = 0 'if n>0
        For j = 1 To k
            u(j) = 1
        Next j
        For i = 1 To (n - k)
            For j = 2 To k 'construct u(0..k) = first k elements of diagonal row i
                u(j) = u(j) * j + u(j - 1)
            Next j
        Next i
        StirlingSubsetNumber = u(k)
    End If
Exit Function
ErrOverflow:
    StirlingSubsetNumber = -1
End Function

'Returns: b(n) Bell Number
'The number of ways to partition n elements into any number of nonempty subsets
Public Function BellNumber(n As Long) As Long
    Dim i As Long, j As Long, u() As Long
    Const table_max As Long = 15
    Static b(table_max) As Long

    Select Case n
    Case Is <= 1
        BellNumber = 1
    Case Is > table_max
        BellNumber = -1 'overflow
    Case Else
        If b(0) = 0 Then         'build table
            ReDim u(table_max - 1)
            u(0) = 1
            b(0) = 1
            b(1) = 1
            For i = 2 To table_max
                u(i - 1) = u(0)  'construct row i of Peirce's triangle
                For j = i - 2 To 0 Step -1
                    u(j) = u(j) + u(j + 1)
                Next j
                b(i) = u(0)      'first element is Bell number
            Next i
        End If
        BellNumber = b(n)
    End Select
End Function

'Returns: <n k> Eulerian Number
'The number of permutations of {1,2,..n} that have k ascents
Public Function EulerianNumber(n As Long, ByVal k As Long) As Long
    Dim i As Long, j As Long, u() As Long
On Error GoTo ErrOverflow

    If ((k <= 0) Or (k >= n)) Then
        If ((k = 0) And (n >= 0)) Then
            EulerianNumber = 1
        Else
            EulerianNumber = 0
        End If
    ElseIf (k = n - 1) Then
        EulerianNumber = 1 'n >= 1
    Else
        If (k > (n - 1 - k)) Then k = (n - 1 - k)
        ReDim u(k) 'recurrence: <n k> = (k+1) * <n-1 k> + (n-k) * <n-1 k-1>
        For j = 0 To k
            u(j) = 1
        Next j
        For i = 2 To (n - k)
            For j = 1 To k 'construct u(0..k) = first k elements of diagonal row i
                u(j) = u(j) * (j + 1) + u(j - 1) * i
            Next j
        Next i
        EulerianNumber = u(k)
    End If
Exit Function
ErrOverflow:
    EulerianNumber = -1
End Function

'Returns: t(n) Involution Count
'The number of permutations of {1,2,..n} that are its own inverse (contain only one cycles and two cycles)
'  number of tableau that can be formed from elements {1..n}
'
't(n) = t(n-1) + (n-1) * t(n-2)
't(0..10) = 1, 1, 2, 4, 10, 26, 76, 232, 764, 2620, 9496
Private Function InvolutionCount(n As Long) As Long
    Dim i As Long
    Const table_max As Long = 18
    Static t(table_max) As Long

    Select Case n
    Case Is < 0:         InvolutionCount = 0
    Case Is > table_max: InvolutionCount = -1 'overflow
    Case Else
        If (t(0) = 0) Then 'build table
            t(0) = 1 'recurrence: t(n) = t(n-1) + (n-1) * t(n-2)
            t(1) = 1
            For i = 2 To table_max
                t(i) = t(i - 1) + (i - 1) * t(i - 2)
            Next i
        End If
       InvolutionCount = t(n)
   End Select
End Function


Private Sub cmdFactorPrime_Click()
    Dim j As Long, m As Long, n As Long, s As String
    Dim p(12) As Long, e(12) As Long

    If Not IsNumeric(txtCountN.Text) Then
        AddItemText "Prime Factorizations"
        AddItemText "error - can not read input" & vbCrLf
        Exit Sub
    End If
    If ((CDbl(txtCountN.Text) < 1) Or (CDbl(txtCountN.Text) > &H7FFFFFFF)) Then
        AddItemText "Prime Factorization"
        AddItemText "error - can not read input" & vbCrLf
        Exit Sub
    End If
    n = CLng(txtCountN.Text)

    m = FactorInteger(p, e, n)
    AddItemText "Prime factorization of " & CStr(n)
    s = CStr(n) & " = "

    For j = 0 To m - 1
        If (e(j) = 1) Then
           s = s & CStr(p(j))
        Else
            s = s & CStr(p(j)) & "^" & CStr(e(j))
        End If
        If (j = m - 1) Then
            s = s & vbCrLf
        Else
            s = s & " * "
        End If
    Next j
    AddItemText s
End Sub

Private Sub cmdFactorAll_Click()
    Dim a As Long, b As Long, j As Long, k As Long, l As Long, x As Long
    Dim c() As Long, f() As Long, u() As Long, v() As Long
    Dim h As Long, t As Long, n As Long, st As String
    Dim p(12) As Long, e(12) As Long, m As Long

    If Not IsNumeric(txtCountN.Text) Then
        AddItemText "All Factorizations"
        AddItemText "error - can not read input" & vbCrLf
        Exit Sub
    End If
    If ((CDbl(txtCountN.Text) < 2) Or (CDbl(txtCountN.Text) > &H7FFFFFFF)) Then
        AddItemText "All Factorizations"
        AddItemText "error - can not read input" & vbCrLf
        Exit Sub
    End If

    n = CLng(txtCountN.Text)
    AddItemText "All factorizations of " & CStr(n)

    'Strategy: Visit all multiset partitions of the factors of the prime factorization of n
    m = FactorInteger(p, e, n)
    k = 0
    For j = m To 1 Step -1
        e(j) = e(j - 1)
        k = k + e(j)
    Next j

    ReDim f(k)
    ReDim c(m * k)
    ReDim u(m * k)
    ReDim v(m * k)

    For j = 0 To m - 1
        c(j) = j + 1
        u(j) = e(j + 1)
        v(j) = e(j + 1)
    Next j
    f(0) = 0
    a = 0
    l = 0
    f(1) = m
    b = m

    Do
        j = a
        k = b
        x = 0
        While (j < b)
            u(k) = u(j) - v(j)
            If (u(k) = 0) Then
                x = 1
                j = j + 1
            ElseIf (x = 0) Then
                c(k) = c(j)
                If v(j) < u(k) Then
                    v(k) = v(j)
                Else
                    v(k) = u(k)
                End If
                If (u(k) < v(j)) Then
                    x = 1
                Else
                    x = 0
                End If
                k = k + 1
                j = j + 1
            Else
                c(k) = c(j)
                v(k) = u(k)
                k = k + 1
                j = j + 1
            End If
        Wend

        If (k > b) Then
            a = b
            b = k
            l = l + 1
            f(l + 1) = b
        Else
            'Visit Partition
            For k = 0 To l 'l+1 partitions
                t = 1
                For j = f(k) To f(k + 1) - 1
                    For h = 1 To v(j)
                        t = t * p(c(j) - 1)
                    Next h
                Next j
                st = st & CStr(t) & IIf(k = l, vbCrLf, " * ")
            Next k
            'End Visit

            j = b - 1
            While (v(j) = 0)
                j = j - 1
            Wend
            While ((j = a) And (v(j) = 1))
                If (l = 0) Then Exit Do
                l = l - 1
                b = a
                a = f(l)
                j = b - 1
                While (v(j) = 0)
                    j = j - 1
                Wend
            Wend
            v(j) = v(j) - 1
            For k = j + 1 To b - 1
                v(k) = u(k)
            Next k
        End If
    Loop

    AddItemText st
End Sub

Private Sub cmdCount_Click(Index As Integer)
    Dim k As Long, n As Long, r As Long

    If IsNumeric(txtCountN.Text) Then
        If ((CDbl(txtCountN.Text) > 0) And (CDbl(txtCountN.Text) < &H7FFFFFFF)) Then
            n = CLng(txtCountN.Text)
            If IsNumeric(txtCountK.Text) Then
                If ((CDbl(txtCountK.Text) > 0) And (CDbl(txtCountK.Text) < &H7FFFFFFF)) Then
                    k = CLng(txtCountK.Text)
                    Select Case Index
                    Case 0 'Binomial
                        r = BinomialCoefficient(n, k)
                        If (r < 0) Then
                            AddItemText "Binomial Coefficient"
                            AddItemText "(" & CStr(n) & " " & CStr(k) & ") = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " ways to select " & CStr(k) & " objects from a set of " & CStr(n) & " objects" & vbCrLf
                        Else
                            AddItemText "Binomial Coefficient"
                            AddItemText "(" & CStr(n) & " " & CStr(k) & ") = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " ways to select " & CStr(k) & " objects from a set of " & CStr(n) & " objects" & vbCrLf
                        End If
                    Case 1 'Multiset
                        r = BinomialCoefficient(n + k - 1, k)
                        If (r < 0) Then
                            AddItemText "Multiset Coefficient"
                            AddItemText "<" & CStr(n) & " " & CStr(k) & "> = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " ways to select " & CStr(k) & " objects from a set of " & CStr(n) & " objects, allowing repetition" & vbCrLf
                        Else
                            AddItemText "Multiset Coefficient"
                            AddItemText "<" & CStr(n) & " " & CStr(k) & "> = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " ways to select " & CStr(k) & " objects from a set of " & CStr(n) & " objects, allowing repetition" & vbCrLf
                        End If
                    Case 2 'Partition
                        r = PartitionCount(n)
                        If (r < 0) Then
                            AddItemText "Partition"
                            AddItemText "p(" & CStr(n) & ") = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " ways to partition " & CStr(n) & " into parts" & vbCrLf
                        Else
                            AddItemText "Partition"
                            AddItemText "p(" & CStr(n) & ") = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " ways to partition " & CStr(n) & " into parts" & vbCrLf
                        End If
                    Case 3 'Partition Parts
                        r = PartitionCountParts(n, k)
                        If (r < 0) Then
                            AddItemText "Partition Parts"
                            AddItemText "|" & CStr(n) & " " & CStr(k) & "| = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " ways to partition " & CStr(n) & " into " & CStr(k) & " parts" & vbCrLf
                        Else
                            AddItemText "Partition Parts"
                            AddItemText "|" & CStr(n) & " " & CStr(k) & "| = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " ways to partition " & CStr(n) & " into " & CStr(k) & " parts" & vbCrLf
                        End If
                    Case 4 'Stirling Cycle
                        r = StirlingCycleNumber(n, k)
                        If (r < 0) Then
                            AddItemText "Stirling Cycle Number"
                            AddItemText "[" & CStr(n) & " " & CStr(k) & "] = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " ways to arrange " & CStr(n) & " objects into " & CStr(k) & " cycles" & vbCrLf
                        Else
                            AddItemText "Stirling Cycle Number"
                            AddItemText "[" & CStr(n) & " " & CStr(k) & "] = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " ways to arrange " & CStr(n) & " objects into " & CStr(k) & " cycles" & vbCrLf
                        End If
                    Case 5 'Stirling Subset
                        r = StirlingSubsetNumber(n, k)
                        If (r < 0) Then
                            AddItemText "Stirling Subset Number"
                            AddItemText "{" & CStr(n) & " " & CStr(k) & "} = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " ways to partition " & CStr(n) & " elements into " & CStr(k) & " nonempty subsets" & vbCrLf
                        Else
                            AddItemText "Stirling Subset Number"
                            AddItemText "{" & CStr(n) & " " & CStr(k) & "} = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " ways to partition " & CStr(n) & " elements into " & CStr(k) & " nonempty subsets" & vbCrLf
                        End If
                    Case 6 'Bell Number
                        r = BellNumber(n)
                        If (r < 0) Then
                            AddItemText "Bell Number"
                            AddItemText "b(" & CStr(n) & ") = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " ways to partition " & CStr(n) & " elements into nonempty subsets" & vbCrLf
                        Else
                            AddItemText "Bell Number"
                            AddItemText "b(" & CStr(n) & ") = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " ways to partition " & CStr(n) & " elements into nonempty subsets" & vbCrLf
                        End If
                    Case 7 'Eulerian
                        r = EulerianNumber(n, k)
                        If (r < 0) Then
                            AddItemText "Eulerian Number"
                            AddItemText "<" & CStr(n) & " " & CStr(k) & "> = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " permutations of {1.." & CStr(n) & "} that have " & CStr(k) & " nonempty subsets" & vbCrLf
                        Else
                            AddItemText "Eulerian Number"
                            AddItemText "<" & CStr(n) & " " & CStr(k) & "> = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " permutations of {1.." & CStr(n) & "} that have " & CStr(k) & " ascents" & vbCrLf
                        End If
                    Case 8 'Involution
                        r = InvolutionCount(n)
                        If (r < 0) Then
                            AddItemText "Involution Count"
                            AddItemText "t(" & CStr(n) & ") = <overflow>"
                            AddItemText "There are more than " & CStr(&H7FFFFFFF) & " permutations of {1.." & CStr(n) & "} that are its own inverse" & vbCrLf
                        Else
                            AddItemText "Involution Count"
                            AddItemText "t(" & CStr(n) & ") = " & CStr(r)
                            AddItemText "There are " & CStr(r) & " permutations of {1.." & CStr(n) & "} that are its own inverse" & vbCrLf
                        End If
                    End Select
                    Exit Sub
                End If
            End If
        End If
    End If
    AddItemText "error - can not read input"
End Sub


Private Sub Form_Load()
    Dim i As Long

    With m_ReserveSpace
        .Left = 120
        .Right = 120
        .top = 120
        .Bottom = 120
    End With
    InitPrimeTable
    For i = 2 To 6
        cboPermutationVisitN.AddItem CStr(i)
    Next i
    For i = 1 To 14
        cboCombinationVisitN.AddItem CStr(i + 1)
        cboCombinationVisitK.AddItem CStr(i)
    Next i
    For i = 1 To 12
        cboMultisetK.AddItem CStr(i)
        cboPartitionVisitN.AddItem CStr(i)
    Next i
    For i = 2 To 4
        cboTupleVisitN.AddItem CStr(i)
    Next i
    For i = 2 To 10
        cboTupleVisitM.AddItem CStr(i)
    Next i
    For i = 2 To 12
        cboGrayCodeVisitN.AddItem CStr(i)
        cboPartitionVisitK.AddItem CStr(i)
    Next i
    cboPermutationVisitN.ListIndex = 2
    cboCombinationVisitN.ListIndex = 4
    cboCombinationVisitK.ListIndex = 2
    cboMultisetK.ListIndex = 2
    cboTupleVisitN.ListIndex = 0
    cboTupleVisitM.ListIndex = 2
    cboGrayCodeVisitN.ListIndex = 2
    cboPartitionVisitN.ListIndex = 3
    cboPartitionVisitK.ListIndex = 0
    Option1(0).ToolTipText = Frame1(0).Caption
    Option1(1).ToolTipText = Frame1(1).Caption
    Option1(2).ToolTipText = Frame1(2).Caption
    cmdClear.ToolTipText = "Clear TextBox to the Left"
    On Error Resume Next
    Text1.Font.Name = "Lucida Console"
    Option1(0).Value = 1
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Dim NewWidth     As Long
    Dim NewHeight    As Long
    Dim FormRect     As RECT

    If (Me.WindowState <> 1) Then       'Only if not minimized
        GetClientRect Me.hwnd, FormRect 'Find available space on Form

        NewWidth = FormRect.Right * Screen.TwipsPerPixelX - m_ReserveSpace.Left - m_ReserveSpace.Right - Frame1(0).Width - 120
        NewHeight = FormRect.Bottom * Screen.TwipsPerPixelY - m_ReserveSpace.top - m_ReserveSpace.Bottom
        If (NewWidth > 0) Then
            If (NewHeight > Option1(0).Height + 120) Then
                Text1.Move m_ReserveSpace.Left, m_ReserveSpace.top, NewWidth, NewHeight
                Option1(0).Move m_ReserveSpace.Left + NewWidth + 120, m_ReserveSpace.top + NewHeight - Option1(0).Height
                Option1(1).Move Option1(0).Left + Option1(0).Width + 60, Option1(0).top
                Option1(2).Move Option1(1).Left + Option1(1).Width + 60, Option1(0).top
                Frame1(0).Move m_ReserveSpace.Left + NewWidth + 120, m_ReserveSpace.top, Frame1(0).Width, NewHeight - Option1(0).Height - 120
                Frame1(1).Move Frame1(0).Left, Frame1(0).top, Frame1(0).Width, Frame1(0).Height
                Frame1(2).Move Frame1(0).Left, Frame1(0).top, Frame1(0).Width, Frame1(0).Height
                cmdClear.Move Frame1(0).Left + Frame1(0).Width - cmdClear.Width, Option1(0).top
            End If
        End If
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Frame1(0).Visible = False
    Frame1(1).Visible = False
    Frame1(2).Visible = False
    Frame1(Index).Visible = True
End Sub

Private Sub cmdClear_Click()
    Text1.Text = ""
End Sub



'                        ************  PERMUTATIONS  ************

'Algorithm 7.2.1.2 P - Plain changes
'Visit all permutations of n distinct elements a(0..n-1) such that each iteration swaps two adjacent elements
'Required: 0 < n
Private Sub PermutationVisitPlainChanges(a() As Long, n As Long)
    Dim i As Long, j As Long, k As Long, q As Long
    Dim c() As Long, o() As Long
    Dim stb() As Byte, sti As Long

    'To demonstrate this algorithm, each visited permutation is added to a list and displayed.  A pre-sized
    'byte array is used, in place of a string, to avoid multiple slow string append statements.
    'Displaying a list
    'A useful application would test/examine/other a permutation at each visit
    'Building a list and displaying it shows what this does

    'simply listing visit data merely shows that it works
    AddItemText "Visit all permutations of " & CStr(n) & " distinct elements such that each iteration swaps two adjacent elements"
    If ((n < 1) Or (n > 12)) Then
        k = -1
    Else
        k = 1 'k = n!
        For i = 2 To n
            k = k * i
        Next i
    End If
    AddItemText "Number of permutations: " & IIf(k = -1, "<overflow>" & vbCrLf, CStr(k))
    If ((k = -1) Or (k > 6000)) Then Exit Sub
    ReDim stb(2 * k * (2 * n + 1) - 1)
    sti = 0

    ReDim c(1 To n)
    ReDim o(1 To n)
    For j = 1 To n
        c(j) = 0 'inversion table
        o(j) = 1 'direction by which c(i) changes
    Next j

    Do
        'Visit Permutation a(0..n-1)
        For i = 0 To n - 1
            If (i <> 0) Then 'insert space character
                stb(sti) = 32
                sti = sti + 2
            End If
            stb(sti) = a(i)
            sti = sti + 2
        Next i
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        i = n - 1
        For j = n To 1 Step -1
            q = c(j) + o(j)
            Select Case q
            Case -1:   i = i - 1
            Case j:    If (j = 1) Then Exit Do
            Case Else: Exit For
            End Select
            o(j) = -o(j)
        Next j
        k = a(i - c(j))
        a(i - c(j)) = a(i - q)
        c(j) = q
        a(i - q) = k
    Loop

    If (n > 1) Then 'final swap (not necessary) restores original a(0..n-1)
        k = a(0)
        a(0) = a(1)
        a(1) = k
    End If

    AddItemText CStr(stb)
End Sub

'Algorithm 7.2.1.2 E - Ehrlich swaps
'Visit all permutations of n distinct elements a(0..n-1) such that each iteration swaps a(0) with another element
'Required: 1 < n
Private Sub PermutationVisitEhrlichSwap(a() As Long, n As Long)
    Dim i As Long, j As Long, k As Long
    Dim b() As Long, c() As Long
    Dim stb() As Byte, sti As Long

    'Visiting a permutation here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    AddItemText "Visit all permutations of " & CStr(n) & " distinct elements such that each iteration swaps the first element with another element"
    If ((n <= 1) Or (n > 12)) Then
        k = -1
    Else
        k = 1 'k = n!
        For i = 2 To n
            k = k * i
        Next i
    End If
    AddItemText "Number of permutations: " & IIf(k = -1, "<overflow>" & vbCrLf, CStr(k))
    If ((k = -1) Or (k > 6000)) Then Exit Sub
    ReDim stb(2 * k * (2 * n + 1) - 1)
    sti = 0

    ReDim b(1 To n - 1)
    ReDim c(1 To n - 1)
    For j = 1 To n - 1
        b(j) = j
        c(j) = 0  'control table
    Next j

    Do
        'Visit Permutation a(0..n-1)
        For i = 0 To n - 1
            If (i <> 0) Then 'insert space character
                stb(sti) = 32
                sti = sti + 2
            End If
            stb(sti) = a(i)
            sti = sti + 2
        Next i
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        For i = 1 To n - 1
            If (c(i) < i) Then Exit For
            c(i) = 0
        Next i
        If (i = n) Then Exit Do

        k = a(0)    'swap a(0), a(b(i))
        a(0) = a(b(i))
        c(i) = c(i) + 1
        a(b(i)) = k

        i = i - 1   'reverse the order of b(1..i-1)
        j = 1
        While (i > j)
            k = b(i)
            b(i) = b(j)
            i = i - 1
            b(j) = k
            j = j + 1
        Wend
    Loop

    AddItemText CStr(stb)
End Sub

'Algorithm 7.2.1.2 T - Plain change transitions
'Compute table t(0..n!-1) such that the actions of PermutationVisitPlainChanges() are equivalent to the successive interchanges
'swap(a(t(k)), a(t(k)-1)) for 0 <= k <= n!-2.  [ for 1 based arrays: swap(a(t(k)), a(t(k)+1)) ]
'Required: 1 < n
'A final swap with element t(n!-1) will return elements back to their original order
'Returns n! (number of ways n distinct elements can be permutated)
Private Function PermutationVisitPlainChangesIndex(t() As Long, n As Long) As Long
    Dim d As Long, j As Long, k As Long, m As Long, nf As Long

    Select Case n
    Case Is <= 0
        nf = 0
    Case 1
        nf = 1
    Case Else
        nf = 1 'nf = n!
        For j = 2 To n
            nf = nf * j
        Next j
        d = nf \ 2
        ReDim t(nf - 1)
        t(d - 1) = 1
        t(nf - 1) = 1
        For m = 3 To n
            d = d \ m
            k = d - 1
            While (k < nf)
                For j = m - 1 To 1 Step -1
                    t(k) = j
                    k = k + d
                Next j
                t(k) = t(k) + 1
                k = k + d
                For j = 1 To m - 1
                    t(k) = j
                    k = k + d
                Next j
                k = k + d
            Wend
        Next m
    End Select

    PermutationVisitPlainChangesIndex = nf
End Function

'Visit all permutations of n elements a(n-1..0) in lexicographic order
'Elements need not be distinct
Private Sub PermutationVisitMultiset(a() As Long, n As Long)
    Dim h As Long, i As Long, j As Long, k As Long
    Dim stb() As Byte, sti As Long

    'Visiting a permutation here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    h = MultinomialCoefficient(a, n)
    AddItemText "Visit all permutations of a " & CStr(n) & " element multiset in lexicographic order"
    Select Case h
    Case -1
        AddItemText "Number of permutations: <overflow>"
        AddItemText "(abort list)" & vbCrLf
        Exit Sub
    Case Is > 6000
        AddItemText "Number of permutations: " & CStr(h)
        AddItemText "(abort list)" & vbCrLf
        Exit Sub
    End Select
    AddItemText "Number of permutations: " & CStr(h)
    ReDim stb(2 * h * (n + 2) - 1)
    sti = 0

    While (h > 0)
        h = h - 1

        'Visit Permutation a(n-1..0)
        For i = n - 1 To 0 Step -1
            stb(sti) = a(i)
            sti = sti + 2
        Next i
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        j = 0
        For i = 0 To n - 2
            If (a(i + 1) < a(i)) Then    'if there exists an index i such that a(i + 1) < a(i),
                k = a(i + 1)             '  swap a(i+1) with the minimum a(j) (0 <= j <= i) whose value
                While (a(i + 1) >= a(j)) '  is greater than a(i+1).  otherwise set i = n-1
                    j = j + 1
                Wend
                a(i + 1) = a(j)
                a(j) = k
                j = 0
                Exit For
            End If
        Next i
        While (i > j) 'reverse the order of a(0..i)
            k = a(i)
            a(i) = a(j)
            i = i - 1
            a(j) = k
            j = j + 1
        Wend
    Wend

    AddItemText CStr(stb)
End Sub

'The following two functions convert between index i and the i-th lexicographic permutation of a(0..n-1)
'Returns: index of permutation p(0..n-1) in lexicographic order
Public Function PermutationLexicographicIndex(a() As Long, n As Long) As Long
    Dim i As Long, j As Long, k As Long, c() As Long

    ReDim c(n - 1) 'c(0..n-1) = inversion vector
    For i = 0 To n - 1
        k = 0
        For j = i + 1 To n - 1
            If (a(i) > a(j)) Then k = k + 1
        Next j
        c(i) = k
    Next i

    j = 1
    k = 0
    For i = 1 To n - 1
        j = j * i
        k = k + c(n - 1 - i) * j
    Next i
    PermutationLexicographicIndex = k
End Function

Public Sub PermutationLexicographicIndexInv(a() As Long, n As Long, ByVal idx As Long)
    Dim h As Long, d As Long, z As Long, i As Long, j As Long, k As Long, c() As Long
    Dim p() As Long

    ReDim p(n - 1)
    p(0) = 0
    p(1) = 1
    h = 1 'h = (n-1)!
    For i = 2 To n - 1
        h = h * i
        p(i) = i
    Next i

    For i = 0 To n - 2
        j = 0
        While (idx >= h)
            idx = idx - h
            j = j + 1
        Wend
        h = h \ (n - 1 - i)

        'element i is jth largest
        z = p(i + j)
        For d = i + j To i + 1 Step -1
            p(d) = p(d - 1)
        Next d
        p(i) = z
    Next i
    
    'sort a() to lexicographic minimum
    For i = 0 To n - 2
        h = i
        For j = i + 1 To n - 1
            If a(h) > a(j) Then h = j
        Next j
        k = a(i)
        a(i) = a(h)
        a(h) = k
    Next i
    
    'Apply permutation p(0..n-1) to elements a(0..n-1)
    For i = 0 To n - 1
        If (p(i) >= 0) Then     'element p(i) has not been visited yet
            If (p(i) <> i) Then 'element p(i) is in cycle > 1
                h = a(i)
                j = i
                k = p(i)
                While (k <> i)
                    a(j) = a(k)
                    j = k
                    k = p(k)
                    p(j) = Not p(j)
                Wend
                a(j) = h
            End If
            p(i) = Not p(i) 'permutation elements p(0..n-1) are inverted after applying permutation
        End If
    Next i
End Sub


Private Sub cmdPermutationVisit_Click(Index As Integer)
    Dim i As Long, n As Long
    Dim a() As Long

    n = CLng(cboPermutationVisitN.Text)
    ReDim a(n - 1)
    For i = 0 To n - 1
        a(i) = 49 + i 'Asc("1") ...
    Next i
    Select Case Index
    Case 0: PermutationVisitPlainChanges a, n
    Case 1: PermutationVisitEhrlichSwap a, n
    End Select
End Sub

Private Sub cmdPermutationVisitIdx_Click()
    Dim i As Long, j As Long, k As Long, n As Long, nf As Long
    Dim a() As Long, t() As Long
    Dim stb() As Byte, sti As Long

    n = CLng(cboPermutationVisitN.Text)
    ReDim a(n - 1)
    For i = 0 To n - 1
        a(i) = 49 + i 'Asc("1") ...
    Next i

    'Visiting a permutation here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    AddItemText "Visit all permutations of " & CStr(n) & " distinct elements such that each iteration swaps two adjacent elements (via pre-calculated swap indexes)"
    If ((n <= 1) Or (n > 12)) Then
        k = -1
    Else
        k = 1 'k = n!
        For i = 2 To n
            k = k * i
        Next i
    End If
    AddItemText "Number of permutations: " & IIf(k = -1, "<overflow>" & vbCrLf, CStr(k))
    If ((k < 1) Or (k > 6000)) Then Exit Sub
    ReDim stb(2 * k * (2 * n + 1) - 1)
    sti = 0

    'get swap index list
    nf = PermutationVisitPlainChangesIndex(t, n)

    For j = 0 To nf - 1
        'Visit Permutation a(0..n-1)
        For i = 0 To n - 1
            If (i <> 0) Then 'insert space character
                stb(sti) = 32
                sti = sti + 2
            End If
            stb(sti) = a(i)
            sti = sti + 2
        Next i
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        k = a(t(j)) 'make the swap
        a(t(j)) = a(t(j) - 1)
        a(t(j) - 1) = k
    Next j

    AddItemText CStr(stb)
End Sub

Private Sub cmdPermutationVisitMultiset_Click()
    Dim a() As Long, j As Long, n As Long

    'copy ASCII values of txtMultiset.Text to a(0..n-1) and call PermutationVisitMultiset()
    n = Len(txtMultiset.Text)
    If (n = 0) Then
        AddItemText "No Elements" & vbCrLf
    Else
        ReDim a(n - 1)
        For j = 0 To n - 1
            a(j) = Asc(Mid$(txtMultiset.Text, n - j, 1))
        Next j
        PermutationVisitMultiset a, n
    End If
End Sub

Private Sub cmdPermutationCountMultiset_Click()
    Dim a() As Long, j As Long, n As Long

    n = Len(txtMultiset.Text)
    If (n = 0) Then
        AddItemText "No Elements" & vbCrLf
    Else
        ReDim a(n - 1)
        For j = 0 To n - 1
            a(j) = Asc(Mid$(txtMultiset.Text, n - j, 1))
        Next j
        AddItemText "Multiset: " & txtMultiset.Text
        AddItemText "Number of permutations: " & CStr(MultinomialCoefficient(a, n)) & vbCrLf
    End If
End Sub



'                        ************ SET COMBINATIONS  ************

'Algorithm 7.2.1.3 T - Lexicographic combinations
'Visit all "n choose k" combinations of distinct elements
'Required: 0 < k < n
Private Sub CombinationVisit(a() As Long, n As Long, k As Long)
    Dim i As Long, j As Long, c() As Long
    Dim stb() As Byte, sti As Long

    'Visiting a combination here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    i = BinomialCoefficient(n, k)
    ReDim stb(i * (4 * k + 2) - 1)
    sti = 0

    AddItemText "Visit all """ & CStr(n) & " choose " & CStr(k) & """ combinations of distinct elements in lexicographic order"
    AddItemText "Number of combinations: " & CStr(i)

    ReDim c(k + 1)
    For j = 0 To k - 1
        c(j) = j
    Next j
    c(k) = n
    c(k + 1) = 0

    j = k - 1
    Do
        'Visit Combination c(0..k-1)
        For i = 0 To k - 1
            If (i <> 0) Then 'insert space character
                stb(sti) = 32
                sti = sti + 2
            End If
            stb(sti) = a(c(i))  'c(i) = {0..n-1}
            sti = sti + 2      'c(0) < c(1) < .. < c(k-1)
        Next i
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        If (j < 0) Then
            If (c(0) + 1 < c(1)) Then
                c(0) = c(0) + 1
            Else
                For j = 0 To k - 2
                    c(j) = j
                    i = c(j + 1) + 1
                    If (i <> c(j + 2)) Then Exit For
                Next j
                If (j = k - 1) Then Exit Do
                c(j + 1) = i
            End If
        Else
            c(j) = j + 1
            j = j - 1
        End If
    Loop

    AddItemText CStr(stb)
End Sub

'Algorithm 7.2.1.3 R - Revolving-door combinations
'Visit all "n choose k" combinations of distinct elements
'Genlex order - all permutations with a common prefix occur consecutively
'Each iteration swaps a selected element with an unselected element
'Required: 1 < k < n
Private Sub CombinationVisitRevolvingDoor(a() As Long, n As Long, k As Long)
    Dim i As Long, j As Long, c() As Long
    Dim stb() As Byte, sti As Long

    'Visiting a combination here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    i = BinomialCoefficient(n, k)
    ReDim stb(2 * i * (2 * k + 1) - 1)
    sti = 0

    AddItemText "Visit all """ & CStr(n) & " choose " & CStr(k) & """ combinations of distinct elements such that each iteration swaps a selected/unselected pair"
    AddItemText "Number of combinations: " & CStr(i)

    ReDim c(k)
    For j = 0 To k - 1
        c(j) = j
    Next j
    c(j) = n

    'Algorithm has two variants, depending on whether k is even or odd
    If ((k And 1) = 0) Then 'k is even
        Do
            'Visit Combination c(0..k-1)  *procedure instance 1 of 2*
            For i = 0 To k - 1
                If (i <> 0) Then 'insert space character
                    stb(sti) = 32
                    sti = sti + 2
                End If
                stb(sti) = a(c(i)) 'c(i) = {0..n-1}
                sti = sti + 2      'c(0) < c(1) < .. < c(k-1)
            Next i
            stb(sti) = 13
            stb(sti + 2) = 10
            sti = sti + 4
            'End Visit

            If (c(0) > 0) Then
                c(0) = c(0) - 1
            Else
                For j = 1 To k - 1
                    If (c(j) + 1 < c(j + 1)) Then
                        c(j - 1) = c(j)
                        c(j) = c(j) + 1
                        Exit For
                    End If
                    j = j + 1
                    If (j = k) Then Exit Do
                    If (c(j) > j) Then
                        c(j) = c(j - 1)
                        c(j - 1) = j - 1
                        Exit For
                    End If
                Next j
            End If
        Loop
    Else 'k is odd
        Do
            'Visit Combination c(0..k-1)  *procedure instance 2 of 2*
            For i = 0 To k - 1
                If (i <> 0) Then 'insert space character
                    stb(sti) = 32
                    sti = sti + 2
                End If
                stb(sti) = a(c(i)) 'c(i) = {0..n-1}
                sti = sti + 2      'c(0) < c(1) < .. < c(k-1)
            Next i
            stb(sti) = 13
            stb(sti + 2) = 10
            sti = sti + 4
            'End Visit

            If (c(0) + 1 < c(1)) Then
                c(0) = c(0) + 1
            Else
                For j = 1 To k - 1
                    If (c(j) > j) Then
                        c(j) = c(j - 1)
                        c(j - 1) = j - 1
                        Exit For
                    End If
                    j = j + 1
                    If (c(j) + 1 < c(j + 1)) Then
                        c(j - 1) = c(j)
                        c(j) = c(j) + 1
                        Exit For
                    End If
                Next j
                If (j = k) Then Exit Do
            End If
        Loop
    End If

    AddItemText CStr(stb)
End Sub

'Algorithm 7.2.1.3 C - Chase's sequence
'Visit all "n choose k" combinations of distinct elements
'Genlex order - all permutations with a common prefix occur consecutively
'Each iteration swaps a selected element with an unselected element
'Near-perfect order of Chase's sequence - each iteration changes selected index by at most 2
'Required:  0 < k <= n
Private Sub CombinationVisitChaseSequence(a() As Long, n As Long, k As Long)
    Dim i As Long, j As Long, r As Long, s As Long
    Dim c() As Long, w() As Long
    Dim stb() As Byte, sti As Long

    s = n - k
    i = BinomialCoefficient(n, k)
    AddItemText "Visit all """ & CStr(n) & " choose " & CStr(k) & """ combinations of distinct elements in the near-perfect order of Chase's sequence"
    AddItemText "Number of combinations: " & CStr(i)

    'Visiting a combination here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    ReDim stb(2 * i * (2 * k + 1) - 1)
    sti = 0

    ReDim c(n - 1)
    ReDim w(n)
    j = 0
    While (j < s)
        c(j) = 0
        w(j) = 1
        j = j + 1
    Wend
    While (j < n)
        c(j) = 1
        w(j) = 1
        j = j + 1
    Wend
    w(j) = 1
    If (s > 0) Then
        r = s
    Else
        r = k
    End If

    Do
        'Visit Combination c(0..n-1)
        j = 0
        For i = 0 To n - 1
            If (c(i) <> 0) Then 'c(i)=1
                If (j <> 0) Then 'insert space character
                    stb(sti) = 32
                    sti = sti + 2
                End If
                j = j + 1
                stb(sti) = a(i)
                sti = sti + 2
            End If
        Next i
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        j = r
        While (w(j) = 0)
            w(j) = 1
            j = j + 1
        Wend
        If (j = n) Then Exit Do
        w(j) = 0
        If (c(j) = 0) Then
            If ((j And 1) = 1) Then
                If (c(j - 1) = 0) Then
                    c(j) = 1
                    c(j - 2) = 0
                    If (r = j - 2) Then
                        r = j
                    ElseIf (r = j - 1) Then
                        r = j - 2
                    End If
                Else
                    c(j) = 1 'Begin snippet C6
                    c(j - 1) = 0
                    If ((r = j) And (j > 1)) Then
                        r = j - 1
                    ElseIf (r = j - 1) Then
                        r = j
                    End If   'End snippet C6
                End If
            Else
                    c(j) = 1 'Begin snippet C6
                    c(j - 1) = 0
                    If ((r = j) And (j > 1)) Then
                        r = j - 1
                    ElseIf (r = j - 1) Then
                        r = j
                    End If   'End snippet C6
            End If
        Else
            If ((j And 1) = 0) Then
                If (c(j - 2) = 0) Then
                    c(j - 2) = 1
                    c(j) = 0
                    If (r = j) Then
                        r = IIf(j - 2 > 1, j - 2, 1)
                    ElseIf (r = j - 2) Then
                        r = j - 1
                    End If
                Else
                    c(j - 1) = 1 'Begin snippet C4
                    c(j) = 0
                    If ((r = j) And (j > 1)) Then
                        r = j - 1
                    ElseIf (r = j - 1) Then
                        r = j
                    End If       'End snippet C4
                End If
            Else
                    c(j - 1) = 1 'Begin snippet C4
                    c(j) = 0
                    If ((r = j) And (j > 1)) Then
                        r = j - 1
                    ElseIf (r = j - 1) Then
                        r = j
                    End If       'End snippet C4
            End If
        End If
    Loop

    AddItemText CStr(stb)
End Sub

'              ************  MULTISET COMBINATIONS, BOUNDED COMPOSITIONS  ************

'Counting bounded compositions is done by visiting integer partitions
'e(i) is number of element i, for 0 <= i < n
Private Function BoundedCompositionsCount(e() As Long, n As Long, ByVal k As Long) As Long
    Dim i As Long, j As Long, q As Long, x As Long
    Dim c() As Long
    Dim f As Long, g As Long, h As Long, t As Long, total As Long, v() As Long

On Error GoTo ErrOverflow

    'generate v(), a sorted index of e
    ReDim v(n - 1)
    For f = 0 To n - 1
        v(f) = f
    Next f
    For f = 1 To n - 1
        h = v(f)
        i = e(v(f))
        For j = f - 1 To 0 Step -1
            If (e(v(j)) >= i) Then Exit For
            v(j + 1) = v(j)
        Next j
        v(j + 1) = h
    Next f

    ReDim c(k) 'initialize
    For i = k To 2 Step -1
        c(i) = 1
    Next i
    c(0) = 0
    total = 0
    Do
        c(i) = k 'store final part
        If (k = 1) Then
            q = i - 1
        Else
            q = i
        End If

        Do
            'Visit Partition c(1..i)
            f = 0         'c(1) >= c(2) >= ... >= c(i) >= 1
            g = 0         'c(1) + c(2) + ... + c(i) = k
            t = 1         '1 <= i <= k
            For j = 1 To i
                h = 1 'h = number of equal c's
                Do While (j < i)
                    If (c(j) <> c(j + 1)) Then Exit Do
                    j = j + 1
                    h = h + 1
                Loop
                Do While (f < n)
                    If (e(v(f)) < c(j)) Then Exit Do
                    f = f + 1
                    g = g + 1
                Loop
                t = t * BinomialCoefficient(g, h)
                g = g - h
            Next j
            total = total + t
            'End Visit

            If (c(q) <> 2) Then Exit Do
            c(q) = 1 'change 2 to 1+1
            q = q - 1
            i = i + 1
        Loop
        If (q = 0) Then Exit Do 'decrease c(q)
        x = c(q) - 1
        c(q) = x
        k = i - q + 1
        i = q + 1
        While (k > x) 'copy x if necessary
            c(i) = x
            i = i + 1
            k = k - x
        Wend
    Loop

    BoundedCompositionsCount = total
Exit Function
ErrOverflow:
If Err.Number = 6 Then
BoundedCompositionsCount = -1
Else
Err.Raise Err.Number
End If
End Function

'Visit all bounded compositions c(0..n-1)
'k = c(0) + .. + c(n-1)
'0 <=c(j) <= e(j)    for 0 <= j < n
Private Sub BoundedCompositionsVisit(e() As Long, n As Long, ByVal k As Long)
    Dim i As Long, j As Long
    Dim c() As Long
    Dim st As String

'check valid input
Debug.Assert (n > 1)
i = 0
For j = 0 To n - 1
    Debug.Assert (e(j) > 0)
    i = i + e(j)
Next j
Debug.Assert (k > 0) And (k <= i)

    j = BoundedCompositionsCount(e, n, k)
    AddItemText "Visit bounded compositions of " & CStr(i) & " elements"
    AddItemText "Number of bounded compositions: " & CStr(j)

    'initial distribute
    ReDim c(n - 1)
    For j = n - 1 To 1 Step -1
        c(j) = 0
    Next j 'j=0
    While (k > e(j))
        c(j) = e(j)
        k = k - e(j)
        j = j + 1
    Wend
    c(j) = k

    Do
        'Visit bounded composition c(0..n-1)
        For i = 0 To n - 2
            st = st & CStr(c(i)) & " "
        Next i
        st = st & CStr(c(i)) & vbCrLf
        'End Visit

        If ((j = 0) Or (c(0) = 0)) Then
            k = c(j) - 1                'Pick up the rightmost units
            If (j <> 0) Then c(j) = 0
            j = j + 1

            If (j = n) Then Exit Do 'Full?
            While (c(j) = e(j))
                k = k + e(j)
                c(j) = 0
                j = j + 1
                If (j = n) Then Exit Do
            Wend

            c(j) = c(j) + 1 'Increase c(j)
            If (k = 0) Then
                c(0) = 0 'status: c(0..j-1) = 0
            Else
                j = 0 'Distribute
                While (k > e(j))
                    c(j) = e(j)
                    k = k - e(j)
                    j = j + 1
                Wend
                c(j) = k
            End If
        Else 'Increase and decrease
            While (c(j) = e(j)) 'status: c(i) = e(i)   for 0 <= i < j
                j = j + 1
                If (j = n) Then Exit Do
            Wend
            c(j) = c(j) + 1
            j = j - 1
            c(j) = c(j) - 1
            If (c(0) = 0) Then j = 1
        End If
    Loop

    AddItemText st
End Sub

'Visit all "choose k" combinations of a multiset, input as:
'  p(0..n-1) are distinct elements in the multiset
'  there are e(i) instances of element p(i) in the multiset, for i <= 0 < n
'Required:
'  n > 1
'  0 < k <= Sum(e(0..n-1))
'  p(0..n-1) are distinct
'  e(0..n-1) > 0
Private Sub CombinationVisitMultiset(p() As Long, e() As Long, n As Long, ByVal k As Long)
    Dim h As Long, i As Long, j As Long, c() As Long
    Dim stb() As Byte, sti As Long

    i = 0
    For j = 0 To n - 1
        i = i + e(j)
    Next j
    h = BoundedCompositionsCount(e, n, k)
    AddItemText "Visit all ""choose " & CStr(k) & """ combinations of a " & CStr(i) & " element multiset"
    Select Case h
    Case -1
        AddItemText "Number of combinations: <overflow>"
        AddItemText "(abort list)" & vbCrLf
        Exit Sub
    Case Is > 6000
        AddItemText "Number of combinations: " & CStr(h)
        AddItemText "(abort list)" & vbCrLf
        Exit Sub
    End Select
    AddItemText "Number of combinations: " & CStr(h)
    ReDim stb(2 * h * (k + 2) - 1)
    sti = 0

    'initial distribute
    ReDim c(n - 1)
    For j = n - 1 To 1 Step -1
        c(j) = 0
    Next j 'j=0
    While (k > e(j))
        c(j) = e(j)
        k = k - e(j)
        j = j + 1
    Wend
    c(j) = k

    Do
        'Visit bounded composition c(0..n-1)
        For i = 0 To n - 1
            For h = 1 To c(i)
                stb(sti) = p(i)
                sti = sti + 2
            Next h
        Next i
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        If ((j = 0) Or (c(0) = 0)) Then
            k = c(j) - 1                'Pick up the rightmost units
            If (j <> 0) Then c(j) = 0
            j = j + 1

            If (j = n) Then Exit Do 'Full?
            While (c(j) = e(j))
                k = k + e(j)
                c(j) = 0
                j = j + 1
                If (j = n) Then Exit Do
            Wend

            c(j) = c(j) + 1 'Increase c(j)
            If (k = 0) Then
                c(0) = 0 'status: c(0..j-1) = 0
            Else
                j = 0 'Distribute
                While (k > e(j))
                    c(j) = e(j)
                    k = k - e(j)
                    j = j + 1
                Wend
                c(j) = k
            End If
        Else 'Increase and decrease
            While (c(j) = e(j)) 'status: c(i) = e(i)   for 0 <= i < j
                j = j + 1
                If (j = n) Then Exit Do
            Wend
            c(j) = c(j) + 1
            j = j - 1
            c(j) = c(j) - 1
            If (c(0) = 0) Then j = 1
        End If
    Loop

    AddItemText CStr(stb)
End Sub


Private Sub cmdCombinationVisit_Click(Index As Integer)
    Dim i As Long, k As Long, m As Long, n As Long
    Dim a() As Long

    n = CLng(cboCombinationVisitN.Text)
    k = CLng(cboCombinationVisitK.Text)
    m = BinomialCoefficient(n, k)

    ReDim a(n - 1)
    If (n < 10) Then
        For i = 0 To n - 1
            a(i) = i + 49 'asc("1"), asc("2"), ...
        Next i
    Else
        For i = 0 To n - 1
            a(i) = i + 97 'asc("a"), asc("b"), ...
        Next i
    End If

    Select Case Index
    Case 0 'CombinationVisit
        If ((k <= 0) Or (k >= n)) Then
            AddItemText "Required: 0 < k < n" & vbCrLf
        ElseIf (m > 6000) Then
            AddItemText "Number of combinations: " & CStr(m) & vbCrLf
        Else
            CombinationVisit a, n, k
        End If
    Case 1 'CombinationVisitRevolvingDoor
        If ((k <= 1) Or (k >= n)) Then
            AddItemText "Required: 1 < k < n" & vbCrLf
        ElseIf (m > 6000) Then
            AddItemText "Number of combinations: " & CStr(m) & vbCrLf
        Else
            CombinationVisitRevolvingDoor a, n, k
        End If
    Case 2 'CombinationVisitChaseSequence
        If ((k <= 0) Or (k > n)) Then
            AddItemText "Required: 0 < k <= n" & vbCrLf
        ElseIf (m > 6000) Then
            AddItemText "Number of combinations: " & CStr(m) & vbCrLf
        Else
            CombinationVisitChaseSequence a, n, k
        End If
    End Select
End Sub

'Visit Multiset Combinations
'Example:  A bin contains 4 blue, 3 green, 3 red, and 2 white marbles.
'  List all possible ways to select 5 marbles from the bin.  (there are 37 ways)
'  Call this function with multiset (txtPermutate.Text) = "bbbbgggrrrww" and k (cboCombinationMultisetK.ListIndex) = 5
Private Sub cmdCombinationVisitMultiset_Click()
    Dim h As Long, i As Long, j As Long, k As Long, n As Long
    Dim p() As Long, e() As Long

    'copy ASCII values of txtPermutate.Text to p(0..n-1)
    n = Len(txtMultiset.Text)
    If (n = 0) Then
        AddItemText "No Elements" & vbCrLf
        Exit Sub
    End If
    ReDim p(n - 1)
    ReDim e(n - 1)
    For j = 0 To n - 1
        p(j) = Asc(Mid$(txtMultiset.Text, j + 1, 1))
    Next j

    'sort p(0..n-1)
    For i = 1 To n - 1
        k = p(i)
        For j = i - 1 To 0 Step -1
            If (p(j) <= k) Then Exit For
            p(j + 1) = p(j)
        Next j
        p(j + 1) = k
    Next i

    h = 1
    j = 0
    For i = 1 To n - 1
        If (p(i) = p(j)) Then
            h = h + 1
        Else
            e(j) = h     'e(j) is number of instances of p(j)
            j = j + 1
            h = 1
            p(j) = p(i)  'p(j) is distinct in p(0..j)
        End If
    Next i
    e(j) = h
    n = j + 1 'number of distinct elements
    k = CLng(cboMultisetK.List(cboMultisetK.ListIndex))

    'Call procedure CombinationVisitMultiset() if input is valid
    i = 0
    For j = 0 To n - 1
        If (e(j) <= 0) Then Exit For
        i = i + e(j)
    Next j
    If (j <> n) Or (k <= 0) Or (k > i) Or (n <= 1) Then
        AddItemText "Invalid Input" & vbCrLf
    Else
        CombinationVisitMultiset p, e, n, k
    End If
End Sub

Private Sub cmdCombinationCountMultiset_Click()
    Dim h As Long, i As Long, j As Long, k As Long, n As Long
    Dim p() As Long, e() As Long

    'copy ASCII values of txtPermutate.Text to p(0..n-1)
    n = Len(txtMultiset.Text)
    If (n = 0) Then
        AddItemText "No Elements" & vbCrLf
        Exit Sub
    End If
    ReDim p(n - 1)
    ReDim e(n - 1)
    For j = 0 To n - 1
        p(j) = Asc(Mid$(txtMultiset.Text, j + 1, 1))
    Next j

    'sort p(0..n-1)
    For i = 1 To n - 1
        k = p(i)
        For j = i - 1 To 0 Step -1
            If (p(j) <= k) Then Exit For
            p(j + 1) = p(j)
        Next j
        p(j + 1) = k
    Next i

    h = 1
    j = 0
    For i = 1 To n - 1
        If (p(i) = p(j)) Then
            h = h + 1
        Else
            e(j) = h     'e(j) is number of instances of p(j)
            j = j + 1
            h = 1
            p(j) = p(i)  'p(j) is distinct in p(0..j)
        End If
    Next i
    e(j) = h
    n = j + 1 'number of distinct elements
    k = CLng(cboMultisetK.List(cboMultisetK.ListIndex))

    AddItemText "Multiset: " & txtMultiset.Text
    AddItemText "Number of ""choose " & CStr(k) & """ combinations: " & CStr(BoundedCompositionsCount(e, n, k)) & vbCrLf
End Sub


'Algorithm 7.2.1.1 H - Loopless reflected mixed-radix Gray generation
Public Sub LooplessGrayMixedReflect(m() As Long, n As Long)
    Dim j As Long, k As Long
    Dim c() As Long, f() As Long, o() As Long
    Dim stb() As Byte, sti As Long

    'Visiting a Gray code here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    k = 1
    For j = 0 To n - 1
        k = k * m(j)
    Next j
    ReDim stb(2 * k * (2 * n + 1) - 1)
    sti = 0

    AddItemText "Visit all mixed-radix n-tuples of " & CStr(n) & " components such that each iteration increments or decrements one component by 1"
    AddItemText "Number of n-tuples: " & CStr(k)

    ReDim c(n - 1)
    ReDim f(n)
    ReDim o(n - 1)
    For j = 0 To n - 1
        c(j) = 0
        f(j) = j
        o(j) = 1
    Next j
    f(n) = n

    Do
        'Visit n-tuple c(0..n-1)
        For j = n - 1 To 0 Step -1
            If (j <> (n - 1)) Then 'insert space character
                stb(sti) = 32
                sti = sti + 2
            End If
            stb(sti) = c(j) + 48
            sti = sti + 2
        Next j
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        j = f(0)
        f(0) = 0
        If (j = n) Then Exit Do
        c(j) = c(j) + o(j)
        If ((c(j) = 0) Or (c(j) = m(j) - 1)) Then
            o(j) = -o(j)
            f(j) = f(j + 1)
            f(j + 1) = j + 1
        End If
    Loop

    AddItemText CStr(stb)
End Sub

Private Sub LooplessGrayMixedMod(m() As Long, n As Long)
    Dim j As Long, k As Long
    Dim c() As Long, f() As Long, o() As Long
    Dim stb() As Byte, sti As Long

    'Visiting a Gray code here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    k = 1
    For j = 0 To n - 1
        k = k * m(j)
    Next j
    ReDim stb(2 * k * (2 * n + 1) - 1)
    sti = 0

    AddItemText "Visit all mixed-radix n-tuples of " & CStr(n) & " components such that each iteration modulo-increments one component by 1"
    AddItemText "Number of n-tuples: " & CStr(k)

    ReDim c(n - 1)
    ReDim f(n)
    ReDim o(n - 1)
    For j = 0 To n - 1
        c(j) = 0
        f(j) = j
        o(j) = m(j) - 1
    Next j
    f(n) = n

    Do
        'Visit n-tuple c(0..n-1)
        For j = n - 1 To 0 Step -1
            If (j <> (n - 1)) Then 'insert space character
                stb(sti) = 32
                sti = sti + 2
            End If
            stb(sti) = c(j) + 48
            sti = sti + 2
        Next j
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        j = f(0)
        f(0) = 0
        If (j = n) Then Exit Do
        c(j) = (c(j) + 1) Mod m(j)
        If (c(j) = o(j)) Then
            If (o(j) = 0) Then
                o(j) = o(j) - 1 + m(j)
            Else
                o(j) = o(j) - 1
            End If
            f(j) = f(j + 1)
            f(j + 1) = j + 1
        End If
    Loop

    AddItemText CStr(stb)
End Sub

Private Sub cmdTupleVisit_Click(Index As Integer)
    Dim i As Long, n As Long, m() As Long

    n = CLng(cboTupleVisitN.Text)
    ReDim m(n - 1)
    For i = 0 To n - 1
        m(i) = CLng(cboTupleVisitM.Text)
    Next i
    Select Case Index
    Case 0: LooplessGrayMixedReflect m, n
    Case 1: LooplessGrayMixedMod m, n
    End Select
End Sub



'If bits of a Long represent a Gray code, they can be generated like this:
'code = index xor (index\2)
Public Function GrayDecode(ByVal code As Long) As Long
    Dim i As Long, k As Long

    k = 2
    Do
        i = code
        code = code Xor (code \ k)
        If ((i <= 1) Or (k = &H10000)) Then Exit Do
        k = k * k
    Loop
    GrayDecode = code
End Function

'Algorithm 7.2.1.1 L - Loopless Gray binary generation
Private Sub LooplessGrayBinary(n As Long)
    Dim j As Long
    Dim c() As Long, f() As Long
    Dim stb() As Byte, sti As Long

    'Visiting a Gray code here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    j = 2 ^ n
    ReDim stb(2 * j * (n + 2) - 1)
    sti = 0

    AddItemText "Visit all binary tuples of " & CStr(n) & " components such that each iteration toggles one bit"
    AddItemText "Number of binary tuples: " & CStr(j)

    ReDim c(n - 1)
    ReDim f(n)
    For j = 0 To n - 1
        c(j) = 0
        f(j) = j 'focus pointer
    Next j
    f(n) = n
    Do
        'Visit n-tuple c(0..n-1)
        For j = n - 1 To 0 Step -1 'msb..lsb
            stb(sti) = c(j) + 48  'c(j) = {0, 1}
            sti = sti + 2
        Next j
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        j = f(0)
        f(0) = 0
        If (j = n) Then Exit Do
        f(j) = f(j + 1)
        f(j + 1) = j + 1
        c(j) = 1 - c(j)
    Loop

    AddItemText CStr(stb)
End Sub

'calculate delta sequence u(0..2^n-1) for balanced n-bit Gray code
'n >= 2
Public Function BalancedGrayDelta(u() As Long, n As Long) As Long
    Dim i As Long, j As Long, k As Long, l As Long, m As Long, p As Long
    Dim b() As Long, c() As Long, ji() As Long
    Dim maxx As Long, maxi As Long, maxis As Long

    ReDim b(n - 1)
    ReDim c(n - 1)
    ReDim u(2 ^ n - 1)

    If ((n And 1) = 0) Then 'even - seed u(0..4) with Gray code delta sequence for n=2
        m = 2
        u(0) = 0: u(1) = 1: u(2) = 0: u(3) = 1
    Else                    'odd  - seed u(0..8) with Gray code delta sequence for n=3
        m = 3
        u(0) = 0: u(1) = 1: u(2) = 2: u(3) = 1: u(4) = 0: u(5) = 1: u(6) = 2: u(7) = 1
    End If

    While (m < n)
        For j = 0 To m - 1
            b(j) = 0
            c(j) = 0
        Next j
        For j = 0 To 2 ^ m - 1
            c(u(j)) = c(u(j)) + 1
        Next j
        For j = 0 To m - 1
            c(j) = c(j) * 4
        Next j

        l = 1
        c(u(2 ^ m - 1)) = c(u(2 ^ m - 1)) - 4

        maxx = -1
        maxis = 0
        Do
            'find max(c(j))  0 <= j < m
            For j = 0 To m - 1
                If (maxx < c(j)) Then
                    maxx = c(j)
                    maxi = j
                End If
            Next j
            If ((maxx - l - 1) <= 2) Then Exit Do
            b(maxi) = b(maxi) + 1 'b(maxi) = number of maxi's to underline
            c(maxi) = c(maxi) - 2
            l = l + 1 '2
            maxis = maxis + 1
            maxx = maxx - 2
        Loop

        ReDim ji(l)
        ji(l - 1) = 2 ^ m - 1

        k = 0
        i = 0
        Do While (i < 2 ^ m)
            For j = m - 1 To 0 Step -1
                If (b(u(i)) <> 0) Then
                    ji(k) = i
                    k = k + 1
                    b(u(i)) = b(u(i)) - 1
                    If (maxis = k) Then Exit Do
                    Exit For
                End If
            Next j
            i = i + 1
        Loop

For i = l To 1 Step -1
    ji(i) = ji(i - 1)
Next i
ji(i) = -1

        'Expand m-bit Gray code delta sequence u(0..2^m-1) to an (m+2)-bit Gray code delta sequence u(0..2^(m+2)-1)
        i = 2 ^ (m + 2) - 1
        u(i) = m
        For j = 1 To l - 1
            For p = ji(j - 1) + 1 To ji(j) - 1
                i = i - 1
                u(i) = u(p)
            Next p
            i = i - 1
            u(i) = u(ji(j))
        Next j
        For p = ji(j - 1) + 1 To ji(j) - 1
            i = i - 1
            u(i) = u(p)
        Next p
        If ((l And 1) = 0) Then
            i = i - 1
            u(i) = m
        Else
            i = i - 1
            u(i) = m + 1
        End If
        j = l
        Do
            For p = ji(j) - 1 To ji(j - 1) + 1 Step -1
                i = i - 1
                u(i) = u(p)
            Next p
            If (((l And j) And 1) = 0) Then
                i = i - 1
                u(i) = m + 1
            Else
                i = i - 1
                u(i) = m
            End If
            For p = ji(j - 1) + 1 To ji(j) - 1
                i = i - 1
                u(i) = u(p)
            Next p
            If (((l And j) And 1) = 0) Then
                i = i - 1
                u(i) = m
            Else
                i = i - 1
                u(i) = m + 1
            End If
            For p = ji(j) - 1 To ji(j - 1) + 1 Step -1
                i = i - 1
                u(i) = u(p)
            Next p
            If (j = 1) Then Exit Do
            j = j - 1
            i = i - 1
            u(i) = u(ji(j))
        Loop
        m = m + 2
        Debug.Assert (i = 0)
    Wend
End Function


Private Sub cmdGrayVisit_Click()
    Dim n As Long

    n = CLng(cboGrayCodeVisitN.Text)
    LooplessGrayBinary n
End Sub

Private Sub cmdGrayLongVisit_Click()
    Dim i As Long, k As Long, n As Long, code As Long, s As String

    n = CLng(cboGrayCodeVisitN.Text)
    AddItemText "Visit all binary tuples of " & CStr(n) & " components such that each iteration toggles one bit, bits are in Long variable type"
    code = 0
    For i = 0 To 2 ^ n - 1
        code = i Xor (i \ 2) 'i-th Gray code
        k = 2 ^ (n - 1) 'add the binary value of code to the list
        While (k <> 0)
            s = s & IIf((code And k) = 0, "0", "1")
            k = k \ 2
        Wend
        s = s & "  " & CStr(GrayDecode(code)) & vbCrLf 'decode the Gray code back to its index
    Next i
    AddItemText s
End Sub

Private Sub cmdGrayBalancedVisit_Click()
    'GrayBalancedBinary 0
    Dim i As Long, j As Long, n As Long, p As Long, a() As Long, u() As Long
    Dim stb() As Byte, sti As Long

    n = CLng(cboGrayCodeVisitN.Text)

    'Visiting a Gray code here simply appends it to a list to be displayed.  A byte array
    'is used while building the list because appending to a string is very slow.
    j = 2 ^ n
    ReDim stb(2 * j * (n + 2) - 1)
    sti = 0

    AddItemText "Visit all binary tuples of " & CStr(n) & " components such that each iteration toggles one bit, and each " & _
                "component toggles an equal (or near equal) number of times during the sequence"
    AddItemText "Number of binary n-tuples: " & CStr(j)

    p = BalancedGrayDelta(u, n)
    p = 2 ^ n

    ReDim a(n - 1)
    For i = 0 To p - 1
        'Visit n-tuple c(0..n-1)
        For j = 0 To n - 1
            stb(sti) = a(j) + 48  'a(j) = {0, 1}
            sti = sti + 2
        Next j
        stb(sti) = 13
        stb(sti + 2) = 10
        sti = sti + 4
        'End Visit

        a(u(i)) = 1 - a(u(i))
    Next i

    AddItemText CStr(stb)
End Sub


'                       ************  PARTITIONS  ************

'Algorithm 7.2.1.4 P - Partitions in reverse lexicographic order
'Visit all integer partitions of n
'Required: 1 <= n
Private Sub PartitionVisit(ByVal n As Long)
    Dim i As Long, j As Long, q As Long, x As Long
    Dim c() As Long
    Dim st As String

    AddItemText "Visit all integer partitions of " & CStr(n) & " elements in reverse lexicographic order"
    AddItemText "Number of partitions: " & CStr(PartitionCount(n))

    ReDim c(n) 'initialize
    For i = n To 2 Step -1
        c(i) = 1
    Next i
    c(0) = 0
    Do
        c(i) = n 'store final part
        If (n = 1) Then
            q = i - 1
        Else
            q = i
        End If
        Do
            'Visit Partition c(1..i)
            For j = 1 To i                                       'c(1) >= c(2) >= ... >= c(i) >= 1
                st = st & CStr(c(j)) & IIf(j = i, vbCrLf, " + ") 'c(1) + c(2) + ... + c(i) = n
            Next j                                               '1 <= i <= n
            'End Visit

            If (c(q) <> 2) Then Exit Do
            c(q) = 1 'change 2 to 1+1
            q = q - 1
            i = i + 1
        Loop
        If (q = 0) Then Exit Do 'decrease c(q)
        x = c(q) - 1
        c(q) = x
        n = i - q + 1
        i = q + 1
        While (n > x) 'copy x if necessary
            c(i) = x
            i = i + 1
            n = n - x
        Wend
    Loop

    AddItemText st
End Sub

'Algorithm 7.2.1.4 H - Partitions into m parts
'Visit all k-part integer partitions of n
'Required: 2 <= k <= n
Private Sub PartitionVisitParts(n As Long, k As Long)
    Dim j As Long, s As Long, x As Long
    Dim c() As Long
    Dim st As String

    AddItemText "Visit all " & CStr(k) & "-part integer partitions of " & CStr(n) & " elements"
    AddItemText "Number of partitions: " & CStr(PartitionCountParts(n, k))

    ReDim c(1 To k) 'initialize
    c(1) = n - k + 1
    For j = 2 To k
        c(j) = 1
    Next j
    Do
        Do
            'Visit Partition c(1..k)
            For j = 1 To k                                       'c(1) >= c(2) >= ... >= c(k) >= 1
                st = st & CStr(c(j)) & IIf(j = k, vbCrLf, " + ") 'c(1) + c(2) + ... + c(k) = n
            Next j                                               '2 <= k <= n
            'End Visit

            If (c(2) >= c(1) - 1) Then Exit Do
            c(1) = c(1) - 1 'tweak c(1..2)
            c(2) = c(2) + 1
        Loop
        s = c(1) + c(2) - 1 'find j
        For j = 3 To k
            If (c(j) < c(1) - 1) Then Exit For
            s = s + c(j)
        Next j
        If (j > k) Then Exit Do 'increase c(j)
        x = c(j) + 1
        c(j) = x
        While (j > 2) 'tweak c(1..j)
            j = j - 1
            c(j) = x
            s = s - x
        Wend
        c(1) = s
    Loop

    AddItemText st
End Sub

'Algorithm 7.2.1.5 H - Restricted growth strings in lexicographic order
'Visit all Set Partitions of n elements
'a(0..n-1) should be distinct elements
'Required: 1 < n
Private Sub SetPartitionVisit(a() As Long, n As Long)
    Dim i As Long, j As Long, k As Long, m As Long
    Dim c() As Long, b() As Long
    Dim st As String
    Const lbracket As String = "{", rbracket As String = "}"

    AddItemText "Visit all set partitions of " & CStr(n) & " distinct elements"
    AddItemText "Number of partitions: " & CStr(BellNumber(n))

Debug.Assert (n >= 2) And (n <= 7) '(n >= 2) required, (n <= 7) limits size to prevent long process
If ((n < 2) Or (n > 7)) Then Exit Sub

    ReDim b(1 To n)
    ReDim c(1 To n)
    For j = 1 To n - 1
        b(j) = 1
        c(j) = 0
    Next j
    c(j) = 0
    m = 1
    Do
        'Visit Partition
        i = 0 'partition index
        k = n 'elements not yet assigned to a partition
        st = st & lbracket
        Do
            For j = 1 To n 'scan elements
                If (c(j) = i) Then 'element j belongs to partition i
                    st = st & IIf(Right$(st, 1) = lbracket, "", " ") & Chr$(a(j - 1))
                    k = k - 1
                    If (k = 0) Then
                        st = st & rbracket & vbCrLf
                        Exit Do
                    End If
                End If
            Next j
            st = st & rbracket & " " & lbracket
            i = i + 1
        Loop While (i <= n)
        'End Visit

        If (c(n) = m) Then
            j = n - 1 'find j
            While (c(j) = b(j))
                j = j - 1
            Wend
            If (j = 1) Then Exit Do 'increase a(j)
            c(j) = c(j) + 1
            If (c(j) = b(j)) Then 'zero out a(j+1..n)
                m = b(j) + 1
            Else
                m = b(j)
            End If
            j = j + 1
            While (j < n)
                c(j) = 0
                b(j) = m
                j = j + 1
            Wend
            c(n) = 0
        Else
            c(n) = c(n) + 1 'increase a(n)
        End If
    Loop

    AddItemText st
End Sub

'Algorithm 7.2.1.5 M - Multipartitions in decreasing lexicographic order
'm > 0
'e(1..m) > 0
Private Sub PartitionVisitMultiset(p() As Long, e() As Long, m As Long)
    Dim a As Long, b As Long, j As Long, k As Long, l As Long, x As Long
    Dim c() As Long, f() As Long, u() As Long, v() As Long
    Const lbracket As String = "{", rbracket As String = "}"
Dim st As String
Dim h As Long

    k = 0
    For j = 1 To m
        k = k + e(j)
    Next j

    AddItemText "Visit all partitions of a " & CStr(k) & " element multiset"

    ReDim f(k)
    ReDim c(m * k)
    ReDim u(m * k)
    ReDim v(m * k)

    For j = 0 To m - 1
        c(j) = j + 1
        u(j) = e(j + 1)
        v(j) = e(j + 1)
    Next j
    f(0) = 0
    a = 0
    l = 0
    f(1) = m
    b = m

Dim crashguard As Long
crashguard = GetTickCount + 5000

    Do
        j = a
        k = b
        x = 0
        While (j < b)
            u(k) = u(j) - v(j)
            If (u(k) = 0) Then
                x = 1
                j = j + 1
            ElseIf (x = 0) Then
                c(k) = c(j)
                If v(j) < u(k) Then
                    v(k) = v(j)
                Else
                    v(k) = u(k)
                End If
                If (u(k) < v(j)) Then
                    x = 1
                Else
                    x = 0
                End If
                k = k + 1
                j = j + 1
            Else
                c(k) = c(j)
                v(k) = u(k)
                k = k + 1
                j = j + 1
            End If
        Wend

        If (k > b) Then
            a = b
            b = k
            l = l + 1
            f(l + 1) = b
        Else
            'Visit Partition
            For k = 0 To l 'l+1 partitions
                st = st & IIf(k = 0, lbracket, " " & lbracket)
                For j = f(k) To f(k + 1) - 1
                    For h = 1 To v(j)
                        st = st & IIf(Right$(st, 1) = lbracket, "", " ") & Chr$(p(c(j) - 1))
                    Next h
                Next j
                st = st & rbracket & IIf(k = l, vbCrLf, "")
            Next k
            'End Visit

If (GetTickCount > crashguard) Then
    st = st & "EXIT DUE TO TIMEOUT" & vbCrLf
    Exit Do
End If
            j = b - 1
            While (v(j) = 0)
                j = j - 1
            Wend
            While ((j = a) And (v(j) = 1))
                If (l = 0) Then Exit Do                    '<<<
                l = l - 1
                b = a
                a = f(l)
                j = b - 1
                While (v(j) = 0)
                    j = j - 1
                Wend
            Wend
            v(j) = v(j) - 1
            For k = j + 1 To b - 1
                v(k) = u(k)
            Next k
        End If
    Loop

    AddItemText st
End Sub


Private Sub cmdPartitionVisit_Click(Index As Integer)
    Dim i As Long, k As Long, n As Long, a() As Long

    n = CLng(cboPartitionVisitN.Text)
Select Case Index
Case 0
    Select Case n
    Case Is < 1, Is > 22
        AddItemText "Number of partitions: " & CStr(PartitionCount(n)) & vbCrLf
    Case Else
        PartitionVisit n 'n >= 1
    End Select
Case 1
    'all m part partitions of n
    k = CLng(cboPartitionVisitK.Text)
    PartitionVisitParts n, k 'n >= m >= 2
Case 2
    If (n < 2) Then
        AddItemText "Invalid Input" & vbCrLf
    Else
        ReDim a(n - 1)
        For i = 0 To n - 1 'elements of a(0..n-1) should be distinct
            a(i) = i + 97  'asc("a"), asc("b"), ...
        Next i
        SetPartitionVisit a, n
    End If
End Select
End Sub

Private Sub cmdPartitionMultisetVisit_Click()
    Dim h As Long, i As Long, j As Long, k As Long, n As Long
    Dim a() As Long, m() As Long

    'copy ASCII values of txtPermutate.Text to a(0..n-1)
    n = Len(txtPartitionMultisetVisit.Text)
    If (n = 0) Then
        AddItemText "No Elements" & vbCrLf
        Exit Sub
    End If
    ReDim a(n - 1)
    ReDim m(n - 1)
    For j = 0 To n - 1
        a(j) = Asc(Mid$(txtPartitionMultisetVisit.Text, j + 1, 1))
    Next j

    'sort a(0..n-1)
    For i = 1 To n - 1
        k = a(i)
        For j = i - 1 To 0 Step -1
            If (a(j) <= k) Then Exit For
            a(j + 1) = a(j)
        Next j
        a(j + 1) = k
    Next i

    h = 1
    j = 0
    For i = 1 To n - 1
        If (a(i) = a(j)) Then
            h = h + 1
        Else
            m(j) = h     'm(j) is number of instances of a(j)
            j = j + 1
            h = 1
            a(j) = a(i)  'a(j) is distinct in a(0..j)
        End If
    Next i
    m(j) = h
    n = j + 1 'number of distinct elements

    ReDim Preserve m(n)
    For i = n To 1 Step -1
        m(i) = m(i - 1)
    Next i

    PartitionVisitMultiset a, m, n
End Sub
