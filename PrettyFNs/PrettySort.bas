Attribute VB_Name = "mPrettySort"
Option Explicit                       ' -©Rd 04/08-

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' Pretty Sort string array sorting algorithms.
'
' Intuitive/natural sorting and comparison functions.
'
' The following are intended for sorting filenames that
' may contain numbers in a more intuitive order:
'
'  strPrettyFileNames
'  strPrettyFileNamesIndexed
'  StrCompFileNames
'
' The following are intended for sorting strings that
' may contain numbers in a more intuitive order:
'
'  strPrettyNumSort
'  strPrettyNumSortIndexed
'  StrCompNumbers
'
' The following are intended for sorting strings that
' do not contain numbers in a more intuitive order:
'
'  strPrettySort
'  strPrettySortIndexed
'
' You are free to use any part or all of this code even for
' commercial purposes in any way you wish under the one condition
' that no copyright notice is moved or removed from where it is.
'
' For comments, suggestions or bug reports you can contact me at:
' rd•edwards•bigpond•com.
'
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' This constant defines the maximum allowed occurence of numeric character
' groups (1 or more) within any string item being passed to these routines:
Private Const MAX_DISCRETE_OCCUR_NUMS As Long = 256&

' For example, this string has 4 occurences: a6string 07with 89 four3occurs

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Declare some CopyMemory Alias's (thanks Bruce :)
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lLenB As Long)
Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lLenB As Long)
Private Declare Function AllocStrB Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal lpszStr As Long, ByVal lLenB As Long) As Long

' More efficient repeated use of numeric literals
Private Const n0 = 0&, n1 = 1&, n2 = 2&, n3 = 3&, n4 = 4&, n5 = 5&, n6 = 6&
Private Const n7 = 7&, n8 = 8&, n12 = 12&, n16 = 16&, n32 = 32&, n64 = 64&
Private Const nKB As Long = 1024&

Private Enum SAFEATURES
    FADF_AUTO = &H1              ' Array is allocated on the stack
    FADF_STATIC = &H2            ' Array is statically allocated
    FADF_EMBEDDED = &H4          ' Array is embedded in a structure
    FADF_FIXEDSIZE = &H10        ' Array may not be resized or reallocated
    FADF_BSTR = &H100            ' An array of BSTRs
    FADF_UNKNOWN = &H200         ' An array of IUnknown*
    FADF_DISPATCH = &H400        ' An array of IDispatch*
    FADF_VARIANT = &H800         ' An array of VARIANTs
    FADF_RESERVED = &HFFFFF0E8   ' Bits reserved for future use
    #If False Then
        Dim FADF_AUTO, FADF_STATIC, FADF_EMBEDDED, FADF_FIXEDSIZE, FADF_BSTR, FADF_UNKNOWN, FADF_DISPATCH, FADF_VARIANT, FADF_RESERVED
    #End If
End Enum
Private Const VT_BYREF = &H4000& ' Tests whether the InitedArray routine was passed a Variant that contains an array, rather than directly an array in the former case ptr already points to the SA structure. Thanks to Monte Hansen for this fix

Private Type SAFEARRAY
    cDims       As Integer       ' Count of dimensions in this array (only 1 supported)
    fFeatures   As Integer       ' Bitfield flags indicating attributes of a particular array
    cbElements  As Long          ' Byte size of each element of the array
    cLocks      As Long          ' Number of times the array has been locked without corresponding unlock
    pvData      As Long          ' Pointer to the start of the array data (use only if cLocks > 0)
End Type
Private Type SABOUNDS            ' This module supports single dimension arrays only
    cElements As Long            ' Count of elements in this dimension
    lLBound   As Long            ' The lower-bounding index of this dimension
    lUBound   As Long            ' The upper-bounding index of this dimension
End Type

Private qs4Lb() As Long, qs4Ub() As Long ' Non-stable non-recursive quicksort stacks
Private ss2Lb() As Long, ss2Ub() As Long ' Stable non-recursive quicksort stacks
Private tw4Lb() As Long, tw4Ub() As Long ' Stable insert/binary runner stacks
Private lA_1() As Long, lA_2() As Long   ' Stable quicksort working buffers
Private qs4Max As Long, ss2Max As Long
Private tw4Max As Long, bufMax As Long

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Type tPrettySort
    occurs As Long
    idxn(1 To MAX_DISCRETE_OCCUR_NUMS) As Long
    cnums(1 To MAX_DISCRETE_OCCUR_NUMS) As Long
    cpads(1 To MAX_DISCRETE_OCCUR_NUMS) As Long
End Type

Public Enum ePrettyFiles
    GroupByExtension = 0&
    GroupByFolder = 1&
    #If False Then
        Dim GroupByExtension, GroupByFolder
    #End If
End Enum

Public Enum eCompare
    Lesser = -1&
    Equal = 0&
    Greater = 1&
    #If False Then
        Dim Lesser, Equal, Greater
    #End If
End Enum

Public Enum eSortOrder
    Descending = -1&
    Default = 0&
    Ascending = 1&
    #If False Then
        Dim Descending, Default, Ascending
    #End If
End Enum

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const Default_Direction As Long = Ascending

Private mComp As eCompare
Private mCriteria As VbCompareMethod
Private mSortOrder As eSortOrder

Private padZs As String

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

' The following properties should be set before sorting.

Property Get SortMethod() As VbCompareMethod
    SortMethod = mCriteria
End Property

Property Let SortMethod(ByVal NewMethod As VbCompareMethod)
    mCriteria = NewMethod
End Property

Property Get SortOrder() As eSortOrder
    If mSortOrder = Default Then mSortOrder = Default_Direction
    SortOrder = mSortOrder
End Property

Property Let SortOrder(ByVal NewDirection As eSortOrder)
    If NewDirection = Default Then
        If mSortOrder = Default Then mSortOrder = Default_Direction
    Else
        mSortOrder = NewDirection
    End If
End Property

' + Pretty File Names Sorter +++++++++++++++++++++++++++

' This sub will sort string array items containing numeric
' characters in a more intuitive order. It will take into
' account all occurences of numbers in the string item.

' It is intended for sorting filenames, and will apply the
' same intuitive order for folders in the file path if they
' contain numeric characters.

' Specifying a path is not required, neither do they need to
' have extensions, in fact, they do not need to be filenames;
' just strings that may contain numbers grouped together
' within the string text.

' In other words, it can be used for normal pretty sorting
' operations. But please note, this is intended for sorting
' filenames and is a little slower than the strPrettyNumSort
' sub below because of extra code to handle the extensions.

' If GroupByExtension is specified it will order the items by
' extension, whilst still producing an intuitive order of the
' file path and names that may have numbers within them and
' that have the same extension.

' Note: this will also sort pure numbers, which will group
' positive and negative numbers together from small values
' to large values (or reversed) irrespective of their sign.

' This is version two of the Pretty File Names routine which
' also handles numbers in the extension.

Sub strPrettyFileNames(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True, Optional ByVal Group As ePrettyFiles = GroupByExtension) '-©Rd-
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim aIdx() As Long, lAbuf() As Long
    Dim lpS As Long, lpL As Long
    Dim walk As Long, cnt As Long

    cnt = ubA - lbA + n1            ' Grab array item count
    If (cnt < n1) Then Exit Sub     ' If nothing to do then exit

    strPrettyFileNamesIndexed sA, aIdx, lbA, ubA, CapsFirst, Group

    ' Now use copymemory to copy the string pointers to a long array buffer for later
    ' reference to the original strings. We need this to re-order the strings in the
    ' last step that would over-write needed items if the buffer was not used.

    ReDim lAbuf(lbA To ubA) As Long
    lpS = VarPtr(sA(lbA))
    lpL = VarPtr(lAbuf(lbA))
    CopyMemByV lpL, lpS, cnt * n4

    ' Next we do the actual re-ordering of the array items by referencing the string
    ' pointers with the index array, and assigning back into the index array ready to
    ' be copied across to the string array in one copy process.

    For walk = lbA To ubA
        aIdx(walk) = lAbuf(aIdx(walk))
    Next

    ' The last step assigns the string pointers back into the original array
    ' from the pointer array buffer.

    lpL = VarPtr(aIdx(lbA))
    CopyMemByV lpS, lpL, cnt * n4
End Sub

' + Indexed Pretty File Names Sorter ++++++++++++++++++++++

Sub strPrettyFileNamesIndexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True, Optional ByVal Group As ePrettyFiles = GroupByExtension) '-©Rd-
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim numsN() As Long, numsE() As Long
    Dim sAtemp() As String, lpS As Long
    Dim atFN() As tPrettySort
    Dim atEXT() As tPrettySort
    Dim period() As Long, extens() As Long
    Dim filenmP() As Long, extensP() As Long
    Dim prevMethod As eCompare
    Dim prevOrder As eSortOrder
    Dim lenP() As Long, lpads As Long
    Dim walk As Long, lpad As Long
    Dim lPos As Long, clen As Long
    Dim lenFN As Long, lenExt As Long
    Dim item As String, lpStr As Long

    If (ubA - lbA < n1) Then Exit Sub           ' If nothing to do then exit

    lpStr = VarPtr(item)                        ' Cache pointer to the string variable
    lpS = VarPtr(sA(lbA)) - (lbA * n4)          ' Cache pointer to the string array

    ReDim sAtemp(lbA To ubA) As String          ' ReDim buffers
    ReDim atFN(lbA To ubA) As tPrettySort
    ReDim atEXT(lbA To ubA) As tPrettySort
    ReDim period(lbA To ubA) As Long
    ReDim extens(lbA To ubA) As Long
    ReDim lenP(lbA To ubA) As Long
    ReDim filenmP(lbA To ubA) As Long
    ReDim extensP(lbA To ubA) As Long

    ' First build an array of data about the filenames to use for the padding process.

    ' Because filename and extension are distinctive entities in the comparison process
    ' it is neccessary to do a good deal of preperation to produce the resulting padded
    ' items for comparison. Pre-calculating the padding data is relatively fast compared
    ' to string manipulation.

    For walk = lbA To ubA    ' Loop thru the array items one by one

        CopyMemByV lpStr, lpS + (walk * n4), n4   ' Cache current item
        clen = Len(item)                          ' Cache the items length

        lPos = InStrRev(item, "\")                    ' Determine position of the last backslash instance
        If lPos = n0 Then lPos = InStrRev(item, "/")  ' If no backslash then maybe it's a forward slash?
        period(walk) = InStrRev(item, ".")            ' Determine position of the last period character

        If period(walk) = n0 Or period(walk) < lPos Then   ' If no period or it's before the last slash
            period(walk) = clen + n1                       ' Set to item length instead, + phantom period pos
            extens(walk) = n0
        Else                                          ' Record length of this items extension
            extens(walk) = clen - period(walk)
            If extens(walk) > lenExt Then lenExt = extens(walk)
        End If

        If period(walk) - n1 > lenFN Then lenFN = period(walk) - n1
    Next
    CopyMemByR ByVal lpStr, 0&, n4   ' De-reference pointer to item variable

    ' Set nums to the maximum position that numbers can occur in the strings.

    ReDim numsN(n0 To lenFN) As Long
    ReDim numsE(n0 To lenExt) As Long

    ' Find all occurences of numeric chars in the filename portion of the items.

    For walk = lbA To ubA
        GoNumLoop atFN(walk), sA(walk), numsN, period(walk)
    Next

    ' Next calculate the padding length for all num inst's in each filename, and
    ' add them together to determine the total padding needed for each filename.

    ' The total lengths are compared to identify the longest length that will
    ' be used to pre-allocate the string lengths for faster string operations.

    For walk = lbA To ubA
        lpad = GoPadLoop(atFN(walk), numsN)   ' Calc the padding length for this filename
        If lpad > lpads Then lpads = lpad     ' Set lpads to longest padding length
        clen = period(walk) - n1 + lpad       ' Calc the length of this filename when padded
        If clen > lenFN Then lenFN = clen     ' Set lenFN to longest filename length

        filenmP(walk) = clen    ' Record the new length of this filename when padded
    Next

    ' Find all occurences of numeric chars in the extension portion of the items.

    For walk = lbA To ubA
        GoNumLoop atEXT(walk), Mid$(sA(walk), period(walk) + n1), numsE, extens(walk) + n1
    Next

    ' Next calculate the padding length for all num inst's in each extension, and
    ' add them together to determine the total padding needed for each extension.

    ' The total lengths are compared to identify the longest length that will
    ' be used to pre-allocate the string lengths for faster string operations.

    For walk = lbA To ubA
        lpad = GoPadLoop(atEXT(walk), numsE)  ' Calc the padding length for this extension
        If lpad > lpads Then lpads = lpad     ' Set lpads to longest padding length
        clen = extens(walk) + lpad            ' Calc the length of this extension when padded
        If clen > lenExt Then lenExt = clen   ' Set lenExt to longest extension length

        extensP(walk) = clen    ' Record the new length of this extension when padded
    Next

    ' Pre-allocate the buffer array strings without assigning any data into them.

    If Group = GroupByExtension Then
        For walk = lbA To ubA
            clen = lenExt + filenmP(walk)
            CopyMemByV VarPtr(sAtemp(walk)), VarPtr(AllocStrB(n0, clen + clen)), n4
            lenP(walk) = clen   ' Record total length of this item
        Next
    Else ' Group = GroupByFolder
        For walk = lbA To ubA
            clen = lenFN + extensP(walk)
            CopyMemByV VarPtr(sAtemp(walk)), VarPtr(AllocStrB(n0, clen + clen)), n4
            lenP(walk) = clen   ' Record total length of this item
        Next
    End If

    padZs = String$(lpads, "0")  ' Create pad to longest padding length

    ' Next, pad all filenames containing numeric characters, based on the
    ' longest number for that position, into the temp string array.

    ' Step through each item building the temp string with padded numeric
    ' chars using recorded info in atFN and atEXT.

    ' If GroupByExtension is specified then pad the extensions and prefix
    ' them to the temp array filenames, else pad the filenames and append
    ' the extensions.

    If Group = GroupByExtension Then
        For walk = lbA To ubA     ' Loop thru the array items one by one
            GoBufLoop atEXT(walk), Mid$(sA(walk), period(walk) + n1), sAtemp(walk), n1, lenP(walk)
            clen = extensP(walk)
            lpad = lenExt - clen  ' Write padding before filename
            If lpad > n0 Then Mid$(sAtemp(walk), clen + n1) = Space$(lpad)
            If period(walk) > n1 Then
                GoBufLoop atFN(walk), Left$(sA(walk), period(walk) - n1), sAtemp(walk), lenExt + n1, lenP(walk)
            End If
        Next
    Else ' Group = GroupByFolder
        For walk = lbA To ubA     ' Loop thru the array items one by one
            GoBufLoop atFN(walk), Left$(sA(walk), period(walk) - n1), sAtemp(walk), n1, lenP(walk)
            clen = filenmP(walk)
            lpad = lenFN - clen   ' Write padding before extension
            If lpad > n0 Then Mid$(sAtemp(walk), clen + n1) = Space$(lpad)
            If extens(walk) > n0 Then
                GoBufLoop atEXT(walk), Mid$(sA(walk), period(walk) + n1), sAtemp(walk), lenFN + n1, lenP(walk)
            End If
        Next
    End If

    ' To sort numeric values in a more intuitive order we sort with an indexing
    ' sorter to index the string array which will provide sorted indicies into all
    ' the working arrays.

    prevOrder = SortOrder        ' Cache SortOrder property
    prevMethod = SortMethod      ' Cache SortMethod property

    ' First sort with binary comparison to seperate upper and lower case letters
    ' in the order specified by CapsFirst.

    'CapsFirst: False(0) >> Descending(-1) : True(-1) >> Ascending(1)
    SortOrder = (CapsFirst * -2&) - n1
    SortMethod = vbBinaryCompare
    strSwapSort4Indexed sAtemp, idxA, lbA, ubA

    ' Next sort in the desired direction with case-insensitive comparison to group
    ' upper and lower case letters together, but with a stable sorter to preserve the
    ' requested caps-first or lower-first order.
    
    ' Notice we pass on the indexed array in its pre-sorted state to be further modified.
    
    ' Notice also we are comparing the padded items in the temp string array whose items
    ' are still in their original positions, which of course corresponds to the indices
    ' of the source string array.

    SortOrder = prevOrder        ' Reset SortOrder property
    SortMethod = vbTextCompare
    strStableSort2Indexed sAtemp, idxA, lbA, ubA

    SortMethod = prevMethod      ' Reset SortMethod property
End Sub

' + Pretty File Names Compare Function ++++++++++++++++++++++++++++++

' This function will compare two string items containing numeric characters in
' a more intuitive order. It will take into account all occurences of numbers in
' the string items including in their extensions.

' It is intended for comparing filenames, and will apply the same intuitive order
' for folders in the file path if they contain numeric characters.

' Specifying a path is not required, neither do they need to have extensions, in
' fact, they do not need to be filenames; just strings that may contain numbers
' grouped together within the string text.

' In other words, it can be used for normal pretty sorting comparisons.
' But please note, this is intended for comparing filenames and is a little
' slower than the StrCompNumbers compare function below because of extra code
' to handle the extensions.

' If SortByExtension is specified it will compare the items by extension, whilst
' still producing an intuitive order of the file path and names that may have
' numbers within them and that have the same extension.

' Note: this will also sort pure numbers, which will group positive and negative
' numbers together from small values to large values (or reversed) irrespective
' of their sign.

' This is version two of this function which also handles numbers in the extension.

Public Function StrCompFileNames(sThis As String, sThan As String, Optional ByVal CapsFirst As Boolean, Optional ByVal Group As ePrettyFiles = GroupByFolder) As eCompare '-©Rd-
    Dim periodThis As Long, extensThis As Long
    Dim periodThan As Long, extensThan As Long
    Dim tFNthis As tPrettySort, tEXTthis As tPrettySort
    Dim tFNthan As tPrettySort, tEXTthan As tPrettySort
    Dim numsN() As Long, numsE() As Long
    Dim lpad As Long, lpads As Long, clen As Long
    Dim filenmPthis As Long, extensPthis As Long
    Dim filenmPthan As Long, extensPthan As Long
    Dim lenPthis As Long, lenPthan As Long
    Dim sTempThis As String, sTempThan As String
    Dim eComp As eCompare, lPos As Long
    Dim lenFN As Long, lenExt As Long

    ' First, we gather information about our filenames to use for the
    ' padding process and to access the extensions as needed later.

    clen = Len(sThis)                              ' Cache the items length
    lPos = InStrRev(sThis, "\")                    ' Determine position of the last backslash instance
    If lPos = n0 Then lPos = InStrRev(sThis, "/")  ' If no backslash then maybe it's a forward slash?
    periodThis = InStrRev(sThis, ".")              ' Determine position of the last period character

    If periodThis = n0 Or periodThis < lPos Then   ' If no period or it's before the last slash
        periodThis = clen + n1                     ' Set to item length instead, + phantom period pos
        extensThis = n0
    Else                                           ' Record length of this items extension
        extensThis = clen - periodThis
        lenExt = extensThis
    End If
    lenFN = periodThis - n1

    clen = Len(sThan)                              ' Cache the items length
    lPos = InStrRev(sThan, "\")                    ' Determine position of the last backslash instance
    If lPos = n0 Then lPos = InStrRev(sThan, "/")  ' If no backslash then maybe it's a forward slash?
    periodThan = InStrRev(sThan, ".")              ' Determine position of the last period character

    If periodThan = n0 Or periodThan < lPos Then   ' If no period or it's before the last slash
        periodThan = clen + n1                     ' Set to item length instead, + phantom period pos
        extensThan = n0
    Else                                           ' Record length of than items extension
        extensThan = clen - periodThan
        If extensThan > lenExt Then lenExt = extensThan
    End If
    If periodThan - n1 > lenFN Then lenFN = periodThan - n1

    ' Set nums to the maximum position that numbers can occur in the strings.

    ReDim numsN(n0 To lenFN) As Long
    ReDim numsE(n0 To lenExt) As Long

    ' Find all occurences of numeric chars in the filename portion of the items.

    GoNumLoop tFNthis, sThis, numsN, periodThis
    GoNumLoop tFNthan, sThan, numsN, periodThan

    ' Next calculate the padding length for all num inst's in each filename, and
    ' add them together to determine the total padding needed for each filename.

    ' The total lengths are compared to identify which is longest, and that will
    ' be used to pre-allocate the string lengths for faster string operations.

    lpad = GoPadLoop(tFNthis, numsN)    ' Calc the padding length for this filename
    lpads = lpad                        ' Set lpads to this padding length
    clen = periodThis - n1 + lpad       ' Calc the length of this filename when padded
    lenFN = clen                        ' Set lenFN to this filename length

    filenmPthis = clen    ' Record the new length of this filename when padded

    lpad = GoPadLoop(tFNthan, numsN)    ' Calc the padding length for than filename
    If lpad > lpads Then lpads = lpad   ' Set lpads to longest padding length
    clen = periodThan - n1 + lpad       ' Calc the length of than filename when padded
    If clen > lenFN Then lenFN = clen   ' Set lenFN to longest filename length

    filenmPthan = clen    ' Record the new length of than filename when padded

    ' Find all occurences of numeric chars in the extension portion of the items.

    GoNumLoop tEXTthis, Mid$(sThis, periodThis + n1), numsE, extensThis + n1
    GoNumLoop tEXTthan, Mid$(sThan, periodThan + n1), numsE, extensThan + n1

    ' Next calculate the padding length for all num inst's in each extension, and
    ' add them together to determine the total padding needed for each extension.

    ' The total lengths are compared to identify the longest length that will
    ' be used to pre-allocate the string lengths for faster string operations.

    lpad = GoPadLoop(tEXTthis, numsE)   ' Calc the padding length for this extension
    If lpad > lpads Then lpads = lpad   ' Set lpads to longest padding length
    clen = extensThis + lpad            ' Calc the length of this extension when padded
    lenExt = clen                       ' Set lenExt to this extension length

    extensPthis = clen    ' Record the new length of this extension when padded

    lpad = GoPadLoop(tEXTthan, numsE)   ' Calc the padding length for than extension
    If lpad > lpads Then lpads = lpad   ' Set lpads to longest padding length
    clen = extensThan + lpad            ' Calc the length of than extension when padded
    If clen > lenExt Then lenExt = clen ' Set lenExt to longest extension length

    extensPthan = clen    ' Record the new length of than extension when padded

    ' Pre-allocate the buffer strings for faster string building operations.

    If Group = GroupByExtension Then
        clen = lenExt + filenmPthis
    Else 'Group = GroupByFolder
        clen = lenFN + extensPthis
    End If

    sTempThis = Space$(clen)
    lenPthis = clen   ' Record total length of this item

    If Group = GroupByExtension Then
        clen = lenExt + filenmPthan
    Else 'Group = GroupByFolder
        clen = lenFN + extensPthan
    End If

    sTempThan = Space$(clen)
    lenPthan = clen   ' Record total length of than item

    padZs = String$(lpads, "0")  ' Create pad to longest padding length

    ' Next, pad the filenames containing numeric characters, based on the
    ' longest number for that position, into the temp string variables.

    ' Step through both items building the temp string with padded numeric
    ' chars using recorded info in tFN and tEXT.

    ' If GroupByExtension is specified then pad the extensions and prefix them
    ' to the temp filenames, else pad the filenames and append the extensions.

    If Group = GroupByExtension Then

        GoBufLoop tEXTthis, Mid$(sThis, periodThis + n1), sTempThis, n1, lenPthis
        lpad = lenExt - extensPthis  ' Write padding before filename
        If lpad > n0 Then Mid$(sTempThis, extensPthis + n1) = Space$(lpad)
        If periodThis > n1 Then
            GoBufLoop tFNthis, Left$(sThis, periodThis - n1), sTempThis, lenExt + n1, lenPthis
        End If

        GoBufLoop tEXTthan, Mid$(sThan, periodThan + n1), sTempThan, n1, lenPthan
        lpad = lenExt - extensPthan  ' Write padding before filename
        If lpad > n0 Then Mid$(sTempThan, extensPthan + n1) = Space$(lpad)
        If periodThan > n1 Then
            GoBufLoop tFNthan, Left$(sThan, periodThan - n1), sTempThan, lenExt + n1, lenPthan
        End If

    Else 'Group = GroupByFolder

        GoBufLoop tFNthis, Left$(sThis, periodThis - n1), sTempThis, n1, lenPthis
        lpad = lenFN - filenmPthis   ' Write padding before extension
        If lpad > n0 Then Mid$(sTempThis, filenmPthis + n1) = Space$(lpad)
        If extensThis > n0 Then
            GoBufLoop tEXTthis, Mid$(sThis, periodThis + n1), sTempThis, lenFN + n1, lenPthis
        End If

        GoBufLoop tFNthan, Left$(sThan, periodThan - n1), sTempThan, n1, lenPthan
        lpad = lenFN - filenmPthan   ' Write padding before extension
        If lpad > n0 Then Mid$(sTempThan, filenmPthan + n1) = Space$(lpad)
        If extensThan > n0 Then
            GoBufLoop tEXTthan, Mid$(sThan, periodThan + n1), sTempThan, lenFN + n1, lenPthan
        End If

    End If

    ' Next, we compare the padded items with case-insensitive comparison.

    eComp = StrComp(sTempThis, sTempThan, vbTextCompare)

    ' If the items are equal with case-insensitive comparison we return the
    ' order specified by CapsFirst, else we return the case-insensitive result.

    If eComp = Equal Then

        ' To order items that are spelled the same in a more consistent order we
        ' compare with binary comparison to seperate upper and lower case letters
        ' in the order specified by CapsFirst.

        'CapsFirst: False(0) >> Descending(-1) : True(-1) >> Ascending(1)
        lPos = (CapsFirst * -2&) - n1
        StrCompFileNames = StrComp(sTempThis, sTempThan, vbBinaryCompare) * lPos

    Else
        StrCompFileNames = eComp
    End If

End Function

' + Pretty Number Sorter ++++++++++++++++++++++++++++++++++++++++++++

' This routine will sort string array items containing numeric characters
' in a more intuitive order. It will take into account all occurences of
' numbers in string items of any length.

' It first sorts with binary comparison to seperate upper and lower case
' letters in the order specified by CapsFirst.

' It then sorts in the desired direction with case-insensitive comparison
' to group upper and lower case letters together, but with a stable sorter
' to preserve the requested caps-first or lower-first order.

Sub strPrettyNumSort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True) '-©Rd-
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim aIdx() As Long, lAbuf() As Long
    Dim lpS As Long, lpL As Long
    Dim walk As Long, cnt As Long

    cnt = ubA - lbA + n1            ' Grab array item count
    If (cnt < n1) Then Exit Sub     ' If nothing to do then exit

    strPrettyNumSortIndexed sA, aIdx, lbA, ubA, CapsFirst

    ' Now use copymemory to copy the string pointers to a long array buffer for later
    ' reference to the original strings. We need this to re-order the strings in the
    ' last step that would over-write needed items if the buffer was not used.

    ReDim lAbuf(lbA To ubA) As Long
    lpS = VarPtr(sA(lbA))
    lpL = VarPtr(lAbuf(lbA))
    CopyMemByV lpL, lpS, cnt * n4

    ' Next we do the actual re-ordering of the array items by referencing the string
    ' pointers with the index array, and assigning back into the index array ready to
    ' be copied across to the string array in one copy process.

    For walk = lbA To ubA
        aIdx(walk) = lAbuf(aIdx(walk))
    Next

    ' The last step assigns the string pointers back into the original array
    ' from the pointer array buffer.

    lpL = VarPtr(aIdx(lbA))
    CopyMemByV lpS, lpL, cnt * n4
End Sub

' + Indexed Pretty Number Sorter ++++++++++++++++++++++++++++++++++

Sub strPrettyNumSortIndexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True) '-©Rd-
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim sAtemp() As String, lpS As Long
    Dim nums() As Long, citems() As Long
    Dim atPS() As tPrettySort
    Dim prevMethod As eCompare
    Dim prevOrder As eSortOrder
    Dim walk As Long, clen As Long
    Dim lpad As Long, lpads As Long

    If (ubA - lbA < n1) Then Exit Sub           ' If nothing to do then exit

    ReDim sAtemp(lbA To ubA) As String          ' ReDim buffers
    ReDim atPS(lbA To ubA) As tPrettySort
    ReDim citems(lbA To ubA) As Long

    lpS = VarPtr(sAtemp(lbA)) - (lbA * n4)      ' Cache pointer to the temp array

    ' First build an array of data about the items to use for the padding process. The
    ' resulting padded items will be used in the sorting process in place of the src array.
    ' Pre-calculating the padding data is relatively fast compared to string manipulation.

    For walk = lbA To ubA
        clen = Len(sA(walk))
        If clen > lpad Then lpad = clen     ' Set lmax to longest item length
        citems(walk) = clen                 ' Cache items length
    Next

    ' Set nums to the maximum position that numbers can occur in the strings.

    ReDim nums(n0 To lpad) As Long

    ' Find all occurences of numeric chars in the string array items.

    For walk = lbA To ubA
        GoNumLoop atPS(walk), sA(walk), nums, citems(walk) + n1
    Next

    ' Next calculate the padding length for all num inst's in each item, and
    ' add them together to determine the total padding needed for each item.

    ' The total lengths are compared to identify the longest length that will
    ' be used to pre-allocate the string lengths for faster string operations.

    For walk = lbA To ubA
        lpad = GoPadLoop(atPS(walk), nums)  ' Calc the padding length for this item
        If lpad > lpads Then lpads = lpad   ' Set lpads to longest padding length
        clen = citems(walk) + lpad          ' Calc the length of this item when padded
        citems(walk) = clen                 ' Record the new length of this item when padded
    Next

    ' Pre-allocate the buffer array strings without assigning any data into them.

    For walk = lbA To ubA
        clen = citems(walk)
        CopyMemByV lpS + (walk * n4), VarPtr(AllocStrB(n0, clen + clen)), n4
    Next

    padZs = String$(lpads, "0")  ' Create pad to longest padding length

    ' Next, pad all string items containing numeric characters, based on
    ' the longest number for that position, into the temp string array.

    ' Step through each item building the temp string with padded numeric
    ' chars using recorded info in atPS.

    For walk = lbA To ubA
        GoBufLoop atPS(walk), sA(walk), sAtemp(walk), n1, citems(walk)
    Next

    ' To sort numeric values in a more intuitive order we sort with an indexing
    ' sorter to index the string array which will provide sorted indicies into all
    ' the working arrays.

    prevOrder = SortOrder        ' Cache SortOrder property
    prevMethod = SortMethod      ' Cache SortMethod property
    
    ' First sort with binary comparison to seperate upper and lower case letters
    ' in the order specified by CapsFirst.

    'CapsFirst: False(0) >> Descending(-1) : True(-1) >> Ascending(1)
    SortOrder = (CapsFirst * -2&) - n1
    SortMethod = vbBinaryCompare
    strSwapSort4Indexed sAtemp, idxA, lbA, ubA

    ' Next sort in the desired direction with case-insensitive comparison to group
    ' upper and lower case letters together, but with a stable sorter to preserve the
    ' requested caps-first or lower-first order.
    
    ' Notice we pass on the indexed array in its pre-sorted state to be further modified.
    
    ' Notice also we are comparing the padded items in the temp string array whose items
    ' are still in their original positions, which of course corresponds to the indices
    ' of the source string array.

    SortOrder = prevOrder        ' Reset SortOrder property
    SortMethod = vbTextCompare
    strStableSort2Indexed sAtemp, idxA, lbA, ubA

    SortMethod = prevMethod      ' Reset SortMethod property
End Sub

' + Pretty Number Compare Function +++++++++++++++++++++++++

' This function will compare two string items containing numeric
' characters in a more intuitive order. It will take into account
' all occurences of numbers in string items of any length.

Public Function StrCompNumbers(sThis As String, sThan As String, Optional ByVal CapsFirst As Boolean) As eCompare '-©Rd-
    Dim tPNthis As tPrettySort, tPNthan As tPrettySort
    Dim sTempThis As String, sTempThan As String
    Dim lenPthis As Long, lenPthan As Long
    Dim lpad As Long, lpads As Long
    Dim nums() As Long, eComp As eCompare

    ' First, gather information about the string items to use for the padding process.

    lenPthis = Len(sThis)
    lenPthan = Len(sThan)

    ' Set nums to the maximum position that numbers can occur in the strings.

    If lenPthis > lenPthan Then
        ReDim nums(n0 To lenPthis) As Long
    Else
        ReDim nums(n0 To lenPthan) As Long
    End If

    ' Find all occurences of numeric chars in the items.

    GoNumLoop tPNthis, sThis, nums, lenPthis + n1
    GoNumLoop tPNthan, sThan, nums, lenPthan + n1

    ' Next calculate the padding length for all num inst's in each item, and
    ' add them together to determine the total padding needed for each item.

    ' The total lengths are calculated to identify the length that will be
    ' used to pre-allocate the string lengths for faster string operations.

    lpad = GoPadLoop(tPNthis, nums)   ' Calc the padding length for this item
    lpads = lpad                      ' Set lpads to this item pad length
    lenPthis = lenPthis + lpad        ' Record the new length of this item when padded

    lpad = GoPadLoop(tPNthan, nums)   ' Calc the padding length for than item
    If lpad > lpads Then lpads = lpad ' Set lpads to longest padding length
    lenPthan = lenPthan + lpad        ' Record the new length of than item when padded

    padZs = String$(lpads, "0")  ' Create pad to longest padding length

    ' Pre-allocate the buffer strings for faster string building operations.

    sTempThis = Space$(lenPthis)
    sTempThan = Space$(lenPthan)

    ' Next, pad all string items containing numeric characters, based on
    ' the longest number for that position, into the temp string items.

    ' Step through each item building the temp string with padded numeric
    ' chars using recorded info in tPNthis and tPNthan.

    GoBufLoop tPNthis, sThis, sTempThis, n1, lenPthis
    GoBufLoop tPNthan, sThan, sTempThan, n1, lenPthan

    ' Next, we compare the padded items with case-insensitive comparison.

    eComp = StrComp(sTempThis, sTempThan, vbTextCompare)

    ' If the items are equal with case-insensitive comparison we return the
    ' order specified by CapsFirst, else we return the case-insensitive result.

    If eComp = Equal Then

        ' To order items that are spelled the same in a more consistent order we
        ' compare with binary comparison to seperate upper and lower case letters
        ' in the order specified by CapsFirst.

        'CapsFirst: False(0) >> Descending(-1) : True(-1) >> Ascending(1)
        lpad = (CapsFirst * -2&) - n1
        StrCompNumbers = StrComp(sTempThis, sTempThan, vbBinaryCompare) * lpad

    Else
        StrCompNumbers = eComp
    End If

End Function

' + Pretty Sorter ++++++++++++++++++++++++++++++++++++++++++

' Sort with binary comparison to seperate upper and lower
' case letters in the order specified by CapsFirst.

' Then sort in the desired direction with case-insensitive
' comparison to group upper and lower case letters together,
' but with a stable sort to preserve the requested caps-first
' or lower-first order.

Sub strPrettySort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True) '-©Rd-
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim aIdx() As Long, lAbuf() As Long
    Dim lpS As Long, lpL As Long
    Dim walk As Long, cnt As Long

    cnt = ubA - lbA + n1            ' Grab array item count
    If (cnt < n1) Then Exit Sub     ' If nothing to do then exit

    strPrettySortIndexed sA, aIdx, lbA, ubA, CapsFirst

    ReDim lAbuf(lbA To ubA) As Long
    lpS = VarPtr(sA(lbA))
    lpL = VarPtr(lAbuf(lbA))
    CopyMemByV lpL, lpS, cnt * n4

    For walk = lbA To ubA
        aIdx(walk) = lAbuf(aIdx(walk))
    Next

    lpL = VarPtr(aIdx(lbA))
    CopyMemByV lpS, lpL, cnt * n4
End Sub

' + Indexed Pretty Sorter ++++++++++++++++++++++++++++++++++

Sub strPrettySortIndexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True) '-©Rd-
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim lAbuf() As Long
    Dim lpS As Long, lpL As Long
    Dim walk As Long, cnt As Long
    Dim prevMethod As eCompare
    Dim prevOrder As eSortOrder

    cnt = ubA - lbA + n1            ' Grab array item count
    If (cnt < n2) Then Exit Sub     ' If nothing to do then exit

    prevOrder = SortOrder           ' Cache SortOrder property
    prevMethod = SortMethod         ' Cache SortMethod property

    'False(0) >> Descending(-1) : True(-1) >> Ascending(1)
    SortOrder = (CapsFirst * -2) - 1
    SortMethod = vbBinaryCompare
    strSwapSort4Indexed sA, idxA, lbA, ubA

    SortOrder = prevOrder           ' Reset SortOrder property
    SortMethod = vbTextCompare
    strStableSort2Indexed sA, idxA, lbA, ubA

    SortMethod = prevMethod         ' Reset SortMethod property
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub GoNumLoop(tPS As tPrettySort, item As String, nums() As Long, ByVal itemEnd As Long) '-©Rd-
    Dim lPos As Long, oc As Long     ' Read pointer and occurence counter
    Dim lInst As Long, clen As Long  ' Inst start and length variables

    Do While lPos < itemEnd                         ' Do while position is before item end
        Do: lPos = lPos + n1                        ' Increment position
            If lPos = itemEnd Then Exit Do          ' Search only up to item end
        Loop Until IsNumeric(Mid$(item, lPos, n1))  ' Loop thru item until we find a numeric char

        If lPos < itemEnd Then           ' If numeric char found before item end
            oc = oc + n1                 ' Increment num inst occurence count
            lInst = lPos                 ' Cache num inst start position

            Do: lPos = lPos + n1                        ' Increment position
            Loop While IsNumeric(Mid$(item, lPos, n1))  ' Find end of num inst

            clen = lPos - lInst          ' Cache nums count of num inst
            If clen > nums(lInst) Then   ' Compare the length of num inst's at this pos and
                nums(lInst) = clen       ' set nums(startpos) to longest num inst length
            End If
            'Assert oc <= MAX_DISCRETE_OCCUR_NUMS
            tPS.idxn(oc) = lInst  ' Record start of this num inst
            tPS.cnums(oc) = clen  ' Record length of this num inst
    End If: Loop
    tPS.occurs = oc   ' Record occurences of num inst's within this item
End Sub

Private Function GoPadLoop(tPS As tPrettySort, nums() As Long) As Long '-©Rd-
    Dim clen As Long, lInst As Long  ' Write length and index variables
    Dim lpads As Long, lpad As Long  ' Total padding and occurence counter
    Dim oc As Long: oc = n1          ' Set occurence counter

    Do Until oc > tPS.occurs      ' Do until no more occurences of numeric chars
        lInst = tPS.idxn(oc)      ' Grab index of next num inst
        clen = tPS.cnums(oc)      ' Grab nums count of next num inst
        lpad = nums(lInst) - clen ' Calc pad len (longest at this pos-clen)
        tPS.cpads(oc) = lpad      ' Store pad len for this item pos
        lpads = lpads + lpad      ' Calc total padding needed for this item
        oc = oc + n1              ' Increment num inst count
    Loop
    GoPadLoop = lpads  ' Return total padding for this item
End Function

Private Sub GoBufLoop(tPS As tPrettySort, item As String, temp As String, ByVal lPos As Long, ByVal ctemp As Long) '-©Rd-
    Dim lRead As Long, oc As Long    ' Read pointer and occurence counter
    Dim clen As Long, lInst As Long  ' Write length and index variables
    lRead = n1: oc = n1              ' Set pointer and occurence variables

    Do Until oc > tPS.occurs                        ' Do until no more occurences of numeric chars
        lInst = tPS.idxn(oc)                        ' Grab index of next num inst
        clen = lInst - lRead                        ' Calc sub-str length up to next num inst
        Mid$(temp, lPos) = Mid$(item, lRead, clen)  ' Grab sub-str up to next num inst
        lPos = lPos + clen                          ' Reset write pointer pos
        clen = tPS.cpads(oc)                        ' Grab pad len for this pos
        Mid$(temp, lPos) = Left$(padZs, clen)       ' Write padding before nums
        lPos = lPos + clen                          ' Reset write pointer pos
        clen = tPS.cnums(oc)                        ' Grab nums count of this num inst
        Mid$(temp, lPos) = Mid$(item, lInst, clen)  ' Write numeric chars
        lPos = lPos + clen                          ' Reset write pointer pos
        lRead = lInst + clen                        ' Reset read pointer pos
        oc = oc + n1                                ' Increment num inst count
    Loop
    If Not lPos > ctemp Then
        Mid$(temp, lPos) = Mid$(item, lRead)        ' Assign the rest of the item to temp array
    End If
End Sub

' + Validate Index Array +++++++++++++++++++++++++++++++++++++++

' This will initialize the passed index array if it is not already.

' This sub-routine requires that the index array be passed either
' prepared for the sort process (see the For loop) or that it be
' uninitialized (or Erased).

' This permits subsequent sorting of the data without interfering
' with the index array if it is already sorted (based on criteria
' that may differ from the current process) and so is not in its
' uninitialized or primary pre-sort state produced by the For loop.

Sub ValidateIdxArray(lIdxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim bReDim As Boolean, lb As Long, ub As Long, j As Long
    lb = &H80000000: ub = &H7FFFFFFF
    bReDim = Not InitedArray(lIdxA, lb, ub)
    If bReDim = False Then
        bReDim = lbA < lb Or ubA > ub
    End If
    If bReDim Then
        ReDim lIdxA(lbA To ubA) As Long
        For j = lbA To ubA
            lIdxA(j) = j
        Next
    End If
End Sub

' + Stable QuickSort 2.2 Indexed Version ++++++++++++++++++++

' This is an indexed stable non-recursive quicksort.

' This is the latest version of my stable Avalanche algorithm, which
' is a non-recursive quicksort based algorithm that has been written
' from the ground up as a stable alternative to the blindingly fast
' quicksort.

' It has the benifit of indexing which allows the source array to
' remain unchanged. This also allows the index array to be passed
' on to other sort processes to be further manipulated.

' It uses a long array that holds references to the string arrays
' indices. This is known as an indexed sort. No changes are made
' to the source string array.

' After a sort procedure is run the long array is ready as a sorted
' index to the string array items.

' E.G sA(idxA(lo)) returns the lo item in the string array whose
' index may be anywhere in the string array.

Sub strStableSort2Indexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    ' This is my indexed stable non-recursive quick sort
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim item As String, lpStr As Long, lpS As Long
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, run As Long, cast As Long
    Dim idx As Long, optimal As Long, pvt As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim ceil As Long, mezz As Long, lpB As Long
    Dim inter1 As Long, inter2 As Long
    Dim lpL_1 As Long, lpL_2 As Long
    Dim idxItem As Long, lpI As Long
    Dim lPrettyReverse As Long

    idx = ubA - lbA + n1                  ' Grab array item count
    If (idx < n2) Then Exit Sub           ' If nothing to do then exit
    mComp = SortOrder                     ' Initialize compare variable
    pvt = (idx \ nKB) + n32               ' Allow for worst case senario + some

    ValidateIdxArray idxA, lbA, ubA            ' Initialize the index array
    InitializeStacks ss2Lb, ss2Ub, ss2Max, pvt ' Initialize pending boundary stacks
    InitializeStacks tw4Lb, tw4Ub, tw4Max, pvt ' Initialize pending runner stacks
    InitializeStacks lA_1, lA_2, bufMax, idx   ' Initialize working buffers

    lpL_1 = VarPtr(lA_1(n0))              ' Cache pointer to lower buffer
    lpL_2 = VarPtr(lA_2(n0))              ' Cache pointer to upper buffer
    lpStr = VarPtr(item)                  ' Cache pointer to the string variable
    lpS = VarPtr(sA(lbA)) - (lbA * n4)    ' Cache pointer to the string array
    lpI = VarPtr(idxA(lbA)) - (lbA * n4)  ' Cache pointer to the index array

    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA    ' Get pivot index position
        idxItem = idxA(pvt)               ' Grab current value into item
        CopyMemByV lpStr, lpS + (idxItem * n4), n4

        For idx = lbA To pvt - n1
            If (StrComp(sA(idxA(idx)), item, mCriteria) = mComp) Then ' (idx > item)
                lA_2(ptr2) = idxA(idx)    ' 3
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = idxA(idx)    ' 1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For idx = pvt + n1 To ubA
            If (StrComp(item, sA(idxA(idx)), mCriteria) = mComp) Then ' (idx < item)
                lA_1(ptr1) = idxA(idx)    ' 2
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = idxA(idx)    ' 4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2i ©Rd-
        lpB = VarPtr(idxA(lbA))           ' Cache pointer to current lo
        CopyMemByV lpB, lpL_1, ptr1 * n4
        idxA(lbA + ptr1) = idxItem
        CopyMemByV lpB + ((ptr1 + n1) * n4), lpL_2, ptr2 * n4

        If (ubA - lbA < n64) Then         ' Ignore false indicators
        ElseIf (inter2 = n0) And (inter1 = ptr1) Then                         ' Delegate to built-in Insert/Binary hybrid on ideal data state
            walk = lbA: mezz = ubA: idx = n0                                  ' Initialize our walker variables
            Do While walk < mezz ' ----==================================---- ' Do the twist while there's more items
                walk = walk + n1                                              ' Walk up the array and use binary search to insert each item down into the sorted lower array)
                CopyMemByV lpStr, lpS + (idxA(walk) * n4), n4                 ' Grab current value into item
                find = walk                                                   ' Default to current position
                ceil = walk - n1                                              ' Set ceiling to current position - 1
                base = lbA                                                    ' Set base to lower bound
                Do While StrComp(sA(idxA(ceil)), item, mCriteria) = mComp '   ' While current item must move down
                    midd = (base + ceil) \ n2                                 ' Find mid point
                    Do Until StrComp(sA(idxA(midd)), item, mCriteria) = mComp ' Step back up if below
                        base = midd + n1                                      ' Bring up the base
                        midd = (base + ceil) \ n2                             ' Find mid point
                        If midd = ceil Then Exit Do                           ' If we're up to ceiling
                    Loop                                                      ' Out of loop >= target pos
                    find = midd                                               ' Set provisional to new ceiling
                    If find = base Then Exit Do                               ' If we're down to base
                    ceil = midd - n1                                          ' Bring down the ceiling
                Loop '-Twister v4i ©Rd-    .       . ...   .               .  ' Out of binary search loops
                If (find < walk) Then                                         ' If current item needs to move down
                    CopyMemByV lpStr, lpS + (idxA(find) * n4), n4
                    run = walk + n1
                    Do Until run > mezz Or run - walk > n32                   ' Runner do loop
                        If Not (StrComp(item, sA(idxA(run)), mCriteria) = mComp) Then Exit Do
                        run = run + 1
                    Loop: cast = (run - walk)
                    CopyMemByV lpL_2, lpI + (walk * n4), cast * n4            ' Grab current value(s)
                    CopyMemByV lpI + ((find + cast) * n4), lpI + (find * n4), (walk - find) * n4 ' Move up items
                    CopyMemByV lpI + (find * n4), lpL_2, cast * n4            ' Re-assign current value(s) into found pos
                    If cast > n1 Then
                        If Not run > mezz Then
                            idx = idx + n1
                            tw4Lb(idx) = run - n1
                            tw4Ub(idx) = mezz
                        End If
                        walk = find
                        mezz = find + cast - n1
                End If: End If
                If walk = mezz Then
                    If idx Then
                        walk = tw4Lb(idx)
                        mezz = tw4Ub(idx)
                        idx = idx - n1
            End If: End If: Loop          ' Out of walker do loop
            ' ----===========================----
            ptr1 = n0: ptr2 = n0: inter2 = n0
        ElseIf (lPrettyReverse = n0) Then
            If (inter1 = n0) Then
                If (inter2 = ptr2) Then   ' Reverse
                    lPrettyReverse = 10000
                ElseIf (ptr1 = n0) Then   ' Pretty
                    lPrettyReverse = 50000
                End If
            ElseIf (inter2 = n0) Then
                If (ptr2 = n0) Then       ' Pretty
                    lPrettyReverse = 50000
                End If
        End If: End If

        If (lPrettyReverse) Then
            If (ptr1 > inter1) And (inter1 < lPrettyReverse) Then        ' Runners dislike super large ranges
                CopyMemByV lpStr, lpS + (idxA(lbA + ptr1 - n1) * n4), n4
                optimal = lbA + (inter1 \ n2)
                run = lbA + inter1
                Do While run > optimal                                   ' Runner do loop
                    If Not (StrComp(sA(idxA(run - n1)), item, mCriteria) = mComp) Then Exit Do
                    run = run - n1
                Loop: cast = lbA + inter1 - run
                If cast Then
                    CopyMemByV lpL_1, lpI + (run * n4), cast * n4        ' Grab items that stayed below current that should also be above items that have moved down below current
                    CopyMemByV lpI + (run * n4), lpI + ((lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move down items
                    CopyMemByV lpI + ((lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                End If
            End If '1 2 i 3 4
            If (inter2) And (ptr2 - inter2 < lPrettyReverse) Then
                base = lbA + ptr1 + n1
                CopyMemByV lpStr, lpS + (idxA(base) * n4), n4
                pvt = lbA + ptr1 + inter2
                optimal = pvt + ((ptr2 - inter2) \ n2)
                run = pvt
                Do While run < optimal                                   ' Runner do loop
                    If Not (StrComp(sA(idxA(run + n1)), item, mCriteria) = mComp) Then Exit Do
                    run = run + n1
                Loop: cast = run - pvt
                If cast Then
                    CopyMemByV lpL_1, lpI + ((pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                    CopyMemByV lpI + ((base + cast) * n4), lpI + (base * n4), inter2 * n4 ' Move up items
                    CopyMemByV lpI + (base * n4), lpL_1, cast * n4       ' Re-assign items into position immediately above current item
        End If: End If: End If

        If (ptr1 > n1) Then
            If (ptr2 > n1) Then cnt = cnt + n1: ss2Lb(cnt) = lbA + ptr1 + n1: ss2Ub(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > n1) Then
            lbA = lbA + ptr1 + n1
        Else
            If (cnt = n0) Then Exit Do
            lbA = ss2Lb(cnt): ubA = ss2Ub(cnt): cnt = cnt - n1
        End If
    Loop: CopyMemByR ByVal lpStr, 0&, n4 ' De-reference pointer to item variable
End Sub

' + SwapSort 4 Indexed Version ++++++++++++++++++++++++

' This is my indexed non-recursive swapsort - a super fast
' quicksort algorithm using variable pointers and copymemory.

' The heart of the algorithm has been completely re-written
' and bears little resemblance to the original quicksort
' algorithm and is much much faster.

' This algorithm I have dubbed the Blizzard©.

' The Blizzard algorithm is completely unfazed when re-sorting
' very large arrays that are already sorted and reverse-sorting
' of very large sorted arrays, unlike many other quicksorts.
' In fact, it is very very fast at it.

' It has the benifit of indexing which allows the source array
' to remain unchanged. This also allows the index array to be
' passed on to other sort processes to be further manipulated.

' This is my outright fastest array sorting algorithm.

Sub strSwapSort4Indexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long) '-©Rd-
    ' This is my non-recursive Quick-Sort, and is very very fast!
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim lo As Long, hi As Long, cnt As Long
    Dim item As String, lpStr As Long
    Dim idxItem As Long, lpS As Long
    hi = ubA - lbA: If hi < n1 Then Exit Sub
    mComp = SortOrder                                ' Initialize compare variable
    lo = (hi \ nKB) + n32                            ' Allow for worst case senario + some
    ValidateIdxArray idxA, lbA, ubA                  ' Initialize the index array
    InitializeStacks qs4Lb, qs4Ub, qs4Max, lo        ' Initialize pending boundary stacks
    lpStr = VarPtr(item)                             ' Cache pointer to the string variable
    lpS = VarPtr(sA(lbA)) - (lbA * n4)               ' Cache pointer to the string array
    Do: hi = ((ubA - lbA) \ n2) + lbA                ' Get pivot index position
        CopyMemByV lpStr, lpS + (idxA(hi) * n4), n4  ' Grab current value into item
        idxItem = idxA(hi): idxA(hi) = idxA(ubA)     ' Grab current index
        lo = lbA: hi = ubA                           ' Set bounds
        Do Until (hi = lo)                           ' Storm right in
            If Not (StrComp(item, sA(idxA(lo)), mCriteria) = mComp) Then ' (item <= lo)
                idxA(hi) = idxA(lo)
                hi = hi - n1
                Do Until (hi = lo)
                    If Not (StrComp(sA(idxA(hi)), item, mCriteria) = mComp) Then ' (hi <= item)
                        idxA(lo) = idxA(hi)
                        Exit Do
                    End If
                    hi = hi - n1
                Loop
                If (hi = lo) Then Exit Do ' Found swaps or out of loop
            End If
            lo = lo + n1
        Loop '-Blizzard v4i ©Rd-
        idxA(hi) = idxItem                ' Re-assign current
        If (lbA < lo - n1) Then
            If (ubA > lo + n1) Then cnt = cnt + n1: qs4Lb(cnt) = lo + n1: qs4Ub(cnt) = ubA
            ubA = lo - n1
        ElseIf (ubA > lo + n1) Then
            lbA = lo + n1
        Else
            If cnt = n0 Then Exit Do
            lbA = qs4Lb(cnt): ubA = qs4Ub(cnt): cnt = cnt - n1
        End If
    Loop: CopyMemByR ByVal lpStr, 0&, n4
End Sub

' + Inited Array +++++++++++++++++++++++++++++++++++++++++++

' This function determines if the passed array is initialized,
' and if so will return -1.

' It will also optionally indicate whether the array can be
' redimmed - in which case it will return -2.

' If the array is uninitialized (has never been redimmed or
' has been erased) it will return 0 (zero).

Function InitedArray(Arr, lbA As Long, ubA As Long, Optional ByVal bTestRedimable As Boolean) As Long
    ' Thanks to Francesco Balena who solved the Variant
    ' headache, and to Monte Hansen for the ByRef fix
    Dim tSA As SAFEARRAY, tSAB As SABOUNDS, lpSA As Long
    Dim iDataType As Integer, lOffset As Long
    On Error GoTo UnInit
    CopyMemByR iDataType, Arr, n2                   ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then       ' if a valid string array was passed
        CopyMemByR lpSA, ByVal VarPtr(Arr) + n8, n4 ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array
        If (iDataType And VT_BYREF) Then            ' see whether the function was passed a Variant that contains an array, rather than directly an array in the former case lpSA already points to the SA structure. Thanks to Monte Hansen for this fix
            CopyMemByR lpSA, ByVal lpSA, n4         ' lpSA is a discripter (pointer) to the safearray structure
        End If
        InitedArray = Not (lpSA = n0)
        If InitedArray Then
            CopyMemByR tSA.cDims, ByVal lpSA, n4
            If bTestRedimable Then ' Return -2 if redimmable
                InitedArray = InitedArray + _
                 (Not (tSA.fFeatures And FADF_FIXEDSIZE) = FADF_FIXEDSIZE)
            End If '-©Rd-
            lOffset = n16 + ((tSA.cDims - n1) * n8)
            CopyMemByR tSAB.cElements, ByVal lpSA + lOffset, n8
            tSAB.lUBound = tSAB.lLBound + tSAB.cElements - n1
            If (lbA < tSAB.lLBound) Then lbA = tSAB.lLBound
            If (ubA > tSAB.lUBound) Then ubA = tSAB.lUBound
    End If: End If
UnInit:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub InitializeStacks(LBstack() As Long, UBstack() As Long, ByRef pCurMax As Long, ByVal NewMax As Long)
    If NewMax > pCurMax Then
        ReDim LBstack(n0 To NewMax) As Long   ' Stack to hold pending lower boundries
        ReDim UBstack(n0 To NewMax) As Long   ' Stack to hold pending upper boundries
        pCurMax = NewMax
    End If
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Rd - crYptic but cRaZy!
