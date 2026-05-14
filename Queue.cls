VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Generic Queue Class."
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, MultipleDeclarations, UseMeaningfulName, ProcedureNotUsed
'@Folder("Class")
'@ModuleDescription("Generic Queue Class.")

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Generic Queue Class
' Enumeration: For Each item in object
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.10.17
'
' Methods
' Clear             Removes all items from the queue
' Enqueue item      Adds an item to the rear of the queue
' Dequeue           Removes and returns the item in front of the queue
' Items([base])     Exports all queue items to an array
'
' Properties (Get)
' Count             Returns the number of items contained in the queue
' IsEmpty           Returns True if the queue is empty or False if not
' Peek              Returns the item in front of the queue (default)
'
' This generic Queue Class is implemented as a circular encapsulated VB array.
'
' Timings (ms) for one enqueue + one dequeue.
'     count      Array   Collection
' 1      10    0.00072      0.00046
' 2     100    0.00068      0.00046
' 3    1000    0.00034      0.00046
' 4   10000    0.00035      0.00046
' 5  100000    0.00040      0.00046

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Performance Notes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Array-based queue is more memory-efficient than using a Collection.
' Circular buffer avoids shifting data on Dequeue.
' Shrink logic avoids fragmentation and excessive resizing.
' Enumeration is rarely used in queue workflows, but supported for completeness.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
End Enum

Private Const ChunkSize As Long = 100
Private Const ShrinkFactor As Long = 4 ' Shrink if usage < 1/4 capacity.

' Wrapper for private data.
Private Type TPRIVATE
    Head As Long
    Tail As Long
    Count As Long
    Capacity As Long
    CircularQueue() As Variant
    List As Collection
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    this.Capacity = ChunkSize
    ReDim this.CircularQueue(0 To this.Capacity - 1)
    this.Head = 0
    this.Tail = 0
    this.Count = 0
    Set this.List = Nothing

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Enumerator
'@Description "Enumerates queue from front to rear."
Public Property Get Enumerate() As IEnumVARIANT
' Called by For Each by setting Attribute Enumerate.VB_UserMemId = -4.
' [_NewEnum] is the name of the hidden method to enumerate a VB Collection.
' The internal queue array is first copied to an intermediate VB Collection.
'
' After a For Each loop completes, this.List still exists. VB does not provide
' a standard hook to detect when enumeration has finished. To keep it simple,
' any call to Enqueue or Dequeue will reset the enumerator.
'
' Although the current enumerator remains valid, invoking Enqueue or Dequeue
' during iteration may yield unpredictable results. This mirrors VB's native
' behavior: modifying a collection during iteration can lead to undefined or
' erratic runtime behavior.

    Set this.List = New Collection
    Dim i As Long
    For i = this.Head To this.Head + this.Count - 1
        this.List.Add this.CircularQueue(i Mod this.Capacity)
    Next
    Set Enumerate = this.List.[_NewEnum]

End Property


'@Description "Removes all items from the queue."
Public Sub Clear()

    ReDim this.CircularQueue(0 To ChunkSize - 1)
    this.Capacity = ChunkSize
    this.Head = 0
    this.Tail = 0
    this.Count = 0
    Set this.List = Nothing

End Sub


'@Description "Returns the number of items contained in the queue."
Public Property Get Count() As Long

    Count = this.Count

End Property


'@Description "Returns True if the queue is empty or False if not."
Public Property Get IsEmpty() As Boolean

    IsEmpty = (this.Count = 0)

End Property


'@DefaultMember
'@Description "Returns the item at the front of the queue."
Public Property Get Peek() As Variant

    If IsEmpty Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Peek", "Queue is empty"
    If VBA.IsObject(this.CircularQueue(this.Head)) Then
        Set Peek = this.CircularQueue(this.Head)
    Else
        Peek = this.CircularQueue(this.Head)
    End If

End Property


'@Description "Returns and removes the item at the front of the queue."
Public Function Dequeue() As Variant

    If IsEmpty Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Dequeue", "Queue is empty"
    If VBA.IsObject(this.CircularQueue(this.Head)) Then
        Set Dequeue = this.CircularQueue(this.Head)
    Else
        Dequeue = this.CircularQueue(this.Head)
    End If

    this.CircularQueue(this.Head) = Empty
    this.Head = (this.Head + 1) Mod this.Capacity
    this.Count = this.Count - 1

    If this.Capacity > ChunkSize And this.Capacity > ShrinkFactor * this.Count Then
        DecreaseCircularBuffer
    End If

    ' Reset the enumerator list.
    Set this.List = Nothing

End Function


'@Description "Adds an item to the rear of the queue."
Public Sub Enqueue(ByVal Item As Variant)

    If this.Count = this.Capacity Then
        IncreaseCircularBuffer
    End If

    If VBA.IsObject(Item) Then
        Set this.CircularQueue(this.Tail) = Item
    Else
        this.CircularQueue(this.Tail) = Item
    End If

    this.Tail = (this.Tail + 1) Mod this.Capacity
    this.Count = this.Count + 1

    ' Reset the enumerator list.
    Set this.List = Nothing

End Sub


'@Description "Exports all queue items to an array."
Public Function Items(Optional ByVal base As Long) As Variant()

    If this.Count = 0 Then
        ' Returns initialized empty array allowing For Each on object.Items.
        Items = Array()
        Exit Function
    End If

    Dim arr() As Variant: ReDim arr(base To base + this.Count - 1)
    Dim j As Long: j = base
    Dim i As Long, idx As Long
    For i = this.Head To this.Head + this.Count - 1
        idx = i Mod this.Capacity
        If VBA.IsObject(this.CircularQueue(idx)) Then
            Set arr(j) = this.CircularQueue(idx)
        Else
            arr(j) = this.CircularQueue(idx)
        End If
        j = j + 1
    Next

    Items = arr

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Decrease the circular buffer."
Private Sub DecreaseCircularBuffer()

    If this.Count = 0 Then
        Clear
        Exit Sub
    End If

    Dim NewCapacity As Long: NewCapacity = ((this.Count + ChunkSize) \ ChunkSize) * ChunkSize
    ' If this wouldn't actually shrink the capacity, skip the work.
    If NewCapacity >= this.Capacity Then Exit Sub

    Dim NewCircularQueue() As Variant: ReDim NewCircularQueue(0 To NewCapacity - 1)

    Dim i As Long, idx As Long
    For i = 0 To this.Count - 1
        idx = (this.Head + i) Mod this.Capacity
        If VBA.IsObject(this.CircularQueue(idx)) Then
            Set NewCircularQueue(i) = this.CircularQueue(idx)
        Else
            NewCircularQueue(i) = this.CircularQueue(idx)
        End If
    Next

    this.CircularQueue = NewCircularQueue
    this.Capacity = NewCapacity
    this.Head = 0
    this.Tail = this.Count

End Sub


'@Description "Increase the circular buffer."
Private Sub IncreaseCircularBuffer()

    Dim NewCapacity As Long: NewCapacity = this.Capacity + ChunkSize
    Dim NewCircularQueue() As Variant: ReDim NewCircularQueue(0 To NewCapacity - 1)

    Dim i As Long, idx As Long
    For i = 0 To this.Count - 1
        idx = (this.Head + i) Mod this.Capacity
        If VBA.IsObject(this.CircularQueue(idx)) Then
            Set NewCircularQueue(i) = this.CircularQueue(idx)
        Else
            NewCircularQueue(i) = this.CircularQueue(idx)
        End If
    Next

    this.CircularQueue = NewCircularQueue
    this.Capacity = NewCapacity
    this.Head = 0
    this.Tail = this.Count

End Sub
