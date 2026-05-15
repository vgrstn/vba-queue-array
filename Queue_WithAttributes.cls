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
' 1      10    0.00073      0.00045
' 2     100    0.00072      0.00045
' 3    1000    0.00034      0.00045
' 4   10000    0.00035      0.00045
' 5  100000    0.00040      0.00045

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

' InitialSize is the initial allocation size of the circular buffer.
Private Const InitialSize As Long = 100

' GrowthFactor is the multiplier applied on both grow and shrink.
Private Const GrowthFactor As Double = 2

' ShrinkTrigger is the occupancy fraction below which a shrink fires.
' Must be < 1 / GrowthFactor to prevent thrashing on alternating enqueue/dequeue.
Private Const ShrinkTrigger As Double = 0.5 / GrowthFactor

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

    ReDim this.CircularQueue(0 To InitialSize - 1)
    this.Capacity = InitialSize

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Enumerate() As IEnumVARIANT
Attribute Enumerate.VB_UserMemId = -4
Attribute Enumerate.VB_Description = "Enumerates queue from front to rear."
' Called by For Each via Attribute Enumerate.VB_UserMemId = -4.
' [_NewEnum] is the name of the hidden method to enumerate a VB Collection.
' The internal queue array is first copied to an intermediate VB Collection.
'
' A snapshot Collection is built on each call and stored in this.List.
' Enqueue or Dequeue sets this.List = Nothing, but the enumerator holds its own
' reference, so the snapshot survives until the For Each loop completes.
' Mutation during iteration is therefore safe but iterates the old snapshot.

    Set this.List = New Collection
    Dim i As Long
    For i = this.Head To this.Head + this.Count - 1
        this.List.Add this.CircularQueue(i Mod this.Capacity)
    Next
    Set Enumerate = this.List.[_NewEnum]

End Property


Public Sub Clear()
Attribute Clear.VB_Description = "Removes all items from the queue."

    ReDim this.CircularQueue(0 To InitialSize - 1)
    this.Capacity = InitialSize
    this.Head = 0
    this.Tail = 0
    this.Count = 0
    Set this.List = Nothing

End Sub


Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items contained in the queue."

    Count = this.Count

End Property


Public Property Get IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Returns True if the queue is empty or False if not."

    IsEmpty = (this.Count = 0)

End Property


Public Property Get Peek() As Variant
Attribute Peek.VB_UserMemId = 0
Attribute Peek.VB_Description = "Returns the item at the front of the queue."

    If IsEmpty Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Peek", "Queue is empty"
    If VBA.IsObject(this.CircularQueue(this.Head)) Then
        Set Peek = this.CircularQueue(this.Head)
    Else
        Peek = this.CircularQueue(this.Head)
    End If

End Property


Public Function Dequeue() As Variant
Attribute Dequeue.VB_Description = "Returns and removes the item at the front of the queue."

    If IsEmpty Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Dequeue", "Queue is empty"
    If VBA.IsObject(this.CircularQueue(this.Head)) Then
        Set Dequeue = this.CircularQueue(this.Head)
    Else
        Dequeue = this.CircularQueue(this.Head)
    End If

    this.CircularQueue(this.Head) = Empty
    this.Head = (this.Head + 1) Mod this.Capacity
    this.Count = this.Count - 1

    If this.Count < ShrinkTrigger * this.Capacity Then
        DecreaseCircularBuffer
    End If

    Set this.List = Nothing

End Function


Public Sub Enqueue(ByVal Item As Variant)
Attribute Enqueue.VB_Description = "Adds an item to the rear of the queue."

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

    Set this.List = Nothing

End Sub


Public Function Items(Optional ByVal base As Long) As Variant()
Attribute Items.VB_Description = "Exports all queue items to an array."

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

Private Sub DecreaseCircularBuffer()
Attribute DecreaseCircularBuffer.VB_Description = "Decrease the circular buffer."

    If this.Count = 0 Then Clear: Exit Sub

    Dim NewCapacity As Long: NewCapacity = Int(this.Capacity / GrowthFactor)
    If NewCapacity < InitialSize Then NewCapacity = InitialSize
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


Private Sub IncreaseCircularBuffer()
Attribute IncreaseCircularBuffer.VB_Description = "Increase the circular buffer."

    Dim NewCapacity As Long: NewCapacity = Int(GrowthFactor * this.Capacity)
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
