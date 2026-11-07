# vba-queue-array
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

Generic Queue Class based on a circular encapsulated VB Array

A lightweight, **circular array-backed FIFO queue** for VBA with:
- O(1) enqueue at the **rear** and dequeue from the **front** via circular buffer
- Safe **enumeration** (`For Each`) via a proper COM enumerator
- Clear error semantics (`Peek` / `Dequeue` on empty queue raise error 5)
- Zero dependencies (pure VBA)

---

## 📁 Files

| File | Description |
|---|---|
| `Queue.cls` | Source file with [Rubberduck](https://rubberduckvba.com/) annotations (`'@Description`, `'@DefaultMember`, `'@Enumerator`) |
| `Queue_WithAttributes.cls` | Ready-to-import version with VB attributes baked in — no Rubberduck required |

Both files are identical in behaviour. Import `Queue_WithAttributes.cls` if you are not using Rubberduck.

---

## 📦 Features

- **Circular buffer** avoids data shifting on every `Dequeue`
- **Dynamic resizing** doubles on grow (`GrowthFactor = 2`, in place when the buffer starts at position 0); shrinks by half when occupancy drops below `ShrinkTrigger` (25%), never below `InitialSize`
- **Enumeration**: `For Each item In Queue` (via hidden `[_NewEnum]`)
- **Utility export**: `Items([base])` returns a 0- or 1-based array copy
- Pure VBA, no external references, Rubberduck-friendly annotations

---

## ⚙️ Public Interface

| Member               | Type       | Description |
|---------------------|------------|-------------|
| `Enqueue(Item)`      | `Sub`      | Adds an item to the **rear** of the queue. |
| `Dequeue()`          | `Function` | Returns **and removes** the front item. Raises error 5 if empty. |
| `Peek` *(Default)*   | `Property` | Returns the front item **without** removing it. Raises error 5 if empty. |
| `Count`              | `Property` | Number of items. |
| `IsEmpty`            | `Property` | `True` if empty, else `False`. |
| `Clear`              | `Sub`      | Removes all items. |
| `Items([base])`      | `Function` | Returns all items as a `Variant()` array; `base` default = `0`; an empty queue returns an empty array (`0 To -1`). |
| `For Each`           | Enumerator | Iterates **front → rear** on a snapshot. (Mutation during iteration is safe but iterates the old snapshot.) |

**Error behavior**  
- Empty queue on `Peek` / `Dequeue` raises **`vbErrorInvalidProcedureCall (=5)`** with source `"Queue.Peek"` or `"Queue.Dequeue"`.

---

## 🚀 Quick Start

```vb
Dim q As New Queue

q.Enqueue "alpha"
q.Enqueue "beta"
Debug.Print q.Peek        ' -> alpha  (front)
Debug.Print q.Dequeue     ' -> alpha  (removed)
Debug.Print q.Dequeue     ' -> beta   (removed)
Debug.Print q.IsEmpty     ' -> True
```

---

## ⏱️ Performance

Timings (ms) for one `Enqueue` + one `Dequeue`, measured on Windows x64:

| # | Count   | vba-queue-array | vba-queue (Collection) |
|---|---------|-----------------|------------------------|
| 1 | 10      | 0.00034         | 0.00053                |
| 2 | 100     | 0.00036         | 0.00051                |
| 3 | 1,000   | 0.00035         | 0.00052                |
| 4 | 10,000  | 0.00036         | 0.00053                |
| 5 | 100,000 | 0.00037         | 0.00056                |

The array-based version is more memory-efficient and needs about two thirds of the time of the Collection-based [vba-queue](https://github.com/vgrstn/VBA-Queue) at every size. The circular buffer avoids data shifting on every `Dequeue`. `For Each` first copies all items to a snapshot `Collection`, which makes it about four times slower than in the Collection-based version.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
