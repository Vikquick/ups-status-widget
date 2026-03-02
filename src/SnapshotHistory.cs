#nullable disable
using System;
using System.Collections.Generic;

namespace UpsStatusWidget;

sealed class SnapshotHistory
{
    readonly object _lock = new();
    readonly int _capacity;
    readonly Queue<UpsSnapshot> _queue;

    public SnapshotHistory(int capacity)
    {
        _capacity = Math.Max(1, capacity);
        _queue = new Queue<UpsSnapshot>(_capacity);
    }

    public void Add(UpsSnapshot snapshot)
    {
        lock (_lock) {
            if (_queue.Count >= _capacity) _queue.Dequeue();
            _queue.Enqueue(snapshot);
        }
    }

    public UpsSnapshot[] GetRecent(int limit)
    {
        if (limit <= 0) return Array.Empty<UpsSnapshot>();

        lock (_lock) {
            int take = Math.Min(limit, _queue.Count);
            if (take <= 0) return Array.Empty<UpsSnapshot>();

            var arr = _queue.ToArray();
            var result = new UpsSnapshot[take];
            Array.Copy(arr, arr.Length - take, result, 0, take);
            return result;
        }
    }
}

