using System;
using System.Collections;
using System.Collections.Generic;

namespace DotNetCoreTypeLibGenerator.Helpers
{
    /// <summary>
    /// Encapsulates the disposals of a group of disposable types; disposing the list will dispose the elements
    /// contained within. 
    /// </summary>
    /// <typeparam name="T">Any types that implements <see cref="IDisposable"/></typeparam>
    public class DisposableList<T> : IList<T>, IDisposable
        where T : IDisposable
    {
        private bool _disposed;
        private List<T> _list = new List<T>();

        public int Count => _list.Count;

        public bool IsReadOnly => false;

        public T this[int index] { get => _list[index]; set => _list[index] = value; }

        public int IndexOf(T item)
        {
            return _list.IndexOf(item);
        }

        public void Insert(int index, T item)
        {
            _list.Insert(index, item);
        }

        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }

        public void Add(T item)
        {
            _list.Add(item);
        }

        public void Clear()
        {
            _list.Clear();
        }

        public bool Contains(T item)
        {
            return _list.Contains(item);
        }

        public void CopyTo(T[] array, int arrayIndex)
        {
            _list.CopyTo(array, arrayIndex);
        }

        public bool Remove(T item)
        {
            return _list.Remove(item);
        }

        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    foreach (IDisposable item in _list)
                    {
                        item.Dispose();
                    }
                }

                _list = null;
                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
