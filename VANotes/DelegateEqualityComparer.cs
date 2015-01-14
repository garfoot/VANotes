using System;
using System.Collections.Generic;

namespace VANotes
{
    public class DelegateEqualityComparer<T> : IEqualityComparer<T>
    {
        private readonly Func<T,T, bool> _equals;
        private readonly Func<T, int> _getHashCode;

        public DelegateEqualityComparer(Func<T, T, bool> equals, Func<T, int> getHashCode)
        {
            _equals = equals;
            _getHashCode = getHashCode;
        }

        public bool Equals(T x, T y)
        {
            return _equals(x, y);
        }

        public int GetHashCode(T obj)
        {
            return _getHashCode(obj);
        }
    }
}