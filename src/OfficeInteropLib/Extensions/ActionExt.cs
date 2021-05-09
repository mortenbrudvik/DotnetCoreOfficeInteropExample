using System;
using LanguageExt;

namespace OfficeInteropLib.Extensions
{
    public static class ActionExt
    {
        public static Func<Unit> ToFunc(this Action action)
            => () => { action(); return new Unit(); };

        public static Func<T, Unit> ToFunc<T>(this Action<T> action)
            => t => { action(t); return new Unit(); };
    }
}