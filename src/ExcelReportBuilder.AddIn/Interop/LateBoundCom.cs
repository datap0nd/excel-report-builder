using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelReportBuilder.AddIn.Interop
{
    internal static class LateBoundCom
    {
        public static object? GetProperty(object target, string propertyName)
        {
            if (target == null)
            {
                throw new ArgumentNullException(nameof(target));
            }

            return target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                binder: null,
                target,
                args: null,
                CultureInfo.InvariantCulture);
        }

        public static bool TryGetProperty(object? target, string propertyName, out object? value)
        {
            value = null;
            if (target == null)
            {
                return false;
            }

            try
            {
                value = GetProperty(target, propertyName);
                return true;
            }
            catch (Exception exception) when (IsDispatchFailure(exception))
            {
                return false;
            }
        }

        public static void SetProperty(object target, string propertyName, object? value)
        {
            if (target == null)
            {
                throw new ArgumentNullException(nameof(target));
            }

            target.GetType().InvokeMember(
                propertyName,
                BindingFlags.SetProperty,
                binder: null,
                target,
                new[] { value },
                CultureInfo.InvariantCulture);
        }

        public static bool TrySetProperty(object? target, string propertyName, object? value)
        {
            if (target == null)
            {
                return false;
            }

            try
            {
                SetProperty(target, propertyName, value);
                return true;
            }
            catch (Exception exception) when (IsDispatchFailure(exception))
            {
                return false;
            }
        }

        public static object? Invoke(object target, string methodName, params object?[] arguments)
        {
            if (target == null)
            {
                throw new ArgumentNullException(nameof(target));
            }

            return target.GetType().InvokeMember(
                methodName,
                BindingFlags.InvokeMethod,
                binder: null,
                target,
                arguments,
                CultureInfo.InvariantCulture);
        }

        public static bool TryInvoke(object? target, string methodName, params object?[] arguments)
        {
            if (target == null)
            {
                return false;
            }

            try
            {
                Invoke(target, methodName, arguments);
                return true;
            }
            catch (Exception exception) when (IsDispatchFailure(exception))
            {
                return false;
            }
        }

        public static void FinalRelease(object? comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch (InvalidComObjectException)
            {
                // A peer may already have released this RCW during Excel shutdown.
            }
        }

        private static bool IsDispatchFailure(Exception exception)
        {
            return exception is COMException
                || exception is TargetInvocationException
                || exception is MissingMemberException
                || exception is ArgumentException
                || exception is InvalidOperationException;
        }
    }
}
