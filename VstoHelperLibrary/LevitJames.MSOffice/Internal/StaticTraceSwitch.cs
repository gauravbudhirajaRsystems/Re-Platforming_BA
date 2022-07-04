// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace LevitJames.MSOffice
{
    internal static class StaticTraceSwitch
    {
        private static TraceSwitch _traceSwitch;


        [SuppressMessage("Microsoft.Performance", "CA1810:InitializeReferenceTypeStaticFieldsInline")]
        static StaticTraceSwitch()
        {
            //Create here so we do not have to keep checking the 
            _traceSwitch = new TraceSwitch("Word Extensions", "Word Extensions default trace switch");
        }


        public static TraceSwitch TraceSwitch
        {
            get => _traceSwitch;
            set
            {
                lock (_traceSwitch)
                {
                    if (value == null)
                    {
                        throw new ArgumentNullException(nameof(value));
                    }

                    _traceSwitch = value;
                }
            }
        }


        public static bool TraceError => _traceSwitch.TraceError;


        public static bool TraceWarning => _traceSwitch.TraceWarning;


        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public static bool TraceInfo => _traceSwitch.TraceInfo;


        public static bool TraceVerbose => _traceSwitch.TraceVerbose;
    }
}