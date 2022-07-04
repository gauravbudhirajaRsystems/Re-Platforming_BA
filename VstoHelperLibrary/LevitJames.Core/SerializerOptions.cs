// © Copyright 2018 Levit & James, Inc.

using System;

namespace LevitJames.Core
{
    /// <summary>
    ///     Serialization Options used by the Serializer class
    /// </summary>
    [Flags]
    public enum SerializerOptions
    {
        /// <summary>
        ///     No options. Include both the Version number and the Public Key token when Serializing types.
        /// </summary>
        None = 0,

        /// <summary>
        ///     Don't include the version number of the Assembly when serializing a Type.
        /// </summary>
        NoVersion = 0x1,

        /// <summary>
        ///     Don't include the PublicKey of the Assembly when serializing a Type.
        /// </summary>
        NoPublicKey = 0x2
    }
}