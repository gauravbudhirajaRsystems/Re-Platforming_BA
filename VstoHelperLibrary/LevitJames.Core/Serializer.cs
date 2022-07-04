// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    ///     A helper class for serializing and de-serializing classes
    /// </summary>
    /// <remarks>
    ///     Note: The ToXml\FromXml methods currently use the SoapFormatter, which has now been marked as obsolete by
    ///     Microsoft. The internal implementation of these methods will be changed in due course.
    /// </remarks>
    public static class Serializer
    {
        /// <summary>
        ///     The string name to use when adding a Version number during serializing and de-serializing.
        /// </summary>
        public static string SerializeVersionName { get; set; } = ".Version";


        /// <summary>
        ///     The Serialization options to use when Serializing an object.
        /// </summary>
        public static SerializerOptions Options { get; set; }


        /// <summary>
        ///     Serializes an object to a MemoryStream using Binary Serialization.
        /// </summary>
        /// <param name="value">The object to serialize</param>
        /// <param name="context">Context data to pass to the object being serialized.</param>
        public static Stream ToBinary(object value, object context = null)
        {
            var stream = new MemoryStream();
            try
            {
                ToBinary(value, stream, context);
                return stream;
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>
        ///     Serializes an object to a MemoryStream using Binary Serialization.
        /// </summary>
        /// <param name="value">The object to serialize</param>
        /// <param name="formatter">The formatter used to serialize the object.</param>
        public static Stream ToBinary(object value, IFormatter formatter)
        {
            var stream = new MemoryStream();
            try
            {
                ToBinary(value, stream, formatter);
                return stream;
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>
        ///     Serializes an object to the supplied stream using Binary Serialization.
        /// </summary>
        /// <param name="value">The object to serialize.</param>
        /// <param name="stream">The stream to serialize the object to.</param>
        /// <param name="context">Context data to pass to the object being serialized.</param>
        public static void ToBinary([NotNull] object value, Stream stream, object context = null)
        {
            Check.NotNull(stream, nameof(stream));

            var bf = new BinaryFormatter();
            bf.Context = new StreamingContext(bf.Context.State, context);

            SerializeClass(value, stream, bf);
        }

        /// <summary>
        ///     Serializes an object to the supplied stream using Binary Serialization.
        /// </summary>
        /// <param name="value">The object to serialize.</param>
        /// <param name="stream">The stream to serialize the object to.</param>
        /// <param name="formatter">Context data to pass to the object being serialized.</param>
        private static void ToBinary([NotNull] object value, Stream stream, IFormatter formatter)
        {
            Check.NotNull(stream, nameof(stream));

            SerializeClass(value, stream, formatter);
        }


        /// <summary>
        ///     Serializes an object to a Base64 string.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="context">Context data to pass to the object being serialized.</param>
        public static string ToBase64(object value, object context = null)
        {
            using (var stream = (MemoryStream) ToBinary(value, context))
            {
                stream.Position = 0;
                return Convert.ToBase64String(stream.ToArray());
            }
        }

        /// <summary>
        ///     Serializes an object to a Base64 string.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="formatter">IFormatter.</param>
        public static string ToBase64(object value, IFormatter formatter)
        {
            using (var stream = (MemoryStream) ToBinary(value, formatter))
            {
                stream.Position = 0;
                return Convert.ToBase64String(stream.ToArray());
            }
        }

        /// <summary>
        ///     Serializes an object to a Base64 string.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="stream">The stream to serialize to</param>
        /// <param name="context">Context data to pass to the object being serialized.</param>
        public static void ToBase64(object value, Stream stream, object context = null)
        {
            var b64 = ToBase64(value, context);
            var sw = new StreamWriter(stream);
            sw.Write(b64);
            sw.Flush();
        }


        /// <summary>
        ///     De-serializes the class of type TClass from the provided stream using the binary formatter.
        /// </summary>
        /// <typeparam name="TClass"></typeparam>
        /// <param name="stream">The stream containing the serialized class</param>
        public static TClass FromBinary<TClass>(Stream stream) where TClass : class
        {
            return FromBinary<TClass>(stream, new BinaryFormatter());
        }

        /// <summary>
        ///     De-serializes the class of type TClass from the provided stream using the binary formatter.
        /// </summary>
        /// <typeparam name="TClass"></typeparam>
        /// <param name="stream">The stream containing the serialized class</param>
        /// <param name="context">Additional context used when deserializing the class</param>
        public static TClass FromBinary<TClass>(Stream stream, object context) where TClass : class
        {
            return FromBinary<TClass>(stream,
                                      new BinaryFormatter(selector: null,
                                                          context:
                                                          new StreamingContext(StreamingContextStates.All, context)));
        }

        /// <summary>
        ///     De-serializes the class of type TClass from the provided stream using the binary formatter.
        /// </summary>
        /// <param name="stream">The stream containing the serialized class</param>
        /// <param name="context">Additional context used when deserializing the class</param>
        public static object FromBinary([NotNull] Stream stream, object context = null)
        {
            Check.NotNull(stream, "stream");

            var formatter = new BinaryFormatter(selector: null,
                                                context: new StreamingContext(StreamingContextStates.All, context));

            return formatter.Deserialize(stream);
        }

        /// <summary>
        ///     De-serializes the class of type TClass from the provided stream using the binary formatter.
        /// </summary>
        /// <param name="stream">The stream containing the serialized class</param>
        /// <param name="formatter">Additional formatter used when deserializing the class</param>
        public static TClass FromBinary<TClass>([NotNull] Stream stream, [NotNull] IFormatter formatter) where TClass : class
        {
            Check.NotNull(stream, "stream");
            Check.NotNull(formatter, "formatter");

            return (TClass) formatter.Deserialize(stream);
        }


        /// <summary>
        ///     De-serializes a class previously serialized using the ToBase64 method.
        /// </summary>
        /// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        /// <param name="data">The serialized data represented as a string.</param>
        /// <returns>
        ///     A deserialize instance of the class.
        ///     The class is de-serialized using the
        ///     <see cref="System.Runtime.Serialization.Formatters.Binary.BinaryFormatter">BinaryFormatter</see>.
        ///     If the data string is null or empty then the class is created passing the context in the formatter if possible,
        ///     else the default constructor is called.
        /// </returns>
        public static TClass FromBase64<TClass>(string data) where TClass : class
        {
            return FromBase64<TClass>(data, new BinaryFormatter());
        }


        /// <summary>
        ///     De-serializes a class previously serialized using the ToBase64 method.
        /// </summary>
        /// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        /// <param name="data">The serialized data represented as a string.</param>
        /// <param name="context">An object used to store custom data to be passed into the constructor.</param>
        /// <returns>
        ///     A deserialize instance of the class.
        ///     The class is de-serialized using the
        ///     <see cref="System.Runtime.Serialization.Formatters.Binary.BinaryFormatter">BinaryFormatter</see>.
        ///     When the class is If the data string is null or empty the class is created, passing through the context into the
        ///     constructor. If no appropriate constructor is available then
        ///     the class is created without passing in the context.
        /// </returns>
        public static TClass FromBase64<TClass>(string data, object context) where TClass : class
        {
            return FromBase64<TClass>(data,
                                      new BinaryFormatter(selector: null,
                                                          context:
                                                          new StreamingContext(StreamingContextStates.All, context)));
        }

        /// <summary>
        ///     De-serializes a class previously serialized using the ToBase64 method.
        /// </summary>
        /// <param name="data">The serialized data represented as a string.</param>
        /// <returns>
        ///     A deserialize instance of the class.
        ///     The class is de-serialized using the
        ///     <see cref="System.Runtime.Serialization.Formatters.Binary.BinaryFormatter">BinaryFormatter</see>.
        ///     If the data string is null or empty then the class is created passing the context in the formatter if possible,
        ///     else the default constructor is called.
        /// </returns>
        public static object FromBase64(string data)
        {
            return FromBase64(data, new BinaryFormatter());
        }

        /// <summary>
        ///     De-serializes a class previously serialized using the ToBase64 method.
        /// </summary>
        /// <param name="data">The serialized data represented as a string.</param>
        /// <param name="context">An object used to store custom data to be passed into the constructor.</param>
        /// <returns>
        ///     A deserialize instance of the class.
        ///     The class is de-serialized using the
        ///     <see cref="System.Runtime.Serialization.Formatters.Binary.BinaryFormatter">BinaryFormatter</see>.
        ///     If the data string is null or empty then the class is created passing the context in the formatter if possible,
        ///     else the default constructor is called.
        /// </returns>
        public static object FromBase64(string data, object context)
        {
            return FromBase64(data, new BinaryFormatter(selector: null,
                                                        context:
                                                        new StreamingContext(StreamingContextStates.All, context)));
        }

        /// <summary>
        ///     De-serializes a class previously serialized using the ToBase64 method.
        /// </summary>
        /// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        /// <param name="data">The serialized data represented as a string.</param>
        /// <param name="formatter">The formatter to use to deserialize the class.</param>
        /// <returns>
        ///     A de-serialized instance of the class.
        ///     If the data string is null or empty then the class is created passing the context in the formatter if possible,
        ///     else the default constructor is called.
        /// </returns>
        public static TClass FromBase64<TClass>(string data, [NotNull] IFormatter formatter) where TClass : class
        {
            Check.NotNull(formatter, nameof(formatter));

            if (string.IsNullOrWhiteSpace(data))
            {
                return (TClass) CreateDefaultClass(typeof(TClass), formatter.Context.Context);
            }

            using (var stream = new MemoryStream(Convert.FromBase64String(data)))
            {
                stream.Position = 0;
                //Deserialize and return the object.
                var obj = formatter.Deserialize(stream);
                return (TClass) obj;
            }
        }

        /// <summary>
        ///     De-serializes a class previously serialized using the ToBase64 method.
        /// </summary>
        /// <param name="data">The serialized data represented as a string.</param>
        /// <param name="formatter">The formatter to use to deserialize the class.</param>
        /// <returns>
        ///     A de-serialized instance of the class.
        ///     If the data string is null or empty then the class is created passing the context in the formatter if possible,
        ///     else the default constructor is called.
        /// </returns>
        public static object FromBase64(string data, [NotNull] IFormatter formatter)
        {
            if (string.IsNullOrEmpty(data))
                return null;

            Check.NotNull(formatter, nameof(formatter));

            using (var stream = new MemoryStream(Convert.FromBase64String(data)))
            {
                stream.Position = 0;
                //Deserialize and return the object.
                return formatter.Deserialize(stream);
            }
        }


        ///// <summary>
        ///// De-serializes a class previously serialized using the ToXml method.
        ///// </summary>
        ///// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        ///// <param name="data">The serialized data represented as an xml string.</param>
        ///// <returns>A deserialize instance of the class.
        ///// <remarks>The class is de-serialized using the <see cref="System.Runtime.Serialization.Formatters.Soap.SoapFormatter">SoapFormatter</see>.</remarks>
        ///// If the data string is null or empty then the class is created passing the context in the formatter if possible, else the default constructor is called.</returns>
        //public static TClass FromXml<TClass>(string data) where TClass : class
        //{
        //    return FromXml<TClass>(data, new System.Runtime.Serialization.Formatters.Soap.SoapFormatter());
        //}

        ///// <summary>
        ///// De-serializes a class previously serialized using the ToXml method.
        ///// </summary>
        ///// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        ///// <param name="data">The serialized data represented as an xml string.</param>
        ///// <param name="context">An object used to store custom data to be passed into the constructor.</param>
        ///// <returns>A deserialize instance of the class.
        ///// <remarks>The class is de-serialized using the <see cref="System.Runtime.Serialization.Formatters.Soap.SoapFormatter">SoapFormatter</see>.</remarks>
        ///// If the data string is null or empty then the class is created passing the context in the formatter if possible, else the default constructor is called.</returns>
        //public static TClass FromXml<TClass>(string data, object context) where TClass : class
        //{
        //    return FromXml<TClass>(data, new System.Runtime.Serialization.Formatters.Soap.SoapFormatter(null, new StreamingContext(StreamingContextStates.All, context)));
        //}

        ///// <summary>
        ///// De-serializes a class previously serialized using the ToXml method.
        ///// </summary>
        ///// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        ///// <param name="data">The serialized data represented as an xml string.</param>
        ///// <param name="formatter"></param>
        ///// <returns>A deserialize instance of the class.
        ///// <remarks>The class is de-serialized using the <see cref="System.Runtime.Serialization.Formatters.Soap.SoapFormatter">SoapFormatter</see>.</remarks>
        ///// If the data string is null or empty then the class is created passing the context in the formatter if possible, else the default constructor is called.</returns>
        //public static TClass FromXml<TClass>(string data, IFormatter formatter) where TClass : class
        //{
        //    if (formatter == null)
        //        throw new ArgumentNullException(nameof(formatter));

        //    if (string.IsNullOrWhiteSpace(data))
        //    {
        //        return CreateDefaultClass<TClass>(formatter.Context.Context);
        //    }
        //    var ms = new MemoryStream();
        //    try
        //    {
        //        using (var sw = new StreamWriter(ms))
        //        {
        //            ms = null;
        //            sw.Write(data);
        //            sw.Flush();
        //            sw.BaseStream.Position = 0;
        //            return ((TClass)formatter.Deserialize(sw.BaseStream));
        //        }
        //    }
        //    catch
        //    {
        //        ms?.Dispose();
        //        throw;
        //    }

        //}

        ///// <summary>
        ///// De-serializes a class previously serialized using the ToXml method.
        ///// </summary>
        ///// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        ///// <param name="stream">The serialized stream.</param>
        ///// <returns>A de-serialized instance of the class.</returns> 
        ///// <remarks>The class is de-serialized using the <see cref="System.Runtime.Serialization.Formatters.Soap.SoapFormatter">SoapFormatter</see>.</remarks>
        //public static TClass FromXml<TClass>(Stream stream) where TClass : class
        //{
        //    return FromBinary<TClass>(stream, new System.Runtime.Serialization.Formatters.Soap.SoapFormatter());
        //}

        ///// <summary>
        ///// De-serializes a class previously serialized using the ToXml method.
        ///// </summary>
        ///// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        ///// <param name="stream">The serialized stream.</param>
        ///// <param name="context">An object used to store custom data to be passed into the constructor.</param>
        ///// <returns>A de-serialized instance of the class.</returns>
        ///// <remarks>The class is de-serialized using the <see cref="System.Runtime.Serialization.Formatters.Soap.SoapFormatter">SoapFormatter</see>.</remarks>
        //public static TClass FromXml<TClass>(Stream stream, object context) where TClass : class
        //{
        //    return FromBinary<TClass>(stream, new System.Runtime.Serialization.Formatters.Soap.SoapFormatter(null, new StreamingContext(StreamingContextStates.All, context)));
        //}

        ///// <summary>
        ///// De-serializes a class previously serialized using the ToXml method.
        ///// </summary>
        ///// <typeparam name="TClass">The class to deserialize and return.</typeparam>
        ///// <param name="stream">The serialized stream.</param>
        ///// <param name="formatter">The formatter to use to deserialize the class.</param>
        ///// <returns>A de-serialized instance of the class.</returns>
        //
        //public static TClass FromXml<TClass>(Stream stream, IFormatter formatter) where TClass : class
        //{

        //    if (stream == null)
        //    {
        //        throw new ArgumentNullException(nameof(stream));
        //    }
        //    if (formatter == null)
        //    {
        //        throw new ArgumentNullException(nameof(formatter));
        //    }

        //    return (TClass)formatter.Deserialize(stream);

        //}


        /// <summary>
        ///     Serializes a collection of items.
        /// </summary>
        /// <param name="info">The SerializationInfo containing the serialized items.</param>
        /// <param name="items">The collection of items to serialize.</param>
        public static void SerializeCollection([NotNull] SerializationInfo info, [NotNull] IEnumerable items)
        {
            Check.NotNull(info, nameof(info));
            Check.NotNull(items, nameof(items));

            var counter = 0;
            foreach (var itm in items)
            {
                info.AddValue("Item" + counter.ToString(CultureInfo.InvariantCulture), itm);
                counter++;
            }

            info.AddValue("Count", counter);
        }


        /// <summary>
        ///     De-serializes a collection of items.
        /// </summary>
        /// <param name="info">The SerializationInfo containing the serialized items.</param>
        /// <param name="items">The collection of items to serialize.</param>
        public static void DeserializeCollection(SerializationInfo info, [NotNull] IList items)
        {
            Check.NotNull(items, nameof(items));

            DeserializeCollection(info, (object item) => items.Add(item));
        }

        /// <summary>
        ///     De-serializes a collection of TItem into the supplied list.
        /// </summary>
        /// <typeparam name="TItem">The Type of items to de-serialize</typeparam>
        /// <param name="info">A SerializationInfo containing the serialized items.</param>
        /// <param name="items">the list to add the items to.</param>
        public static void DeserializeCollection<TItem>(SerializationInfo info, ICollection<TItem> items)
        {
            Check.NotNull(items, nameof(items));

            DeserializeCollection<TItem>(info, items.Add);
        }

        /// <summary>
        ///     De-serializes a collection of items
        /// </summary>
        /// <typeparam name="TAction">The action to perform as each item is de-serialized.</typeparam>
        /// <param name="info">The SerializationInfo containing the serialized items</param>
        /// <param name="action">
        ///     The action to perform after the item is deserialize, such as adding the item to a collection, or
        ///     dictionary
        /// </param>
        public static void DeserializeCollection<TAction>([NotNull] SerializationInfo info, [NotNull] Action<TAction> action)
        {
            Check.NotNull(info, nameof(info));
            Check.NotNull(action, nameof(action));

            var count = info.GetInt32("Count");
            for (var i = 0; i < count; i++)
            {
                action.Invoke((TAction) info.GetValue("Item" + i.ToString(CultureInfo.InvariantCulture), typeof(TAction)));
            }
        }


        /// <summary>
        ///     Adds a Version value to the <see cref="SerializationInfo">SerializationInfo</see> class.
        /// </summary>
        /// <param name="info">A valid <see cref="SerializationInfo">SerializationInfo</see> instance to add the version number to</param>
        /// <instance></instance>
        /// <remarks>
        ///     This method adds the Calling Assembly Version string to the
        ///     <see cref="SerializationInfo">SerializationInfo</see> using the SerializeVersionName constant and it sets the
        ///     <see cref="SerializationInfo.AssemblyName">SerializationInfo.AssemblyName</see> to a version independent name.
        /// </remarks>
        public static void SetVersionInfo([NotNull] SerializationInfo info)
        {
            Check.NotNull(info, nameof(info));
            SetVersionInfo(info, new AssemblyName(info.AssemblyName).Version);
        }

        /// <summary>
        ///     Adds the Version information to the supplied SerializationInfo.
        ///     If the SerializerOptions.NoVersion is set then the Version information is not added.
        ///     It also sets the SerializationInfo.AssemblyName member according to the SeializationOption.NoPublicKey.
        /// </summary>
        /// <param name="info"></param>
        /// <param name="version"></param>
        public static void SetVersionInfo([NotNull] SerializationInfo info, [NotNull] Version version)
        {
            Check.NotNull(info, nameof(info));
            Check.NotNull(version, nameof(version));

            //var assemblyName = Assembly.GetCallingAssembly().GetName();
            //info.AssemblyName = GetAssemblyName(assemblyName);

            if ((Options & SerializerOptions.NoVersion) == SerializerOptions.None)
            {
                info.AddValue(SerializeVersionName, version.ToString());
            }
        }


        /// <summary>
        ///     Reads the Version number previously set using the <see cref="SetVersionInfo(SerializationInfo)">AddVersion</see>
        ///     method.
        /// </summary>
        /// <param name="info">A valid <see cref="SerializationInfo">SerializationInfo</see> instance to add the version number to</param>
        /// <returns>An integer value representing the version of the serialized data.</returns>
        /// <remarks>
        ///     The name used to retrieve the version number is taken from the
        ///     <see cref="SerializeVersionName">SerializeVersionName</see> member.
        /// </remarks>
        public static string ReadVersion([NotNull] SerializationInfo info)
        {
            Check.NotNull(info, nameof(info));
            //var value = info.GetValue(SerializeVersionName);

            foreach (var item in info)
            {
                if (item.Name == SerializeVersionName)
                    return item.Value?.ToString();
            }

            return null;
        }

        /// <summary>
        ///     Reads the Serialized Version information from the SerializationInfo. If the Version information does not exist then
        ///     the provided defaultVersion string can be used.
        ///     The Version information should have been serialized using the <see cref="SetVersionInfo(SerializationInfo)" /> call
        /// </summary>
        /// <param name="info">The SerializationInfo containing the Version.</param>
        /// <param name="defaultVersion">
        ///     The default version to use if the Version information does not exist. This string should
        ///     be in able to be parsed by the Version class
        /// </param>
        /// <returns>A new Version instance</returns>
        public static Version ReadVersion([NotNull] SerializationInfo info, [NotNull] string defaultVersion)
        {
            var versionString = ReadVersion(info);
            if (string.IsNullOrEmpty(versionString))
            {
                versionString = defaultVersion;
            }

            if (versionString.Contains(".") == false)
            {
                versionString += ".0";
            }

            Version version;

            if (Version.TryParse(versionString, out version) == false)
            {
                if (defaultVersion != versionString)
                {
                    if (Version.TryParse(defaultVersion, out version) == false)
                    {
                        throw new LJException("Invalid Version Information");
                    }
                }
            }

            return version;
        }


        // private members


        private static object CreateDefaultClass(Type classType, object context)
        {
            //Try to create the object 
            //Note: System.Activator.CreateInstance only works with public constructors so we use reflection instead

            //If we have a formatter.Context.Context object then try and find a constructor that can take that object
            if (context != null)
            {
                var ctor = classType.GetConstructor(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
                                                    binder: null, types: new[] {context.GetType()}, modifiers: null);
                if (ctor != null)
                {
                    return ctor.Invoke(new[] {context});
                    //Else
                    //Debug.Assert(False, "Error could not get constructor info")
                }
            }

            //The formatter.Context.Context was null or no valid constructors exist that can take the context object.
            // so try and create the object as normal.
            //ctor = classType.GetConstructor(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance,
            //                                     binder: null, types: Type.EmptyTypes, modifiers: null);
            //return (TClass) ctor?.Invoke(parameters: null);

            return Activator.CreateInstance(classType);
        }


        /// <summary>
        ///     Serializes a instance of a class to a string
        /// </summary>
        private static void SerializeClass(object value, [NotNull] Stream stream, [NotNull] IFormatter formatter)
        {
            Check.NotNull(stream, nameof(stream));
            Check.NotNull(formatter, nameof(formatter));

            formatter.Serialize(stream, value);
            stream.Flush();
        }


        private static byte[] StreamToByteArray(Stream input)
        {
            if (input.Length == 0)
                return new byte[] { };

            if (input is MemoryStream ms)
                return ms.ToArray();

            using (ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}