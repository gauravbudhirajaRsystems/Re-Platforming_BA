// © Copyright 2018 Levit & James, Inc.

using System;
using System.ComponentModel;
using System.Linq;
using System.Security.AccessControl;
using Microsoft.Win32;

namespace LevitJames.Core
{
    public static partial class CompilerExtensions
    {
        /// <summary>
        /// Copies the Key and its tree to the supplied registryKey.
        /// </summary>
        /// <param name="source">The key to copy from. The key must have been opened with at least read access rights</param>
        /// <param name="destination">The destination key. The calling process must have create sub key access rights to the key</param>
        /// <param name="sourceSubKey">A subkey of the source key parameter</param>
        /// <param name="deleteSourceSubKey">True to delete the source key after copying; false to keep the existing tree.</param>
        /// <param name="throwOnMissingSourceKey">If deleteSource is true, but the source cannot be deleted throw an exception; else ignore.</param>
        
        public static bool CopyTo(this RegistryKey source, RegistryKey destination, string sourceSubKey = null, bool deleteSourceSubKey = false, bool throwOnMissingSourceKey = true)
        {
            Check.NotNull(source, nameof(source));
            Check.NotNull(destination, nameof(destination));
 
            if (throwOnMissingSourceKey && !string.IsNullOrEmpty(sourceSubKey) && !source.Exists(sourceSubKey))
                throw new ArgumentException("sourceSubKey not found", nameof(sourceSubKey));
 
            if (source.Equals(destination))
                throw new InvalidOperationException();
            var retVal = NativeMethods.RegCopyTree(source.Handle, sourceSubKey, destination.Handle);

            if (retVal == 5)
                throw new UnauthorizedAccessException();

            if (retVal != 0)
                return false;

            if (deleteSourceSubKey && !string.IsNullOrEmpty(sourceSubKey))
                source.DeleteSubKeyTree(sourceSubKey, throwOnMissingSourceKey);

            return true;
        }

        ///// <summary>
        ///// Saves the key using supplied file name.
        ///// </summary>
        ///// <param name="source">The key, and subtree to save</param>
        ///// <param name="fileName">The name of the file to save the tree too.</param>
        //public static void Save(this RegistryKey source, string fileName)
        //{
        //    Check.NotNull(source, nameof(source));
        //    Check.NotEmpty(fileName, nameof(fileName));

        //    uint retVal;
        //    if (source == Registry.ClassesRoot)
        //        retVal = NativeMethods.RegSaveKey(source.Handle , fileName, IntPtr.Zero);
        //    else
        //        retVal = NativeMethods.RegSaveKeyEx(source.Handle, fileName, IntPtr.Zero, 2); //REG_STANDARD_FORMAT = 1

        //    if (retVal != 0)
        //        throw new Win32Exception();
        //}
 
        /// <summary>Adds a value to the registry, creating the key path as necessary.</summary>
        /// <param name="hive">The top-level registry key.</param>
        /// <param name="subKey">The sub-path from the hive to the reg key containing the value to be added.</param>
        /// <param name="valueName">The name of the new value.</param>
        /// <param name="value">The value of the new value.</param>
        /// <param name="valueKind">The kind of the new value.</param>
        /// <returns>Returns true if the value was added successfully.</returns>
        public static bool SetValue(this RegistryKey source, string subKey, string valueName, object value,
                                    RegistryValueKind valueKind)
        {
            Check.NotNull(source, nameof(source));

            using (var regKey = source.CreateSubKey(subKey, RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                regKey?.SetValue(valueName, value, valueKind);
            }

            return true;
        }


        /// <summary>Gets a value from the registry.</summary>
        /// <param name="hive">The top-level registry key.</param>
        /// <param name="subKey">The sub-path from the hive to the reg key containing the value.</param>
        /// <param name="valueName">The name of the existing value.</param>
        /// <param name="defaultValue">The default value, if the key does not exist.</param>
        /// <returns>Returns the data from the value.</returns>
        public static object GetKeyValue(this RegistryKey source, string subKey, string valueName, object defaultValue = null)
        {
            Check.NotNull(source, nameof(source));

            using (var regKey = source.OpenSubKey(subKey, RegistryKeyPermissionCheck.ReadSubTree, RegistryRights.ReadKey))
            {
                return regKey != null
                           ? regKey.GetValue(valueName, defaultValue)
                           : defaultValue;
            }
        }


        /// <summary>Deletes a value from the Registry.</summary>
        /// <param name="hive">The top-level key.</param>
        /// <param name="subKey">The sub-path from the hive to the key containing the value to be deleted.</param>
        /// <param name="valueName">The name of the value to be deleted.</param>
        /// <returns>Returns true if the value was deleted successfully.</returns>
        public static bool DeleteValue(this RegistryKey source, string subKey, string valueName)
        {
            Check.NotNull(source, nameof(source));

            using (var regKey = source.OpenSubKey(subKey, writable: true))
            {
                regKey?.DeleteValue(valueName, throwOnMissingValue: false);
            }

            return true;
        }


        /// <summary>Determines if a reg key exists.</summary>
        /// <param name="key">The sub-path from the hive to the last key in the path.</param>
        /// <param name="subKey">The path of the sub key</param>
        /// <returns>Returns true if the registry key exists in the path given.</returns>
        public static bool Exists(this RegistryKey source, string subKey)
        {
            Check.NotNull(source, nameof(source));

            using (var regKey = source.OpenSubKey(subKey, writable: false))
            {
                return regKey != null;
            }
        }


        /// <summary>Detemines if a value in the registry exists within a specified registry key.</summary>
        /// <param name="hive">The top-level key.</param>
        /// <param name="subKey">The sub-path from the hive to the key containing the value.</param>
        /// <param name="valueName">The name of the value being queried.</param>
        /// <returns>Returns true if the value exists for the specified registry key.</returns>
        public static bool HasValue(this RegistryKey hive, string subKey, string valueName)
        {
            Check.NotNull(hive, nameof(hive));

            using (var regkey = hive.OpenSubKey(subKey, writable: false))
            {
                if (regkey != null)
                {
                    return regkey.GetValueNames().Any(vName => string.Compare(vName, valueName, StringComparison.OrdinalIgnoreCase) == 0);
                }
            }

            return false;
        }
    }
}