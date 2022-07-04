// © Copyright 2018 Levit & James, Inc.

using System;
using Microsoft.Win32;

namespace LevitJames.Core
{
    /// <summary>
    ///     Helper methods used to determine the OS version + other versioning functionality.
    /// </summary>
    public static class OSVersionHelper
    {
        //private static int _win10ReleaseId;

        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows Vista
        /// </summary>
        public static bool IsWindowsVistaOrGreater()
            => IsOsAtLeastVersion(majorVerion: 6, greaterThanOrEqualToMinorVersion: 0);

        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows Vista SP1
        /// </summary>
        public static bool IsWindowsVistaSp1OrGreater()
            => IsOsAtLeastVersion(majorVerion: 6, greaterThanOrEqualToMinorVersion: 0, greaterThanOrEqualToBuildVersion: 6001);

        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows Vista SP2
        /// </summary>
        public static bool IsWindowsVistaSp2OrGreater()
            => IsOsAtLeastVersion(majorVerion: 6, greaterThanOrEqualToMinorVersion: 0, greaterThanOrEqualToBuildVersion: 6002);


        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows 7.0
        /// </summary>
        public static bool IsWindows7OrGreater()
            => IsOsAtLeastVersion(majorVerion: 6, greaterThanOrEqualToMinorVersion: 1);

        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows 7.0 SP1
        /// </summary>
        public static bool IsWindows7SP1OrGreater()
            => IsOsAtLeastVersion(majorVerion: 6, greaterThanOrEqualToMinorVersion: 1,greaterThanOrEqualToBuildVersion: 7601);

        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows 8.0
        /// </summary>
        public static bool IsWindows8OrGreater()
            => IsOsAtLeastVersion(majorVerion: 6, greaterThanOrEqualToMinorVersion: 2);

        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows 8.1
        /// </summary>
        public static bool IsWindows8Point1OrGreater()
            => IsOsAtLeastVersion(majorVerion: 6, greaterThanOrEqualToMinorVersion: 3);

        /// <summary>Determines if the version of Windows is equal to or greater than Windows 10</summary>
        public static bool IsWindows10OrGreater()
            => IsOsAtLeastVersion(majorVerion: 10, greaterThanOrEqualToMinorVersion: 0);

        //https://docs.microsoft.com/en-us/windows/uwp/whats-new/windows-10-build-14393-api-diff
        /// <summary>
        ///     Determines if the version of Windows is equal to or greater than Windows 10 version 1607 (releaseId). This
        ///     version is commonly known as the Windows 10 Anniversary Addition
        /// </summary>
        public static bool IsWindows10AnniversaryAdditionOrGreater()
            => IsOsAtLeastVersion(majorVerion: 10, greaterThanOrEqualToMinorVersion: 0, greaterThanOrEqualToBuildVersion: 14393);

        //https://docs.microsoft.com/en-us/windows/uwp/whats-new/windows-10-build-15063-api-diff
        /// <summary>Determines if the version of Windows is equal to or greater than Windows 10 build 15063 (Creators Addition).</summary>
        public static bool IsWindows10CreatorsAdditionOrGreater()
            => IsOsAtLeastVersion(majorVerion: 10, greaterThanOrEqualToMinorVersion: 0, greaterThanOrEqualToBuildVersion: 15063);

        private static bool IsOsAtLeastVersion(int majorVerion, int greaterThanOrEqualToMinorVersion,
                                              int greaterThanOrEqualToBuildVersion = 0)
        {
            var osVersion = Environment.OSVersion.Version;
            //return osVersion.Major > majorVerion ||
            //       (osVersion.Major == majorVerion &&
            //        osVersion.Minor >= greaterThanOrEqualToMinorVersion &&
            //        osVersion.Build >= build);

            if ((osVersion.Major < majorVerion))
                return false;

            if ((osVersion.Major > majorVerion))
                return true;

            //Major equals

            if (osVersion.Minor < greaterThanOrEqualToMinorVersion)
                return false;

            if (osVersion.Minor > greaterThanOrEqualToMinorVersion)
                return true;

            //Major equals

            //So if build is greater or equals then we are good
            return (osVersion.Build >= greaterThanOrEqualToBuildVersion);

        }

        /// <summary>
        /// Gets the friendly name of the OS from the version. i.e. "Windows Vista", "Windows 7", etc.
        /// </summary>
        /// <returns>A string representing the name of the OS</returns>
        public static string FriendlyName()
        {
            var version = Environment.OSVersion.Version;
            if (version.Major == 6)
            {
                switch (version.Minor)
                {
                case 0:
                    return "Windows Vista";
                case 1:
                    return "Windows 7";
                case 2:
                    return "Windows 8";
                case 3:
                    return "Windows 8.1";
                }
            }
            else if (version.Major == 10)
            {

                if (IsWindows10AnniversaryAdditionOrGreater())
                    return "Windows 10 (with Anniversary Addition)";

                if (IsWindows10CreatorsAdditionOrGreater())
                    return "Windows 10 (with Creators Addition)";

                return "Windows 10";
            }

            return "Unknown";
        }

    }
}