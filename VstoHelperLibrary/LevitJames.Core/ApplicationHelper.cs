// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using JetBrains.Annotations;
 
namespace LevitJames.Core
{
	/// <summary>
	///     Provides application level functions.
	/// </summary>
	public static class ApplicationHelper
	{
   
        /// <summary>
        ///     Returns whether the code in the entry Assembly has been obfuscated or not.
        /// </summary>

        public static bool IsObfuscated() => IsObfuscated(Assembly.GetEntryAssembly() ?? Assembly.GetCallingAssembly());

	    /// <summary>
	    ///     Returns whether the passed Assembly has been obfuscated or not.
	    /// </summary>
        public static bool IsObfuscated([NotNull] Assembly assembly)
        {

            Check.NotNull(assembly, nameof(assembly));

	        var isObfuscated = false;

	        foreach (CustomAttributeData attributeData in assembly.CustomAttributes)
	        {
	            if (attributeData.AttributeType.Name == "DotfuscatorAttribute")
	            {
	                isObfuscated = true;
	                break;
	            }
	        }
	        return isObfuscated;
        }


	    /// <summary>
	    /// Returns if the assembly is a Debug assembly or a Release assembly.
	    /// </summary>
	    /// <param name="assembly">The assembly to check</param>
	    /// <returns></returns>
        public static bool IsDebugBuild(Assembly assembly) => assembly.GetCustomAttribute<DebuggableAttribute>()?.IsJITTrackingEnabled == true;

        /// <summary>Formats the product version into a user displayable format. For example "1.6.322 RC1 64 bit (O) "</summary>
        ///<remarks>
        /// The 'RC1' (or any other string) is taken from the AssemblyInformationalVersionAttribute of the entry assembly.
        /// An '(O)' string is appended if the Assembly is obfuscated.
        /// An '(Debug)' string is appended if the build type is a debug build.
        /// The method first checks the Assembly.EntryAssembly, if this is null then Assembly.GetCallingAssembly is used.
        /// </remarks>
        public static string VersionDisplayString()
	        => VersionDisplayString(Assembly.GetEntryAssembly() ?? Assembly.GetCallingAssembly());


        /// <summary>Formats the product version into a user displayable format. For example "1.6.322 RC1 64 bit (O) "</summary>
        /// <param name="assembly">The assembly to get the version information from</param>
        ///<remarks>
        /// The 'RC1' (or any other string) is taken from the AssemblyInformationalVersionAttribute of the entry assembly.
        /// An '(O)' string is appended if the Assembly is obfuscated.
        /// An '(Debug)' string is appended if the build type is a debug build.
        /// The method first checks the Assembly.EntryAssembly, if this is null then Assembly.GetCallingAssembly is used.
        /// </remarks>
        public static string VersionDisplayString([NotNull] Assembly assembly)
	    {
            Check.NotNull(assembly, nameof(assembly));
	   
	        var version = assembly.GetName().Version.ToString();
	        var versionInfo = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
	        var isObfuscated = IsObfuscated(assembly);
	        var isDebugBuild = IsDebugBuild(assembly);
	        var isX64 = Environment.Is64BitProcess;

	        return version + (isX64 ? " 64 bit " : " 32 bit ")
	                       + (versionInfo)
	                       + (isDebugBuild ? " (Debug)" : null)
	                       + (isObfuscated ? " (O)" : null);

	    }

        /// <summary>Formats the product version into a user displayable format. For example "1.6.322 64 bit"</summary>
        /// <param name="version">
        ///     <para>A Version instance.</para>
        /// </param>
        /// <param name="isDebugBuild">Appends a '((Debug))' string if the build type is a debug build.</param>
        /// /// <param name="isObfuscated">Appends a '(Debug)' string if the Assembly is obfuscated.</param>
        public static string VersionDisplayString(string version, bool isDebugBuild, bool isObfuscated)
	    {
	        Check.NotNull(version, "version");
	        
            return version + (Environment.Is64BitProcess ? " 64 bit " : " 32 bit ") 
                           + (isDebugBuild ? " (Debug)" : null) 
	                       + (isObfuscated ? " (O)" : null);
        }



		/// <summary>
		///     Returns the bitness of a .NET assembly
		/// </summary>
		/// <param name="assemblyOrNativeFileName">The file name of the application dll or executable to check</param>
		/// <returns>A Bitness value indicating the bitness of the Assembly</returns>
		public static Bitness GetBitness(string assemblyOrNativeFileName)
		{
			if (!File.Exists(assemblyOrNativeFileName))
				throw new FileNotFoundException(assemblyOrNativeFileName);
			var bitness = Bitness.Unknown;
			try
			{
				bitness = GetNativeBitness(assemblyOrNativeFileName);
				if (bitness == Bitness.Unknown)
					bitness = GetAssemblyBitness(assemblyOrNativeFileName);
			}
#pragma warning disable CA1031 // Do not catch general exception types
            catch
#pragma warning restore CA1031 // Do not catch general exception types
            {
				//Ignore
			}

			return bitness;
		}

		/// <summary>
		///     Returns the bitness of a .NET assembly
		/// </summary>
		/// <param name="assemblyFileName">The file name of the assembly to check</param>
		/// <returns>A Bitness value indicating the bitness of the Assembly</returns>
		public static Bitness GetAssemblyBitness(string assemblyFileName)
		{
			var assembly = Assembly.ReflectionOnlyLoadFrom(assemblyFileName);
			return GetAssemblyBitness(assembly);
		}


		/// <summary>
		///     Returns the bitness of a .NET assembly
		/// </summary>
		/// <param name="assembly">The assembly to check</param>
		/// <returns>A Bitness value indicating the bitness of the Assembly</returns>
		public static Bitness GetAssemblyBitness(Assembly assembly)
		{
            Check.NotNull(assembly,nameof(assembly));

			assembly.ManifestModule.GetPEKind(out PortableExecutableKinds kinds, out ImageFileMachine _);

			if ((kinds & PortableExecutableKinds.Required32Bit) == PortableExecutableKinds.Required32Bit)
				return Bitness.x86;

			if ((kinds & PortableExecutableKinds.PE32Plus) == PortableExecutableKinds.PE32Plus)
				return Bitness.x64;

			if ((kinds & PortableExecutableKinds.ILOnly) == PortableExecutableKinds.ILOnly)
				return Bitness.AnyCPU;

			return Bitness.Unknown;
		}


		/// <summary>
		///     Returns the bitness of a Native (None .Net) file
		/// </summary>
		/// <param name="fileName">The native File to check</param>
		/// <returns>A Bitness value indicating the bitness of the File</returns>
		public static Bitness GetNativeBitness(string fileName)
		{
			var retVal = GetImageArchitecture(fileName);
			if (retVal == 0x10b)
				return Bitness.x86;

			if (retVal == 0x20B)
				return Bitness.x64;

			return Bitness.Unknown;
		}


		/// <summary>
		///     Searches for and retrieves a file or protocol association-related string from the registry.
		/// </summary>
		/// <param name="extensionOrProgramId">An application's file extension or ProgID, such as '.docx' or  'Word.Document.8.'</param>
		public static string GetAssociatedApplication(string extensionOrProgramId)
        {
            var bufferCapacity = 0;

            var hr = NativeMethods.AssocQueryString(NativeMethods.ASSOCF_IGNOREUNKNOWN, NativeMethods.ASSOCSTR_EXECUTABLE,
                extensionOrProgramId, null, null, ref bufferCapacity);
            if (bufferCapacity == 0) // Don't use hr here
                throw Marshal.GetExceptionForHR(hr);

            var buffer = new StringBuilder(bufferCapacity);
            hr = NativeMethods.AssocQueryString(NativeMethods.ASSOCF_IGNOREUNKNOWN, 
                NativeMethods.ASSOCSTR_EXECUTABLE, extensionOrProgramId,  null,
                buffer,ref bufferCapacity);
            if (hr != 0)
                throw Marshal.GetExceptionForHR(hr);

            buffer.Capacity = bufferCapacity;
            return buffer.ToString();
        }


		/// <summary>Determines if link provided is a valid link address to the web.</summary>
		/// <param name="candidateUrl">Uri of the web address.</param>
		/// <param name="defaultPrefix">If missing a prefix, the input prefix will be assumed.</param>
		/// <returns>True, if the link is a valid web address.</returns>
		public static bool IsValidWebAddress(Uri candidateUrl, string defaultPrefix = "https")
		{
			return IsValidWebAddress(candidateUrl?.OriginalString, defaultPrefix);
		}

		/// <summary>Determines if link provided is a valid link address to the web.</summary>
		/// <param name="candidateUrl">Text of the web address.</param>
		/// <param name="defaultPrefix">If missing a prefix, the input prefix will be assumed.</param>
		/// <returns>True, if the link is a valid web address.</returns>
		public static bool IsValidWebAddress(string candidateUrl, string defaultPrefix = "https")
		{
			if (string.IsNullOrEmpty(candidateUrl))
				return false;

            candidateUrl = candidateUrl.Trim();

			// Try #3
			// Use Regex
			// Reference https://mathiasbynens.be/demo/url-regex
			// Use Diego Perini regex
			const string pattern = @"^" +
								   // protocol identifier
								   @"(?:(?:https?|ftp)://)" +
								   // user:pass authentication
								   @"(?:\S+(?::\S*)?@)?" +
								   "(?:" +
								   // IP address exclusion
								   // private & local networks
								   @"(?!(?:10|127)(?:\.\d{1,3}){3})" +
								   @"(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})" +
								   @"(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})" +
								   // IP address dotted notation octets
								   // excludes loopback network 0.0.0.0
								   // excludes reserved space >= 224.0.0.0
								   // excludes network & broacast addresses
								   // (first & last IP address of each class)
								   @"(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])" +
								   @"(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}" +
								   @"(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))" +
								   "|" +
								   // host name
								   @"(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)" +
								   // domain name
								   @"(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*" +
								   // TLD identifier
								   @"(?:\.(?:[a-z\u00a1-\uffff]{2,}))" +
								   // TLD may end with dot
								   @"\.?" +
								   ")" +
								   // port number
								   @"(?::\d{2,5})?" +
								   // resource path
								   @"(?:[/?#]\S*)?" +
								   "$";

			// Validate prefix
			var webAddress = candidateUrl;
			if (!string.IsNullOrEmpty(defaultPrefix))
			{
				// If user left off http://, https://, then add it now before validating
				var prefixIndex = webAddress.IndexOf("://", StringComparison.InvariantCulture);
				if (prefixIndex == -1)
				{
					// Not found
					webAddress = $"{defaultPrefix}://{webAddress}";
				}
				else
				{
					var prefix = webAddress.Substring(0, prefixIndex).ToLower();
					switch (prefix)
					{
						case "http":
						case "https":
						case "ftp":
							// Prefix is good
							break;
						default:
							webAddress = $"{defaultPrefix}{webAddress.Substring(prefixIndex)}";
							break;
					}
				}
			}

			var rgx = new Regex(pattern, RegexOptions.IgnoreCase);
			var matches = rgx.Matches(webAddress);
			return matches.Count > 0;

			// Try #2
			// KDP NOTE: The IsWellFormedUriString determines if there are illegal characters, not if the url is valid syntactically
			// This call uses built-in framework call. Relative allows "google.com" as a valid address.
			//return Uri.IsWellFormedUriString(link, UriKind.RelativeOrAbsolute);

			// Try #1
			//if (string.IsNullOrEmpty(link)) return false;
			//var webAddressPrefixes = new[] { "http:", "https:", "www.", "ftp:", "ftp." };
			//return webAddressPrefixes.Any(prefix => link.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
		}


        /// <summary>Pings the website for a valid return code.</summary>
        /// <param name="webAddress">Web address to ping.</param>
        /// <param name="token"></param>
        /// <returns>Returns true if ping generates successful return code (IPStatus.Success).</returns>
        public async static System.Threading.Tasks.Task<string> WebsiteStatusTextAsync(Uri webAddress, CancellationToken token = default(CancellationToken))

        {

            // Code from https://stackoverflow.com/questions/924679/c-sharp-how-can-i-check-if-a-url-exists-is-valid

            //NJKA: Upadted to support async and cancellation.

            if (!(WebRequest.Create(webAddress) is HttpWebRequest request))
				return null;

            CancellationTokenRegistration cancelRegister = default(CancellationTokenRegistration);
            if (token != default(CancellationToken))
                cancelRegister = token.Register(() => request.Abort(), useSynchronizationContext: false);
 
            try {
                request.Timeout = 5000; //set the timeout to 5 seconds to keep the user from waiting too long for the page to load
				request.Method = "HEAD"; //Get only the header information -- no need to download any content


				if (!(await request.GetResponseAsync()
					.ConfigureAwait(false) is HttpWebResponse response))
					return string.Empty;

				var statusCode = (int)response.StatusCode;
				if (statusCode >= 100 && statusCode < 400) //Good requests

					return null;

				if (statusCode >= 500 && statusCode <= 510) //Server Errors
					return $"The remote server has thrown an internal error. Url is not valid: {webAddress}";

            } catch (WebException ex) {
                return ex.Status == WebExceptionStatus.ProtocolError
                           ? $"The Url could not be resolved: {webAddress}"
                           : $"Unhandled status [{ex.Status}] returned for url: {webAddress}";
#pragma warning disable CA1031 // Do not catch general exception types
            } catch (Exception ex) {
#pragma warning restore CA1031 // Do not catch general exception types
                return $"Error trying to test Url: {webAddress}; Error: {ex.Message}";
            }
            finally {
                cancelRegister.Dispose();
            }
            return $"Unknown response trying to test Url: {webAddress}";
		}


		private static ushort GetImageArchitecture(string filepath)
		{
            using (var stream = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            using (var reader = new BinaryReader(stream))
            { 
                //check the MZ signature to ensure it's a valid Portable Executable image
                if (reader.ReadUInt16() != 23117)
                    return 0;

                // seek to, and read, e_lfanew then advance the stream to there (start of NT header)
                stream.Seek(0x3A, SeekOrigin.Current);
                stream.Seek(reader.ReadUInt32(), SeekOrigin.Begin);

                // Ensure the NT header is valid by checking the "PE\0\0" signature
                if (reader.ReadUInt32() != 17744)
                    return 0;

                // seek past the file header, then read the magic number from the optional header
                stream.Seek(20, SeekOrigin.Current);
                return reader.ReadUInt16();
            }
        }


        /// <summary>
        ///     Shows a basic message box without loading System.Windows.Forms
        /// </summary>
        /// <param name="message">The message to display</param>
        /// <param name="title">The title to display</param>
        /// <param name="flags">flags that specify the buttons and icon for the message</param>
        /// <param name="owner">The handle of the message box owner</param>
        public static void ShowMessage(string message, string title, int flags = 0, IntPtr owner = default(IntPtr))
		{
            // ReSharper disable InconsistentNaming
            // ReSharper disable IdentifierTypo
		    const int MB_TASKMODAL = 0x00002000;
            const int MB_SETFOREGROUND = 0x00010000;
		    // ReSharper restore IdentifierTypo
		    // ReSharper restore InconsistentNaming

            if (owner == IntPtr.Zero)
                flags |= MB_TASKMODAL;

		    flags |= MB_SETFOREGROUND;

            _ = NativeMethods.MessageBox(owner, message, title, flags);
		}
        /// <summary>
        /// Shows a MessageBox with an Error icon
        /// </summary>
        /// <param name="message">The message to display</param>
        /// <param name="title">The title to display</param>
        /// <param name="flags">flags that specify the buttons and icon for the message</param>
        /// <param name="owner">The handle of the message box owner</param>
	    public static void ShowErrorMessage(string message, string title, int flags = 0, IntPtr owner = default(IntPtr))
        {
            // ReSharper disable InconsistentNaming
            // ReSharper disable IdentifierTypo
            const int MB_ICONERROR = 0x00000010;
            // ReSharper restore IdentifierTypo
            // ReSharper restore InconsistentNaming
            ShowMessage(message, title, flags | MB_ICONERROR, owner);
        }

        ///// <summary>
        ///// Returns all the instances in the Running object table (ROT) that match the provided Program Identifiers (ProgID).
        ///// If not Program Identifiers are provided then all instances in the ROT are returned.
        ///// </summary>
        ///// <param name="progIds"></param>
        //
        //public static IEnumerable<object> EnumRunningInstances([NotNull] params string[] progIds)
        //{
        //    var clsIds = new List<string>();

        //    foreach (var progId in progIds)
        //    {
        //        var type = Type.GetTypeFromProgID(progId);
        //        if (type != null)
        //            clsIds.Add(type.GUID.ToString("B"));
        //    }

        //    if (progIds.Length > 0 && clsIds.Count == 0)
        //        yield break; // No Types returned from the provided ProgIds.

        //    // get Running Object Table ...
        //    IRunningObjectTable rot;
        //    IBindCtx bindCtx;
        //    NativeMethods.CreateBindCtx(0, out bindCtx);
        //    bindCtx.GetRunningObjectTable(out rot);

        //    if (rot == null)
        //        yield break;

        //    // get enumerator for ROT entries

        //    IEnumMoniker monikerEnumerator;
        //    rot.EnumRunning(out monikerEnumerator);

        //    if (monikerEnumerator == null)
        //        yield break;

        //    monikerEnumerator.Reset();

        //    var pNumFetched = new IntPtr();
        //    var monikers = new IMoniker[1];

        //    // go through all entries and identifies app instances

        //    while (monikerEnumerator.Next(1, monikers, pNumFetched) == 0)
        //    {

        //        string displayName;
        //        monikers[0].GetDisplayName(bindCtx, null, out displayName);

        //        dynamic comObject = null;
        //        if (clsIds.Count == 0)
        //        {
        //            rot.GetObject(monikers[0], out comObject);
        //            if (comObject != null)
        //                yield return comObject;
        //        }
        //        else
        //        {
        //            if (clsIds.Any(clsId=> displayName.EndsWith(clsId, StringComparison.OrdinalIgnoreCase)))
        //                rot.GetObject(monikers[0], out comObject);
        //        }

        //        if (comObject == null)
        //            continue;

        //        yield return comObject;
        //    }

        //}
    }
}