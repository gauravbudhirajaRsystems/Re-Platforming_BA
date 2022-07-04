// © Copyright 2018 Levit & James, Inc.

using System;
using System.Reflection;
using LevitJames.Core;

//� Copyright 2009 Levit & James, Inc.

namespace LevitJames.MSOffice.Addin
{
    /// <summary>
    ///     This singleton class provides application information about the addin that is loaded.
    /// </summary>
    /// <remarks>
    ///     It is primarily used to provide support for the AddinSettingsProvider so it can identify
    ///     the assembly where the add-in was started from.
    /// </remarks>
    public sealed class AddinApplicationInfo
    {
        private string _companyName;
        private string _productName;

 
        internal AddinApplicationInfo(OfficeAddinConnection connection)
        {
            Check.NotNull(connection, "connection");
            Connection = connection;
             
        }


        // public members


        /// <summary>
        ///     Returns the assembly that set as the entry assembly for the add-in.
        /// </summary>
        public Assembly EntryAssembly => Connection.GetType().Assembly;


        /// <summary>
        ///     Returns the value of the AssemblyProductAttribute for the Addin Assembly
        /// </summary>
        public string ProductName => _productName ?? (_productName = GetProductName());


        /// <summary>
        ///     Returns the value of the AssemblyCompanyAttribute for the Addin Assembly
        /// </summary>
        public string CompanyName => _companyName ?? (_companyName = GetCompanyName());


        public Version EntryAssemblyVersion => EntryAssembly.GetName().Version;


        public OfficeAddinConnection Connection { get; private set; }


        private string GetProductName()
        {
            string productName = null;
            var attr = EntryAssembly.GetCustomAttribute<AssemblyProductAttribute>();
            if (attr != null)
            {
                productName = attr.Product;
            }

            if (string.IsNullOrEmpty(productName))
            {
                productName = EntryAssembly.GetName().Name;
            }

            return productName?.Trim();
        }


        private string GetCompanyName()
        {
            string companyName = null;
            var attr = EntryAssembly.GetCustomAttribute<AssemblyCompanyAttribute>();
            if (attr != null)
            {
                companyName = attr.Company;
            }

            if (string.IsNullOrEmpty(companyName))
            {
                companyName = ProductName;
            }

            return companyName?.Trim();
        }


        internal void Reset()
        {
            Connection = null;
        }
    }
}