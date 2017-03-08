//------------------------------------------------------------------------------
// <copyright>
//   Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace Microsoft.VisualStudio.Data.Tools
{
    /// <summary>
    /// This interface is used to expose the SqlServerObjectExplorerService, which can be used
    /// to perform various actions found in the SQL Server Object Explorer.
    /// All calls to the methods on this interface must be performed from the UI thread.
    /// Calls that are not made on the UI thread may throw exceptions.
    /// 
    /// This is v2 of the interface, and supports all existing methods plus an additional
    /// <see cref="IsSupportedSqlServerVersion"/> method.
    /// </summary>
    [ComImport]
    // Uncomment the line below for delivery to a customer in lieu of an interop assembly - Delete this comment as well
    [TypeIdentifier]
    [Guid("A5392196-68CF-4A64-A723-25F5EDEEF9A3")]
    public interface ISqlServerObjectExplorerService2 : ISqlServerObjectExplorerService
    {
        /// <summary>
        /// Checks if a connection to a specified SQL Server is supported by the SQL tooling.
        /// </summary>
        /// <param name="connection">The connection information.</param>
        /// <exception cref="System.ArgumentException">If the InitialCatalog in the connection exceeds the
        /// maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If connection is null.</exception>
        bool IsSupportedSqlServerVersion(SqlConnectionStringBuilder connection);

    }
    
}
