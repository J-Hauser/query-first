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
    /// </summary>
    [ComImport]
    // Uncomment the line below for delivery to a customer in lieu of an interop assembly - Delete this comment as well
    [TypeIdentifier]
    [Guid("33C78F10-FB16-423C-850D-C44F3C78C736")]
    public interface ISqlServerObjectExplorerService
    {
        /// <summary>
        /// Creates a new connected T-SQL editor window populated with the new object template
        /// for the given object type.
        /// </summary>
        /// <param name="connection">The connection information for the new editor window.</param>
        /// <param name="objectType">The type of object.</param>
        /// <exception cref="System.ArgumentException">If the InitialCatalog in the connection exceeds the
        /// maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If connection is null.</exception>
        /// <exception cref="System.InvalidOperationException">If there was a problem opening up the editor
        /// or designer for the new item template.</exception>
        /// <exception cref="System.IO.FileNotFoundException">If there was a problem finding the item template
        /// for the object type.</exception>
        /// <exception cref="System.NotSupportedException">If the objectType is not supported.</exception>
        void Add(SqlConnectionStringBuilder connection, SqlServerObjectType objectType);

        /// <summary>
        /// Shows the SQL Server Object Explorer and expands to the instance provided in the connection.
        /// </summary>
        /// <param name="connection">The connection information for the database.</param>
        /// <exception cref="System.ArgumentException">If the InitialCatalog in the connection exceeds the
        /// maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If connection is null.</exception>
        /// <exception cref="System.InvalidOperationException">If there was a problem adding the connection.</exception>
        void Browse(SqlConnectionStringBuilder connection);

        /// <summary>
        /// Deletes objects and dependent objects from a database.  A confirmation dialog is presented to the user.
        /// </summary>
        /// <param name="connection">The connection information for the database where the object exists.</param>
        /// <param name="objectType">The type of object.</param>
        /// <param name="name">The full name of the object separated into its individual parts.</param>
        /// <exception cref="System.ArgumentException">If the name or any of its parts are empty or white space.
        /// This is also thrown if the InitialCatalog in the connection exceeds the maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If the connection or name parameters are null.</exception>
        /// <exception cref="System.NotSupportedException">If the objectType is not supported.</exception>
        void Delete(SqlConnectionStringBuilder connection, SqlServerObjectType objectType, params string[] name);

        /// <summary>
        /// Executes the given programmability object.  A parameter dialog is presented first, followed by
        /// the creation of a new connected T-SQL editor window where the statement is then executed.
        /// </summary>
        /// <param name="connection">The connection information for the database where the programmability object exists.</param>
        /// <param name="name">The full name of the object separated into its individual parts.</param>
        /// <exception cref="System.ArgumentException">If the name or any of its parts are empty or white space.
        /// This is also thrown if the InitialCatalog in the connection exceeds the maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If any argument is null.</exception>
        void Execute(SqlConnectionStringBuilder connection, params string[] name);

        /// <summary>
        /// Creates a new connected T-SQL editor window.
        /// </summary>
        /// <param name="connection">The connection information for the new T-SQL editor window.</param>
        /// <exception cref="System.ArgumentException">If the InitialCatalog in the connection exceeds the
        /// maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If connection is null.</exception>
        /// <exception cref="System.InvalidOperationException">If there was a problem using the connection.</exception>
        void OpenQuery(SqlConnectionStringBuilder connection);

        /// <summary>
        /// Opens the CREATE statement for an object in a new connected T-SQL editor window.
        /// </summary>
        /// <param name="connection">The connection information for the object.</param>
        /// <param name="objectType">The type of object.</param>
        /// <param name="name">The full name of the object in its individual parts.</param>
        /// <exception cref="System.ArgumentException">If the name or any of its parts are empty or white space.
        /// This is also thrown if the InitialCatalog in the connection exceeds the maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If the connection or name parameters are null.</exception>
        /// <exception cref="System.InvalidOperationException">If there was a problem opening the T-SQL editor.</exception>
        /// <exception cref="System.NotSupportedException">If the objectType is not supported.</exception>
        void ViewCode(SqlConnectionStringBuilder connection, SqlServerObjectType objectType, params string[] name);

        /// <summary>
        /// Opens an editable data grid for the object's data.
        /// </summary>
        /// <param name="connection">The connection information for the object with the data.</param>
        /// <param name="objectType">The type of object. Only Table and View are supported.</param>
        /// <param name="name">The full name of the object in its individual parts.</param>
        /// <exception cref="System.ArgumentException">If the name or any of its parts are empty or white space.
        /// This is also thrown if the InitialCatalog in the connection exceeds the maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If the connection or name parameters are null.</exception>
        /// <exception cref="System.NotSupportedException">If the objectType is not supported.</exception>
        void ViewData(SqlConnectionStringBuilder connection, SqlServerObjectType objectType, params string[] name);

        /// <summary>
        /// Opens the designer for the provided object.
        /// </summary>
        /// <param name="connection">The connection information for the object.</param>
        /// <param name="objectType">The type of object. Only Table is supported.</param>
        /// <param name="name">The full name of the object in its individual parts.</param>
        /// <exception cref="System.ArgumentException">If the name or any of its parts are empty or white space.
        /// This is also thrown if the InitialCatalog in the connection exceeds the maximum allowed length.</exception>
        /// <exception cref="System.ArgumentNullException">If the connection or name parameters are null.</exception>
        /// <exception cref="System.NotSupportedException">If the objectType is not Table.</exception>
        void ViewDesigner(SqlConnectionStringBuilder connection, SqlServerObjectType objectType, params string[] name);
    }
    
    /// <summary>
    /// This enum provides the available SQL Server objects that may be used
    /// in an implementation of the ISqlServerObjectExplorerService
    /// </summary>
    // Uncomment the TypeIdentifier below for delivery to customer in lieu of an interop assembly - Delete this comment as well
    [TypeIdentifier("2F20C022-820F-482D-95A8-55A46DA2DBDD", "Microsoft.VisualStudio.Data.Tools.SqlServerObjectType")]
    public enum SqlServerObjectType
    {
        /// <summary>
        /// Inline User-Defined Function
        /// </summary>
        FunctionInline,
        /// <summary>
        /// Scalar Function
        /// </summary>
        FunctionScalar,
        /// <summary>
        /// Table-Valued User-Defined Function
        /// </summary>
        FunctionTableValue,
        /// <summary>
        /// Stored Procedure
        /// </summary>
        StoredProcedure,
        /// <summary>
        /// Synonym
        /// </summary>
        Synonym,
        /// <summary>
        /// Table
        /// </summary>
        Table,
        /// <summary>
        /// DML Trigger
        /// </summary>
        Trigger,
        /// <summary>
        /// User-Defined Data Type
        /// </summary>
        UserDefinedDataType,
        /// <summary>
        /// View
        /// </summary>
        View,
    }
}
