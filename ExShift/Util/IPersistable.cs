/*
 * Author: Jan Baumann
 * Version: 27.09.2020
 */

namespace ExShift.Util
{
    /// <summary>
    /// This interface requires the functionality for entities to be persistable.
    /// </summary>
    public interface IPersistable
    {
        /// <summary>
        /// Returns the name of the table in which the entities are stored.
        /// </summary>
        /// <returns>Table name</returns>
        string GetTableName();

        /// <summary>
        /// Gets the short name for subtables. This is important since the excel worksheets
        /// have name length boundaries.
        /// </summary>
        /// <returns>Short name</returns>
        string GetShortName();

        /// <summary>
        /// Gets the primary key.
        /// </summary>
        /// <returns>Primary key</returns>
        double GetIndex();

        /// <summary>
        /// Sets the primary key.
        /// </summary>
        /// <param name="index">Primary key</param>
        void SetIndex(double index);
    }
}
