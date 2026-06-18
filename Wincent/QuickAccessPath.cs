using System;

namespace Wincent
{
    /// <summary>
    /// Represents a Quick Access path without requiring the file system entry to exist.
    /// </summary>
    public sealed class QuickAccessPath : IEquatable<QuickAccessPath>
    {
        /// <summary>
        /// Initializes a path wrapper.
        /// </summary>
        /// <param name="fullName">The Explorer path string.</param>
        public QuickAccessPath(string fullName)
        {
            FullName = fullName ?? string.Empty;
        }

        /// <summary>
        /// Gets the Explorer path string.
        /// </summary>
        public string FullName { get; }

        /// <summary>
        /// Gets the Explorer path string.
        /// </summary>
        public override string ToString()
        {
            return FullName;
        }

        /// <summary>
        /// Compares this path to another path using ordinal string equality.
        /// </summary>
        public bool Equals(QuickAccessPath other)
        {
            return other != null && string.Equals(FullName, other.FullName, StringComparison.Ordinal);
        }

        /// <summary>
        /// Compares this path to another object.
        /// </summary>
        public override bool Equals(object obj)
        {
            return Equals(obj as QuickAccessPath);
        }

        /// <summary>
        /// Gets a hash code for the path string.
        /// </summary>
        public override int GetHashCode()
        {
            return StringComparer.Ordinal.GetHashCode(FullName);
        }
    }
}
