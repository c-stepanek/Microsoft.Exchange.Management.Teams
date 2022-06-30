//-----------------------------------------------------------------------
// <copyright file="Recipients.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.Management.Teams.SearchTeamsMessage
{
    /// <summary>
    /// Defines the <see cref="Recipient"/> class.
    /// </summary>
    public class Recipients
    {
        /// <summary>
        /// Gets or sets the 'Name' property
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the 'EmailAddress' property
        /// </summary>
        public string EmailAddress { get; set; }
    }
}
