//-----------------------------------------------------------------------
// <copyright file="RecipientsPreview.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.Management.Teams.SearchTeamsMessage
{
    /// <summary>
    /// Defines the <see cref="RecipientsPreview"/> class.
    /// </summary>
    public class RecipientsPreview
    {
        /// <summary>
        /// Gets or sets the 'Sender' property
        /// </summary>
        public Sender Sender { get; set; }

        /// <summary>
        /// Gets or sets the 'Recipients' property
        /// </summary>
        public Recipients[] Recipients { get; set; }

        /// <summary>
        /// Gets or sets the 'RecipientsCount' property
        /// </summary>
        public int RecipientsCount { get; set; }
    }
}
