//-----------------------------------------------------------------------
// <copyright file="SearchResult.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
using System;

namespace Microsoft.Exchange.Management.Teams.SearchTeamsMessage
{
    /// <summary>
    /// Defines the <see cref="SearchResult"/> cmdlet class.
    /// </summary>
    public class SearchResult
    {
        /// <summary>
        /// Gets or sets the 'Sender' property
        /// </summary>
        public string Sender { get; set; }

        /// <summary>
        /// Gets or sets the 'Recipient' property
        /// </summary>
        public string Recipient { get; set; }

        /// <summary>
        /// Gets or sets the 'Message' property
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Gets or sets the 'CreatedDateTime' property
        /// </summary>
        public DateTime CreatedDateTime { get; set; }

        /// <summary>
        /// Gets or sets the 'HasCardAttachments' property
        /// </summary>
        public bool HasCardAttachments { get; set; }

        /// <summary>
        /// Gets or sets the 'MessageType' property
        /// </summary>
        public string MessageType { get; set; }

        /// <summary>
        /// Gets or sets the 'ThreadId' property
        /// </summary>
        public string ThreadId { get; set; }

        /// <summary>
        /// Gets or sets the 'ThreadType' property
        /// </summary>
        public string ThreadType { get; set; }

        /// <summary>
        /// Gets or sets the 'Topic' property
        /// </summary>
        public string Topic { get; set; }

        /// <summary>
        /// Gets or sets the 'ItemClass' property
        /// </summary>
        public string ItemClass { get; set; }

        /// <summary>
        /// Gets or sets the 'FromSkypeInternalId' property
        /// </summary>
        public string FromSkypeInternalId { get; set; }

        /// <summary>
        /// Gets or sets the 'ParrentMessageId' property
        /// </summary>
        public string ParrentMessageId { get; set; }

        /// <summary>
        /// Gets or sets the 'RecipientsPreview' property
        /// </summary>
        public RecipientsPreview RecipientsPreview { get; set; }

        /// <summary>
        /// Gets or sets the 'SenderTenantId' property
        /// </summary>
        public string SenderTenantId { get; set; }

        /// <summary>
        /// Gets or sets the 'SkypeItemId' property
        /// </summary>
        public string SkypeItemId { get; set; }
    }
}
