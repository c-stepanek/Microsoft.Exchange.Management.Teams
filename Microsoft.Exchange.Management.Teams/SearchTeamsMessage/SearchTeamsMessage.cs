//-----------------------------------------------------------------------
// <copyright file="SearchTeamsMessage.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.Management.Teams.SearchTeamsMessage
{
    using System;
    using System.Linq;
    using System.Management.Automation;
    using System.Threading.Tasks;
    using Microsoft.Exchange.Management.Teams.Common;
    using Microsoft.Exchange.WebServices.Data;
    using Microsoft.Identity.Client;
    using Newtonsoft.Json;

    /// <summary>
    /// Defines the <see cref="SearchTeamsMessage"/> cmdlet class.
    /// </summary>
    [Cmdlet(VerbsCommon.Search, "TeamsMessage")]
    [OutputType(typeof(SearchResult))]
    public sealed class SearchTeamsMessage : PSCmdlet
    {
        #region Internal Class

        /// <summary>
        /// Define parameter sets
        /// </summary>
        private static class ParameterSet
        {
            internal const string MessageContains = "MessageContains";
            internal const string Topic = "Topic";
            internal const string ThreadId = "ThreadId";
        }

        #endregion Internal Class

        #region Constants

        /// <summary>
        /// The name of the folder to be searched.
        /// </summary>
        private const string FolderName = "TeamsMessagesData";

        /// <summary>
        /// The extension promoted properties identifier.
        /// </summary>
        private const string ExtensionPromotedPropertiesId = "{F6E4BA45-C83C-45DA-8F38-43BD3FE76D5C}";

        /// <summary>
        /// The extended property name for CreatedDateTime.
        /// </summary>
        private const string CreatedDateTimeExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#CreatedDateTime";

        /// <summary>
        /// The extended property name for FromSkypeInternalId.
        /// </summary>
        private const string FromSkypeInternalIdExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#FromSkypeInternalId";

        /// <summary>
        /// The extended property name for HasCardAttachments.
        /// </summary>
        private const string HasCardAttachmentsExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#HasCardAttachments";

        /// <summary>
        /// The extended property name for MessageType.
        /// </summary>
        private const string MessageTypeExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#MessageType";

        /// <summary>
        /// The extended property name for ParentMessageId.
        /// </summary>

        private const string ParentMessageIdExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#ParentMessageId";

        /// <summary>
        /// The extended property name for RecipientsPreview.
        /// </summary>
        private const string RecipientsPreviewExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#RecipientsPreview";

        /// <summary>
        /// The extended property name for SenderTenantId.
        /// </summary>
        private const string SenderTenantIdExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#SenderTenantId";

        /// <summary>
        /// The extended property name for SkypeItemId.
        /// </summary>
        private const string SkypeItemIdExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#SkypeItemId";

        /// <summary>
        /// The extended property name for ThreadId.
        /// </summary>
        private const string ThreadIdExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#ThreadId";

        /// <summary>
        /// The extended property name for ThreadType.
        /// </summary>
        private const string ThreadTypeExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#ThreadType";

        /// <summary>
        /// The extended property name for Topic (Group Chat Name)
        /// </summary>
        private const string GroupChatNameExtendedProperty = "IOpenTypedFacet.SkypeSpaces_ConversationPost_Extension#Topic";

        #endregion Constants

        #region Fields

        /// <summary>
        /// The OAuth access token
        /// </summary>
        private PSVariable token;

        /// <summary>
        /// The Guid for ExtensionPromotedProperties
        /// </summary>
        private Guid ExtensionPromotedPropertiesGuid = new Guid(ExtensionPromotedPropertiesId);

        #endregion Fields

        #region Parameters

        /// <summary>
        /// Gets or sets the Sender property
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = ParameterSet.MessageContains)]
        public string Sender { get; set; }

        /// <summary>
        /// Gets or sets the Recipient property
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = ParameterSet.MessageContains)]
        public string Recipient { get; set; }

        /// <summary>
        /// Gets or sets the MessageContains property
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = ParameterSet.MessageContains)]
        public string MessageContains { get; set; }

        /// <summary>
        /// Gets or sets the Topic property
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = ParameterSet.Topic)]
        public string Topic { get; set; }

        /// <summary>
        /// Gets or sets the ThreadId property
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = ParameterSet.ThreadId)]
        public string ThreadId { get; set; }

        #endregion Parameters

        #region Overrides

        /// <summary>
        ///  Performs initialization of command execution
        /// </summary>
        protected override void BeginProcessing()
        {
            // Check and see if access token exists for this session
            this.token = SessionState.PSVariable.Get("token");
            if (this.token == null)
            {
                // Try to aquire a new access token
                try
                {
                    AuthenticationProvider authProvider = new AuthenticationProvider();
                    Task<AuthenticationResult> authResult = authProvider.GetOAuthToken();
                    string accessToken = authResult.GetAwaiter().GetResult().AccessToken;
                    SessionState.PSVariable.Set("token", accessToken);
                }
                catch (MsalException ex)
                {
                    this.ThrowTerminatingError(
                        new ErrorRecord(
                            new Exception("Could not acquire access token.\nMake sure you have added your TenantId and ApplicationId to the confiugration file.", ex),
                            string.Empty,
                            ErrorCategory.AuthenticationError,
                            null));
                    return;
                }
            }
        }

        /// <summary>
        /// Provides a record-by-record processing functionality for the cmdlet.
        /// </summary>
        protected override void ProcessRecord()
        {
            // Get the access token stored as a session state variable
            this.token = SessionState.PSVariable.Get("token");

            // Create Exchange Service object (EWS)
            ExchangeService exchangeService = new ExchangeService
            {
                Url = new Uri($"https://outlook.office365.com/EWS/Exchange.asmx"),
                Credentials = new OAuthCredentials(this.token.Value.ToString())
            };

            // Setup folder view
            FolderId rootFolderId = new FolderId(WellKnownFolderName.Root);
            FolderView folderView = new FolderView(1000)
            {
                Traversal = FolderTraversal.Deep
            };

            // Filter for the folder we want
            SearchFilter.IsEqualTo searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, FolderName);
            FindFoldersResults findFolderResults = exchangeService.FindFolders(rootFolderId, searchFilter, folderView);

            if (findFolderResults.Count() == 1)
            {
                Folder folder = findFolderResults.Folders[0];
                Folder teamsMessagesData = Folder.Bind(exchangeService, folder.Id);

                ItemView itemView = new ItemView(1000)
                {
                    Traversal = ItemTraversal.Shallow,
                    PropertySet = new PropertySet(EmailMessageSchema.Sender)
                };

                itemView.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
                FindItemsResults<Item> items;

                switch (this.ParameterSetName)
                {
                    case ParameterSet.Topic:
                        ExtendedPropertyDefinition topicProp = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, GroupChatNameExtendedProperty, MapiPropertyType.String);
                        SearchFilter topicFilter = new SearchFilter.ContainsSubstring(topicProp, this.Topic, ContainmentMode.Substring, ComparisonMode.IgnoreCase);
                        items = exchangeService.FindItems(teamsMessagesData.Id, topicFilter, itemView);
                        break;

                    case ParameterSet.ThreadId:
                        ExtendedPropertyDefinition threadIdProp = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, ThreadIdExtendedProperty, MapiPropertyType.String);
                        SearchFilter threadIdFilter = new SearchFilter.IsEqualTo(threadIdProp, this.ThreadId);
                        items = exchangeService.FindItems(teamsMessagesData.Id, threadIdFilter, itemView);
                        break;

                    case ParameterSet.MessageContains:
                    default:
                        ExtendedPropertyDefinition PR_SENDER_SMTP_ADDRESS = new ExtendedPropertyDefinition(0x5D01, MapiPropertyType.String);
                        ExtendedPropertyDefinition RECIPIENTS_PREVIEW = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, RecipientsPreviewExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition PR_PREVIEW = new ExtendedPropertyDefinition(0x3FD9, MapiPropertyType.String);

                        SearchFilter senderFilter = new SearchFilter.IsEqualTo(PR_SENDER_SMTP_ADDRESS, this.Sender);
                        SearchFilter recipientsFilter = new SearchFilter.ContainsSubstring(RECIPIENTS_PREVIEW, this.Recipient, ContainmentMode.Substring, ComparisonMode.IgnoreCase);
                        SearchFilter messageFilter = new SearchFilter.ContainsSubstring(PR_PREVIEW, this.MessageContains, ContainmentMode.Substring, ComparisonMode.IgnoreCase);
                        SearchFilter.SearchFilterCollection searchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And, senderFilter, recipientsFilter, messageFilter);
                        items = exchangeService.FindItems(teamsMessagesData.Id, searchFilterCollection, itemView);
                        break;
                }

                if (items != null)
                {
                    int matchCount = 0;
                    foreach (Item item in items)
                    {
                        // Add some additional properties to be returned
                        // Property tags above 0x8000 are named properties, which are properties that include a GUID and either a Unicode character string or numeric value.
                        ExtendedPropertyDefinition createdDateTime = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, CreatedDateTimeExtendedProperty, MapiPropertyType.SystemTime);
                        ExtendedPropertyDefinition fromSkypeInternalId = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, SkypeItemIdExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition hasCardAttachments = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, HasCardAttachmentsExtendedProperty, MapiPropertyType.Boolean);
                        ExtendedPropertyDefinition messageType = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, MessageTypeExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition parentMessageId = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, ParentMessageIdExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition recipientsPreview = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, RecipientsPreviewExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition senderTenantId = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, SenderTenantIdExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition skypeItemId = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, SkypeItemIdExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition threadId = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, ThreadIdExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition threadType = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, ThreadTypeExtendedProperty, MapiPropertyType.String);
                        ExtendedPropertyDefinition topic = new ExtendedPropertyDefinition(ExtensionPromotedPropertiesGuid, GroupChatNameExtendedProperty, MapiPropertyType.String);

                        PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties)
                        {
                            ItemSchema.TextBody,
                            createdDateTime,
                            fromSkypeInternalId,
                            hasCardAttachments,
                            messageType,
                            parentMessageId,
                            recipientsPreview,
                            senderTenantId,
                            skypeItemId,
                            threadId,
                            threadType,
                            topic
                        };

                        // Bind to the message and get the properties we want
                        EmailMessage message = EmailMessage.Bind(exchangeService, item.Id, propSet);
                        message.TryGetProperty(createdDateTime, out DateTime createdDateTimeValue);
                        message.TryGetProperty(fromSkypeInternalId, out string fromSkypeInternalIdValue);
                        message.TryGetProperty(hasCardAttachments, out bool hasCardAttachmentsValue);
                        message.TryGetProperty(messageType, out string messageTypeValue) ;
                        message.TryGetProperty(parentMessageId, out string parentMessageIdValue);
                        message.TryGetProperty(recipientsPreview, out string recipientsPreviewStringValue);
                        message.TryGetProperty(senderTenantId, out string senderTenantIdValue);
                        message.TryGetProperty(skypeItemId, out string skypeItemIdValue);
                        message.TryGetProperty(threadId, out string threadIdValue);
                        message.TryGetProperty(threadType, out string threadTypeValue);
                        message.TryGetProperty(topic, out string topicValue);

                        RecipientsPreview recipients = JsonConvert.DeserializeObject<RecipientsPreview>(recipientsPreviewStringValue);

                        string emailAddresses = String.Empty;

                        for (int i = 0; i < message.ToRecipients.Count; i++)
                        { 
                                emailAddresses += $"{message.ToRecipients[i].Address}; ";
                        }

                        emailAddresses = emailAddresses.Substring(0,emailAddresses.Length-2).ToLower();
                        
                        SearchResult searchResult = new SearchResult()
                        {
                            Sender = message.Sender.Address,
                            Recipient = emailAddresses,
                            Message = message.TextBody.ToString().TrimEnd('\n'),
                            ItemClass = message.ItemClass,
                            CreatedDateTime = createdDateTimeValue,
                            FromSkypeInternalId = fromSkypeInternalIdValue,
                            HasCardAttachments = hasCardAttachmentsValue,
                            MessageType = messageTypeValue,
                            ParrentMessageId = parentMessageIdValue,
                            RecipientsPreview = recipients,
                            SenderTenantId = senderTenantIdValue,
                            SkypeItemId = skypeItemIdValue,
                            ThreadId = threadIdValue,
                            ThreadType = threadTypeValue,
                            Topic = topicValue
                        };

                        matchCount++;

                        // Write the search result object to the pipeline
                        this.WriteObject(searchResult);
                    }

                    if (matchCount == 0)
                    {
                        this.WriteWarning($"No messages found for recipient '{this.Recipient}' and message contains '{this.MessageContains}'");
                    }
                }
                else
                {
                    this.WriteWarning($"No messages found for sender '{this.Sender}'");
                }
            }
            else
            {
                this.WriteError(
                    new ErrorRecord(
                        new Exception($"Folder not found '{FolderName}'. The user may not be enabled for Microsoft Teams."),
                        string.Empty,
                        ErrorCategory.ObjectNotFound,
                        null));
            }
        }

        /// <summary>
        /// Performs clean-up after the command execution
        /// </summary>
        protected override void EndProcessing()
        {
            base.EndProcessing();
        }
        #endregion Overrides
    }
}
