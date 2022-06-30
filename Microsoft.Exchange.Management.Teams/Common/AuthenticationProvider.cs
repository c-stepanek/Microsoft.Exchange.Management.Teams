// <copyright file="AuthenticationProvider.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.Management.Teams.Common
{
    using System.Threading.Tasks;
    using Microsoft.Identity.Client;

    /// <summary>
    /// Defines the <see cref="AuthenticationProvider"/> class.
    /// </summary>
    public class AuthenticationProvider
    {
        /// <summary>
        /// Gets the OAuth token
        /// </summary>
        /// <returns>Returns the authentication token.</returns>
        public async Task<AuthenticationResult> GetOAuthToken()
        {
            Configuration config = new Configuration();
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = config.GetConfigKeyValue("ApplicationId"),
                TenantId = config.GetConfigKeyValue("TenantId")
            };

            IPublicClientApplication pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

            string[] scopes = new string[] { $"https://outlook.office365.com/EWS.AccessAsUser.All" };

            AuthenticationResult authResult = await pca.AcquireTokenInteractive(scopes).ExecuteAsync();

            return authResult;
        }
    }
}
