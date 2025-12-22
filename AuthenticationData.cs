using System;

namespace GraphMailer
{
    /// <summary>
    /// Represents the authentication data required for Microsoft Graph API.
    /// This class is serializable to allow for easy storage and retrieval.
    /// </summary>
    [Serializable]
    public class AuthenticationData
    {
        /// <summary>
        /// Gets or sets the Tenant ID for the Azure Active Directory application.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the Client ID for the Azure Active Directory application.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Gets or sets the Client Secret for the Azure Active Directory application.
        /// </summary>
        public string ClientSecret { get; set; }
    }
}
