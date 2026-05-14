using Azure.Core;
using Azure.Core.Diagnostics;
using Azure.Identity;
using Microsoft.Graph;
using System;
using System.IO;
using System.Diagnostics.Tracing;
using System.Reflection;
using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace GraphMailer
{
    /// <summary>
    /// Handles authentication with Microsoft Graph API.
    /// </summary>
    public class GraphAuth
    {
        private readonly AuthenticationData _authData;
        private static AzureEventSourceListener _listener;
        private static readonly ConcurrentDictionary<string, GraphServiceClient> _clientCache = new ConcurrentDictionary<string, GraphServiceClient>();

        static GraphAuth()
        {
            // --- ADDED FOR ASSEMBLY BINDING REDIRECTION ---
            // This handler is called when the CLR fails to load an assembly.
            // It allows us to programmatically load the correct version of an assembly
            // that is present in the bin directory, avoiding the need for app.config binding redirects.
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
            // --- END OF ASSEMBLY BINDING ---
        }

        /// <summary>
        /// Enables verbose logging for the Azure SDK to a specified file.
        /// </summary>
        /// <param name="logFilePath">The full path to the log file.</param>
        public static void EnableLogging(string logFilePath)
        {
            try
            {
                // Dispose existing listener if any
                _listener?.Dispose();

                // Ensure the directory exists
                var directory = Path.GetDirectoryName(logFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                _listener = new AzureEventSourceListener((eventData, message) =>
                {
                    try
                    {
                        File.AppendAllText(logFilePath, $"[{DateTime.UtcNow:HH:mm:ss.fffffff}][{eventData.Level}] {message}\n");
                    }
                    catch
                    {
                        // Ignore logging errors to prevent crashing the application
                    }
                }, EventLevel.Verbose);
            }
            catch
            {
                // If logging setup fails, continue without it.
            }
        }

        /// <summary>
        /// Disables logging.
        /// </summary>
        public static void DisableLogging()
        {
            _listener?.Dispose();
            _listener = null;
        }

        /// <summary>
        /// Handles the AssemblyResolve event to redirect to the version of the assembly in the bin folder.
        /// </summary>
        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var requestedAssembly = new AssemblyName(args.Name);

            // List of assemblies that are known to cause binding redirect issues with the Azure SDK.
            string[] assembliesToRedirect = {
                "Azure.Identity", "System.Deployment","System.Runtime.Serialization.Formatters.Soap","System.Xml","System.Configuration","Accessibility","Azure.Core","System.Threading.Tasks.Extensions", "System.Core","System.ValueTuple","System.Text.Json","System.Text.Encodings.Web","System.Security.Principal.Windows","System.Security.Cryptography.ProtectedData","System.Security.AccessControl","System.Runtime.CompilerServices.Unsafe","System.Numerics.Vectors","System.Net.Http.WinHttpHandler","System.Memory","System.Memory.Data","System.IO.FileSystem.AccessControl","System.IdentityModel.Tokens.Jwt","System.Diagnostics.DiagnosticSource","System.ClientModel","System.Buffers","Std.UriTemplate","Microsoft.Kiota.Serialization.Text","Microsoft.Kiota.Serialization.Multipart","Microsoft.Kiota.Serialization.Json","Microsoft.Kiota.Serialization.Form","Microsoft.Kiota.Http.HttpClientLibrary","Microsoft.Kiota.Authentication.Azure","Microsoft.Kiota.Abstractions","Microsoft.IdentityModel.Validators","Microsoft.IdentityModel.Tokens","Microsoft.IdentityModel.Protocols.OpenIdConnect","Microsoft.IdentityModel.Protocols","Microsoft.IdentityModel.Logging","Microsoft.IdentityModel.JsonWebTokens","Microsoft.IdentityModel.Abstractions","Microsoft.Identity.Client.Extensions.Msal","Microsoft.Identity.Client.Extensions.Msal","Microsoft.Identity.Client","Microsoft.Graph","Microsoft.Graph.Core","Microsoft.Extensions.Logging.Abstractions","Microsoft.Extensions.DependencyInjection.Abstractions","Microsoft.Bcl.TimeProvider","Microsoft.Bcl.AsyncInterfaces"

            };

            foreach (var assemblyName in assembliesToRedirect)
            {
                if (requestedAssembly.Name == assemblyName)
                {
                    try
                    {
                        // Load the assembly from the application's base directory.
                        // This loads whatever version is available, effectively redirecting the binding.
                        return Assembly.Load(requestedAssembly.Name);
                    }
                    catch
                    {
                        // If it fails, return null to allow the default load behavior to continue.
                        return null;
                    }
                }
            }

            return null;
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="GraphAuth"/> class.
        /// </summary>
        /// <param name="authData">The authentication data required for Graph API.</param>
        public GraphAuth(AuthenticationData authData)
        {
            if (authData == null)
            {
                throw new ArgumentNullException(nameof(authData));
            }
            _authData = authData;
        }

        /// <summary>
        /// Acquires a fresh access token for the Microsoft Graph API and returns the raw JWT string.
        /// Always fetches a new token, bypassing the MSAL cache, to reflect the current Azure AD state.
        /// </summary>
        /// <returns>The raw JWT access token string.</returns>
        public async Task<string> GetTokenAsync()
        {
            var credentials = new ClientSecretCredential(
                _authData.TenantId,
                _authData.ClientId,
                _authData.ClientSecret);

            var tokenResult = await credentials.GetTokenAsync(
                new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }))
                .ConfigureAwait(false);

            return tokenResult.Token;
        }

        /// <summary>
        /// Acquires a fresh access token for the Microsoft Graph API and returns the raw JWT string.
        /// Always fetches a new token, bypassing the MSAL cache, to reflect the current Azure AD state.
        /// </summary>
        /// <returns>The raw JWT access token string.</returns>
        public string GetToken()
        {
            return GetTokenAsync().GetAwaiter().GetResult();
        }

        /// <summary>
        /// Creates and returns an authenticated GraphServiceClient.
        /// </summary>
        /// <returns>An authenticated <see cref="GraphServiceClient"/>.</returns>
        public GraphServiceClient GetAuthenticatedGraphClient()
        {
            // Create a unique key for the credentials to reuse the client
            string key = $"{_authData.TenantId}:{_authData.ClientId}:{_authData.ClientSecret}";

            return _clientCache.GetOrAdd(key, k =>
            {
                var credentials = new ClientSecretCredential(
                    _authData.TenantId,
                    _authData.ClientId,
                    _authData.ClientSecret);

                return new GraphServiceClient(credentials);
            });
        }
    }
}