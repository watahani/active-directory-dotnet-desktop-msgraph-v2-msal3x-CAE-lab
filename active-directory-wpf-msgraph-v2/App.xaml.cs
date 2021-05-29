using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Desktop;
using System;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Windows;

namespace active_directory_wpf_msgraph_v2
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>

    // To change from Microsoft public cloud to a national cloud, use another value of AzureCloudInstance
    public partial class App : Application
    {
        static App()
        {
            //doesn't use wam
            CreateApplication(false);
        }

        public static void CreateApplication(bool useWam)
        {
            var builder = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority($"{Instance}{Tenant}")
                .WithDefaultRedirectUri()
                .WithLogging(Log, LogLevel.Info, true)
                .WithClientCapabilities(new[] { "cp1" });

            if (useWam)
            {
                builder.WithExperimentalFeatures();
                builder.WithWindowsBroker(true);  // Requires redirect URI "ms-appx-web://microsoft.aad.brokerplugin/{client_id}" in app registration
            }
            _clientApp = builder.Build();
            TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);
        }

        // Below are the clientId (Application Id) of your app registration and the tenant information. 
        // You have to replace:
        // - the content of ClientID with the Application Id for your app registration
        // - The content of Tenant by the information about the accounts allowed to sign-in in your application:
        //   - For Work or School account in your org, use your tenant ID, or domain
        //   - for any Work or School accounts, use organizations
        //   - for any Work or School accounts, or Microsoft personal account, use 72aaac3c-90c8-49f1-8c96-6033d3ecd0b5
        //   - for Microsoft Personal account, use consumers
        private static string ClientId = "9ab82d79-c31e-4f9a-8803-087548e387b1";

        // Note: Tenant is important for the quickstart.
        private static string Tenant = "72aaac3c-90c8-49f1-8c96-6033d3ecd0b5";
        private static string Instance = "https://login.microsoftonline.com/";
        private static IPublicClientApplication _clientApp;
        private static void Log(LogLevel level, string message, bool containsPii)
        {
            Trace.WriteLine($"[MSAL]: {level} {message}");
            if (containsPii)
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }
            Console.WriteLine($"{level} {message}");
            Console.ResetColor();
        }
        public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
    }
}
