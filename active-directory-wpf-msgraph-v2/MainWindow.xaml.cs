using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace active_directory_wpf_msgraph_v2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
 
    public partial class MainWindow : Window
    {
        //Set the API Endpoint to Graph 'me' endpoint. 
        // To change from Microsoft public cloud to a national cloud, use another value of graphAPIEndpoint.
        // Reference with Graph endpoints here: https://docs.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints
        string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me/messages?$select=subject&$top=10";

        //Set the scope for API call to user.read
        string[] scopes = new string[] { "user.read", "mail.read" };
        private IAccount currentAccount;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Call AcquireToken - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            ResultText.Text = string.Empty;
            TokenInfoText.Text = string.Empty;

            IAccount firstAccount;

            switch(howToSignIn.SelectedIndex)
            {
                // 0: Use account used to signed-in in Windows (WAM)
                case 0:
                    // WAM will always get an account in the cache. So if we want
                    // to have a chance to select the accounts interactively, we need to
                    // force the non-account
                    firstAccount = PublicClientApplication.OperatingSystemAccount;
                    break;

                //  1: Use one of the Accounts known by Windows(WAM)
                case 1:
                    // We force WAM to display the dialog with the accounts
                    firstAccount = null;
                    break;

                //  Use any account(Azure AD). It's not using WAM
                default:
                    var accounts = await app.GetAccountsAsync();
                    firstAccount = accounts.FirstOrDefault();
                    this.currentAccount = firstAccount;
                    break;
            }

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                Trace.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(firstAccount)
                        .WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    ResultText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {
                ResultText.Text = await GetHttpContentWithToken(graphAPIEndpoint, authResult.AccessToken);
                DisplayBasicTokenInfo(authResult);
                this.SignOutButton.Visibility = Visibility.Visible;
                currentAccount = authResult.Account;
            }
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);

                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized && response.Headers.WwwAuthenticate.Any())
                {
                    var c = await response.Content.ReadAsStringAsync();
                    AuthenticationHeaderValue bearer = response.Headers.WwwAuthenticate.First
                        (v => v.Scheme == "Bearer");
                    IEnumerable<string> parameters = bearer.Parameter.Split(',').Select(v => v.Trim()).ToList();
                    var error = GetParameter(parameters, "error");
                    var headers = String.Join(Environment.NewLine, response.Headers.Select(h => h.Key + ":" + String.Join(",",h.Value.ToArray())).ToArray());
                    TokenInfoText.Text += $"{System.Environment.NewLine}[CAE] headers: {headers}";
                    Trace.WriteLine($"[CAE] headers: {headers}");

                    if (null != error && "insufficient_claims" == error)
                    {
                        TokenInfoText.Text += $"{System.Environment.NewLine}[CAE]Graph access has been blocked CAE";
                        Trace.WriteLine($"[CAE]Graph access has been blocked CAE");
                        var claimChallengeParameter = GetParameter(parameters, "claims");
                        if (null != claimChallengeParameter)
                        {
                            var claimChallengebase64Bytes = System.Convert.FromBase64String(claimChallengeParameter);
                            var claimChallenge = System.Text.Encoding.UTF8.GetString(claimChallengebase64Bytes);
                            var newAccessToken = await GetAccessTokenWithClaimChallenge(scopes, claimChallenge);
                            Trace.WriteLine($"[CAE]token renewal with claimChallenge: {claimChallengeParameter}");

                        }
                        else
                        {
                            //claimChallenge not found
                        }
                    }
                }
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        private async Task<string> GetAccessTokenWithClaimChallenge(string[] scopes, string claimChallenge)
        {
            var app = App.PublicClientApp;
            AuthenticationResult authResult = null;

            ResultText.Text = string.Empty;
            TokenInfoText.Text += "Refresh Token with claimChallenge...";

            //https://docs.microsoft.com/ja-jp/azure/active-directory/develop/app-resilience-continuous-access-evaluation
            try
            {
                authResult = await app.AcquireTokenSilent(scopes, currentAccount)
                                            .WithClaims(claimChallenge)
                                            .ExecuteAsync()
                                            .ConfigureAwait(false);
            }
            catch (MsalUiRequiredException)
            {
                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithClaims(claimChallenge)
                        .WithAccount(currentAccount)
                        .ExecuteAsync()
                        .ConfigureAwait(false);
                }
                catch (MsalException msalex)
                {
                    ResultText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return null;
            }
            return authResult?.AccessToken;
            
        }

        private static string GetParameter(IEnumerable<string> parameters, string paramName)
        {
            return parameters.Select(p => p.Split('=')).Where(p => p[0] == paramName).Select(p => p[1].Trim().Replace("\"", "")).First();
        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    this.ResultText.Text = "User has signed-out";
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// Display basic information contained in the token
        /// </summary>
        private void DisplayBasicTokenInfo(AuthenticationResult authResult)
        {
            TokenInfoText.Text = "";
            if (authResult != null)
            {
                TokenInfoText.Text += $"Username: {authResult.Account.Username}" + Environment.NewLine;
                TokenInfoText.Text += $"Token Expires: {authResult.ExpiresOn.ToLocalTime()}" + Environment.NewLine;
            }
        }

        private void UseWam_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            //SignOutButton_Click(sender, e);
            //App.CreateApplication(howToSignIn.SelectedIndex != 2); // Not Azure AD accounts (that is use WAM accounts)
        }
    }
}
