using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Identity.Client;

class GraphHelper
{
// --------initial config--------
// Settings object
private static Settings? _settings;
// User auth token credential
private static DeviceCodeCredential? _deviceCodeCredential;
// Client configured with user authentication
private static GraphServiceClient? _userClient;

public static void InitializeGraphForUserAuth(Settings settings,
    Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
{
    _settings = settings;

    _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
        settings.TenantId, settings.ClientId);

    _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
}


// ------------greet the user once authenticated------------
public static async Task<string> GetUserTokenAsync()
{
    // Ensure credential isn't null
    _ = _deviceCodeCredential ??
        throw new System.NullReferenceException("Graph has not been initialized for user auth");

    // Ensure scopes isn't null
    _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

    // Request token with given scopes
    var context = new TokenRequestContext(_settings.GraphUserScopes);
    var response = await _deviceCodeCredential.GetTokenAsync(context);
    return response.Token;
}


// ------get the user---------
public static Task<User> GetUserAsync()
{
    // Ensure client isn't null
    _ = _userClient ??
        throw new System.NullReferenceException("Graph has not been initialized for user auth");

    return _userClient.Me
        .Request()
        .Select(u => new
        {
            // Only request specific properties
            u.DisplayName,
            u.Mail,
            u.UserPrincipalName
        })
        .GetAsync();
}

// ---------------list user inbox -------
public static Task<IMailFolderMessagesCollectionPage> GetInboxAsync()
{
    // Ensure client isn't null
    _ = _userClient ??
        throw new System.NullReferenceException("Graph has not been initialized for user auth");

    return _userClient.Me
        // Only messages from Inbox folder
        .MailFolders["Inbox"]
        .Messages
        .Request()
        .Select(m => new
        {
            // Only request specific properties
            m.From,
            m.IsRead,
            m.ReceivedDateTime,
            m.Subject
        })
        // Get at most 25 results
        .Top(25)
        // Sort by received time, newest first
        .OrderBy("ReceivedDateTime DESC")
        .GetAsync();
}

// ---send mail--------
public static async Task SendMailAsync(string subject, string body, string recipient)
{
    // Ensure client isn't null
    _ = _userClient ??
        throw new System.NullReferenceException("Graph has not been initialized for user auth");

    // Create a new message
    var message = new Message
    {
        Subject = subject,
        Body = new ItemBody
        {
            Content = body,
            ContentType = BodyType.Text
        },
        ToRecipients = new Recipient[]
        {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }
        }
    };

    // Send the message
    await _userClient.Me
        .SendMail(message)
        .Request()
        .PostAsync();
}


// This function serves as a playground for testing Graph snippets
// or other code
public async static Task MakeGraphCallAsync()
{
    // INSERT YOUR CODE HERE
}

}