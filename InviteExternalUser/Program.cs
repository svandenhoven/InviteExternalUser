using Azure.Identity;
using Microsoft.Graph;

var scopes = new[] { "https://graph.microsoft.com/.default" };
var tenantId = "your own";
var clientId = "your own";
var clientSecret = "your own";
var userId = "your own";
var invitedUserDisplayName = "your own";
var invitedUserEmail = "your own";
var senderName = "your own";
var senderEmail = "your own";


var options = new TokenCredentialOptions
{
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
};

var clientSecretCredential = new ClientSecretCredential(
    tenantId, clientId, clientSecret, options);

var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

var invitation = new Invitation
{
	InvitedUserDisplayName = invitedUserDisplayName,
	InvitedUserEmailAddress = invitedUserEmail,
	SendInvitationMessage = false, //default is false but assigned for clarity
	InviteRedirectUrl = "https://myapp.contoso.com",

};

var result = await graphClient.Invitations
	.Request()
	.AddAsync(invitation);

Console.WriteLine($"user is invited and this is redeem url {result.InviteRedeemUrl}");

var message = new Message
{
	Sender = new Recipient { EmailAddress = new EmailAddress { Address = senderEmail, Name = senderName } },
	Subject = "Your are invited for Contoso Community, please confirm your email address as account for Contoso Community.",
	Body = new ItemBody
	{
		ContentType = BodyType.Html,
		Content = "<html><body>" +
				  $"<img src=''/><p><H1>Confirm your email address {result.InvitedUserEmailAddress}</H1>" +
				  "<p>This is email from Contoso. You are eligiable to to access Community. Please confirm your email account with below link</p>" +
				  $"<a href='{result.InviteRedeemUrl}'>Confirm your emailaddress as account for Community</a>" +
				  "<p>For more information see <a href='https://unit4.com'/>" +
				  "</body>" +
				  "</html>"
	},
	ToRecipients = new List<Recipient>()
	{
		new Recipient
		{
			EmailAddress = new EmailAddress
			{
				Address = result.InvitedUserEmailAddress
			}
		}
	}
};

var saveToSentItems = true;

await graphClient.Users[userId]
	.SendMail(message, saveToSentItems)
	.Request()
	.PostAsync();

Console.WriteLine($"Invitation Email is send");