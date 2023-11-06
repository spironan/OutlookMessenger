package outlook

import (
	"context"
	"fmt"
	"log"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	auth "github.com/microsoft/kiota-authentication-azure-go"
	msgraphsdk "github.com/microsoftgraph/msgraph-sdk-go"
	"github.com/microsoftgraph/msgraph-sdk-go/models"
	"github.com/microsoftgraph/msgraph-sdk-go/users"
)

type BodyType int

const (
	TEXT_BODYTYPE BodyType = iota
	HTML_BODYTYPE
)

type OutlookEmail struct {
	Sender     string
	Subject    string
	MailType   BodyType
	Body       string
	Recipients []string
}

// Config holds the configuration for OutlookMessenger
type Config struct {
	ClientID     string
	TenantID     string
	ClientSecret string
}

type OutlookMessenger struct {
	clientSecretCredential *azidentity.ClientSecretCredential
	appClient              *msgraphsdk.GraphServiceClient
}

func NewOutlookMessenger(config Config) *OutlookMessenger {
	g := &OutlookMessenger{}
	initializeOutlookGraph(g, config)
	return g
}

func initializeOutlookGraph(outlookMessenger *OutlookMessenger, config Config) {
	err := outlookMessenger.initializeGraphForAppAuth(config)
	if err != nil {
		log.Panicf("Error initializing Graph for user auth: %v\n", err)
	}
}

func DisplayAccessToken(outlookMessenger *OutlookMessenger) {
	token, err := outlookMessenger.GetAppToken()
	if err != nil {
		log.Panicf("Error getting user token: %v\n", err)
	}

	fmt.Printf("App-only token: %s\n", *token)
	fmt.Println()
}

func (g *OutlookMessenger) initializeGraphForAppAuth(config Config) error {
	clientId := config.ClientID
	tenantId := config.TenantID
	clientSecret := config.ClientSecret

	// fmt.Println("clientId:", clientId)
	// fmt.Println("tenantId:", tenantId)
	// fmt.Println("clientSecret:", clientSecret)

	credential, err := azidentity.NewClientSecretCredential(tenantId, clientId, clientSecret, nil)
	if err != nil {
		return err
	}

	g.clientSecretCredential = credential

	// Create an auth provider using the credential
	authProvider, err := auth.NewAzureIdentityAuthenticationProviderWithScopes(g.clientSecretCredential, []string{
		"https://graph.microsoft.com/.default",
	})
	if err != nil {
		return err
	}

	// Create a request adapter using the auth provider
	adapter, err := msgraphsdk.NewGraphRequestAdapter(authProvider)
	if err != nil {
		return err
	}

	// Create a Graph client using request adapter
	client := msgraphsdk.NewGraphServiceClient(adapter)
	g.appClient = client

	return nil
}

func (g *OutlookMessenger) GetAppToken() (*string, error) {
	token, err := g.clientSecretCredential.GetToken(context.Background(), policy.TokenRequestOptions{
		Scopes: []string{
			"https://graph.microsoft.com/.default",
		},
	})
	if err != nil {
		return nil, err
	}

	return &token.Token, nil
}

func (g *OutlookMessenger) SendMail(mail OutlookEmail) error {
	// set up variables
	sender := mail.Sender
	subject := mail.Subject
	bodyType := mail.MailType
	body := mail.Body
	recipients := mail.Recipients

	// Create a new message
	message := models.NewMessage()
	message.SetSubject(&subject)

	messageBody := models.NewItemBody()
	messageBody.SetContent(&body)
	var contentType models.BodyType
	switch bodyType {
	case HTML_BODYTYPE:
		contentType = models.HTML_BODYTYPE
	case TEXT_BODYTYPE:
		contentType = models.TEXT_BODYTYPE
	}
	messageBody.SetContentType(&contentType)
	message.SetBody(messageBody)

	toRecipients := []models.Recipientable{}

	for i, recipientAddress := range recipients {
		recipient := models.NewRecipient()
		emailAddress := models.NewEmailAddress()
		temp := recipientAddress
		emailAddress.SetAddress(&temp)
		recipient.SetEmailAddress(emailAddress)
		fmt.Printf("Recipients [%d]: <%s>\n", i, recipientAddress)
		toRecipients = append(toRecipients, recipient)
	}
	message.SetToRecipients(toRecipients)

	// for i, receipient := range toRecipients {
	// 	fmt.Printf("Recipient [%d]: <%s>\n", i, *receipient.GetEmailAddress().GetAddress())
	// }

	sendMailBody := users.NewItemSendMailPostRequestBody()
	sendMailBody.SetMessage(message)

	// Send the message
	return g.appClient.Users().ByUserId(sender).SendMail().Post(context.Background(), sendMailBody, nil)
}
