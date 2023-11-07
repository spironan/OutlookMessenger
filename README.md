# OutlookMessenger
Simple go package for app-only messaging system

# Quickstart
```go
package main

import (
	"fmt"
	"log"

	outlook "github.com/spironan/outlookmessenger"
)

func main() {
	// set up app-only microsoft azure configuration
	config := outlook.Config{
		ClientID:     "exampleClientID",
		TenantID:     "exampleTenantID",
		ClientSecret: "exampleClientSecret",
	}
	// Initialize outlook messenger
	messenger := NewOutlookMessenger(config)
	// Compose Email
	email := outlook.OutlookEmail{
		Sender:     "sender@outlook.com",
		Subject:    "Title of email",
		MailType:   outlook.TEXT_BODYTYPE,
		Body:       "This is a sample email",
		Recipients: {"recipient1@outlook.com", "recipient2@gmail.com", "recipient3@hotmail.com"},
	}
	// Send Email
	err := messenger.SendMail(email)
	if err != nil {
		log.Panicf(err)
	}

	// Debug log access token
	token, err := outlookMessenger.GetAppToken()
	if err != nil {
		log.Panicf("Error getting user token: %v\n", err)
	}

	fmt.Printf("App-only token: %s\n", *token)
	fmt.Println()
}
```