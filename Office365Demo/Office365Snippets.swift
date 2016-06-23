//
//  Office365Snippets.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/23/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import Foundation

/***************************************************************************************
 These helpers are used to pupulate new Outlook item objects and are used while creating
 new items on the server
 ***************************************************************************************/
class Office365Snippets{
    
    let clientFetcher : Office365ClientFetcher
    
    init () {
        
        clientFetcher = Office365ClientFetcher()
    }
    
    //Populates a new email message item
    func outlookMessageWithProperties(recipients: NSArray,
                                      subject: String,
                                      body: String) -> MSOutlookMessage {
        
        let message : MSOutlookMessage = MSOutlookMessage()
        message.Subject = subject
        
        message.Body = MSOutlookItemBody()
        message.Body.Content = body
        message.Body.ContentType = MSOutlookBodyType.BodyType_Text
        
        let toRecipients : NSMutableArray = NSMutableArray()
        
        for emailAdress in recipients {
            
            let recipient: MSOutlookRecipient = MSOutlookRecipient()
            
            recipient.EmailAddress = MSOutlookEmailAddress()
            recipient.EmailAddress.Address = emailAdress as! String
            
            toRecipients.addObject(recipient)
        }
        
        message.ToRecipients = toRecipients
        
        return message
    }
    
    //Populates a new calendar event item
    func outlookEventWithProperties(attendees: NSArray,
                                    subject: String,
                                    body: String,
                                    start: NSDate,
                                    end: NSDate) -> MSOutlookEvent {
        
        let event: MSOutlookEvent = MSOutlookEvent()
        
        event.Subject = subject
        event.Start = start
        event.End = end
        event.Type = MSOutlookEventType.EventType_SingleInstance
        
        event.Body = MSOutlookItemBody()
        event.Body.Content = body
        event.Body.ContentType = MSOutlookBodyType.BodyType_Text
        
        let toAttendees : NSMutableArray = NSMutableArray()
        
        for emailAddress in attendees {
            
            let attendee: MSOutlookAttendee = MSOutlookAttendee()
            
            attendee.EmailAddress = MSOutlookEmailAddress()
            attendee.EmailAddress.Address = emailAddress as! String
            
            toAttendees.addObject(attendee)
        }
        
        event.Attendees = toAttendees
        
        return event
    }
    
    //Populates a new contact
    func outlookContactWithProperties(emailAddresses: NSArray,
                                      givenName : String,
                                      displayName: String,
                                      surname: String,
                                      title: String,
                                      mobilePhone1: String
        ) -> MSOutlookContact {
        
        let contact : MSOutlookContact = MSOutlookContact()
        
        contact.GivenName = givenName
        contact.Surname = surname
        contact.DisplayName = displayName
        contact.Title = title
        contact.MobilePhone1 = mobilePhone1
        
        let contactEmailAddresses: NSMutableArray = NSMutableArray()
        
        for emailAddress in emailAddresses {
            
            let email: MSOutlookEmailAddress = MSOutlookEmailAddress()
            email.Address = emailAddress as! String
            
            contactEmailAddresses.addObject(email)
        }
        
        contact.EmailAddresses = contactEmailAddresses
        
        return contact
    }
    
    //Get the 10 most recent email messages in the user's inbox
    func fetchMailMessages(completionHandler:((NSArray, NSError) -> Void)) {
    
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            // Retrieve mail messages from O365 and pass the status to the callback. Uses a default page size of 10
            // This results in a call to the service
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
           
            let task = messageCollectionFetcher.readWithCallback{(messages:[AnyObject]!, error:MSODataException!) -> Void in
                completionHandler(messages, error)
            }
            
            task.resume()
        }
    }
    
    //Sends a new email message to the user
    func sendMailMessage(message: MSOutlookMessage, completionHandler:((Bool, NSError) -> Void)) {
    
        // Get the client and get the operations for sending a mail
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            let userFetcher = outlookClient.getMe()
            let userOperations = (userFetcher.operations as MSOutlookUserOperations)
            
            // The returnValue is the HTTP status code
            let task = userOperations.sendMailWithMessage(message, saveToSentItems: true) {
                (returnValue:Int32, error:MSODataException!) -> Void in
                
                let sucess: Bool = (returnValue == 0)
                completionHandler(sucess, error)
            }
            
            task.resume()
        }
    }
    
    //Creates a new email message in the user's Drafts folder
    //Does not send the email
    func createDraftMailMessage(message: MSOutlookMessage, completionHandler:((MSOutlookMessage, NSError) -> Void )){
        
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            
            let task = messageCollectionFetcher.addMessage(message){ (addedMessage: MSOutlookMessage!, error: MSODataException!) -> Void in
                completionHandler(addedMessage, error)
            }
            
            task.resume()
        }
    }
    
    // Creates and sends an email message to a single recipient, with a subject, an HTML body and save a copy in the sender's
    // sentitems folder
    func createAndSendHTMLMailMessage(toRecipients: NSMutableArray, completionHandler:((Bool, NSError) -> Void)) {
        
        // Get the client and get the operations for sending a mail
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let userOperations = (userFetcher.operations as MSOutlookUserOperations)
        
            // Form the email message
            let message: MSOutlookMessage = MSOutlookMessage()
            
            message.Subject = "Here's the subject"
            
            message.Body = MSOutlookItemBody()
            message.Body.Content = "<!DOCTYPE html><html><body>Here's the body.</body></html>"
            message.Body.ContentType = MSOutlookBodyType.BodyType_HTML
            message.ToRecipients = toRecipients
            
            // The returnValue is the HTTP status code
            let task = userOperations.sendMailWithMessage(message, saveToSentItems: true) {
                (returnValue:Int32, error:MSODataException!) -> Void in
                
                let sucess: Bool = (returnValue == 0)
                completionHandler(sucess, error)
            }
            
            task.resume()
            
        }
    }
    
    //Updates an email message on the server
    func updateMailMessage(message: MSOutlookMessage, completionHandler:((MSOutlookMessage, NSError) -> Void )){
        
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.Id)
            
            let task = messageFetcher.updateMessage(message){ (updatedMessage: MSOutlookMessage!, error: MSODataException!) -> Void in
                completionHandler(updatedMessage, error)
            }
            
            task.resume()

        }
    }
    
    //Deletes an email message from the server
    func deleteMailMessage(message: MSOutlookMessage, completionHandler:((Bool, NSError) -> Void)) {
        
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.Id)
            
            let task = messageFetcher.deleteMessage { (status: Int32, error: MSODataException!) -> Void in
            
                let success: Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task.resume()
        
        }
        
    }
    
    //Replies to a single recipient in a mail message
    func replyToMailMessage(message: MSOutlookMessage, completionHandler:((Int32, NSError) -> Void)) {
    
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.Id)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            //TO DO: a reply all (multiple recipients) you can replace replyWithComment to replyAllWithComment
            let task = messageOperations.replyWithComment("Testing reply snippet"){ (returnValue:Int32, error:MSODataException!) -> Void in
                
                completionHandler(returnValue, error)
            }
            
            task.resume()
        }
    }
    
    
    //Create a draft reply email message in inbox
    func createDraftReplyMessage(message: MSOutlookMessage, completionHandler:((MSOutlookMessage, NSError) -> Void )){
    
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.Id)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.createReplyWithCallback(){ (replyMessage:MSOutlookMessage!, error:MSODataException!) -> Void in
                
                completionHandler(replyMessage, error)
            }
            
            task.resume()
        }
    }
    
    /**
     *  Copy a mail message to the deleted items folder.
     *  https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations#MoveorcopymessagesCopyamessageREST
     *
     *  @param itemId     Identifier of the message that will be copied
     *  @param completion
     */
    func copyMessage(messageId: String, completionHandler:((Bool, MSODataException) -> Void)) {
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(messageId)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.copyWithDestinationId("DeletedItems"){ (msg:MSOutlookMessage!, error:MSODataException!) -> Void in
            
                // You now have the copied MSOutlookMessage named msg.
                let success: Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task.resume()
        }
    }
    
    /**
     *  Move a mail message to the deleted items folder.
     *  https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations#MoveorcopymessagesCopyamessageREST
     *
     *  @param itemId     Identifier of the message that will be moved
     *  @param completion
     */
    func moveMessage(messageId: String, completionHandler:((Bool, MSODataException) -> Void)) {
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(messageId)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.moveWithDestinationId("DeletedItems"){ (msg:MSOutlookMessage!, error:MSODataException!) -> Void in
                
                // You now have the copied MSOutlookMessage named msg.
                let success: Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task.resume()
        }
    }
    
    /**
     *  Fetch up to the first 10 unread messages in your inbox that have been marked as important.
     *  https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations#MoveorcopymessagesCopyamessageREST
     *
     *  @param completion
     */
    func fetchUnreadImportantMessages(completionHandler:((NSArray, NSError) -> Void)){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            
            // Identify which properties to return. Only request properties that you will use.
            // The identifier is always returned.
            // https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesMessage
            messageCollectionFetcher.select("Subject,Importance,IsRead,Sender,DateTimeReceived")
            
            // Search for items that are both unread and marked as important. The library will URL encode this for you.
            // https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters
            //[messageCollectionFetcher addCustomParametersWithName:@"filter" value:@"IsRead eq true AND Importance eq 'High'"];
            messageCollectionFetcher.filter("IsRead eq true AND Importance eq 'High'")
            
            let task = messageCollectionFetcher.readWithCallback(){ (messages: [AnyObject]!, error: MSODataException!) -> Void in
                
                if error == nil {
                    completionHandler(messages, error)
                }else{
                    completionHandler([], error)
                }
            }
            
            task.resume()
            
        }
    }
    
    
    /**
     *  Get the weblink to the first item in your inbox. This example assumes you have at least one item in your inbox.
     *  https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesMessage
     *
     *  @param completion
     */
    func fetchMessagesWebLink(completionHandler:((String, NSError) -> Void)){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
        
            // Identify which properties to return. Only request properties that you will use.
            // The identifier is always returned.
            // https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesMessage
            messageCollectionFetcher.select("Subject,WebLink")
            
            // Identify how many items to return in the page.
            messageCollectionFetcher.top(1)
            
            let task = messageCollectionFetcher.readWithCallback(){ (messages: [AnyObject]!, error: MSODataException!) -> Void in
            
                //You now have an NSArray of MSOutlookMessage objects. Let's get the first and ony object
                let message: MSOutlookMessage = messages.first as! MSOutlookMessage
                
                if error == nil {
                    completionHandler(message.WebLink, error)
                }else{
                    completionHandler("", error)
                }
            }
            
            task.resume()
        
        }
    }
    
}