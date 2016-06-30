//
//  Office365Manager.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/29/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import Foundation

class Office365Manager {
    
    let clientFetcher : Office365ClientFetcher
    var lastrefreshdate: NSDate!
    var allMessages: [MSOutlookMessage] = [MSOutlookMessage]()
    var allConversations: [Conversation] = [Conversation]()
    
    
    init () {
        
        clientFetcher = Office365ClientFetcher()
    }
    
    func getConversationsFromMessages(messages: [MSOutlookMessage]) -> [Conversation] {
        
        let messagesByConversationID : NSMutableDictionary = [:]
        //let filteredMessages: NSArray = messages.filteredArrayUsingPredicate(NSPredicate(format: "isHidden == false"))
        
        for message in messages {
            var messagQue = messagesByConversationID[message.ConversationId]
            if(messagQue == nil){
                messagQue = NSMutableArray()
                messagesByConversationID[message.ConversationId] = messagQue
            }
            messagQue?.addObject(message)
        }
     
        
        let conversations: NSMutableArray = NSMutableArray()
        
        for value  in messagesByConversationID.allValues  {
            let messages: NSMutableArray = value as! NSMutableArray
            let conversation : Conversation = Conversation(messages: messages)
            conversations.addObject(conversation)
        }
    
        return conversations.sortedArrayUsingSelector(#selector(NSNumber.compare(_:))) as! [Conversation]
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
    
    
    
    /***************************************************  MAIL SNIPPETS **************************************************/
    
    //Get the 10 most recent email messages in the user's inbox
    func fetchMailMessages(completionHandler:((NSArray, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            // Retrieve mail messages from O365 and pass the status to the callback. Uses a default page size of 10
            // This results in a call to the service
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            messageCollectionFetcher.top(100)
            messageCollectionFetcher.orderBy("DateTimeReceived desc")
            
            let task = messageCollectionFetcher.readWithCallback{(messages:[AnyObject]!, error:MSODataException!) -> Void in
                self.lastrefreshdate = NSDate()
                self.allMessages = messages as! [MSOutlookMessage]
                self.allConversations = self.getConversationsFromMessages(self.allMessages)
//                print(self.allMessages)
//                print(self.allConversations)
                
                completionHandler(messages, error)
            }
            
            task.resume()
        }
    }
    
    //fetch email messages based on pagenumber
    func fetchMailMessagesForPageNumber(pageNumber: Int32, pageSize: Int32, orderBy: String, completionHandler:((NSArray, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            messageCollectionFetcher.orderBy(orderBy)
            messageCollectionFetcher.top(pageSize)
            messageCollectionFetcher.skip(pageNumber * pageSize)
            
            //retrieve messages
            let task = messageCollectionFetcher.readWithCallback{(messages:[AnyObject]!, error:MSODataException!) -> Void in
                self.lastrefreshdate = NSDate()
                
                //if more pages are called, then append them
                if(pageNumber > 0){
                    
                    //append additional messages
                    let additionalMessages: [MSOutlookMessage] = messages as! [MSOutlookMessage]
                    for additionalMessage in additionalMessages {
                        self.allMessages.append(additionalMessage)
                    }
                    
                    //append additional conversations
                    let additionalConversations = self.getConversationsFromMessages(additionalMessages)
                    for additionalConversation in additionalConversations {
                        self.allConversations.append(additionalConversation)
                    }
                    
                }else{//otherwise initialize the first page (pageNumber=0)
                    self.allMessages = messages as! [MSOutlookMessage]
                    self.allConversations = self.getConversationsFromMessages(self.allMessages)
                }
                
                completionHandler(messages, error)
            }
            
            task.resume()
        }
        
    }
    
    //Sends a new email message to the user
    func sendMailMessage(message: MSOutlookMessage, completionHandler:((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            let userFetcher = outlookClient.getMe()
            let userOperations = (userFetcher.operations as MSOutlookUserOperations)
            
            // The returnValue is the HTTP status code
            let task = userOperations.sendMailWithMessage(message, saveToSentItems: true) {
                (returnValue: Int32, error: MSODataException!) -> Void in
                
                let success: Bool = (returnValue == 0)
                completionHandler(success, error)
            }
            
            task.resume()
        }
    }
    
    //Creates a new email message in the user's Drafts folder
    //Does not send the email
    func createDraftMailMessage(message: MSOutlookMessage, completionHandler:((MSOutlookMessage, MSODataException?) -> Void )){
        
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
    func createAndSendHTMLMailMessage(toRecipients: NSMutableArray, completionHandler:((Bool, MSODataException?) -> Void)) {
        
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
    func updateMailMessage(message: MSOutlookMessage, completionHandler:((MSOutlookMessage, MSODataException?) -> Void )){
        
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
    func deleteMailMessage(message: MSOutlookMessage, completionHandler:((Bool, MSODataException?) -> Void)) {
        
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
    func replyToMailMessage(message: MSOutlookMessage, completionHandler:((Int32, MSODataException?) -> Void)) {
        
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
    func createDraftReplyMessage(message: MSOutlookMessage, completionHandler:((MSOutlookMessage, MSODataException?) -> Void )){
        
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
    func copyMessage(messageId: String, completionHandler:((Bool, MSODataException?) -> Void)) {
        
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
    func moveMessage(messageId: String, completionHandler:((Bool, MSODataException?) -> Void)) {
        
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
    func fetchUnreadImportantMessages(completionHandler:((NSArray, MSODataException?) -> Void)){
        
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
    func fetchMessagesWebLink(completionHandler:((String, MSODataException?) -> Void)){
        
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
                let weblink = (error == nil) ? message.WebLink : ""
                completionHandler(weblink, error)
            }
            
            task.resume()
            
        }
    }
    
    //Send a draft message
    func sendDraftMessage(messageId: String, completionHandler:((MSOutlookMessage, MSODataException?) -> Void)){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(messageId)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            
        }
        
    }
    
    //Add an attachment to a message
    func addAttachment(message: MSOutlookMessage, contentType: String, contentBytes: NSData, completionHandler:((MSOutlookMessage, MSODataException?) -> Void)){
    }
    
    //Send mail with an attachment
    func sendMailMessage(message: MSOutlookMessage,
                         AttachmentContentType: String,
                         contentBytes: NSData, completionHandler:((Bool, MSODataException?) -> Void)){
    }
    
    /************************************************* END:  MAIL SNIPPETS ************************************************/
    
    
    
    
    
    /****************************************************  CALENDAR **************************************************/
    
    //Gets the 10 most recent calendar events from the user's calendar
    func fetchCalendarEvents(completionHandler:((NSArray, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            // Retrieve events from O365 and pass the status to the callback. Uses a default page size of 10.
            // This results in a call to the service.
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher.getEvents()
            
            let task = eventCollectionFetcher.readWithCallback{(events:[AnyObject]!, error:MSODataException!) -> Void in
                completionHandler(events, error)
            }
            
            task.resume()
        }
    }
    
    //Creates a new event in the user's calendar
    func createCalendarEvent(event: MSOutlookEvent, completionHandler:((MSOutlookEvent, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher.getEvents()
            
            let task = eventCollectionFetcher.addEvent(event){ (addedEvent: MSOutlookEvent!, error: MSODataException!) -> Void in
                completionHandler(addedEvent, error)
            }
            
            task.resume()
        }
    }
    
    //Updates an event in the user's calendar
    func updateCalendarEvent(event: MSOutlookEvent, completionHandler:((MSOutlookEvent, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher.getEvents()
            let eventFetcher : MSOutlookEventFetcher = eventCollectionFetcher.getById(event.Id)
            
            let task = eventFetcher.updateEvent(event){ (updatedEvent: MSOutlookEvent!, error: MSODataException!) -> Void in
                completionHandler(updatedEvent, error)
            }
            
            task.resume()
        }
    }
    
    //Deletes an event from ther user's calendar
    func deleteCalendarEvent(event: MSOutlookEvent, completionHandler:((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher.getEvents()
            let eventFetcher : MSOutlookEventFetcher = eventCollectionFetcher.getById(event.Id)
            
            let task = eventFetcher.deleteEvent{(status: Int32, error: MSODataException!) -> Void in
                
                let success : Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task.resume()
            
        }
    }
    
    //Accepts an event with comment - comment can be nil
    func acceptCalendarMeetingEvent(event: MSOutlookEvent, comment: String, completionHandler:((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventFetcher : MSOutlookEventFetcher = userFetcher.getEventsById(event.Id)
            let operations: MSOutlookEventOperations = (eventFetcher.operations as MSOutlookEventOperations)
            
            let task = operations.acceptWithComment(comment){(returnValue: Int32, error: MSODataException!) -> Void in
                
                let success: Bool = (returnValue == 0)
                completionHandler(success, error)
            }
            
            task.resume()
        }
        
    }
    
    //Declines an event with a comment - comment can be nil
    func declineCalendarMeetingEvent(event: MSOutlookEvent, comment: String, completionHandler:((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventFetcher : MSOutlookEventFetcher = userFetcher.getEventsById(event.Id)
            let operations: MSOutlookEventOperations = (eventFetcher.operations as MSOutlookEventOperations)
            
            let task = operations.declineWithComment(comment){(returnValue: Int32, error: MSODataException!) -> Void in
                
                let success: Bool = (returnValue == 0)
                completionHandler(success, error)
            }
            
            task.resume()
        }
    }
    
    //Tentatively accepts an event with comment - comment can be nil
    func tentativelyAcceptCalendarMeetingEvent(event: MSOutlookEvent, comment: String, completionHandler:((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventFetcher : MSOutlookEventFetcher = userFetcher.getEventsById(event.Id)
            let operations: MSOutlookEventOperations = (eventFetcher.operations as MSOutlookEventOperations)
            
            let task = operations.tentativelyAcceptWithComment(comment){(returnValue: Int32, error: MSODataException!) -> Void in
                
                let success: Bool = (returnValue == 0)
                completionHandler(success, error)
            }
            
            task.resume()
        }
    }
    
    // Fetches the first 10 event instances in the specified date range
    // For more information about calendar view, visit https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations#GetCalendarView
    func fetchCalendarViewFrom(start: NSDate, end: NSDate, completionHandler:((NSArray, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher.getCalendarView()
            eventCollectionFetcher.select("Subject")
            
            let numDaysToCheck : Double = 10
            let secondsBackToConsider: Double = numDaysToCheck * 60 * 60 * 24
            
            eventCollectionFetcher.addCustomParametersWithName("startDateTime", value: NSDate(timeIntervalSinceNow: -secondsBackToConsider))
            eventCollectionFetcher.addCustomParametersWithName("endDateTime", value: NSDate())
            
            let task = eventCollectionFetcher.readWithCallback{(events:[AnyObject]!, error:MSODataException!) -> Void in
                
                completionHandler(events, error)
            }
            
            task.resume()
            
        }
    }
    
    
    /*************************************************** END CALENDAR **************************************************/
    
    
    
    
    /*************************************************** CONTACTS **************************************************/
    
    //Gets the 10 most recently added user's contacts from Office 365
    func fetchContacts(completionHandler:((NSArray, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactFetcher = userFetcher.getContacts()
            
            let task = contactFetcher.readWithCallback{(contacts:[AnyObject]!, error:MSODataException!) -> Void in
                
                completionHandler(contacts, error)
            }
            
            task.resume()
        }
    }
    
    //Creates a new contact for the user
    func createContact(contact: MSOutlookContact, completionHandler:((MSOutlookContact, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactCollectionFetcher: MSOutlookContactCollectionFetcher = userFetcher.getContacts()
            
            let task = contactCollectionFetcher.addContact(contact){(addedContact: MSOutlookContact!, error: MSODataException!) in
                completionHandler(addedContact, error)
            }
            
            task.resume()
        }
    }
    
    //Updates a contact in Office 365
    func updateContact(contact: MSOutlookContact, completionHandler:((MSOutlookContact, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactCollectionFetcher: MSOutlookContactCollectionFetcher = userFetcher.getContacts()
            let contactFetcher: MSOutlookContactFetcher = contactCollectionFetcher.getById(contact.Id)
            
            let task = contactFetcher.updateContact(contact){(updatedContact: MSOutlookContact!, error: MSODataException!) in
                completionHandler(updatedContact, error)
            }
            
            task.resume()
        }
    }
    
    //Deletes a contact from Office 365
    func deleteContact(contact: MSOutlookContact, completionHandler:((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactCollectionFetcher: MSOutlookContactCollectionFetcher = userFetcher.getContacts()
            let contactFetcher: MSOutlookContactFetcher = contactCollectionFetcher.getById(contact.Id)
            
            let task = contactFetcher.deleteContact({ (status: Int32, error: MSODataException!) in
                
                let success: Bool = (error == nil)
                completionHandler(success, error)
            })
            
            task.resume()
            
        }
    }
    
    
    /*************************************************** END CONTACTS **************************************************/
    
    
    
    
    /*************************************************** FILES *****************************************************/
    
    //Get 10 files or folders from the user's OneDrive for Business folder
    func fetchFiles(completionHandler:((NSArray, MSODataException?) -> Void)) {
        
        // Get the SharePoint client. This object contains access tokens and methods to call the service.
        clientFetcher.fetchSharePointClient{ (sharePointClient) -> Void in
            
            let fileFetcher: MSSharePointItemCollectionFetcher = sharePointClient.getfiles()
            
            // Retrieve files from O365 and pass the status to the callback. Uses a default page size of 10
            let task = fileFetcher.readWithCallback({ (files: [AnyObject]!, error: MSODataException!) in
                completionHandler(files, error)
            })
            
            task.resume()
            
        }
    }
    
    
    /*************************************************** END FILES **************************************************/
    
}