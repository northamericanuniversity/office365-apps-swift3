//
//  Office365Manager.swift
//  OyventIOSApp
//
//  Created by Mehmet Sen on 9/13/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import Foundation

class Office365Manager {
    
    let clientFetcher : Office365ClientFetcher
    var lastrefreshdate: Date!
    var allMessages: [MSOutlookMessage] = [MSOutlookMessage]()
    var allConversations: [Conversation] = [Conversation]()
    
    
    init () {
        
        clientFetcher = Office365ClientFetcher()
    }
    
    var messagesByConversationID : NSMutableDictionary = [:]
    var conversations: NSMutableArray = NSMutableArray()
    
    func getConversationsFromMessages(_ messages: [MSOutlookMessage]) -> [Conversation] {
        
        for (_,message) in messages.enumerated() {
            var messagQue = messagesByConversationID[message.conversationId!]
            if(messagQue == nil){
                messagQue = NSMutableArray()
                messagesByConversationID[message.conversationId!] = messagQue
            }
            (messagQue as! NSMutableArray).add(message)
        }
        
        for value  in messagesByConversationID.allValues {
            let messages: NSMutableArray = value as! NSMutableArray
            let conversation : Conversation = Conversation(messages: messages)
            conversations.add(conversation)
        }
        
        return conversations.sortedArray(using: #selector(NSNumber.compare(_:))) as! [Conversation]
    }
    
    //Populates a new email message item
    func outlookMessageWithProperties(_ recipients: NSArray,
                                      subject: String,
                                      body: String) -> MSOutlookMessage {
        
        let message : MSOutlookMessage = MSOutlookMessage()
        message.subject = subject
        
        message.body = MSOutlookItemBody()
        message.body.content = body
        message.body.contentType = MSOutlookBodyType.bodyType_Text
        
        let toRecipients : NSMutableArray = NSMutableArray()
        
        for emailAdress in recipients {
            
            let recipient: MSOutlookRecipient = MSOutlookRecipient()
            
            recipient.emailAddress = MSOutlookEmailAddress()
            recipient.emailAddress.address = emailAdress as! String
            
            toRecipients.add(recipient)
        }
        
        message.toRecipients = toRecipients
        
        return message
    }
    
    //Populates a new calendar event item
    func outlookEventWithProperties(_ attendees: NSArray,
                                    subject: String,
                                    body: String,
                                    start: Date,
                                    end: Date) -> MSOutlookEvent {
        
        let event: MSOutlookEvent = MSOutlookEvent()
        
        event.subject = subject
        event.start = start
        event.end = end
        event.type = MSOutlookEventType.eventType_SingleInstance
        
        event.body = MSOutlookItemBody()
        event.body.content = body
        event.body.contentType = MSOutlookBodyType.bodyType_Text
        
        let toAttendees : NSMutableArray = NSMutableArray()
        
        for emailAddress in attendees {
            
            let attendee: MSOutlookAttendee = MSOutlookAttendee()
            
            attendee.emailAddress = MSOutlookEmailAddress()
            attendee.emailAddress.address = emailAddress as! String
            
            toAttendees.add(attendee)
        }
        
        event.attendees = toAttendees
        
        return event
    }
    
    //Populates a new contact
    func outlookContactWithProperties(_ emailAddresses: NSArray,
                                      givenName : String,
                                      displayName: String,
                                      surname: String,
                                      title: String,
                                      mobilePhone1: String
        ) -> MSOutlookContact {
        
        let contact : MSOutlookContact = MSOutlookContact()
        
        contact.givenName = givenName
        contact.surname = surname
        contact.displayName = displayName
        contact.title = title
        contact.mobilePhone1 = mobilePhone1
        
        let contactEmailAddresses: NSMutableArray = NSMutableArray()
        
        for emailAddress in emailAddresses {
            
            let email: MSOutlookEmailAddress = MSOutlookEmailAddress()
            email.address = emailAddress as! String
            
            contactEmailAddresses.add(email)
        }
        
        contact.emailAddresses = contactEmailAddresses
        
        return contact
    }
    
    
    
    /***************************************************  MAIL SNIPPETS **************************************************/
    
    func fetchMessageDetailForMessage(_ message: MSOutlookMessage, completionHandler:@escaping ((MSOutlookMessage, MSODataException?) -> Void)){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher : MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.id)
            //messageFetcher.select("Id","Body","UniqueBody","Attachments")
            messageFetcher.expand("@Attachments")
            
            let task = messageFetcher.read(callback: { (messageDetail: MSOutlookMessage?, error: MSODataException?) in
                completionHandler(messageDetail!, error)
            })
            
            task?.resume()
            
        }
    }
    
    
    //Get the 10 most recent email messages in the user's inbox
    func fetchMailMessagesByConversationId(_ message: MSOutlookMessage, completionHandler:@escaping (([Any]?, MSODataException?) -> Void)){
    
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            //userFetcher!.getFolders(). filter("Id eq 'Inbox' or Id eq 'SentItems'")
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getFolders().getById("Inbox").getMessages()
            messageCollectionFetcher.order(by: "DateTimeReceived desc")
            messageCollectionFetcher.select("*")
            messageCollectionFetcher.filter("ConversationId eq '\(message.conversationId!)'")
            
           
            let task = messageCollectionFetcher.read{(messages:[Any]?, error:MSODataException?) -> Void in
            
                completionHandler(messages, error)
            }
            
            task?.resume()
        
        }
        
    
    }
    
    //Get the 10 most recent email messages in the user's inbox
    func fetchMailMessages(_ completionHandler:@escaping (([Any]?, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            // Retrieve mail messages from O365 and pass the status to the callback. Uses a default page size of 10
            // This results in a call to the service
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getFolders().getById("Inbox").getMessages()
            messageCollectionFetcher.top(10)
            messageCollectionFetcher.order(by: "DateTimeReceived desc")
            messageCollectionFetcher.select("*")
            
            
            let task = messageCollectionFetcher.read{(messages:[Any]?, error:MSODataException?) -> Void in
                self.lastrefreshdate = Date()
                self.allMessages = messages as! [MSOutlookMessage]
                self.allConversations = self.getConversationsFromMessages(self.allMessages)
                
                completionHandler(messages, error)
            }
            
            task?.resume()
        }
    }
    
    //fetch email messages based on pagenumber
    func fetchMailMessagesForPageNumber(_ pageNumber: Int32, pageSize: Int32, orderBy: String, folder: String, completionHandler:@escaping (([Any]?, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getFolders().getById(folder).getMessages()
            messageCollectionFetcher.order(by: orderBy)
            messageCollectionFetcher.top(pageSize)
            messageCollectionFetcher.skip(pageNumber * pageSize)
            messageCollectionFetcher.select("*")
            
            //retrieve messages
            let task = messageCollectionFetcher.read{(messages:[Any]?, error:MSODataException?) -> Void in
                self.lastrefreshdate = Date()
                
                //append additional messages
                let additionalMessages: [MSOutlookMessage] = messages as! [MSOutlookMessage]
                for additionalMessage in additionalMessages {
                        self.allMessages.append(additionalMessage)
                }
                
                //append additional conversations
                let additionalConversations = self.getConversationsFromMessages(additionalMessages)
                for additionalConversation in additionalConversations {
                    if(!self.allConversations.contains(where: {$0.newestMessage().hash == additionalConversation.newestMessage().hash})){
                        self.allConversations.append(additionalConversation)
                    }
                }
        
                completionHandler(messages, error)
            }
            
            task?.resume()
        }
        
    }
    
    //Sends a new email message to the user
    func sendMailMessage(_ message: MSOutlookMessage, completionHandler:@escaping ((Int32, MSODataException?) -> Void)) {
        
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            let userFetcher = outlookClient.getMe()
            let userOperations = ((userFetcher?.operations)! as MSOutlookUserOperations)
            
            // The returnValue is the HTTP status code
            let task = userOperations.sendMail(with: message, saveToSentItems: true) {
                (returnValue: Int32, error: MSODataException?) -> Void in
                completionHandler(returnValue, error)
            }
            
            task?.resume()
        }
    }
    
    
    //Creates a new email message in the user's Drafts folder
    //Does not send the email
    func createDraftMailMessage(_ message: MSOutlookMessage, completionHandler:@escaping ((MSOutlookMessage, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            
            let task = messageCollectionFetcher.add(message){ (addedMessage: MSOutlookMessage?, error: MSODataException?) -> Void in
                completionHandler(addedMessage!, error)
            }
            
            task?.resume()
        }
    }
    
    // Creates and sends an email message to a single recipient, with a subject, an HTML body and save a copy in the sender's
    // sentitems folder
    func createAndSendHTMLMailMessage(_ toRecipients: NSMutableArray, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        // Get the client and get the operations for sending a mail
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let userOperations = ((userFetcher?.operations)! as MSOutlookUserOperations)
            
            // Form the email message
            let message: MSOutlookMessage = MSOutlookMessage()
            
            message.subject = "Here's the subject"
            
            message.body = MSOutlookItemBody()
            message.body.content = "<!DOCTYPE html><html><body>Here's the body.</body></html>"
            message.body.contentType = MSOutlookBodyType.bodyType_HTML
            message.toRecipients = toRecipients
            
            // The returnValue is the HTTP status code
            let task = userOperations.sendMail(with: message, saveToSentItems: true) {
                (returnValue:Int32, error:MSODataException?) -> Void in
                
                let sucess: Bool = (returnValue == 0)
                completionHandler(sucess, error)
            }
            
            task?.resume()
            
        }
    }
    
    //Updates an email message on the server
    func updateMailMessage(_ message: MSOutlookMessage, completionHandler:@escaping ((MSOutlookMessage, MSODataException?) -> Void )){
        
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.id)
            
          
            
            let task = messageFetcher.update(message){ (updatedMessage: MSOutlookMessage?, error: MSODataException?) -> Void in
                completionHandler(updatedMessage!, error)
            }
            
            task?.resume()
            
        }
    }
    
    //Deletes an email message from the server
    func deleteMailMessage(_ message: MSOutlookMessage, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        //let clientFetcher : Office365ClientFetcher = Office365ClientFetcher()
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.id)
            
            let task = messageFetcher.deleteMessage { (status: Int32, error: MSODataException?) -> Void in
                
                let success: Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task?.resume()
            
        }
        
    }
    
    //Replies to a single recipient in a mail message
    func replyToMailMessage(_ message: MSOutlookMessage, body: String!, completionHandler:@escaping ((Int32, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.id)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.reply(withComment: body){ (returnValue:Int32, error:MSODataException?) -> Void in
                completionHandler(returnValue, error)
            }
            
            task?.resume()
        }
    }
    
    //Replies to all recipients in a mail message
    func replyAllToMailMessage(_ message: MSOutlookMessage, body: String!, completionHandler:@escaping ((Int32, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.id)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.replyAll(withComment: body){ (returnValue:Int32, error:MSODataException?) -> Void in
                completionHandler(returnValue, error)
            }
            
            task?.resume()
        }
    }
    
    //Forwards a mail message to a specific Recipient
    func forwardMailMessage(_ message: MSOutlookMessage, body:String!, toRecipient: MSOutlookRecipient!, completionHandler:@escaping ((Int32, MSODataException?) -> Void)) {
    
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.id)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.forward(withComment: body, toRecipients: toRecipient, callback: { (returnValue: Int32, error: MSODataException?) in
                completionHandler(returnValue, error)
            })
            
            task?.resume()
        }
    
    }
    
    
    //Create a draft reply email message in inbox
    func createDraftReplyAllMessage(_ message: MSOutlookMessage, completionHandler:@escaping ((MSOutlookMessage, MSODataException?) -> Void )){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(message.id)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.createReply(){ (replyMessage:MSOutlookMessage?, error:MSODataException?) -> Void in
                completionHandler(replyMessage!, error)
            }
            
            task?.resume()
        }
    }
    
    /**
     *  Copy a mail message to the deleted items folder.
     *  https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations#MoveorcopymessagesCopyamessageREST
     *
     *  @param itemId     Identifier of the message that will be copied
     *  @param completion
     */
    func copyMessage(_ messageId: String, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(messageId)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.copy(withDestinationId: "DeletedItems"){ (msg:MSOutlookMessage?, error:MSODataException?) -> Void in
                
                // You now have the copied MSOutlookMessage named msg.
                let success: Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task?.resume()
        }
    }
    
    /**
     *  Move a mail message to the deleted items folder.
     *  https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations#MoveorcopymessagesCopyamessageREST
     *
     *  @param itemId     Identifier of the message that will be moved
     *  @param completion
     */
    func moveMessage(_ messageId: String, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(messageId)
            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            let task = messageOperations.move(withDestinationId: "DeletedItems"){ (msg:MSOutlookMessage?, error:MSODataException?) -> Void in
                
                // You now have the copied MSOutlookMessage named msg.
                let success: Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task?.resume()
        }
    }
    
    /**
     *  Fetch up to the first 10 unread messages in your inbox that have been marked as important.
     *  https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations#MoveorcopymessagesCopyamessageREST
     *
     *  @param completion
     */
    func fetchUnreadImportantMessages(_ completionHandler:@escaping (([Any]?, MSODataException?) -> Void)){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            
            // Identify which properties to return. Only request properties that you will use.
            // The identifier is always returned.
            // https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesMessage
            messageCollectionFetcher.select("Subject,Importance,IsRead,Sender,DateTimeReceived")
            
            // Search for items that are both unread and marked as important. The library will URL encode this for you.
            // https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#UseODataqueryparameters
            //[messageCollectionFetcher addCustomParametersWithName:@"filter" value:@"IsRead eq true AND Importance eq 'High'"];
            messageCollectionFetcher.filter("IsRead eq true AND Importance eq 'High'")
            
            let task = messageCollectionFetcher.read(){ (messages: [Any]?, error: MSODataException?) -> Void in
                
                if error == nil {
                    completionHandler(messages, error)
                }else{
                    completionHandler([], error)
                }
            }
            
            task?.resume()
            
        }
    }
    
    
    /**
     *  Get the weblink to the first item in your inbox. This example assumes you have at least one item in your inbox.
     *  https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesMessage
     *
     *  @param completion
     */
    func fetchMessagesWebLink(_ completionHandler:@escaping ((String, MSODataException?) -> Void)){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            
            // Identify which properties to return. Only request properties that you will use.
            // The identifier is always returned.
            // https://msdn.microsoft.com/office/office365/APi/complex-types-for-mail-contacts-calendar#RESTAPIResourcesMessage
            messageCollectionFetcher.select("Subject,WebLink")
            
            // Identify how many items to return in the page.
            messageCollectionFetcher.top(1)
            
            let task = messageCollectionFetcher.read(){ (messages: [Any]?, error: MSODataException?) -> Void in
                
                //You now have an NSArray of MSOutlookMessage objects. Let's get the first and ony object
                let message: MSOutlookMessage = messages!.first as! MSOutlookMessage
                let weblink = (error == nil) ? message.webLink : ""
                completionHandler(weblink!, error)
            }
            
            task?.resume()
            
        }
    }
    
    //Send a draft message
    func sendDraftMessage(_ messageId: String, completionHandler:((MSOutlookMessage, MSODataException?) -> Void)){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            //TO DO:
            //            let userFetcher = outlookClient.getMe()
            //            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher.getMessages()
            //            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(messageId)
            //            let messageOperations = (messageFetcher.operations as MSOutlookMessageOperations)
            
            
        }
        
    }
    
    //Add an attachment to a message
    func addAttachment(_ message: MSOutlookMessage, contentType: String, contentBytes: Data, completionHandler:((MSOutlookMessage, MSODataException?) -> Void)){
    }
    
    //Send mail with an attachment
    func sendMailMessage(_ message: MSOutlookMessage,
                         AttachmentContentType: String,
                         contentBytes: Data, completionHandler:((Bool, MSODataException?) -> Void)){
    }
    
    func markAsRead(_ messageid: String, isRead: Bool, completionHandler:@escaping ((String, MSODataException?) -> Void )){
        
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getMessages()
            let messageFetcher: MSOutlookMessageFetcher = messageCollectionFetcher.getById(messageid)
            
            let postString: String =  "{IsRead:\(isRead)}"
            let payload =  String(data: postString.data(using: .utf8)!, encoding: .utf8)
            
            let task = messageFetcher.updateRaw(payload) { (response: String?, error:MSODataException?) in
                
                if(!(response != nil)){
                    completionHandler("", error)
                    return
                }
                
                completionHandler(response!, error)
            }
            
            task?.resume()
        }
        
        
    }
    
    /************************************************* END:  MAIL SNIPPETS ************************************************/
    
    
    
    
    
    /****************************************************  CALENDAR **************************************************/
    
    //Gets the 10 most recent calendar events from the user's calendar
    func fetchCalendarEvents(_ completionHandler:@escaping (([Any]?, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            // Retrieve events from O365 and pass the status to the callback. Uses a default page size of 10.
            // This results in a call to the service.
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher!.getEvents()
            
            let task = eventCollectionFetcher.read{(events:[Any]?, error:MSODataException?) -> Void in
                completionHandler(events, error)
            }
            
            task?.resume()
        }
    }
    
    //Creates a new event in the user's calendar
    func createCalendarEvent(_ event: MSOutlookEvent, completionHandler:@escaping ((MSOutlookEvent, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher!.getEvents()
            
            let task = eventCollectionFetcher.add(event){ (addedEvent: MSOutlookEvent?, error: MSODataException?) -> Void in
                completionHandler(addedEvent!, error)
            }
            
            task?.resume()
        }
    }
    
    //Updates an event in the user's calendar
    func updateCalendarEvent(_ event: MSOutlookEvent, completionHandler:@escaping ((MSOutlookEvent, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher!.getEvents()
            let eventFetcher : MSOutlookEventFetcher = eventCollectionFetcher.getById(event.id)
            
            let task = eventFetcher.update(event){ (updatedEvent: MSOutlookEvent?, error: MSODataException?) -> Void in
                completionHandler(updatedEvent!, error)
            }
            
            task?.resume()
        }
    }
    
    //Deletes an event from ther user's calendar
    func deleteCalendarEvent(_ event: MSOutlookEvent, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher!.getEvents()
            let eventFetcher : MSOutlookEventFetcher = eventCollectionFetcher.getById(event.id)
            
            let task = eventFetcher.deleteEvent{(status: Int32, error: MSODataException?) -> Void in
                
                let success : Bool = (error == nil)
                completionHandler(success, error)
            }
            
            task?.resume()
            
        }
    }
    
    //Accepts an event with comment - comment can be nil
    func acceptCalendarMeetingEvent(_ event: MSOutlookEvent, comment: String, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventFetcher : MSOutlookEventFetcher = userFetcher!.getEventsById(event.id)
            let operations: MSOutlookEventOperations = (eventFetcher.operations as MSOutlookEventOperations)
            
            let task = operations.accept(withComment: comment){(returnValue: Int32, error: MSODataException?) -> Void in
                
                let success: Bool = (returnValue == 0)
                completionHandler(success, error)
            }
            
            task?.resume()
        }
        
    }
    
    //Declines an event with a comment - comment can be nil
    func declineCalendarMeetingEvent(_ event: MSOutlookEvent, comment: String, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventFetcher : MSOutlookEventFetcher = userFetcher!.getEventsById(event.id)
            let operations: MSOutlookEventOperations = (eventFetcher.operations as MSOutlookEventOperations)
            
            let task = operations.decline(withComment: comment){(returnValue: Int32, error: MSODataException?) -> Void in
                
                let success: Bool = (returnValue == 0)
                completionHandler(success, error)
            }
            
            task?.resume()
        }
    }
    
    //Tentatively accepts an event with comment - comment can be nil
    func tentativelyAcceptCalendarMeetingEvent(_ event: MSOutlookEvent, comment: String, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventFetcher : MSOutlookEventFetcher = userFetcher!.getEventsById(event.id)
            let operations: MSOutlookEventOperations = (eventFetcher.operations as MSOutlookEventOperations)
            
            let task = operations.tentativelyAccept(withComment: comment){(returnValue: Int32, error: MSODataException?) -> Void in
                
                let success: Bool = (returnValue == 0)
                completionHandler(success, error)
            }
            
            task?.resume()
        }
    }
    
    // Fetches the first 10 event instances in the specified date range
    // For more information about calendar view, visit https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations#GetCalendarView
    func fetchCalendarViewFrom(_ start: Date, end: Date, completionHandler:@escaping (([Any]?, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let eventCollectionFetcher : MSOutlookEventCollectionFetcher = userFetcher!.getCalendarView()
            eventCollectionFetcher.select("Subject")
            
            let numDaysToCheck : Double = 10
            let secondsBackToConsider: Double = numDaysToCheck * 60 * 60 * 24
            
            eventCollectionFetcher.addCustomParameters(withName: "startDateTime", value: Date(timeIntervalSinceNow: -secondsBackToConsider))
            eventCollectionFetcher.addCustomParameters(withName: "endDateTime", value: Date())
            
            let task = eventCollectionFetcher.read{(events:[Any]?, error:MSODataException?) -> Void in
                
                completionHandler(events, error)
            }
            
            task?.resume()
            
        }
    }
    
    
    /*************************************************** END CALENDAR **************************************************/
    
    
    
    
    /*************************************************** CONTACTS **************************************************/
    
    //Gets the 10 most recently added user's contacts from Office 365
    func fetchContacts(_ completionHandler:@escaping (([Any]?, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactFetcher = userFetcher?.getContacts()
            
            let task = contactFetcher?.read{(contacts:[Any]?, error:MSODataException?) -> Void in
                
                completionHandler(contacts, error)
            }
            
            task?.resume()
        }
    }
    
    //Creates a new contact for the user
    func createContact(_ contact: MSOutlookContact, completionHandler:@escaping ((MSOutlookContact, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactCollectionFetcher: MSOutlookContactCollectionFetcher = userFetcher!.getContacts()
            
            let task = contactCollectionFetcher.add(contact){(addedContact: MSOutlookContact?, error: MSODataException?) in
                completionHandler(addedContact!, error)
            }
            
            task?.resume()
        }
    }
    
    //Updates a contact in Office 365
    func updateContact(_ contact: MSOutlookContact, completionHandler:@escaping ((MSOutlookContact, MSODataException?) -> Void )){
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactCollectionFetcher: MSOutlookContactCollectionFetcher = userFetcher!.getContacts()
            let contactFetcher: MSOutlookContactFetcher = contactCollectionFetcher.getById(contact.id)
            
            let task = contactFetcher.update(contact){(updatedContact: MSOutlookContact?, error: MSODataException?) in
                completionHandler(updatedContact!, error)
            }
            
            task?.resume()
        }
    }
    
    //Deletes a contact from Office 365
    func deleteContact(_ contact: MSOutlookContact, completionHandler:@escaping ((Bool, MSODataException?) -> Void)) {
        
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service.
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            
            let userFetcher = outlookClient.getMe()
            let contactCollectionFetcher: MSOutlookContactCollectionFetcher = userFetcher!.getContacts()
            let contactFetcher: MSOutlookContactFetcher = contactCollectionFetcher.getById(contact.id)
            
            let task = contactFetcher.deleteContact({ (status: Int32, error: MSODataException?) in
                
                let success: Bool = (error == nil)
                completionHandler(success, error)
            })
            
            task?.resume()
            
        }
    }
    
    
    /*************************************************** END CONTACTS **************************************************/
    
    
    
    
    /*************************************************** FILES *****************************************************/
    
    //Get 10 files or folders from the user's OneDrive for Business folder
    func fetchFiles(_ completionHandler:@escaping (([Any]?, MSODataException?) -> Void)) {
        
        // Get the SharePoint client. This object contains access tokens and methods to call the service.
        clientFetcher.fetchSharePointClient{ (sharePointClient) -> Void in
            
            let fileFetcher: MSSharePointItemCollectionFetcher = sharePointClient.getfiles()
            
            // Retrieve files from O365 and pass the status to the callback. Uses a default page size of 10
            let task = fileFetcher.read(callback: { (files: [Any]?, error: MSODataException?) in
                completionHandler(files, error)
            })
            
            task?.resume()
            
        }
    }
    
    
    /*************************************************** END FILES **************************************************/
    
}
