# office365-apps-swift3
Intended to write all Office365 API custom apps in Swift 3 


<h3>SCREENSHOTS</h3>

<table border='0' width='100%' >
<th>
<td>List Emails</td>
<td>Lookup Email Detail</td>
<td>Send Email</td>
<td>Send Email</td>
</th>
<tr>
<td><img src='https://s3-us-west-2.amazonaws.com/s3-random-images/list_messages.png'></td>
<td><img src='https://s3-us-west-2.amazonaws.com/s3-random-images/message_detail.png'></td>
<td><img src='https://s3-us-west-2.amazonaws.com/s3-random-images/sendmessage_1.png'></td>
<td><img src='https://s3-us-west-2.amazonaws.com/s3-random-images/sendmessage_2.png'></td>
</tr>
</table>


<h3>Office 365 Keys</h3>
In AuhtenticationManager.swift file update your redirect url and client id
<pre>
let REDIRECT_URL_STRING = ""
let CLIENT_ID           = ""
let AUTHORITY           = "https://login.microsoftonline.com/common"
</pre>

<h3>POD FILE</h3>
Under Project folder you'll see the pod file
<pre>
source 'https://github.com/CocoaPods/Specs.git'
platform :ios, '8.0'
xcodeproj 'Office365Demo'

target ‘Office365Demo’ do

pod 'ADALiOS', '~> 1.2.2'
pod 'Office365/Outlook', '= 0.9.1'
pod 'Office365/Discovery', '= 0.9.1'
pod 'Office365/Files', '= 0.9.1'

end

</pre>

<h3>INSTALL COCOAPOD</h3>

Assuming you already have some experience with CocoaPod environment, please follow these steps

- Open your shell and go to under Office365Demo folder and run
<code>pod Install</code>

- If you already have this project and updated version of Pod, then run
<code>pod update</code>

<h3>BRIDGING HEADER FILE</h3>
That short bridging file took us a long time to figure out!
<pre>
#ifndef Bridging_Header_h
#define Bridging_Header_h

#import &lt;ADALiOS/ADAuthenticationContext.h&gt;
#import &lt;ADALiOS/ADAuthenticationSettings.h&gt;
#import &lt;ADALiOS/ADAuthenticationError.h&gt;
#import &lt;Office365/office365_discovery_sdk.h&gt;

//odata
#import &lt;office365_odata_base/office365_odata_base.h>
//discovery
#import &lt;office365_discovery_sdk/office365_discovery_sdk.h&gt;
//exchange server -> outlook
#import &lt;office365_exchange_sdk/office365_exchange_sdk.h&gt;
//sharepoint
#import &lt;office365_files_sdk/office365_files_sdk.h&gt;

#endif /* Bridging_Header_h */
</pre>



<h3>API EXAMPLE</h3>

Let's send an email

<pre>
    //from
    let outlookMessage: MSOutlookMessage = MSOutlookMessage()
    let from: MSOutlookRecipient = MSOutlookRecipient()
    let emailfrom: MSOutlookEmailAddress = MSOutlookEmailAddress()
    emailfrom.address = "blablafrom@example.com"
    from.emailAddress = emailfrom
    outlookMessage.from = from
                
    //subject
    outlookMessage.subject = "Test Email"
                
    //to
    let toRecipient: MSOutlookRecipient = MSOutlookRecipient()
    let emailto: MSOutlookEmailAddress = MSOutlookEmailAddress()
    emailto.address = "blablato@example.com"
    toRecipient.emailAddress = emailto
    outlookMessage.toRecipients = NSMutableArray()
    outlookMessage.toRecipients.add(toRecipient)
                
    //body
    outlookMessage.body = MSOutlookItemBody()
    outlookMessage.body.content = "<!DOCTYPE html><html><body>Example Body</body></html>"
    outlookMessage.body.contentType = MSOutlookBodyType.bodyType_HTML
                
    self.office365Manager.sendMailMessage(outlookMessage, completionHandler: { (returnValue: Int32, error: MSODataException?) in
                    
            print("returnValue: \(returnValue)  error: \(error)")
                    
            if(returnValue == 0 && error == nil){//successfull
              //do something
            }
    })

</pre>


Let's see how the sendMailMessage looks like

<pre>
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

</pre>


Office365Manager.swift contains many ready to go functions. 

To fetch to first 10 mails from Inbox
    <pre>
    //Get the 10 most recent email messages in the user's inbox
    func fetchMailMessages(_ completionHandler:@escaping (([Any]?, MSODataException?) -> Void)) {
        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
        // Retrieve mail messages from O365 and pass the status to the callback. Uses a default page size of 10
        // This results in a call to the service
        let userFetcher = outlookClient.getMe()
        let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getFolders().getById("<b>Inbox</b>").getMessages()
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
</pre>


To fetch from "SentItems", just change the folder name <code>userFetcher!.getFolders().getById("<b>SentItems</b>").getMessages()</code>


In the app we actually fetch by page number, page size, order by and folder name

<pre>
    //fetch email messages based on pagenumber
    func fetchMailMessagesForPageNumber(_ pageNumber: Int32, pageSize: Int32, orderBy: String, folder: String, completionHandler:@escaping (([Any]?, MSODataException?) -> Void)) {

        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in
            let userFetcher = outlookClient.getMe()
            let messageCollectionFetcher : MSOutlookMessageCollectionFetcher = userFetcher!.getFolders().getById(folder).getMessages()
            <b>messageCollectionFetcher.order(by: orderBy)
            messageCollectionFetcher.top(pageSize)
            messageCollectionFetcher.skip(pageNumber * pageSize)
            messageCollectionFetcher.select("*")</b>

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
</pre>

We get the conversations of each email on the fly. 
messagesByConversationID and conversations are defined as global
<pre>
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
</pre>

In fact, you don't have the fetch the conversations on the fly as we do in the app. In each email detail lookup you can fetch the conversations from the <b>conversationId</b> retrieved from the MSOutlookMessage object
<pre>
    func fetchConversationMessages(){
        office365Manager.fetchMailMessagesByConversationId(message) { (conversationMessages: [Any]?, error: MSODataException?) in

            let conversations: [MSOutlookMessage] = conversationMessages as! [MSOutlookMessage]

            for conversationMessage in conversations{
                print("\(conversationMessage.from.emailAddress.address!) \(conversationMessage.conversationId!)")
            }
        }
    }
</pre>

<pre>
    //Get the conversation messages from a specific message
    func fetchMailMessagesByConversationId(_ message: MSOutlookMessage, completionHandler:@escaping (([Any]?, MSODataException?) -> Void)){

        // Get the MSOutlookClient. This object contains access tokens and methods to call the service
        clientFetcher.fetchOutlookClient { (outlookClient) -> Void in

            let userFetcher = outlookClient.getMe()
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
</pre>


