# office365-apps-swift3
Intended to write all Office365 apps in Swift 3 

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


<h3>SCREENSHOTS</h3>

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


Examples consists of Reply, ReplyAll, Forward and Compose

