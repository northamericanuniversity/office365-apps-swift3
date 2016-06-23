//
//  InboxViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/21/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class InboxViewController: UIViewController {

    @IBOutlet weak var txtEmail: UITextField!
    @IBOutlet weak var activityIndicator: UIActivityIndicatorView!
    @IBOutlet weak var lblStatus: UILabel!
    
    var baseController = Office365ClientFetcher()
   
    
    override func viewDidLoad() {
        super.viewDidLoad()

        self.lblStatus.text = ""
        activityIndicator.hidden = true
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }
    
  
    @IBAction func sendEmailMessage(sender: AnyObject) {
        
        let message = buildMessage()
            
        // Get the MSOutlookClient. A call will be made to Azure AD and you will be prompted for credentials if you don't
        // have an access or refresh token in your token cache.
            
        baseController.fetchOutlookClient {
            (outlookClient) -> Void in
                
            dispatch_async(dispatch_get_main_queue()) {
                // Show the activity indicator
                self.activityIndicator.hidden = false
                self.activityIndicator.startAnimating()
            }
                
            let userFetcher = outlookClient.getMe()
            let userOperations = (userFetcher.operations as MSOutlookUserOperations)
                
            let task = userOperations.sendMailWithMessage(message, saveToSentItems: true) {
                    (returnValue:Int32, error:MSODataException!) -> Void in
                    
                var statusText: String
                    
                if (error == nil) {
                    statusText = "Check your inbox, you have a new message. :)"
                }
                else {
                    statusText = "The email could not be sent. Check the log for errors."
                    NSLog("%@",[error.localizedDescription])
                }
                    
                // Update the UI.
                    
                dispatch_async(dispatch_get_main_queue()) {
                    self.lblStatus.text = statusText
                    self.activityIndicator .stopAnimating()
                    self.activityIndicator.hidden = true
                }
            }
                
            task.resume()
        } //baseController
    }
        
    // Compose the mail message
    func buildMessage() -> MSOutlookMessage {
        
        // Create a new message. Set properties on the message.
            let  message: MSOutlookMessage  = MSOutlookMessage()
            message.Subject = "Welcome to Office 365 development on iOS with the Office 365 Connect sample"
            
            // Get the recipient's email address.
            // The ToRecipients property is an array of MSOulookRecipient objects.
            // See the helper method getRecipients to understand the usage.
            let toEmail = txtEmail.text
            
            let recipient = MSOutlookRecipient()
            recipient.EmailAddress = MSOutlookEmailAddress()
            recipient.EmailAddress.Address = toEmail!.stringByTrimmingCharactersInSet(NSCharacterSet.whitespaceCharacterSet())
            
            // The mutable array here is required to maintain compatibility with the API
            var recipientArray: [MSOutlookRecipient] = []
            recipientArray.append(recipient as MSOutlookRecipient)
            let mutableRecipientArray = NSMutableArray(array: recipientArray)
            message.ToRecipients = mutableRecipientArray
            
            // Get the email text and put in the email body.
            let filePath = NSBundle.mainBundle().pathForResource("EmailBody", ofType:"html")
            let body = (try? NSString(contentsOfFile: filePath!, encoding: NSUTF8StringEncoding))?.stringByReplacingOccurrencesOfString("\"", withString: "\\\"");
            
            message.Body = MSOutlookItemBody()
            message.Body.ContentType = MSOutlookBodyType.BodyType_HTML
            message.Body.Content = body! as String
            
            return message
    }

    

    /*
    // MARK: - Navigation

    // In a storyboard-based application, you will often want to do a little preparation before navigation
    override func prepareForSegue(segue: UIStoryboardSegue, sender: AnyObject?) {
        // Get the new view controller using segue.destinationViewController.
        // Pass the selected object to the new view controller.
    }
    */

}
