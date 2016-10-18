//
//  MessagesViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 10/17/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class MessagesViewController: UIViewController {

    @IBOutlet weak var lblSubject: UILabel!
    @IBOutlet weak var lblFrom: UILabel!
    @IBOutlet weak var lblSentOn: UILabel!
    @IBOutlet weak var lblTo: UILabel!
    @IBOutlet weak var lblCc: UILabel!
    @IBOutlet weak var lblFiles: UILabel!
    @IBOutlet weak var activityIndicator: UIActivityIndicatorView!
    @IBOutlet weak var bodyWebMessage: UIWebView!
    @IBOutlet weak var barItemReplyAll: UIBarButtonItem!
    @IBOutlet weak var barItemReply: UIBarButtonItem!
    @IBOutlet weak var barItemForward: UIBarButtonItem!
    
    
    let office365Manager: Office365Manager = Office365Manager()
    var message: MSOutlookMessage!
    
    override func viewDidLoad() {
        super.viewDidLoad()

        
        
        if(message != nil){
            office365Manager.markAsRead(message.id, isRead: true, completionHandler: { (response: String, error: MSODataException?) in
                //print("mark as read response: \(response)")
            })
            
            lblSubject.text = message.subject!//subject
            lblFrom.text = "\(message.from.emailAddress.name!)<\(message.from.emailAddress.address!)" //from
            lblTo.text = ""
            lblCc.text = ""
            lblFiles.text = ""
            
            var recipientCount : Int = 0
            
            //to recipients
            if(message.toRecipients) != nil{
                let toRecipients: [MSOutlookRecipient] = message.toRecipients as NSArray as! [MSOutlookRecipient]
                for (index,element) in toRecipients.enumerated() {
                    let recipient:MSOutlookRecipient = element as MSOutlookRecipient
                    let to: String = "\(recipient.emailAddress.name!)<\(recipient.emailAddress.address!)>"
                    lblTo.text = lblTo.text! + (index > 0 ? ",\(to)" : "\(to)")
                }
                recipientCount = recipientCount + message.toRecipients.count
            }
            
            //cc recipients
            if (message.ccRecipients) != nil{
                let ccRecipients: [MSOutlookRecipient] = message.ccRecipients as NSArray as! [MSOutlookRecipient]
                for (index,element) in ccRecipients.enumerated() {
                    let recipient: MSOutlookRecipient = element as MSOutlookRecipient
                    let cc: String = "\(recipient.emailAddress.name!)<\(recipient.emailAddress.address!)>"
                    lblCc.text = lblCc.text! + (index > 0 ? ",\(cc)" : "\(cc)")
                }
                recipientCount = recipientCount + message.ccRecipients.count
            }
            
            
            lblSentOn.text = "\(message.dateTimeReceived.o365_string_from_date())"//sent date
            bodyWebMessage.loadHTMLString(message.body.content, baseURL: nil)//body
            
            barItemReplyAll.isEnabled = recipientCount > 1 ? true : false
        }else{
            barItemReply.isEnabled = false
            barItemReplyAll.isEnabled = false
        }
    }

    func markMessage(_ message: MSOutlookMessage){
        
        
    }
    
    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }
    

    /*
    // MARK: - Navigation

    // In a storyboard-based application, you will often want to do a little preparation before navigation
    override func prepare(for segue: UIStoryboardSegue, sender: Any?) {
        // Get the new view controller using segue.destinationViewController.
        // Pass the selected object to the new view controller.
    }
    */

}
