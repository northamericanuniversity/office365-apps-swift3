//
//  ComposeTableViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 10/19/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

enum ComposeType:String {
    case Compose = "Compose"
    case Reply = "Reply"
    case ReplyAll = "ReplyAll"
    case Forward = "Forward"
}

class ComposeTableViewController: UITableViewController {

    @IBOutlet weak var txtFrom: UITextField!
    @IBOutlet weak var txtTo: UITextField!
    @IBOutlet weak var txtSubject: UITextField!
    @IBOutlet weak var txtBody: UITextView!
    @IBOutlet weak var btnCancel: UIBarButtonItem!
    
    let office365Manager: Office365Manager = Office365Manager()
    var composeType = ComposeType.Compose.rawValue //default is  new email compose
    var message: MSOutlookMessage!
    var activitiyViewController : ActivityViewController!
    var userEmail: String!
    
    override func viewDidLoad() {
        super.viewDidLoad()

        userEmail = UserDefaults.standard.string(forKey: "demo_email")!
        txtFrom.text = userEmail
        txtFrom.isEnabled = false
        
        //reply
        if(composeType == ComposeType.Reply.rawValue){
            message.subject = "Re: \(self.message.subject!)"
            txtSubject.text = message.subject!
            txtSubject.isEnabled = false
            txtTo.isEnabled = false
            txtTo.text = "\(message.from.emailAddress.name!)<\(message.from.emailAddress.address!)>"
           
        }else if(composeType == ComposeType.ReplyAll.rawValue){//replyall
            message.subject = "Re: \(self.message.subject!)"
            txtSubject.text = message.subject!
            txtSubject.isEnabled = false
            txtTo.isEnabled = false
            let toRecipients: [MSOutlookRecipient] = message.toRecipients as NSArray as! [MSOutlookRecipient]
            for (index,element) in toRecipients.enumerated() {
                let recipient:MSOutlookRecipient = element as MSOutlookRecipient
                let to: String = "\(recipient.emailAddress.name!)<\(recipient.emailAddress.address!)>"
                txtTo.text = txtTo.text! + (index > 0 ? ",\(to)" : "\(to)")
            }
            
        }else if(composeType == ComposeType.Forward.rawValue){//forward
            txtSubject.isEnabled = false
            txtTo.isEnabled = true
            message.subject = "Fwd: \(self.message.subject!)"
            txtSubject.text = message.subject!
            //txtBody.loadHTMLString( message.body.content, baseURL: nil)//body
            txtBody.text = "\n\n\n \(message.body.content!)"
            
        }else if(composeType == ComposeType.Compose.rawValue){
            txtSubject.text = ""
            txtSubject.isEnabled = true
            txtBody.text = ""
            txtTo.isEnabled = true
            txtTo.text = ""
        }
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }

    
    @IBAction func sendEmail(_ sender: UIBarButtonItem) {
        
       
        //display loading... text until email posted successfully
        self.activitiyViewController = ActivityViewController(message: "Posting...")
        self.present(self.activitiyViewController, animated: true, completion: {
            
            //reply message
            if(self.composeType == ComposeType.Reply.rawValue){
                self.office365Manager.replyToMailMessage(self.message, body: self.txtBody.text!, completionHandler: { (returnValue: Int32, error: MSODataException?) in
                    //print("returnValue: \(returnValue) error: \(error)")
                    DispatchQueue.main.async(execute: {
                        self.dismiss(animated: true, completion: {
                            self.displayAlertMessage("Email sent successfully!")
                            UIApplication.shared.sendAction(self.btnCancel.action!, to: self.btnCancel.target, from: self, for: nil)
                        })
                    })
                })
            }else if(self.composeType == ComposeType.ReplyAll.rawValue){//replyall
                self.office365Manager.replyAllToMailMessage(self.message, body: self.txtBody.text!, completionHandler: { (returnValue: Int32, error: MSODataException?) in
                    DispatchQueue.main.async(execute: {
                        self.dismiss(animated: true, completion: {
                            self.displayAlertMessage("Email sent successfully!")
                            UIApplication.shared.sendAction(self.btnCancel.action!, to: self.btnCancel.target, from: self, for: nil)
                        })
                    })
                })
            }else if(self.composeType == ComposeType.Forward.rawValue){//forward
                
                var emails: String = self.txtTo.text!
                emails = emails.trimmingCharacters(in: .whitespaces)
                emails = emails.replacingOccurrences(of: " ", with: ";")
                emails = emails.replacingOccurrences(of: ",", with: ";")
                let emailArray : [String] = emails.components(separatedBy: ";")
                
                var toRecipients: [MSOutlookRecipient] = [MSOutlookRecipient]()
                for email in emailArray{
                    let toRecipient: MSOutlookRecipient = MSOutlookRecipient()
                    let emailto: MSOutlookEmailAddress = MSOutlookEmailAddress()
                    emailto.address = email
                    toRecipient.emailAddress = emailto
                    
                    toRecipients.append(toRecipient)
                    
                    self.office365Manager.forwardMailMessage(self.message, body: self.txtBody.text!, toRecipient: toRecipient, completionHandler: { (returnValue: Int32, error: MSODataException?) in
                        print("sent successfully!")
                    })
                }
                
                DispatchQueue.main.async(execute: {
                    self.dismiss(animated: true, completion: {
                        self.displayAlertMessage("Email sent successfully!")
                        UIApplication.shared.sendAction(self.btnCancel.action!, to: self.btnCancel.target, from: self, for: nil)
                    })
                })
            }else if(self.composeType == ComposeType.Compose.rawValue){//compose
                
                //from
                let outlookMessage: MSOutlookMessage = MSOutlookMessage()
                let from: MSOutlookRecipient = MSOutlookRecipient()
                let emailfrom: MSOutlookEmailAddress = MSOutlookEmailAddress()
                emailfrom.address = self.txtFrom.text!
                from.emailAddress = emailfrom
                outlookMessage.from = from
                
                //subject
                outlookMessage.subject = self.txtSubject.text!
                
                //to
                let toRecipient: MSOutlookRecipient = MSOutlookRecipient()
                let emailto: MSOutlookEmailAddress = MSOutlookEmailAddress()
                emailto.address = self.txtTo.text!
                toRecipient.emailAddress = emailto
                outlookMessage.toRecipients = NSMutableArray()
                outlookMessage.toRecipients.add(toRecipient)
                
                //body
                outlookMessage.body = MSOutlookItemBody()
                outlookMessage.body.content = "<!DOCTYPE html><html><body>\(self.txtBody.text!)</body></html>"
                outlookMessage.body.contentType = MSOutlookBodyType.bodyType_HTML
                
                self.office365Manager.sendMailMessage(outlookMessage, completionHandler: { (returnValue: Int32, error: MSODataException?) in
                    print("returnValue: \(returnValue)  error: \(error)")
                    
                    if(returnValue == 0 && error == nil){
                        DispatchQueue.main.async(execute: {
                            self.dismiss(animated: true, completion: {
                                self.displayAlertMessage("Email sent successfully!")
                                UIApplication.shared.sendAction(self.btnCancel.action!, to: self.btnCancel.target, from: self, for: nil)
                            })
                        })
                    }
                })
            }
            
        })//Posting
        
        
    }

    func displayAlertMessage(_ alertMessage:String){
        //shortcut alert message
        let myAlert = UIAlertController(title: "Alert", message: alertMessage, preferredStyle: UIAlertControllerStyle.alert)
        let okAction = UIAlertAction(title: "OK", style: UIAlertActionStyle.default, handler: nil)
        myAlert.addAction(okAction)
        self.present(myAlert, animated:true, completion:nil)
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
