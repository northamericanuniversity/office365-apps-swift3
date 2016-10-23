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
       
        if(composeType == ComposeType.Reply.rawValue){
            print("reply")
            txtTo.text = "\(message.from.emailAddress.name!)<\(message.from.emailAddress.address!)>"
            txtSubject.isEnabled = false
            txtTo.isEnabled = false
        }else if(composeType == ComposeType.ReplyAll.rawValue){
            print("replyall")
            txtSubject.isEnabled = false
            txtTo.isEnabled = false
            
            let toRecipients: [MSOutlookRecipient] = message.toRecipients as NSArray as! [MSOutlookRecipient]
            for (index,element) in toRecipients.enumerated() {
                let recipient:MSOutlookRecipient = element as MSOutlookRecipient
                let to: String = "\(recipient.emailAddress.name!)<\(recipient.emailAddress.address!)>"
                txtTo.text = txtTo.text! + (index > 0 ? ",\(to)" : "\(to)")
            }
            
        }else{
            
        }
        
        txtSubject.text = "\(message.subject!)"
        
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }

    
    @IBAction func sendEmail(_ sender: UIBarButtonItem) {
        
        let outlookMessage: MSOutlookMessage = MSOutlookMessage()
        let from: MSOutlookRecipient = MSOutlookRecipient()
        let emailfrom: MSOutlookEmailAddress = MSOutlookEmailAddress()
        emailfrom.address = txtFrom.text!
        from.emailAddress = emailfrom
        outlookMessage.from = from
       
//        let toRecipients: NSMutableArray = message.toRecipients as NSArray as! NSMutableArray
//        toRecipients.add(txtTo.text!)
//        outlookMessage.toRecipients = toRecipients
        
        let toRecipient: MSOutlookRecipient = MSOutlookRecipient()
        let emailto: MSOutlookEmailAddress = MSOutlookEmailAddress()
        emailto.address = txtTo.text!
        toRecipient.emailAddress = emailto
        outlookMessage.toRecipients = NSMutableArray()
        outlookMessage.toRecipients.add(toRecipient)
        
        outlookMessage.subject = txtSubject.text!
        
//        let body: MSOutlookItemBody = MSOutlookItemBody()
//        body.content = txtBody.text!
//        body.contentType = MSOutlookBodyType.bodyType_HTML
//        outlookMessage.body = body
        
        outlookMessage.body = MSOutlookItemBody()
        outlookMessage.body.content = "<!DOCTYPE html><html><body>\(txtBody.text!)</body></html>"
        outlookMessage.body.contentType = MSOutlookBodyType.bodyType_HTML
        
        
        
       
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
