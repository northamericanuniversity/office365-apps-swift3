//
//  MessagesViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 10/17/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class MessageDetailViewController: UIViewController, UITableViewDelegate, UITableViewDataSource, UIScrollViewDelegate, UIWebViewDelegate {

    @IBOutlet weak var lblSubject: UILabel!
    @IBOutlet weak var lblFrom: UILabel!
    @IBOutlet weak var lblSentOn: UILabel!
    @IBOutlet weak var lblTo: UILabel!
    @IBOutlet weak var lblCc: UILabel!
    @IBOutlet weak var lblFiles: UILabel!
    @IBOutlet weak var activityIndicator: UIActivityIndicatorView!
    @IBOutlet weak var barItemReply: UIBarButtonItem!
    @IBOutlet weak var barItemReplyAll: UIBarButtonItem!
    
    @IBOutlet weak var mTableview: UITableView!
    
    
    let office365Manager: Office365Manager = Office365Manager()
    var message: MSOutlookMessage!
    var conversation: Conversation!
    var contentHeights : [CGFloat] = [0.0, 0.0]
    
    override func viewDidLoad() {
        super.viewDidLoad()

        if(message != nil){
            
            
            office365Manager.markAsRead(message.id, isRead: true, completionHandler: { (response: String, error: MSODataException?) in
                //print("mark as read response: \(response)")
            })
            
            lblSubject.text = message.subject!//subject
            lblFrom.text = "\(message.from.emailAddress.name!)<\(message.from.emailAddress.address!)>" //from
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
        
            
            barItemReplyAll.isEnabled = recipientCount > 1 ? true : false
        }else{
            barItemReply.isEnabled = false
            barItemReplyAll.isEnabled = false
        }
        
        mTableview.delegate = self
        mTableview.rowHeight = UITableViewAutomaticDimension
        mTableview.reloadData()
        print("conversation messages count: \(conversation.messages.count)")

    }

    @IBAction func replyMessage(_ sender: UIBarButtonItem) {
        
        //show navCompose navigation to the ComposeViewController
        let navCompose: UINavigationController = self.storyboard?.instantiateViewController(withIdentifier: "navCompose") as! UINavigationController
        let composeView : ComposeTableViewController = navCompose.viewControllers.first as! ComposeTableViewController
        composeView.message = message
        composeView.composeType = ComposeType.Reply.rawValue
        self.present(navCompose, animated:true, completion:nil)
    }
    
    
    @IBAction func ReplyAllMessage(_ sender: UIBarButtonItem) {
        //show navCompose navigation to the ComposeViewController
        let navCompose: UINavigationController = self.storyboard?.instantiateViewController(withIdentifier: "navCompose") as! UINavigationController
        let composeView : ComposeTableViewController = navCompose.viewControllers.first as! ComposeTableViewController
        composeView.message = message
        composeView.composeType = ComposeType.ReplyAll.rawValue
        self.present(navCompose, animated:true, completion:nil)
    }
    
    
    
    
    func numberOfSections(in tableView: UITableView) -> Int {
        return conversation.messages.count
    }
    
    func tableView(_ tableView: UITableView, numberOfRowsInSection section: Int) -> Int {
        return 1
    }

    func tableView(_ tableView: UITableView, titleForHeaderInSection section: Int) -> String? {
        
        let message : MSOutlookMessage = conversation.messages[section] as! MSOutlookMessage
        return "\(message.from.emailAddress.name!)<\(message.from.emailAddress.address!)>"
    }
    
    func tableView(_ tableView: UITableView, willDisplayHeaderView view: UIView, forSection section: Int) {
        
        let header = view as! UITableViewHeaderFooterView
        
        if let textlabel = header.textLabel {
            textlabel.font = textlabel.font.withSize(10)
        }
    }
    
    func tableView(_ tableView: UITableView, heightForRowAt indexPath: IndexPath) -> CGFloat
    {
        return contentHeights[(indexPath as NSIndexPath).section]
    }
    
    func tableView(_ tableView: UITableView, heightForFooterInSection section: Int) -> CGFloat {
        return UITableViewAutomaticDimension
    }
    
    func tableView(_ tableView: UITableView, cellForRowAt indexPath: IndexPath) -> UITableViewCell {
        
        let cell: MessageDetailTableViewCell = mTableview.dequeueReusableCell(withIdentifier: "messagedetail") as! MessageDetailTableViewCell
        let message: MSOutlookMessage = conversation.messages[(indexPath as NSIndexPath).section] as! MSOutlookMessage
        
        let htmlHeight = contentHeights[(indexPath as NSIndexPath).section]
        cell.mBodyView.tag = (indexPath as NSIndexPath).section
        cell.mBodyView.delegate = self
        cell.mBodyView.loadHTMLString(message.body.content, baseURL: nil)//body
        cell.mBodyView.frame =  CGRect(x: 0, y: 0, width: cell.frame.size.width, height: htmlHeight)
       
        
        
        print("body: \(message.body.content)")
        
        return cell
    
    }
    
    func webViewDidFinishLoad(_ webView: UIWebView)
    {
        if (contentHeights[webView.tag] != 0.0)
        {
            // we already know height, no need to reload cell
            return
        }
        
        contentHeights[webView.tag] = webView.scrollView.contentSize.height
        print("webview height \(webView.scrollView.contentSize.height)")
        mTableview.reloadRows(at: [NSIndexPath(row: 0, section: webView.tag) as IndexPath], with: .automatic)
      
        
        //reloadRowsAtIndexPaths([NSIndexPath(forRow: webView.tag, inSection: 0)], withRowAnimation: .Automatic)
    }
    
    /****************** START: unwind exit or close actions to the viewcontroller **********************/
    @IBAction func unwindToViewController(_ segue: UIStoryboardSegue) {
        
        if(segue.source.isKind(of: Office365Demo.MessageDetailViewController)){//message detail vc
            self.dismiss(animated: true, completion: nil)
        }
    }
    /****************** END: unwind exit or close actions to the viewcontroller *************************/
    
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
