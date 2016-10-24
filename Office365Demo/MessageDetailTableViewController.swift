//
//  MessageDetailTableViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 10/22/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class MessageDetailTableViewController: UITableViewController, UIWebViewDelegate {
    
    @IBOutlet weak var lblSubject: UILabel!

    
    let office365Manager: Office365Manager = Office365Manager()
    var conversation: Conversation!
    var message: MSOutlookMessage!
    var contentHeights : [CGFloat] = [CGFloat]()
    
    override func viewDidLoad() {
        super.viewDidLoad()
        
        
        if(message != nil){
            lblSubject.text = message.subject!
        }
        
        if(conversation != nil){
            tableView.delegate = self
            tableView.rowHeight = UITableViewAutomaticDimension
            tableView.reloadData()
            
            for _ in 0...conversation.messages.count {
                contentHeights.append(0.0)
            }
        }
    }
    
    
    //not used right now
    func fetchConversationMessages(){
        office365Manager.fetchMailMessagesByConversationId(message) { (conversationMessages: [Any]?, error: MSODataException?) in
            
            let conversations: [MSOutlookMessage] = conversationMessages as! [MSOutlookMessage]
            
            for conversationMessage in conversations{
                print("Let's see: \(conversationMessage.from.emailAddress.address!) \(conversationMessage.conversationId!)")
            }
        }
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }

    // MARK: - Table view data source

    override func numberOfSections(in tableView: UITableView) -> Int {
        return conversation.messages.count
    }

    override func tableView(_ tableView: UITableView, numberOfRowsInSection section: Int) -> Int {
        return 1
    }
    
    override func tableView(_ tableView: UITableView, titleForHeaderInSection section: Int) -> String? {
        
        let message : MSOutlookMessage = conversation.messages[section] as! MSOutlookMessage
        
        office365Manager.markAsRead(message.id, isRead: true, completionHandler: { (response: String, error: MSODataException?) in
            //print("mark as read response: \(response)")
        })
        
        
        var to: String = ""//to recipients
        if(message.toRecipients) != nil{
            let toRecipients: [MSOutlookRecipient] = message.toRecipients as NSArray as! [MSOutlookRecipient]
            for (index,element) in toRecipients.enumerated() {
                let recipient:MSOutlookRecipient = element as MSOutlookRecipient
                let temp: String = "\(recipient.emailAddress.name!)<\(recipient.emailAddress.address!)>"
                to = to + (index > 0 ? ",\(temp)" : "\(temp)")
            }
            to = (to != "") ? "To:\(to)" : ""
        }
        
        var cc: String = ""//cc recipients
        if (message.ccRecipients) != nil{
            let ccRecipients: [MSOutlookRecipient] = message.ccRecipients as NSArray as! [MSOutlookRecipient]
            for (index,element) in ccRecipients.enumerated() {
                let recipient: MSOutlookRecipient = element as MSOutlookRecipient
                let temp: String = "\(recipient.emailAddress.name!)<\(recipient.emailAddress.address!)>"
                cc = cc + (index > 0 ? ",\(temp)" : "\(temp)")
            }
            cc = (cc != "") ? "Cc:\(cc)" : ""
        }
        
        return "\(message.from.emailAddress.name!)<\(message.from.emailAddress.address!)> on \(message.dateTimeReceived.o365_string_from_date())  \(to)  \(cc)"
    }
    
    override func tableView(_ tableView: UITableView, willDisplayHeaderView view: UIView, forSection section: Int) {
        
        let header = view as! UITableViewHeaderFooterView
        
        if let textlabel = header.textLabel {
            textlabel.font = textlabel.font.withSize(10)
        }
    }
    
    override func tableView(_ tableView: UITableView, heightForRowAt indexPath: IndexPath) -> CGFloat
    {
        return contentHeights.count <= 0 ? 0 : contentHeights[(indexPath as NSIndexPath).section]
    }
    
    override func tableView(_ tableView: UITableView, heightForFooterInSection section: Int) -> CGFloat {
        return UITableViewAutomaticDimension
    }

    
    override func tableView(_ tableView: UITableView, cellForRowAt indexPath: IndexPath) -> UITableViewCell {

        let cell: MessageDetailTableViewCell = tableView.dequeueReusableCell(withIdentifier: "messagedetail") as! MessageDetailTableViewCell
        let section: Int = (indexPath as NSIndexPath).section
        let message: MSOutlookMessage = conversation.messages[section] as! MSOutlookMessage
        
        print("section \(section)")
        
        //reply button
        cell.mReply.addTarget(self, action: #selector(MessageDetailTableViewController.Reply(sender:)), for: .touchUpInside)
        cell.mReply.tag = section
        
        //replyall button
        cell.mReplyAll.addTarget(self, action: #selector(MessageDetailTableViewController.ReplyAll(sender:)), for: .touchUpInside)
        cell.mReplyAll.tag = section
        cell.mReplyAll.isEnabled = message.toRecipients.count > 1 ? true : false
        
        //forward button
        cell.mForward.addTarget(self, action: #selector(MessageDetailTableViewController.Forward(sender:)), for: .touchUpInside)
        cell.mForward.tag = section
        
        let htmlHeight = contentHeights[section]
        cell.mBodyView.tag = section
        cell.mBodyView.delegate = self
        cell.mBodyView.scrollView.isScrollEnabled = false
        cell.mBodyView.loadHTMLString(message.body.content, baseURL: nil)//body
        cell.mBodyView.frame =  CGRect(x: 0, y: 0, width: cell.frame.size.width, height: htmlHeight)

        return cell
    }
    
    func Reply(sender: UIButton){
        //get section id
        let section: Int = sender.tag
        //get message
        let message: MSOutlookMessage = conversation.messages[section] as! MSOutlookMessage
        //show navCompose navigation to the ComposeViewController
        let navCompose: UINavigationController = self.storyboard?.instantiateViewController(withIdentifier: "navCompose") as! UINavigationController
        let composeView : ComposeTableViewController = navCompose.viewControllers.first as! ComposeTableViewController
        composeView.message = message
        composeView.composeType = ComposeType.Reply.rawValue
        self.present(navCompose, animated:true, completion:nil)
    }
    
    func ReplyAll(sender: UIButton){
        //get section id
        let section: Int = sender.tag
        let message: MSOutlookMessage = conversation.messages[section] as! MSOutlookMessage
        //show navCompose navigation to the ComposeViewController
        let navCompose: UINavigationController = self.storyboard?.instantiateViewController(withIdentifier: "navCompose") as! UINavigationController
        let composeView : ComposeTableViewController = navCompose.viewControllers.first as! ComposeTableViewController
        composeView.message = message
        composeView.composeType = ComposeType.ReplyAll.rawValue
        self.present(navCompose, animated:true, completion:nil)

    }
    
    func Forward(sender: UIButton){
        //get section id
        let section: Int = sender.tag
        let message: MSOutlookMessage = conversation.messages[section] as! MSOutlookMessage
        //show navCompose navigation to the ComposeViewController
        let navCompose: UINavigationController = self.storyboard?.instantiateViewController(withIdentifier: "navCompose") as! UINavigationController
        let composeView : ComposeTableViewController = navCompose.viewControllers.first as! ComposeTableViewController
        composeView.message = message
        composeView.composeType = ComposeType.Forward.rawValue
        self.present(navCompose, animated:true, completion:nil)
    }
    

    func webViewDidFinishLoad(_ webView: UIWebView)
    {
        if (contentHeights[webView.tag] != 0.0)
        {
            // we already know height, no need to reload cell
            return
        }
        
        contentHeights[webView.tag] = webView.scrollView.contentSize.height
        tableView.reloadRows(at: [NSIndexPath(row: 0, section: webView.tag) as IndexPath], with: .automatic)
    }
    
    
    /****************** START: unwind exit or close actions to the viewcontroller **********************/
    @IBAction func unwindToViewController(_ segue: UIStoryboardSegue) {
        
        if(segue.source.isKind(of: Office365Demo.ComposeTableViewController)){//message detail vc
           //TO DO: use if needed in the future
        }
    }
    /****************** END: unwind exit or close actions to the viewcontroller *************************/

  

    /*
    // MARK: - Navigation

    // In a storyboard-based application, you will often want to do a little preparation before navigation
    override func prepare(for segue: UIStoryboardSegue, sender: Any?) {
        // Get the new view controller using segue.destinationViewController.
        // Pass the selected object to the new view controller.
    }
    */

}
