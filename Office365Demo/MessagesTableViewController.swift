//
//  MessagesTableViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/27/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class MessagesTableViewController: UITableViewController {

    @IBOutlet weak var statusBarButtonItem: UIBarButtonItem!
    @IBOutlet weak var activityBarButtonItem: UIBarButtonItem!
    
    var activityIndicator: UIActivityIndicatorView!
    var primaryStatusLabel: UILabel!
    var secondaryStatusLabel: UILabel!
    
    var isLoading : Bool = false
    let office365Manager: Office365Manager = Office365Manager()
    
    var currentPage : Int32 = 0 //default page number is zero (0)
    var actionSheetController : UIAlertController!
    
    var selectedOutlookMessage: MSOutlookMessage?
    
    
    override func viewDidLoad() {
        super.viewDidLoad()

        refreshControl?.addTarget(self, action: #selector(MessagesTableViewController.updateRefreshControl), for: UIControlEvents.valueChanged)
        setupActionSheet()
        setupUI()
    }
    
    override func viewDidAppear(_ animated: Bool) {
        updateRefreshControl()
    }
    
     /****************************** init action sheet ******************************/
    func setupActionSheet(){
       
        actionSheetController = UIAlertController(title: nil, message: "Choose Action", preferredStyle: .actionSheet)
        let cancelAction = UIAlertAction(title: "Cancel", style: .cancel) { (action) in
            
        }
        actionSheetController.addAction(cancelAction)//add action to the alert controller
        
        //sign out
        let signoutAction = UIAlertAction(title: "Sign out", style: .default) { (action) in
            
            let signoutAlert = UIAlertController(title: "Sign out", message: "Do you want to sign out?", preferredStyle: UIAlertControllerStyle.alert)
            
            //delete refused, don't delete this post
            signoutAlert.addAction(UIAlertAction(title: "Cancel", style: .cancel, handler: { (action: UIAlertAction) in
                //don't do anything here, the dialog simply closes itself
            }))
            
            //sign out now
            signoutAlert.addAction(UIAlertAction(title: "Yes", style: .default, handler: { (action: UIAlertAction) in
                
                // Clear the access and refresh tokens from the credential cache. You need to clear cookies
                // since ADAL uses information stored in the cookies to get a new access token.
                let authenticationManager:AuthenticationManager = AuthenticationManager.sharedInstance
                authenticationManager.clearCredentials()
                
                //jump back to the login view
                let loginView : LoginViewController = self.storyboard?.instantiateViewController(withIdentifier: "loginView") as! LoginViewController;
                self.present(loginView, animated: true, completion: nil)
                
            }))
            
            self.present(signoutAlert, animated: true, completion: nil)
            
        }
        actionSheetController.addAction(signoutAction)//add action to the alert controller
       
    }
     /****************************** END: init action sheet *****************************/
    
    /**************************************** init UI  **********************************/
    func setupUI(){
        
        let statusView: UIView = UIView(frame: CGRect(x: 0, y: 0, width: 225, height: 36))
        let primaryStatusLabel: UILabel = UILabel(frame: CGRect(x: 0, y: 0, width: 225, height: 16))
        let secondaryStatusLabel: UILabel = UILabel(frame: CGRect(x: 0, y: 18, width: 225, height: 12))
        
        primaryStatusLabel.font = UIFont.systemFont(ofSize: 13)
        secondaryStatusLabel.font = UIFont.systemFont(ofSize: 10)
        
        primaryStatusLabel.textAlignment = NSTextAlignment.center
        secondaryStatusLabel.textAlignment = NSTextAlignment.center
        
        primaryStatusLabel.textColor = UIColor().o365_PrimaryColor()
        secondaryStatusLabel.textColor = UIColor.gray
        
        statusView.addSubview(primaryStatusLabel)
        statusView.addSubview(secondaryStatusLabel)
        
        self.primaryStatusLabel = primaryStatusLabel
        self.secondaryStatusLabel = secondaryStatusLabel
        
        self.statusBarButtonItem.customView = statusView
        
        let activityIndicator : UIActivityIndicatorView = UIActivityIndicatorView(activityIndicatorStyle: UIActivityIndicatorViewStyle.gray)
        activityIndicator.color = UIColor().o365_PrimaryColor()
        
        self.activityBarButtonItem.customView = activityIndicator
        self.activityIndicator = activityIndicator
        
        refreshControl?.backgroundColor = UIColor().o365_PrimaryColor()
        refreshControl?.tintColor = UIColor.white
    }
    /************************************** END: init UI  ********************************/

    //update refresh control
    func updateRefreshControl(){
      
        if let lastUpdatedDate: Date = office365Manager.lastrefreshdate as Date? {
            let lastUpdatedTitle = "Last updated on \(lastUpdatedDate.o365_string_from_date())"
            refreshControl?.attributedTitle = NSAttributedString(string:   lastUpdatedTitle)
        }
        
        currentPage = 0
        performFetchMailMessages()
        
    }
    
    //update status
    func updateStatusWithPrimaryMessage(_ primaryMessage: String,
                                        secondaryMessage: String,
                                        activityInProgress: Bool){

        DispatchQueue.main.async {
            
            if(primaryMessage != "") {
                self.primaryStatusLabel.text = primaryMessage
            }
            
            if(secondaryMessage != "") {
                self.secondaryStatusLabel.text = secondaryMessage
            }
            
            if(activityInProgress) {
                self.activityIndicator.startAnimating()
            }else {
                self.activityIndicator.stopAnimating()
            }
        }
    }
    
    func performFetchMailMessages(){
        
        //keep always last updated date as secondary message in status tool bar
        var secondaryMessage = ""
        if let lastUpdatedDate: Date = office365Manager.lastrefreshdate as Date? {
            secondaryMessage = "Last updated on \(lastUpdatedDate.o365_string_from_date())"
        }
        
        self.updateStatusWithPrimaryMessage("fetching messages", secondaryMessage: secondaryMessage, activityInProgress: true)
 
        /*********************** Alternative way to get only 10 messags by default ****************
        office365Manager.fetchMailMessages { (messages: NSArray, error: MSODataException?) in
            dispatch_async(dispatch_get_main_queue()) {
                self.tableView.reloadData()
            }
        }****************** END: Alternative way to get only 10 messags by default ****************/
        
        
        //get email messages by page number
        office365Manager.fetchMailMessagesForPageNumber(currentPage, pageSize: 10, orderBy: "DateTimeReceived desc") { (messages: [Any]?, error: MSODataException?) in
            
            DispatchQueue.main.async {
                
                self.tableView.reloadData()
                self.refreshControl?.endRefreshing()
                self.isLoading = false
                var secondaryMessage = ""
                var primaryMessage = ""
                
                if let lastUpdatedDate: Date = self.office365Manager.lastrefreshdate {
                    secondaryMessage = "Last updated on \(lastUpdatedDate.o365_string_from_date())"
                }
                primaryMessage = "fetched latest \(self.office365Manager.allConversations.count) messages"
                self.updateStatusWithPrimaryMessage(primaryMessage, secondaryMessage: secondaryMessage, activityInProgress: false)
            }
        }
        
    }
    

    // MARK: - Table view data source

    override func numberOfSections(in tableView: UITableView) -> Int {
        // #warning Incomplete implementation, return the number of sections
        return office365Manager.allConversations.count
    }

    override func tableView(_ tableView: UITableView, numberOfRowsInSection section: Int) -> Int {
        return 1
    }

    override func tableView(_ tableView: UITableView, heightForHeaderInSection section: Int) -> CGFloat {
        return 10
    }
    
    override func tableView(_ tableView: UITableView, didSelectRowAt indexPath: IndexPath) {
        
        let conversation : Conversation = office365Manager.allConversations[(indexPath as NSIndexPath).section]
        let outlookmessage = conversation.newestMessage() //latest message
        selectedOutlookMessage = outlookmessage
        self.performSegue(withIdentifier: "showemail", sender: nil)
        
    }
    
    override func tableView(_ tableView: UITableView, willDisplay cell: UITableViewCell, forRowAt indexPath: IndexPath) {
        
        if (indexPath as NSIndexPath).section == (office365Manager.allConversations.count - 1) && !isLoading {
            print("come to the last row")
            isLoading = true
            currentPage += 1
            performFetchMailMessages()
            
        }
    }
    
    override func tableView(_ tableView: UITableView, cellForRowAt indexPath: IndexPath) -> MessagesTableViewCell {
        
        // Configure the cell...
        let cell = tableView.dequeueReusableCell(withIdentifier: "messagecell", for: indexPath) as! MessagesTableViewCell
       
        let conversation : Conversation = office365Manager.allConversations[(indexPath as NSIndexPath).section]
        let outlookmessage = conversation.newestMessage() //latest message
        
        cell.lblSubject.text = outlookmessage.subject //subject
        cell.lblSender.text = outlookmessage.from.emailAddress.name //person's name
        cell.lblDateRecieved.text = outlookmessage.dateTimeReceived.o365_string_from_date() // date received
        cell.viewMessageState.backgroundColor = (outlookmessage.isRead && conversation.unreadMessages.count == 0) ? UIColor.clear : UIColor().o365_PrimaryColor() //if new unread email
        //if there is an attachment
        cell.imgAttachment.isHidden = (outlookmessage.hasAttachments) ? false : true
        //hide if importance is not high, otherwie show it
        cell.imgImportance.isHidden = (outlookmessage.importance == MSOutlookImportance.importance_High) ? false : true
        cell.viewMessageCount.isHidden = (conversation.messages.count > 1) ? false : true
        cell.lblMessageCount.text = "\(conversation.messages.count)"
        cell.lblBodyPreviw.text = outlookmessage.bodyPreview
        
        
       return cell
    }
    
    /****************** just pump up the action sheet *******************/
    @IBAction func showActionSheet(_ sender: UIBarButtonItem) {
        self.present(actionSheetController, animated: true, completion: nil)
    }
    /*************** END: just pump up the action sheet *****************/

    
    // MARK: - Navigation

    // In a storyboard-based application, you will often want to do a little preparation before navigation
    override func prepare(for segue: UIStoryboardSegue, sender: Any?) {
       
        if(segue.identifier == "showemail"){
            let messageDetailViewController: MessageDetailViewController =  segue.destination as!  MessageDetailViewController
            messageDetailViewController.message = selectedOutlookMessage
        }
    }
    

}
