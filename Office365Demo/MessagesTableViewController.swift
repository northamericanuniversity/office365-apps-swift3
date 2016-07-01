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
    
    override func viewDidLoad() {
        super.viewDidLoad()

        refreshControl?.addTarget(self, action: #selector(MessagesTableViewController.updateRefreshControl), forControlEvents: UIControlEvents.ValueChanged)
        
        setupActionSheet()
        setupUI()
        updateRefreshControl()
    }
    
     /****************************** init action sheet ******************************/
    func setupActionSheet(){
       
        actionSheetController = UIAlertController(title: nil, message: "Choose Action", preferredStyle: .ActionSheet)
        let cancelAction = UIAlertAction(title: "Cancel", style: .Cancel) { (action) in
            
        }
        actionSheetController.addAction(cancelAction)//add action to the alert controller
        
        //sign out
        let signoutAction = UIAlertAction(title: "Sign out", style: .Default) { (action) in
            
            let signoutAlert = UIAlertController(title: "Sign out", message: "Do you want to sign out?", preferredStyle: UIAlertControllerStyle.Alert)
            
            //delete refused, don't delete this post
            signoutAlert.addAction(UIAlertAction(title: "Cancel", style: .Cancel, handler: { (action: UIAlertAction) in
                //don't do anything here, the dialog simply closes itself
            }))
            
            //sign out now
            signoutAlert.addAction(UIAlertAction(title: "Yes", style: .Default, handler: { (action: UIAlertAction) in
                
                // Clear the access and refresh tokens from the credential cache. You need to clear cookies
                // since ADAL uses information stored in the cookies to get a new access token.
                let authenticationManager:AuthenticationManager = AuthenticationManager.sharedInstance
                authenticationManager.clearCredentials()
                
                //jump back to the login view
                let loginView : LoginViewController = self.storyboard?.instantiateViewControllerWithIdentifier("loginView") as! LoginViewController;
                self.presentViewController(loginView, animated: true, completion: nil)
                
            }))
            
            self.presentViewController(signoutAlert, animated: true, completion: nil)
            
        }
        actionSheetController.addAction(signoutAction)//add action to the alert controller
       
    }
     /****************************** END: init action sheet *****************************/
    
    /**************************************** init UI  **********************************/
    func setupUI(){

        self.navigationController?.toolbarHidden = false //show toolbar
        
        let statusView: UIView = UIView(frame: CGRectMake(0, 0, 225, 36))
        let primaryStatusLabel: UILabel = UILabel(frame: CGRectMake(0, 0, 225, 16))
        let secondaryStatusLabel: UILabel = UILabel(frame: CGRectMake(0, 18, 225, 12))
        
        primaryStatusLabel.font = UIFont.systemFontOfSize(13)
        secondaryStatusLabel.font = UIFont.systemFontOfSize(10)
        
        primaryStatusLabel.textAlignment = NSTextAlignment.Center
        secondaryStatusLabel.textAlignment = NSTextAlignment.Center
        
        primaryStatusLabel.textColor = UIColor().o365_PrimaryColor()
        secondaryStatusLabel.textColor = UIColor.grayColor()
        
        statusView.addSubview(primaryStatusLabel)
        statusView.addSubview(secondaryStatusLabel)
        
        self.primaryStatusLabel = primaryStatusLabel
        self.secondaryStatusLabel = secondaryStatusLabel
        
        self.statusBarButtonItem.customView = statusView
        
        let activityIndicator : UIActivityIndicatorView = UIActivityIndicatorView(activityIndicatorStyle: UIActivityIndicatorViewStyle.Gray)
        activityIndicator.color = UIColor().o365_PrimaryColor()
        
        self.activityBarButtonItem.customView = activityIndicator
        self.activityIndicator = activityIndicator
        
        refreshControl?.backgroundColor = UIColor().o365_PrimaryColor()
        refreshControl?.tintColor = UIColor.whiteColor()
    }
    /************************************** END: init UI  ********************************/

    //update refresh control
    func updateRefreshControl(){
      
        if let lastUpdatedDate: NSDate = office365Manager.lastrefreshdate {
            let lastUpdatedTitle = "Last updated on \(lastUpdatedDate.o365_string_from_date())"
            refreshControl?.attributedTitle = NSAttributedString(string:   lastUpdatedTitle)
        }
        
        currentPage = 0
        performFetchMailMessages()
        
    }
    
    //update status
    func updateStatusWithPrimaryMessage(primaryMessage: String,
                                        secondaryMessage: String,
                                        activityInProgress: Bool){
        
        dispatch_async(dispatch_get_main_queue()) {
            
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
        if let lastUpdatedDate: NSDate = office365Manager.lastrefreshdate {
            secondaryMessage = "Last updated on \(lastUpdatedDate.o365_string_from_date())"
        }
        
        self.updateStatusWithPrimaryMessage("fetching messages", secondaryMessage: secondaryMessage, activityInProgress: true)
 
        /*********************** Alternative way to get only 10 messags by default ****************
        office365Manager.fetchMailMessages { (messages: NSArray, error: MSODataException?) in
            dispatch_async(dispatch_get_main_queue()) {
                self.tableView.reloadData()
            }
        }
        ******************* END: Alternative way to get only 10 messags by default ****************/
        
        
        //get email messages by page number
        office365Manager.fetchMailMessagesForPageNumber(currentPage, pageSize: 10, orderBy: "DateTimeReceived desc") { (messages: NSArray, error: MSODataException?) in
            dispatch_async(dispatch_get_main_queue()) {
                
                self.tableView.reloadData()
                self.refreshControl?.endRefreshing()
                self.isLoading = false
                var secondaryMessage = ""
                var primaryMessage = ""
                
                if let lastUpdatedDate: NSDate = self.office365Manager.lastrefreshdate {
                    secondaryMessage = "Last updated on \(lastUpdatedDate.o365_string_from_date())"
                }
                primaryMessage = "fetched latest \(self.office365Manager.allConversations.count) messages"
                self.updateStatusWithPrimaryMessage(primaryMessage, secondaryMessage: secondaryMessage, activityInProgress: false)
            }
        }
    }
    

    // MARK: - Table view data source

    override func numberOfSectionsInTableView(tableView: UITableView) -> Int {
        // #warning Incomplete implementation, return the number of sections
        return office365Manager.allConversations.count
    }

    override func tableView(tableView: UITableView, numberOfRowsInSection section: Int) -> Int {
        return 1
    }

    override func tableView(tableView: UITableView, heightForHeaderInSection section: Int) -> CGFloat {
        return 10
    }
    
    override func tableView(tableView: UITableView, willDisplayCell cell: UITableViewCell, forRowAtIndexPath indexPath: NSIndexPath) {
        
        if indexPath.section == (office365Manager.allConversations.count - 1) && !isLoading {
            print("come to the last row")
            isLoading = true
            currentPage += 1
            performFetchMailMessages()
            
        }
    }
    
    override func tableView(tableView: UITableView, cellForRowAtIndexPath indexPath: NSIndexPath) -> MessagesTableViewCell {
        
        // Configure the cell...
        let cell = tableView.dequeueReusableCellWithIdentifier("messagecell", forIndexPath: indexPath) as! MessagesTableViewCell
       
        let conversation : Conversation = office365Manager.allConversations[indexPath.section]
        let outlookmessage = conversation.newestMessage() //latest message
        
        cell.lblSubject.text = outlookmessage.Subject //subject
        cell.lblSender.text = outlookmessage.From.EmailAddress.Name //person's name
        cell.lblDateRecieved.text = outlookmessage.DateTimeReceived.o365_string_from_date() // date received
        cell.viewMessageState.backgroundColor = (outlookmessage.IsRead && conversation.unreadMessages.count == 0) ? UIColor.clearColor() : UIColor().o365_PrimaryColor() //if new unread email
        //if there is an attachment
        cell.imgAttachment.hidden = (outlookmessage.HasAttachments) ? false : true
        //hide if importance is not high, otherwie show it
        cell.imgImportance.hidden = (outlookmessage.Importance == MSOutlookImportance.Importance_High) ? false : true
        cell.viewMessageCount.hidden = (conversation.messages.count > 1) ? false : true
        cell.lblMessageCount.text = "\(conversation.messages.count)"
        cell.lblBodyPreviw.text = outlookmessage.BodyPreview
        
        
       return cell
    }
    
    /****************** just pump up the action sheet *******************/
    @IBAction func showActionSheet(sender: UIBarButtonItem) {
        self.presentViewController(actionSheetController, animated: true, completion: nil)
    }
    /*************** END: just pump up the action sheet *****************/

    /*
    // MARK: - Navigation

    // In a storyboard-based application, you will often want to do a little preparation before navigation
    override func prepareForSegue(segue: UIStoryboardSegue, sender: AnyObject?) {
        // Get the new view controller using segue.destinationViewController.
        // Pass the selected object to the new view controller.
    }
    */

}
