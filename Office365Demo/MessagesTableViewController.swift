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
    
    var currentPage : Int32 = 0
    
    
    override func viewDidLoad() {
        super.viewDidLoad()

        // Uncomment the following line to preserve selection between presentations
        // self.clearsSelectionOnViewWillAppear = false

        // Uncomment the following line to display an Edit button in the navigation bar for this view controller.
        // self.navigationItem.rightBarButtonItem = self.editButtonItem()
        
 
        
        refreshControl?.addTarget(self, action: #selector(MessagesTableViewController.updateRefreshControl), forControlEvents: UIControlEvents.ValueChanged)
        
        setupUI()
        updateRefreshControl()
        
        //updateStatusWithPrimaryMessage("Connected Successfully", secondaryMessage: "User successfully authenticated with the server", activityInProgress: false)
       
    }
    
    func setupUI(){
        
        //self.tableView.contentInset = UIEdgeInsetsMake(0, 0, 50, 0)
        
//        self.edgesForExtendedLayout = UIRectEdge.None
//        self.extendedLayoutIncludesOpaqueBars = false
//        self.automaticallyAdjustsScrollViewInsets = false
        
        
        /************************* init comments table view ********************/
        tableView.contentInset.bottom = UIApplication.sharedApplication().statusBarFrame.height + 60
//        tableView.rowHeight = UITableViewAutomaticDimension
//        tableView.estimatedRowHeight = 144
        
        let statusView: UIView = UIView(frame: CGRectMake(0, 0, 225, 32))
        let primaryStatusLabel: UILabel = UILabel(frame: CGRectMake(0, 0, 225, 16))
        let secondaryStatusLabel: UILabel = UILabel(frame: CGRectMake(0, 18, 225, 12))
        
        primaryStatusLabel.font = UIFont(name: "Halvetica", size: CGFloat(13.0))
        secondaryStatusLabel.font = UIFont(name: "Halvetica", size: CGFloat(10.0))
        
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
        
        if let lastUpdatedDate: NSDate = office365Manager.lastrefreshdate {
            let secondaryMessage = "Last updated on \(lastUpdatedDate.o365_string_from_date())"
            self.updateStatusWithPrimaryMessage("fetching messages", secondaryMessage: secondaryMessage, activityInProgress: true)
        }
 
//        office365Manager.fetchMailMessages { (messages: NSArray, error: MSODataException?) in
//            dispatch_async(dispatch_get_main_queue()) {
//                self.tableView.reloadData()
//            }
//        }
        
        office365Manager.fetchMailMessagesForPageNumber(currentPage, pageSize: 10, orderBy: "DateTimeReceived desc") { (messages: NSArray, error: MSODataException?) in
            dispatch_async(dispatch_get_main_queue()) {
                self.tableView.reloadData()
                self.refreshControl?.endRefreshing()
                self.isLoading = false
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
        let outlookmessage = conversation.newestMessage()
        
//        print("Subject: \(outlookmessage.Subject)")
//        print("conversid: \(outlookmessage.ConversationId)")
//        print("Messages Count: \(conversation.messages.count)")
//        print("")
        
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
    
    override func scrollViewWillEndDragging(scrollView: UIScrollView, withVelocity velocity: CGPoint, targetContentOffset: UnsafeMutablePointer<CGPoint>) {
        
//        if(velocity.y > 0){
//            NSLog("dragging up")
//        }else{
//            NSLog("dragging down")
//        }
    }
    
    override func scrollViewDidScroll(scrollView: UIScrollView) {
        
        let height = scrollView.frame.size.height
        let contentYoffset = scrollView.contentOffset.y
        let distanceFromBottom = scrollView.contentSize.height - contentYoffset
        
//        if(distanceFromBottom < height){
//            print("load more")
//        }
        
//        //if we reach the end of the table
//        if((scrollView.contentOffset.y + scrollView.frame.size.height) > scrollView.contentSize.height){
//            //print("reached at the end of the table")
//        }
    }
    

    /*
    // Override to support conditional editing of the table view.
    override func tableView(tableView: UITableView, canEditRowAtIndexPath indexPath: NSIndexPath) -> Bool {
        // Return false if you do not want the specified item to be editable.
        return true
    }
    */

    /*
    // Override to support editing the table view.
    override func tableView(tableView: UITableView, commitEditingStyle editingStyle: UITableViewCellEditingStyle, forRowAtIndexPath indexPath: NSIndexPath) {
        if editingStyle == .Delete {
            // Delete the row from the data source
            tableView.deleteRowsAtIndexPaths([indexPath], withRowAnimation: .Fade)
        } else if editingStyle == .Insert {
            // Create a new instance of the appropriate class, insert it into the array, and add a new row to the table view
        }    
    }
    */

    /*
    // Override to support rearranging the table view.
    override func tableView(tableView: UITableView, moveRowAtIndexPath fromIndexPath: NSIndexPath, toIndexPath: NSIndexPath) {

    }
    */

    /*
    // Override to support conditional rearranging of the table view.
    override func tableView(tableView: UITableView, canMoveRowAtIndexPath indexPath: NSIndexPath) -> Bool {
        // Return false if you do not want the item to be re-orderable.
        return true
    }
    */

    /*
    // MARK: - Navigation

    // In a storyboard-based application, you will often want to do a little preparation before navigation
    override func prepareForSegue(segue: UIStoryboardSegue, sender: AnyObject?) {
        // Get the new view controller using segue.destinationViewController.
        // Pass the selected object to the new view controller.
    }
    */

}
