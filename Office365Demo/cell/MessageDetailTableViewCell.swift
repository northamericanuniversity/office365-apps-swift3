//
//  MessageDetailTableViewCell.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 10/21/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class MessageDetailTableViewCell: UITableViewCell {

   
    @IBOutlet weak var mBodyView: UIWebView!
    @IBOutlet weak var mReply: UIButton!
    @IBOutlet weak var mReplyAll: UIButton!
    @IBOutlet weak var mForward: UIButton!
    
    override func awakeFromNib() {
        super.awakeFromNib()
        // Initialization code
    }

    override func setSelected(_ selected: Bool, animated: Bool) {
        super.setSelected(selected, animated: animated)

        // Configure the view for the selected state
    }

}
