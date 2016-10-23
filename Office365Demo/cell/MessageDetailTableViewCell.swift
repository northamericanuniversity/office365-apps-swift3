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
    
    
    override func awakeFromNib() {
        super.awakeFromNib()
        // Initialization code
    }

    override func setSelected(_ selected: Bool, animated: Bool) {
        super.setSelected(selected, animated: animated)

        // Configure the view for the selected state
    }

}
