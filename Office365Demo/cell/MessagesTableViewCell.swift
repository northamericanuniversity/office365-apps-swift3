//
//  MessagesTableViewCell.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/27/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class MessagesTableViewCell: UITableViewCell {
    @IBOutlet weak var viewMessageState: UIView!
    @IBOutlet weak var lblSender: UILabel!
    @IBOutlet weak var lblSubject: UILabel!
  
    @IBOutlet weak var lblBodyPreviw: UILabel!
    @IBOutlet weak var lblDateRecieved: UILabel!
  
    @IBOutlet weak var imgAttachment: UIImageView!
    @IBOutlet weak var imgImportance: UIImageView!
    
    @IBOutlet weak var viewMessageCount: UIView!
    
    @IBOutlet weak var lblMessageCount: UILabel!
    
    override func awakeFromNib() {
        super.awakeFromNib()
        // Initialization code
    }

    override func setSelected(_ selected: Bool, animated: Bool) {
        super.setSelected(selected, animated: animated)

        // Configure the view for the selected state
    }

}
