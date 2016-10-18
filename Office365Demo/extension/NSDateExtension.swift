//
//  NSDateExtension.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/27/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import Foundation

extension Date{
    
 
    func o365_string_from_date() -> String{
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "MMMM d, YYYY 'at' HH:mm a"
        return dateFormatter.string(from: self)
    }
}
