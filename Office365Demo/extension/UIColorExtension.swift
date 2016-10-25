//
//  UIColorExtension.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/27/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

/*********************************************************************
 How to use:
 
 UIColor().o365_PrimaryColor()
 *********************************************************************/

extension UIColor {

    func o365_PrimaryColor() -> UIColor {
        //return UIColor(red: 1.0, green: 0.6, blue: 0.2, alpha: 1.0)
        return UIColor(red: 0x2f/255, green: 0x51/255, blue: 0x89/255, alpha: 1.0)
    }
    
    func o365_PrimaryHighlightColor() -> UIColor {
        return UIColor(red: 1.0, green: 0.75, blue: 0.35, alpha: 1.0)
    }
    
    func o365_UnreadMessageColor() -> UIColor {
        return UIColor(red: 0.15, green: 0.60, blue: 0.72, alpha: 1.0)
    }
    
    func o365_defaultMessageColor() -> UIColor {
        return UIColor(red: 0.74, green: 0.74, blue: 0.74, alpha: 1.0)
    }
}
