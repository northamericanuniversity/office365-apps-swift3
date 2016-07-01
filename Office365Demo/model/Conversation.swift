//
//  Conversation.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/28/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import Foundation

class Conversation: NSObject {
    
    
    let messages: NSArray!
    let unreadMessages: NSMutableArray! = NSMutableArray()
   
    
    init(messages: NSArray){
        
        self.messages  = messages
        
        for message in messages{
            if(!message.IsRead){
                unreadMessages.addObject(message)
            }
        }
    }
    
    func oldestMessage() -> MSOutlookMessage {
        return self.messages.firstObject as! MSOutlookMessage
    }
    
    func newestMessage() -> MSOutlookMessage {
        return self.messages.lastObject as! MSOutlookMessage
    }
    
    func oldestUnreadMessage() -> MSOutlookMessage {
        return self.unreadMessages.firstObject as! MSOutlookMessage
    }
    
//    func previewMessage() -> MSOutlookMessage {
//        return self.oldestMessage() != nil ? self.oldestMessage() : self.newestMessage()
//    }
    
    func compare(object: Conversation) -> NSComparisonResult {
        return  object.newestMessage().DateTimeReceived.compare(self.newestMessage().DateTimeReceived) //descending order
    }
    
}