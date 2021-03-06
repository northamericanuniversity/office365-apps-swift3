//
//  Conversation.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/28/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import Foundation

class Conversation: NSObject {
    
    var messages: NSArray!
    let unreadMessages: NSMutableArray! = NSMutableArray()
   
    init(messages: NSArray){
        self.messages  = messages
        for message in messages{
            if(!(message as AnyObject).isRead){
                unreadMessages.add(message)
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
    
    func sortMessages(){
        self.messages = self.messages.sorted(by: { ($0 as! MSOutlookMessage).dateTimeReceived > ($1 as! MSOutlookMessage).dateTimeReceived }) as NSArray!
    }
    
    func compare(_ object: Conversation) -> ComparisonResult {
        return  object.newestMessage().dateTimeReceived.compare(self.newestMessage().dateTimeReceived) //descending order
    }
    
}

