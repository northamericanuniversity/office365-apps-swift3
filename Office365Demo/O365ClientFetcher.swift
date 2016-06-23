// *********************************************************
//
// O365-iOS-Connect, https://github.com/OfficeDev/O365-iOS-Connect
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************

import Foundation

/***************************** i.e End Points ************************************
 Capability:  MyFiles
 EndPoint URI:  https://nau3203-my.sharepoint.com/_api/v1.0/me
 ResourceID: https://nau3203-my.sharepoint.com/
 
 Capability: MyFiles
 EndPoint URI: https://nau3203-my.sharepoint.com/_api/v2.0/me
 ResourceID: https://nau3203-my.sharepoint.com/
 
 Capability: RootSite
 EndPoint URI: https://nau3203.sharepoint.com/_api
 ResourceID: https://nau3203.sharepoint.com/
 
 Capability: Mail
 EndPoint URI: https://outlook.office365.com/api/v1.0
 ResourceID: https://outlook.office365.com/
 
 Capability: Contacts
 EndPoint URI: https://outlook.office365.com/api/v1.0
 ResourceID: https://outlook.office365.com/
 
 Capability:  Directory
 EndPoint URI: https://graph.windows.net/na.edu/
 ResourceID: https://graph.windows.net/
 
 Capability:  Notes
 EndPoint URI: https://www.onenote.com/api/beta
 ResourceID: https://onenote.com/
 
 ***********************************************************************************/


class Office365ClientFetcher {
    
    // Gets the Outlook Services client. This will authenticate a user with the service
    // and get the application an access and refresh token to act on behalf of the user.
    // The access and refresh token will be cached. The next time a user attempts
    // to access the service, the access token will be used. If the access token
    // has expired, the client will use the refresh token to get a new access token.
    // If the refresh token has expired, then ADAL will get the authorization code
    // from the cookie cache and use that to get a new access and refresh token.
    
    
    func fetchOutlookClient(completionHandler:((outlookClient: MSOutlookClient) -> Void)) {
        // Get an instance of the authentication controller.
        let authenticationManager = AuthenticationManager.sharedInstance
        
        // The first time this application is run, the authentication manager will send a request
        // to the authority which will redirect you to a login page. You will provide your credentials
        // and the response will contain your refresh and access tokens. The second time this
        // application is run, and assuming you didn't clear your token cache, the authentication
        // manager will use the access or refresh token in the cache to authenticate client requests.
        // This will result in a call to the service if you need to get an access token.
        
        let userDefaults = NSUserDefaults.standardUserDefaults()
        if let serviceEndpoints = userDefaults.dictionaryForKey("O365ServiceEndpoints") {
        
            //serviceEndpoints["MailResourceID"] ==> "https://outlook.office365.com/"
            if let resourceID: AnyObject = serviceEndpoints["MailResourceID"] {
                
                authenticationManager.acquireAuthTokenWithResourceId(resourceID as! String) {
                    (authenticated:Bool) -> Void in
                    
                    if (authenticated) {
                        
                        // serviceEndpoints["Mail"] => "https://outlook.office365.com/api/v1.0"
                        if let serviceEndpointUrl: AnyObject = serviceEndpoints["Mail"] {
                            // Gets the MSOutlookClient with the URL for the Mail service.
                            let outlookClient = MSOutlookClient(url: serviceEndpointUrl as! String, dependencyResolver: authenticationManager.dependencyResolver)
                            completionHandler(outlookClient: outlookClient)
                        }
                        
                    }
                    else {
                        // Display an alert in case of an error
                        dispatch_async(dispatch_get_main_queue()) {
                            NSLog("Error in the authentication")
                            let alert = UIAlertController(title: "Error", message:"Authentication failed. Check the log for errors.", preferredStyle: .Alert)
                            let action = UIAlertAction(title: "OK", style: .Default) { _ in
                                // Put here any code that you would like to execute when
                                // the user taps that OK button (may be empty in your case if that's just
                                // an informative alert)
                            }
                            alert.addAction(action)
                            let rootVC = UIApplication.sharedApplication().keyWindow?.rootViewController
                            rootVC?.presentViewController(alert, animated: true){}
                            
                        }
                    }
                }

            }
            
        }else{
            print("No Service End Points!")
            //TO DO: do Alert
        }
        
    }
    
    // Gets the SharePointClient which is used to access sharepoint services such as Files
    func fetchSharePointClient(completionHandler:((sharePointClient: MSSharePointClient) -> Void)){
        
        // Get an instance of the authentication controller.
        let authenticationManager:AuthenticationManager = AuthenticationManager.sharedInstance
        
        let userDefaults = NSUserDefaults.standardUserDefaults()
        if let serviceEndpoints = userDefaults.dictionaryForKey("O365ServiceEndpoints") {
            
            //serviceEndpoints["MyFilesResourceID"] ==> "https://nau3203-my.sharepoint.com/"
            if let resourceID: AnyObject = serviceEndpoints["MyFilesResourceID"] {
                
                authenticationManager.acquireAuthTokenWithResourceId(resourceID as! String) {
                    (authenticated:Bool) -> Void in
                    
                    if (authenticated) {
                        
                        if let serviceEndpointUrl: AnyObject = serviceEndpoints["MyFiles"] {
                            // Gets the MSSharePointClient with the URL for the MyFiles service.
                            let sharePointClient = MSSharePointClient(url: serviceEndpointUrl as! String, dependencyResolver: authenticationManager.dependencyResolver)
                            completionHandler(sharePointClient: sharePointClient)
                        }
                    }else{
                        // Display an alert in case of an error
                        dispatch_async(dispatch_get_main_queue()) {
                            NSLog("Error in the authentication")
                            let alert = UIAlertController(title: "Error", message:"Authentication failed. Check the log for errors.", preferredStyle: .Alert)
                            let action = UIAlertAction(title: "OK", style: .Default) { _ in
                                // Put here any code that you would like to execute when
                                // the user taps that OK button (may be empty in your case if that's just
                                // an informative alert)
                            }
                            alert.addAction(action)
                            let rootVC = UIApplication.sharedApplication().keyWindow?.rootViewController
                            rootVC?.presentViewController(alert, animated: true){}
                            
                        }
                    }
                }
            }
        }
        
        
        
        
    }
    
    
    // Gets the DiscoveryClient which is used to discover the service endpoints
    func fetchDiscoveryClient(completionHandler:((discoveryClient: MSDiscoveryClient) -> Void)) {
        
        // Get an instance of the authentication controller.
        let authenticationManager:AuthenticationManager = AuthenticationManager.sharedInstance
        
        // The first time this application is run, the authentication manager will send a request
        // to the authority which will redirect you to a login page. You will provide your credentials
        // and the response will contain your refresh and access tokens. The second time this
        // application is run, and assuming you didn't clear your token cache, the authentication
        // manager will use the access or refresh token in the cache to authenticate client requests.
        // This will result in a call to the service if you need to get an access token.
        
        authenticationManager.acquireAuthTokenWithResourceId("https://api.office.com/discovery/") {
            (authenticated:Bool) -> Void in
            
            if (authenticated) {
                // Gets the MSDiscoveryClient with the URL for the Discovery service.
                let discoveryClient = MSDiscoveryClient(url: "https://api.office.com/discovery/v2.0/me/", dependencyResolver: authenticationManager.dependencyResolver)
                completionHandler(discoveryClient: discoveryClient)
            }
            else {
                // Display an alert in case of an error
                dispatch_async(dispatch_get_main_queue()) {
                    NSLog("Error in the authentication")
                    let alert = UIAlertController(title: "Error", message:"Authentication failed. Check the log for errors.", preferredStyle: .Alert)
                    let action = UIAlertAction(title: "OK", style: .Default) { _ in
                        // Put here any code that you would like to execute when
                        // the user taps that OK button (may be empty in your case if that's just
                        // an informative alert)
                    }
                    alert.addAction(action)
                    let rootVC = UIApplication.sharedApplication().keyWindow?.rootViewController
                    rootVC?.presentViewController(alert, animated: true){}
                    
                }
            }
        }
    }
    
    
}
