//
//  LoginViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/21/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class LoginViewController: UIViewController {

    var baseController = Office365ClientFetcher()
    var serviceEndpointLookup = NSMutableDictionary()
    
    override func viewDidLoad() {
        super.viewDidLoad()

        // Do any additional setup after loading the view.
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }
    
    
    @IBAction func connectToOffice365(_ sender: AnyObject) {
        
        // Connect to the service by discovering the service endpoints and authorizing
        // the application to access the user's email. This will store the user's
        // service URLs in a property list to be accessed when calls are made to the
        // service. This results in two calls: one to authenticate, and one to get the
        // URLs. ADAL will cache the access and refresh tokens so you won't need to
        // provide credentials unless you sign out.
        
        // Get the discovery client. First time this is ran you will be prompted
        // to provide your credentials which will authenticate you with the service.
        // The application will get an access token in the response.
        
        baseController.fetchDiscoveryClient { (discoveryClient) -> () in
            let servicesInfoFetcher = discoveryClient.getservices()
            
            // Call the Discovery Service and get back an array of service endpoint information
            
            let servicesTask = servicesInfoFetcher?.read{(serviceEndPointObjects:[Any]?, error:MSODataException?) -> Void in
                
                
                let serviceEndpoints = serviceEndPointObjects as! [MSDiscoveryServiceInfo]
                
                if (serviceEndpoints.count > 0) {
                    // Here is where we cache the service URLs returned by the Discovery Service. You may not
                    // need to call the Discovery Service again until either this cache is removed, or you
                    // get an error that indicates that the endpoint is no longer valid.
                    
                    var serviceEndpointLookup = [AnyHashable: Any]()
                    for serviceEndpoint in serviceEndpoints {
                        serviceEndpointLookup[serviceEndpoint.capability] = serviceEndpoint.serviceEndpointUri
                         serviceEndpointLookup[serviceEndpoint.capability+"ResourceID"] = serviceEndpoint.serviceResourceId
                        
                        print("serviceEndpoint.capability: \(serviceEndpoint.capability) serviceEndpointUri: \(serviceEndpoint.serviceEndpointUri) serviceResourceID: \(serviceEndpoint.serviceResourceId)")
                        
                    }
                    
                    // Keep track of the service endpoints in the user defaults
                    let userDefaults = UserDefaults.standard
                    userDefaults.set(serviceEndpointLookup, forKey: "O365ServiceEndpoints")
                    userDefaults.synchronize()
                    
                    DispatchQueue.main.async {
                        let userEmail = userDefaults.string(forKey: "demo_email")!
                        print("user email: \(userEmail)")
                        
//                        let mainTabBar : UITabBarController = self.storyboard?.instantiateViewControllerWithIdentifier("mainTabBarController") as! UITabBarController;
                        
                        let navMessagesView : UINavigationController = self.storyboard?.instantiateViewController(withIdentifier: "navMessagesView") as! UINavigationController
                        
                        self.present(navMessagesView, animated: true, completion: nil)
                        
                    }
                }
                    
                else {
                    DispatchQueue.main.async {
                        //log("Error in the authentication: \(error)")
                        let alert = UIAlertController(title: "Error", message:"Authentication failed. This may be because the Internet connection is offline  or perhaps the credentials are incorrect. Check the log for errors and try again.", preferredStyle: .alert)
                        let action = UIAlertAction(title: "OK", style: .default) { _ in
                            // Put here any code that you would like to execute when
                            // the user taps that OK button (may be empty in your case if that's just
                            // an informative alert)
                        }
                        alert.addAction(action)
                        
                        self.present(alert, animated: true){}
                        
                    }
                }
            }
            
            servicesTask?.resume()
        }

        
    }
    

    /*
    // MARK: - Navigation

    // In a storyboard-based application, you will often want to do a little preparation before navigation
    override func prepareForSegue(segue: UIStoryboardSegue, sender: AnyObject?) {
        // Get the new view controller using segue.destinationViewController.
        // Pass the selected object to the new view controller.
    }
    */

}
