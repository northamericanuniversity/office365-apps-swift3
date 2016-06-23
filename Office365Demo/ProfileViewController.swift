//
//  ProfileViewController.swift
//  Office365Demo
//
//  Created by Mehmet Sen on 6/21/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

import UIKit

class ProfileViewController: UIViewController {

    override func viewDidLoad() {
        super.viewDidLoad()

        // Do any additional setup after loading the view.
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }
    

    @IBAction func signout(sender: AnyObject) {
        
        // Clear the access and refresh tokens from the credential cache. You need to clear cookies
        // since ADAL uses information stored in the cookies to get a new access token.
        let authenticationManager:AuthenticationManager = AuthenticationManager.sharedInstance
        authenticationManager.clearCredentials()
        
        //jump back to the login view
        let loginView : LoginViewController = self.storyboard?.instantiateViewControllerWithIdentifier("loginView") as! LoginViewController;
        self.presentViewController(loginView, animated: true, completion: nil)

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
