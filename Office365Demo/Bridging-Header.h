//
//  Bridging-Header.h
//  Office365Demo
//
//  Created by Mehmet Sen on 6/20/16.
//  Copyright © 2016 Mehmet Sen. All rights reserved.
//

#ifndef Bridging_Header_h
#define Bridging_Header_h

#import <ADALiOS/ADAuthenticationContext.h>
#import <ADALiOS/ADAuthenticationSettings.h>
#import <ADALiOS/ADAuthenticationError.h>
#import <Office365/office365_discovery_sdk.h>


//odata
#import <office365_odata_base/office365_odata_base.h>


//discovery
#import <office365_discovery_sdk/office365_discovery_sdk.h>




//exchange server -> outlook
#import <office365_exchange_sdk/office365_exchange_sdk.h>
//#import <Office365/MSOutlookClient.h>


//user
//#import <Office365/MSOutlookUserCollectionFetcher.h>
//#import <Office365/MSOutlookUserFetcher.h>
//#import <Office365/MSOutlookUserOperations.h>


//message
//#import <Office365/MSOutlookMessageCollectionFetcher.h>
//#import <Office365/MSOutlookMessageFetcher.h>
//#import <Office365/MSOutlookMessageOperations.h>


//outlook

//#import <Office365/MSOutlookEventCollectionFetcher.h>
//#import <Office365/MSOutlookEventFetcher.h>
//#import <Office365/MSOutlookEventOperations.h>

//contact
//#import <Office365/MSOutlookContactCollectionFetcher.h>
//#import <Office365/MSOutlookContactFetcher.h>

//sharepoint
#import <office365_files_sdk/office365_files_sdk.h>
//#import <Office365/MSSharePointItemCollectionFetcher.h>
//#import <Office365/MSSharePointClient.h>







#endif /* Bridging_Header_h */
