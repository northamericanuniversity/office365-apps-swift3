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

//client
#import <Office365/MSOutlookClient.h>
#import <Office365/MSSharePointClient.h>

//user
#import <Office365/MSOutlookUserCollectionFetcher.h>
#import <Office365/MSOutlookUserFetcher.h>
#import <Office365/MSOutlookUserOperations.h>


//message
#import <Office365/MSOutlookMessageCollectionFetcher.h>
#import <Office365/MSOutlookMessageFetcher.h>
#import <Office365/MSOutlookMessageOperations.h>

//event
#import <Office365/MSOutlookEventCollectionFetcher.h>
#import <Office365/MSOutlookEventFetcher.h>
#import <Office365/MSOutlookEventOperations.h>

//contact
#import <Office365/MSOutlookContactCollectionFetcher.h>
#import <Office365/MSOutlookContactFetcher.h>

//sharepoint
#import <Office365/MSSharePointItemCollectionFetcher.h>

#endif /* Bridging_Header_h */
