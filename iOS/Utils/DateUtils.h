//
//  DateUtils.h
//  GossipGirl
//
//  Created by Woraphot Chokratanasombat on 05/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface DateUtils : NSObject

#pragma mark Web Service Date String
+(NSDate *)dateFromWebServiceString:(NSString *)date;
+(NSString *)stringWebServiceFromDate:(NSDate *)date;

@end
