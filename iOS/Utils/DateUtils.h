//
//  DateUtils.h
//  iOS Utilities
//
//  Created by Woraphot Chokratanasombat on 05/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface DateUtils : NSObject

#pragma mark Web Service Date String
+(NSDate *)dateFromWebServiceString:(NSString *)date;
+(NSString *)stringWebServiceFromDate:(NSDate *)date;

#pragma mark Date Diff Functions
// Timestamp
+(NSDateComponents *)dateDiffFromCurrentTimestamp:(NSDate *)date components:(NSCalendarUnit)components;
+(NSDateComponents *)dateDiffToCurrentTimestamp:(NSDate *)date components:(NSCalendarUnit)components;
// Date
+(NSDateComponents *)dateDiffFromCurrentDate:(NSDate *)date components:(NSCalendarUnit)components;
+(NSDateComponents *)dateDiffToCurrentDate:(NSDate *)date components:(NSCalendarUnit)components;

@end
