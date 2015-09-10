//
//  DateUtils.m
//  iOS Utilities
//
//  Created by Woraphot Chokratanasombat on 05/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import "DateUtils.h"

@implementation DateUtils


static BOOL isInit = NO;
static NSCalendar *gregorian;
static NSDateFormatter *formatWebService;
const NSCalendarUnit dateOnlyUnit = NSYearCalendarUnit | NSMonthCalendarUnit | NSDayCalendarUnit;

+(void)initGlobal
{
	if (isInit) return;

	// Initialize Calendar Instance
	gregorian = [[NSCalendar alloc] initWithCalendarIdentifier:NSGregorianCalendar];

	// Initialize Date Formatter Instance
	formatWebService = [[NSDateFormatter alloc] init];
	[formatWebService setFormatterBehavior:NSDateFormatterBehavior10_4];
	[formatWebService setCalendar:gregorian];
	[formatWebService setDateFormat:@"yyyy-MM-dd HH:mm:ss"];

	isInit = YES;
}

#pragma mark Web Service Date String Function
/**
 * Convert Web Service Date String to NSDate instance (Format yyyy-MM-dd HH:mm:ss) using NSGregorianCalendar
 *
 * @param date Date String in yyyy-MM-dd HH:mm:ss format
 * @return Converted NSDate instance
 * @author Woraphot Chokratanasombat
 * @since 2015-08-07
 * @updated 2015-08-07
 */
+(NSDate *)dateFromWebServiceString:(NSString *)date
{
	[DateUtils initGlobal];
	return [formatWebService dateFromString:date];
}

/**
 * Convert NSDate instance to Web Service Date String (Format yyyy-MM-dd HH:mm:ss) using NSGregorianCalendar
 *
 * @param date NSDate instance
 * @return Converted Date String in yyyy-MM-dd HH:mm:ss format
 * @author Woraphot Chokratanasombat
 * @since 2015-08-07
 * @updated 2015-08-07
 */
+(NSString *)stringWebServiceFromDate:(NSDate *)date
{
	[DateUtils initGlobal];
	return [formatWebService stringFromDate:date];
}

#pragma mark Date Diff Functions
/**
 * Date Diff From Current Timestamp (Current Timestamp < Date)<br />
 * Example : Count down to time out
 *
 * @param date NSDate instance
 * @param components Date Components to be Compared
 * @return Date Diff in NSDateComponents instance
 * @author Woraphot Chokratanasombat
 * @since 2015-08-07
 * @updated 2015-08-07
 */
+(NSDateComponents *)dateDiffFromCurrentTimestamp:(NSDate *)date components:(NSCalendarUnit)components
{
	[DateUtils initGlobal];
	return [gregorian components:components fromDate:[NSDate date] toDate:date options:0];
}

/**
 * Date Diff To Current Timestamp (Current Timestamp > Date)<br />
 * Example : Chat post
 *
 * @param date NSDate instance
 * @param components Date Components to be Compared
 * @return Date Diff in NSDateComponents instance
 * @author Woraphot Chokratanasombat
 * @since 2015-08-07
 * @updated 2015-08-07
 */
+(NSDateComponents *)dateDiffToCurrentTimestamp:(NSDate *)date components:(NSCalendarUnit)components
{
	[DateUtils initGlobal];
	return [gregorian components:components fromDate:date toDate:[NSDate date] options:0];
}

/**
 * Date Diff From Current Date (Current Date < Date)
 *
 * @param date NSDate instance
 * @param components Date Components to be Compared
 * @return Date Diff in NSDateComponents instance
 * @author Woraphot Chokratanasombat
 * @since 2015-08-27
 * @updated 2015-08-27
 */
+(NSDateComponents *)dateDiffFromCurrentDate:(NSDate *)date components:(NSCalendarUnit)components
{
	[DateUtils initGlobal];
	NSDateComponents *to = [gregorian components:dateOnlyUnit fromDate:[NSDate date]];
	return [gregorian components:components fromDate:[gregorian dateFromComponents:to] toDate:date options:0];
}

/**
 * Date Diff To Current Date (Current Date > Date)<br />
 * Example : Calculate Age
 *
 * @param date NSDate instance
 * @param components Date Components to be Compared
 * @return Date Diff in NSDateComponents instance
 * @author Woraphot Chokratanasombat
 * @since 2015-08-27
 * @updated 2015-08-27
 */
+(NSDateComponents *)dateDiffToCurrentDate:(NSDate *)date components:(NSCalendarUnit)components
{
	[DateUtils initGlobal];
	NSDateComponents *to = [gregorian components:dateOnlyUnit fromDate:[NSDate date]];
	return [gregorian components:components fromDate:date toDate:[gregorian dateFromComponents:to] options:0];
}

@end
