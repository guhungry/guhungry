//
//  DateUtils.m
//  GossipGirl
//
//  Created by Woraphot Chokratanasombat on 05/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import "DateUtils.h"

@implementation DateUtils


static BOOL isInit = NO;
static NSCalendar *gregorian;
static NSDateFormatter *formatWebService;

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
 * @param Date String in yyyy-MM-dd HH:mm:ss format
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
 * @param NSDate instance
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

@end
