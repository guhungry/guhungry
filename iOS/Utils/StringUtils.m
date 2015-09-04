//
//  StringUtils.m
//  GossipGirl
//
//  Created by Woraphot Chokratanasombat on 31/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import "StringUtils.h"

@implementation StringUtils

static BOOL isInit = NO;
static NSNumberFormatter *numberFormatter;

+(void)initGlobal
{
	if (isInit) return;

	// Initialize Number Formatter Instance
	numberFormatter = [[NSNumberFormatter alloc] init];
	[numberFormatter setNumberStyle:NSNumberFormatterDecimalStyle];

	isInit = YES;
}

/**
 * Get Not Null String if it is nil or NSNull instance return @"" instead
 *
 * @param number The number in NSNumber instance
 * @return Return original object if not nil and not NSNull, else return @""
 * @author Woraphot Chokratanasombat
 * @since 2015-09-01
 * @updated 2015-09-01
 */
+(NSString *)getNotNullString:(id)object
{
	return [self getNotNullString:object withDefault:@""];
}

/**
 * Get Not Null String if it is nil or NSNull instance return @"" instead
 *
 * @param number The number in NSNumber instance
 * @return Return original object if not nil and not NSNull, else return @""
 * @author Woraphot Chokratanasombat
 * @since 2015-09-01
 * @updated 2015-09-01
 */
+(NSString *)getNotNullString:(id)object withDefault:(NSString *)value
{
	if (object == nil || [object isKindOfClass:[NSNull class]]) return value;
	else return object;
}

/**
 * Format NSNumber instance in to #,##0 format
 *
 * @param number The number in NSNumber instance
 * @return Formatted number string
 * @author Woraphot Chokratanasombat
 * @since 2015-09-04
 * @updated 2015-09-04
 */
+(NSString *)formatNumber:(NSNumber*)number
{
	[StringUtils initGlobal];

	return [numberFormatter stringFromNumber:number];
}

@end
