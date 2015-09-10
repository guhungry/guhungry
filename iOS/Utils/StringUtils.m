//
//  StringUtils.m
//  iOS Utilities
//
//  Created by Woraphot Chokratanasombat on 31/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import "StringUtils.h"

@implementation StringUtils

static BOOL isInit = NO;
static NSNumberFormatter *decimalFormatter;
static NSNumberFormatter *integerFormatter;

+(void)initGlobal
{
	if (isInit) return;

	// Initialize Decimal Number Formatter Instance
	decimalFormatter = [[NSNumberFormatter alloc] init];
	[decimalFormatter setNumberStyle:NSNumberFormatterDecimalStyle];
	[decimalFormatter setMaximumFractionDigits:2];
	[decimalFormatter setMinimumFractionDigits:2];

	// Initialize Integer Number Formatter Instance
	integerFormatter = [[NSNumberFormatter alloc] init];
	[integerFormatter setNumberStyle:NSNumberFormatterDecimalStyle];

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

#pragma mark Format Number Functions
/**
 * Format NSNumber instance in to #,##0.00 format
 *
 * @param number The number in NSNumber instance
 * @return Formatted number string
 * @author Woraphot Chokratanasombat
 * @since 2015-09-04
 * @updated 2015-09-04
 */
+(NSString *)formatDecimalNumber:(NSNumber*)number
{
	[StringUtils initGlobal];

	return [decimalFormatter stringFromNumber:number];
}

/**
 * Format NSString instance in to #,##0.00 format
 *
 * @param number The number in NSString instance
 * @return Formatted number string
 * @author Woraphot Chokratanasombat
 * @since 2015-09-04
 * @updated 2015-09-04
 */
+(NSString *)formatDecimalNumberFromString:(NSString *)number
{
	[StringUtils initGlobal];

	return [StringUtils formatDecimalNumber:[decimalFormatter numberFromString:number]];
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
+(NSString *)formatIntegerNumber:(NSNumber*)number
{
	[StringUtils initGlobal];

	return [integerFormatter stringFromNumber:number];
}

/**
 * Format NSString instance in to #,##0 format
 *
 * @param number The number in NSString instance
 * @return Formatted number string
 * @author Woraphot Chokratanasombat
 * @since 2015-09-04
 * @updated 2015-09-04
 */
+(NSString *)formatIntegerNumberFromString:(NSString *)number
{
	[StringUtils initGlobal];

	return [StringUtils formatIntegerNumber:[integerFormatter numberFromString:number]];
}

@end
