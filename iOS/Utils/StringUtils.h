//
//  StringUtils.h
//  GossipGirl
//
//  Created by Woraphot Chokratanasombat on 31/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface StringUtils : NSObject

+(NSString *)getNotNullString:(id)object;
+(NSString *)getNotNullString:(id)object withDefault:(NSString *)value;
#pragma mark Format Number Functions
+(NSString *)formatDecimalNumber:(NSNumber*)number;
+(NSString *)formatIntegerNumber:(NSNumber*)number;
+(NSString *)formatIntegerNumberFromString:(NSString *)number;

@end
