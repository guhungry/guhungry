//
//  ImageUtils.h
//  iOS Utilities
//
//  Created by Woraphot Chokratanasombat on 05/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface ImageUtils : NSObject

+(UIImageOrientation)imageOrientationToFlipHorizontal:(UIImageOrientation)orientation;
+(UIImageOrientation)imageOrientationToFlipVertical:(UIImageOrientation)orientation;
+(UIImageOrientation)imageOrientationToRotate90Clockwise:(UIImageOrientation)orientation;
+(UIImageOrientation)imageOrientationToRotate90CounterClockwise:(UIImageOrientation)orientation;
+(UIImageOrientation)imageOrientationToRotate180:(UIImageOrientation)orientation;

@end
