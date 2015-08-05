//
//  ImageUtils.m
//  iOS Utilities
//
//  Created by Woraphot Chokratanasombat on 05/08/2015.
//  Copyright (c) 2015 Woraphot Chokratanasombat. All rights reserved.
//

#import "ImageUtils.h"

@implementation ImageUtils

/**
 * Image Orientation to Flip Horizontal Image Orientation
 *
 * @code
 UIImageOrientation newOrientation = [ImageUtils imageOrientationToFlipHorizontal:originalImage.imageOrientation];
 * @endcode
 * @return Flipped Image Orientation
 * @author Woraphot Chokratanasombat
 * @since 2015-08-05
 * @updated 2015-08-05
 */
+(UIImageOrientation)imageOrientationToFlipHorizontal:(UIImageOrientation)orientation
{
	switch (orientation)
	{
		case UIImageOrientationDown:
			return UIImageOrientationDownMirrored;

		case UIImageOrientationDownMirrored:
			return UIImageOrientationDown;

		case UIImageOrientationLeft:
			return UIImageOrientationRightMirrored;

		case UIImageOrientationLeftMirrored:
			return UIImageOrientationRight;

		case UIImageOrientationRight:
			return UIImageOrientationLeftMirrored;

		case UIImageOrientationRightMirrored:
			return UIImageOrientationLeft;

		case UIImageOrientationUp:
			return UIImageOrientationUpMirrored;

		case UIImageOrientationUpMirrored:
			return UIImageOrientationUp;

		default:
			return orientation;
	}
}

/**
 * Image Orientation to Flip Vertical Image Orientation
 *
 * @code
 UIImageOrientation newOrientation = [ImageUtils imageOrientationToFlipVertical:originalImage.imageOrientation];
 * @endcode
 * @return Flipped Image Orientation
 * @author Woraphot Chokratanasombat
 * @since 2015-08-05
 * @updated 2015-08-05
 */
+(UIImageOrientation)imageOrientationToFlipVertical:(UIImageOrientation)orientation
{
	switch (orientation)
	{
		case UIImageOrientationDown:
			return UIImageOrientationUpMirrored;

		case UIImageOrientationDownMirrored:
			return UIImageOrientationUp;

		case UIImageOrientationLeft:
			return UIImageOrientationLeftMirrored;

		case UIImageOrientationLeftMirrored:
			return UIImageOrientationLeft;

		case UIImageOrientationRight:
			return UIImageOrientationRightMirrored;

		case UIImageOrientationRightMirrored:
			return UIImageOrientationRight;

		case UIImageOrientationUp:
			return UIImageOrientationDownMirrored;

		case UIImageOrientationUpMirrored:
			return UIImageOrientationDown;

		default:
			return orientation;
	}
}

/**
 * Image Orientation to Rotate 90° Clockwise
 *
 * @code
 UIImageOrientation newOrientation = [ImageUtils imageOrientationToRotate90Clockwise:originalImage.imageOrientation];
 * @endcode
 * @return Rotated Image Orientation
 * @author Woraphot Chokratanasombat
 * @since 2015-08-05
 * @updated 2015-08-05
 */
+(UIImageOrientation)imageOrientationToRotate90Clockwise:(UIImageOrientation)orientation
{
	switch (orientation)
	{
		case UIImageOrientationUp:
			return UIImageOrientationRight;

		case UIImageOrientationLeft:
			return UIImageOrientationUp;

		case UIImageOrientationDown:
			return UIImageOrientationLeft;

		case UIImageOrientationRight:
			return UIImageOrientationDown;

		case UIImageOrientationUpMirrored:
			return UIImageOrientationRightMirrored;

		case UIImageOrientationLeftMirrored:
			return UIImageOrientationUpMirrored;

		case UIImageOrientationDownMirrored:
			return UIImageOrientationLeftMirrored;

		case UIImageOrientationRightMirrored:
			return UIImageOrientationDownMirrored;

		default:
			return orientation;
	}
}

/**
 * Image Orientation to Rotate 90° Counter Clockwise
 *
 * @code
 UIImageOrientation newOrientation = [ImageUtils imageOrientationToRotate90CounterClockwise:originalImage.imageOrientation];
 * @endcode
 * @return Rotated Image Orientation
 * @author Woraphot Chokratanasombat
 * @since 2015-08-05
 * @updated 2015-08-05
 */
+(UIImageOrientation)imageOrientationToRotate90CounterClockwise:(UIImageOrientation)orientation
{
	switch (orientation)
	{
		case UIImageOrientationUp:
			return UIImageOrientationLeft;

		case UIImageOrientationLeft:
			return UIImageOrientationDown;

		case UIImageOrientationDown:
			return UIImageOrientationRight;

		case UIImageOrientationRight:
			return UIImageOrientationUp;

		case UIImageOrientationUpMirrored:
			return UIImageOrientationLeftMirrored;

		case UIImageOrientationLeftMirrored:
			return UIImageOrientationDownMirrored;

		case UIImageOrientationDownMirrored:
			return UIImageOrientationRightMirrored;

		case UIImageOrientationRightMirrored:
			return UIImageOrientationUpMirrored;

		default:
			return orientation;
	}
}

/**
 * Image Orientation to Rotate 180°
 *
 * @code
 UIImageOrientation newOrientation = [ImageUtils imageOrientationToRotate180:originalImage.imageOrientation];
 * @endcode
 * @return Rotated Image Orientation
 * @author Woraphot Chokratanasombat
 * @since 2015-08-05
 * @updated 2015-08-05
 */
+(UIImageOrientation)imageOrientationToRotate180:(UIImageOrientation)orientation
{
	switch (orientation)
	{
		case UIImageOrientationUp:
			return UIImageOrientationDown;

		case UIImageOrientationLeft:
			return UIImageOrientationRight;

		case UIImageOrientationDown:
			return UIImageOrientationUp;

		case UIImageOrientationRight:
			return UIImageOrientationLeft;

		case UIImageOrientationUpMirrored:
			return UIImageOrientationDownMirrored;

		case UIImageOrientationLeftMirrored:
			return UIImageOrientationRightMirrored;

		case UIImageOrientationDownMirrored:
			return UIImageOrientationUpMirrored;

		case UIImageOrientationRightMirrored:
			return UIImageOrientationLeftMirrored;

		default:
			return orientation;
	}
}

@end
