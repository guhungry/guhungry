#import <UIKit/UIKit.h>

IB_DESIGNABLE

/**
 @brief UITextField that allow set Left Icon in the UI Interface Builder and see result real time.

 @author Woraphot Chokratanasombat
 @since 2015-10-30
 @updated 2015-10-30
 @copyright © 2015 Woraphot Chokratanasombat.
 */
@interface UITextFieldWithLeftIcon : UITextField

@property (nonatomic) IBInspectable UIImage *icon;
@property (nonatomic) IBInspectable CGRect iconRect;

@end
