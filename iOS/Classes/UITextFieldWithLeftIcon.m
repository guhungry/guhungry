/**
 @author Woraphot Chokratanasombat
 @since 2015-10-30
 @updated 2015-10-30
 @copyright © 2015 Woraphot Chokratanasombat.
 */

#import "UITextFieldWithLeftIcon.h"

@interface UITextFieldWithLeftIcon ()
{
	UIView *container;
	UIImageView *view;
}
@end

@implementation UITextFieldWithLeftIcon

#pragma mark Constructor
- (instancetype)initWithCoder:(NSCoder *)aDecoder
{
	self = [super initWithCoder:aDecoder];
	if (self)
	{
		[self setup];
	}
	return self;
}

- (instancetype)initWithFrame:(CGRect)frame
{
	self = [super initWithFrame:frame];
	if (self)
	{
		[self setup];
	}
	return self;
}

- (void)setup
{
	self.iconRect = CGRectMake(0, 0, 0, 0);
}

#pragma mark Layout
- (void)prepareForInterfaceBuilder
{
	[self drawView];
}

- (void)awakeFromNib
{
	[self drawView];
}

- (void)drawView
{
	if (self.icon != nil)
	{
		// Initialize Left Icon + Container
		CGRect frameContainer = CGRectMake(0, 0, self.iconRect.size.width + self.iconRect.origin.x * 2, self.iconRect.size.height + self.iconRect.origin.y * 2);
		if (container == nil)
		{
			container = [[UIView alloc] initWithFrame:frameContainer];
			view = [[UIImageView alloc] initWithFrame:self.iconRect];
			[container addSubview:view];
		}
		else
		{
			container.frame = frameContainer;
			view.frame = self.iconRect;
		}

		// Set up Left Icon
		view.image = self.icon;
		view.contentMode = UIViewContentModeScaleAspectFit;

		// Add Left Icon to Text Field
		self.leftView = container;
		self.leftViewMode = UITextFieldViewModeAlways;
	}
	else
	{
		self.leftView = nil;
		self.leftViewMode = UITextFieldViewModeNever;
	}
}

@end
