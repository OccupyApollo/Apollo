//
//  FlickrImageViewController.m
//  Places
//
//  Created by Ali Fathalian on 2/6/12.
//  Copyright (c) 2012 University of Washington. All rights reserved.
//

#import "FlickrImageViewController.h"
#import "SplitViewBarButtonItemPresenter.h"
@interface FlickrImageViewController() <UIScrollViewDelegate,SplitViewBarButtonItemPresenter>
@property (weak, nonatomic) IBOutlet UIScrollView *scrollView;
@property (weak, nonatomic) IBOutlet UIImageView *imageView;
@property (weak, nonatomic) IBOutlet UIToolbar *toolbar;
@end

@implementation FlickrImageViewController
@synthesize scrollView = _scrollView;
@synthesize imageView = _imageView;
@synthesize toolbar = _toolbar;
@synthesize image = _image;
@synthesize splitViewBarButtonItem = _splitViewBarButtonItem;


-(void)setUpScrollView{
    
    [self.imageView removeFromSuperview];
    self.imageView = [self.imageView initWithImage:self.image];
    float imageWidth = self.imageView.image.size.width;
    float imageHeight = self.imageView.image.size.height;
    if (UIDeviceOrientationIsLandscape([[UIDevice currentDevice] orientation])){
        float temp = imageHeight;
        imageHeight = imageWidth;
        imageWidth = temp;
    }
    self.imageView.frame = CGRectMake(0, 0, imageWidth , imageHeight);
    [self.scrollView addSubview:self.imageView];
    self.scrollView.contentSize = self.imageView.image.size;
    self.scrollView.minimumZoomScale = 0.1;
    self.scrollView.maximumZoomScale = 4.0;
    
}
-(void)viewDidLoad
{
    [super viewDidLoad];
    self.scrollView.delegate = self;
    [self setUpScrollView];

}

- (void)handleSplitViewBarButtonItem:(UIBarButtonItem *)splitViewBarButtonItem
{
    NSMutableArray *toolbarItems = [self.toolbar.items mutableCopy];
    if (_splitViewBarButtonItem) [toolbarItems removeObject:_splitViewBarButtonItem];
    if (splitViewBarButtonItem) [toolbarItems insertObject:splitViewBarButtonItem atIndex:0];
    self.toolbar.items = toolbarItems;
    _splitViewBarButtonItem = splitViewBarButtonItem;
}
- (void) setSplitViewBarButtonItem:(UIBarButtonItem *)splitViewBarButtonItem{
    if (splitViewBarButtonItem != _splitViewBarButtonItem) {
        [self handleSplitViewBarButtonItem:splitViewBarButtonItem];
    }
}

-(void)fitImage{
    CGFloat width = self.imageView.image.size.width;
    CGFloat height = self.imageView.image.size.height;
    CGRect bounds = self.scrollView.bounds;
    float widthScale = bounds.size.width/width;
    float heightScale = bounds.size.height/height;
    float zoomScale = widthScale > heightScale ? widthScale : heightScale;
    if (height < width ){
        zoomScale = zoomScale /(height/width);
    }
    else {
        zoomScale = zoomScale / (width/height);
    }
    
    [self.scrollView setZoomScale:zoomScale];

    
}
-(void)viewWillAppear:(BOOL)animated{
    [super viewWillAppear:animated];
    [self fitImage];
}

- (void)setImage:(UIImage *)image{
    if (image != _image){
        _image = image;
        [self.scrollView setZoomScale:1.0]; 
        [self setUpScrollView];
        [self fitImage];
        [self.scrollView setNeedsDisplay];
        
    }
}


- (UIView *)viewForZoomingInScrollView:(UIScrollView *)scrollView
{
    return self.imageView;
}

- (BOOL)shouldAutorotateToInterfaceOrientation:(UIInterfaceOrientation)interfaceOrientation
{
    return YES;
}
- (void) willRotateToInterfaceOrientation:(UIInterfaceOrientation)toInterfaceOrientation duration:(NSTimeInterval)duration {
    [super willRotateToInterfaceOrientation:toInterfaceOrientation duration:duration];
    float imageWidth = self.imageView.image.size.width;
    float imageHeight = self.imageView.image.size.height;
    if (UIDeviceOrientationIsLandscape([[UIDevice currentDevice] orientation])){
        float temp = imageHeight;
        imageHeight = imageWidth;
        imageWidth = temp;
    }
    self.imageView.frame = CGRectMake(0, 0, imageWidth , imageHeight);
    [self.scrollView addSubview:self.imageView];
    self.scrollView.contentSize = self.imageView.image.size;
    self.scrollView.zoomScale = 1.0;
    self.scrollView.minimumZoomScale = 0.1;
    self.scrollView.maximumZoomScale = 4.0;

}

- (void) didRotateFromInterfaceOrientation:(UIInterfaceOrientation)fromInterfaceOrientation{
    [super didRotateFromInterfaceOrientation:fromInterfaceOrientation];
    [self fitImage];
}



- (void)viewDidUnload
{
    [self setImageView:nil];
    [self setScrollView:nil];
    [self setToolbar:nil];
    [super viewDidUnload];
}

@end
