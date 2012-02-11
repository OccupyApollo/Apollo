//
//  PhotoListViewController.h
//  Places
//
//  Created by Ali Fathalian on 2/6/12.
//  Copyright (c) 2012 University of Washington. All rights reserved.
//

#import <UIKit/UIKit.h>

#define RECENT_PHOTOS_KEY @"PhotoListViewController.RecentPhotos"

@interface PhotoListViewController : UITableViewController
@property (nonatomic, strong) NSArray *photos;
@end
