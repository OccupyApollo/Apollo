//
//  RecentPhotoListViewController.m
//  Places
//
//  Created by Ali Fathalian on 2/8/12.
//  Copyright (c) 2012 University of Washington. All rights reserved.
//

#import "RecentPhotoListViewController.h"

@implementation RecentPhotoListViewController

-(void) viewWillAppear:(BOOL)animated{
    NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
    self.photos = [[defaults objectForKey:RECENT_PHOTOS_KEY] mutableCopy];
}

@end
