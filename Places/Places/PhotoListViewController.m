//
//  PhotoListViewController.m
//  Places
//
//  Created by Ali Fathalian on 2/6/12.
//  Copyright (c) 2012 University of Washington. All rights reserved.
//

#import "PhotoListViewController.h"
#import "FlickrFetcher.h"
#import "FlickrImageViewController.h"


@implementation PhotoListViewController
@synthesize photos = _photos;


-(void)clearPhoto:(NSDictionary *)photo FromArray:(NSMutableArray *)array {
    NSString *photoId = [photo objectForKey:FLICKR_PHOTO_ID];
    int delIndex = -1;
    for (NSDictionary * arrayPhoto in array) {
        if ([[arrayPhoto objectForKey:FLICKR_PHOTO_ID] isEqual:photoId]){
            delIndex = [array indexOfObject:arrayPhoto];
            break;
        }
    }
    if (delIndex >= 0){
        [array removeObjectAtIndex:delIndex];
    }
    
    if ([array count] > 20) [array removeLastObject];
    
}

#define RECENT_PHOTOS_KEY @"PhotoListViewController.RecentPhotos"
-(void)registerRecentPhoto:(NSDictionary *)photo{
    NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
    NSMutableArray *recents = [[defaults objectForKey:RECENT_PHOTOS_KEY] mutableCopy];
    if(!recents) recents = [NSMutableArray array];
    [self clearPhoto:photo FromArray:recents];
    [recents insertObject:photo atIndex:0];
    [defaults setObject:recents forKey:RECENT_PHOTOS_KEY];
    [defaults synchronize];
    
    
}




-(void) prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender{
    if ([segue.identifier isEqualToString:@"ShowImage"]){
        int selectedRow = [self.tableView indexPathForSelectedRow].row;
        
        NSDictionary * photo = [self.photos objectAtIndex:selectedRow];
        if(photo){
            NSURL *imageURL =[FlickrFetcher urlForPhoto:photo format:FlickrPhotoFormatLarge];
            NSData * data = [NSData dataWithContentsOfURL:imageURL];
            UIImage *image = [UIImage imageWithData:data];
            [segue.destinationViewController setImage:image];
            NSString *title = [sender textLabel].text;
            [segue.destinationViewController setTitle:title];
            [self registerRecentPhoto:photo];
            
        }
    }
}


- (void) setPhotos:(NSArray *)photos{
    if (photos != _photos){
        _photos = photos;
        [self.tableView reloadData];
        
    }
}

- (BOOL)shouldAutorotateToInterfaceOrientation:(UIInterfaceOrientation)interfaceOrientation
{
    return YES;
}

#pragma mark - Table view data source

- (NSInteger)numberOfSectionsInTableView:(UITableView *)tableView
{
    return 1;
}

- (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section
{
    return [self.photos count] ;
}

- (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath
{
    static NSString *CellIdentifier = @"PhotoCell";
    
    UITableViewCell *cell = [tableView dequeueReusableCellWithIdentifier:CellIdentifier];
    if (cell == nil) {
        cell = [[UITableViewCell alloc] initWithStyle:UITableViewCellStyleSubtitle reuseIdentifier:CellIdentifier];
    }
    
    NSDictionary *photo = [self.photos objectAtIndex:indexPath.row];
    NSString *title = [photo objectForKey:FLICKR_PHOTO_TITLE];
    NSString *description = [photo valueForKeyPath:FLICKR_PHOTO_DESCRIPTION];
    if (!description || [description isEqualToString:@""]) 
        description = @"Unknown";
    if (!title || [title isEqualToString:@""])
        title = description;
    cell.textLabel.text = title;
    cell.detailTextLabel.text = description;
    return cell;
}


#pragma mark - Table view delegate

- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath
{
    if (self.splitViewController){
        NSDictionary * photo = [self.photos objectAtIndex:indexPath.row];
        if(photo){
            NSURL *imageURL =[FlickrFetcher urlForPhoto:photo format:FlickrPhotoFormatLarge];
            NSData * data = [NSData dataWithContentsOfURL:imageURL];
            UIImage *image = [UIImage imageWithData:data];
            id detailViewController = [[self.splitViewController viewControllers] lastObject];
            [detailViewController setImage:image];
            [self registerRecentPhoto:photo];
            
        }
    }
    
    
}

@end
