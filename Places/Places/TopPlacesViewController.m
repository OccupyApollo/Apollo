//
//  TopPlacesViewController.m
//  Places
//
//  Created by Ali Fathalian on 2/6/12.
//  Copyright (c) 2012 University of Washington. All rights reserved.
//

#import "TopPlacesViewController.h"
#import "FlickrFetcher.h"
#import "PhotoListViewController.h"

@interface TopPlacesViewController() 
@property (nonatomic, strong) NSArray *topPlaces;
@end

@implementation TopPlacesViewController

@synthesize topPlaces = _topPlaces;

- (void) setTopPlaces:(NSArray *)topPlaces{
    if (topPlaces != _topPlaces){
        _topPlaces = topPlaces;
        [self.tableView reloadData];
    }
}

- (void)viewDidLoad{
    NSArray *topPlaces = [FlickrFetcher topPlaces];
    topPlaces = [topPlaces sortedArrayUsingComparator:^(id obj1, id obj2){
        NSString *obj1Value = [obj1 objectForKey:FLICKR_PLACE_NAME];
        NSString *obj2Value = [obj2 objectForKey:FLICKR_PLACE_NAME];
        return [obj1Value caseInsensitiveCompare:obj2Value];
    }];
    self.topPlaces = topPlaces;

}

#define TOP_PHOTO_MAX 50
-(void) prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender{
    if ([segue.identifier isEqualToString:@"ShowPhotosForPlace"]){
        int selectedRow = [self.tableView indexPathForSelectedRow].row;
        NSDictionary * place = [self.topPlaces objectAtIndex:selectedRow];
        if(place){
            NSArray * photos = [FlickrFetcher photosInPlace:place maxResults:TOP_PHOTO_MAX];
            [segue.destinationViewController setPhotos:photos];
        }

        
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
 
    return [self.topPlaces count];
}

- (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath
{
    static NSString *CellIdentifier = @"TopPlaceCell";
    
    UITableViewCell *cell = [tableView dequeueReusableCellWithIdentifier:CellIdentifier];
    if (cell == nil) {
        cell = [[UITableViewCell alloc] initWithStyle:UITableViewCellStyleSubtitle reuseIdentifier:CellIdentifier];
    }
    
    NSDictionary *place = [self.topPlaces objectAtIndex:indexPath.row];
    NSString *placeDescription = [place objectForKey:FLICKR_PLACE_NAME];
    int topIndex = [placeDescription rangeOfCharacterFromSet: [NSCharacterSet characterSetWithCharactersInString:@","]].location;
    cell.textLabel.text = [placeDescription substringToIndex:topIndex];
    cell.detailTextLabel.text = [placeDescription substringFromIndex:topIndex + 2];
    
    return cell;
}


#pragma mark - Table view delegate

- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath
{
    // Navigation logic may go here. Create and push another view controller.
    /*
     <#DetailViewController#> *detailViewController = [[<#DetailViewController#> alloc] initWithNibName:@"<#Nib name#>" bundle:nil];
     // ...
     // Pass the selected object to the new view controller.
     [self.navigationController pushViewController:detailViewController animated:YES];
     */
}

@end
