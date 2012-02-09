"""
Created on Dec 22, 2011
Controller for working with wikipedia API.


The classes in this module include:
SelectionAlgorithm: Enumerates all the algorithms that can be used for selection and ranking of links in an article
AssociationCriteria: The association criteria for finding association between two articles
WikiController: Encapsulates the wikipedia API classes and controls their functionality.

@author: Ali Fathalian
"""
__author__ = 'Ali Fathalian'

import collections

from wikiadapter import Wiki
from wikitextreader import WikiTextReader


class SelectionAlgorithm:
    """Enumerates all the algorithms that can be used for selection and ranking of links in an article


    Currently two algorithms are implemented that can be used to rank links:
    The first one is pageRank and the other
    The other one is the FrequencyRank which is a simple Bag of words ranking based on three level of word
    occurrences in the text.


    Any new algorithm should be added here.
    """

    PageRank = "PageRank"
    FrequencyRank = "FrequencyRank"

class AssociationCriteria:
    """The association criteria for finding association between two articles

    This determines how difficult and hard is the criteria for to thing to be considered associated.


    Right now there are two enumerations of how hard is the criteria for the associations:

    PageFirst: First a page rank algorithm is executed on all the links for the article to find important ones and
    then from those the shared articles are selected
    SharedFirst: First the shared links are found and then if there are more than a certain number of shared items then
    the PageRank Algorithm is executed to select the top shared items.

    Generally, the SharedFirst variation gives both a better performance and finds more shared items between articles
    that are not close (for example Steve Jobs and Jesus). The PageFirst Algorithm works best when the articles are
    close in meaning and there are certainly shared items between them but we want to find the most relevant ones.
    """

    PageFirst = "PageFirst"
    SharedLinksFirst = "SharedLinksFirst"
    ReadArticleFirst = "ReadArticleFirst"


class WikiController:
    """Encapsulates the wikipedia API classes and controls their functionality.


    The main controller to work with API.


    This encapsulates all the complexity of a wikiReader,WikiAdapter, and the ranking algorithms.


    The public API for this class provides the following functions:
    getImportantLinks: Retrieves the most important links in an article based on a specified algorithm
    findAssociations: Finds association between a list of articles
    """

    def getImportantLinks(self, articleTitle, selectionAlgorithm=SelectionAlgorithm.PageRank, outputLimit=15):
        """Retrieves the most important links in an article based on a specified algorithm


        This is the function that retrieves and ranks items from wikipedia. This function always combines the results
        with a bag of words algorithm.


        The bag of words algorithm is run automatically when a wikiReader reads links. It goes through two steps of
        first identifying all links than selecting the most frequent of those links in the wikiText.


        Right now page ranks takes some time to finish but this should not be a problem. A Hadoop server with MapReduce
        and a sophisticated caching mechanisms along with an index database will significantly improve the speed to a
        matter of miliseconds.


        Input Parameters:
        articleTitle : The title of the article to retrieve and rank the links for
        selectionAlgorithm : The algorithm to use for ranking alongside bag of words
        outputLimit: This specifies how many links should be ranked and returned


        Returns:
        A list containing top links titles. (the number or links equals to the outputLimit input parameter passed in)
        """

        #Get article content
        self.wiki = Wiki()
        articleContent = self.wiki.getArticle(articleTitle)

        #Read all the links from the wikiText
        wikiReader = WikiTextReader()
        links = wikiReader.readLinks(articleTitle,articleContent)

        #Select the ranking algorithm and run it in the all links that are retrieved
        selectionAlg = getattr(self, "_selectLinks_%s" % selectionAlgorithm)
        return selectionAlg(links, outputLimit)

    def findAssociations(self, articles, criteria=AssociationCriteria.SharedLinksFirst):
        """ Finds association between a list of articles


        This finds all the associations between a list of articles represented by articles input parameter.


        The associations are found by finding all the important links for all of the links and finding the ones
        that appear in two or more articles.

        Input Parameters:
        articles: A list containing all the article names that you want to find associations between
        criteria: The criteria for finding the association. The default value is shareFirst meaning that
        at first the shared links are found and then a pageRank algorithm is run on them to find the most relevant ones.
        See the "Association Criteria" class for more detail

        Returns:
        A List of all the shared articles and associations between all the articles in the inputparameter
        """

        findAssociationAlg = getattr(self,"_findAssociation_%s" % criteria)
        return findAssociationAlg(articles)

    def _findSharedLinks(self,allLinksMultiSet,articles,rankLimit):

        #Find the intersection of all the multi sets for all the articles
        mainArticleTitle = articles.pop()
        mainSet = allLinksMultiSet[mainArticleTitle]
        for articleTitle in articles:
            mainSet = mainSet & allLinksMultiSet[articleTitle]
        sharedLinks = list(mainSet.elements())

        #If more than rankLimit items are found. Use PageRank to find the top rankLimit Items
        if len(sharedLinks) > rankLimit:
            interimLinks = [(link,self._calculatePageRankLinkWeight(link)) for link in sharedLinks]
            sortedLinks = sorted(interimLinks, key= lambda item:item[1])
            length = len(sortedLinks)
            sharedLinks = [link for (link,freq) in sortedLinks[ length-rankLimit - 1: length]]

        return sharedLinks
    def _findAssociation_ReadArticleFirst(self,articles,rankLimit =7):

        self.wiki = Wiki()
        allLinksMultiSet = {}

        wikiReader = WikiTextReader()


        for articleTitle in articles:
            content = self.wiki.getArticle(articleTitle)
            links = wikiReader.readLinks(articleTitle,content,0,0,100000)
            onlyLinks = [link for (link,freq) in links]
            allLinksMultiSet[articleTitle] = collections.Counter(onlyLinks)

        return self._findSharedLinks(allLinksMultiSet,articles,rankLimit)




    def _findAssociation_SharedLinksFirst(self, articles, rankLimit=7):
        """ The algorithm for finding soft associations between a list of articles

        Input Parameters:
        articles: A list containing all the article names that you want to find associations between
        rankLimit: Determines how many top shared article associations should be returned for the list of articles
        given. The function may return at most rankLimit items.

        returns:
        An Adjacency list representation of the graph that associates all the articles with intermediate articles and
        the original articles as vertices.
        """

        self.wiki = Wiki()
        allLinksMultiSet = {}

        #Create a multi set for each article links.
        for articleTitle in articles:
            allLinksMultiSet[articleTitle] = collections.Counter(self.wiki.getLinks(articleTitle))

        return self._findSharedLinks(allLinksMultiSet,articles,rankLimit)

    def _selectLinks_PageRank(self, links, rankLimit):
        """
        This is the page rank algorithm for selecting at most rankLimit items from a list of (link,freq) tuples.
        The freq in the tuples corresponds to the bag of words frequencies.


        Because the number of frequencies returned by the bag of words is usually in the order of ten and the frequencies
        returned by the pageRank are in the order of thousands. Both of these should be normalized so that they can be
        combined


        Input parameters:
        links: a list of tuples in the form of (Link,freq) in which Link is the title of Link and freq is the result of
        the bag of words algorithm
        rankLimit: The number of items that should be returned. In other words, the top rankLimit number of items.


        Returns:
        A list of link titles containing the top rankLimit items.
        """

        #Normalize the bag of words frequencies
        links = self._normalize(links)

        #Calculate pageRankWeights by counting the backLinks
        interimLinks = [(link, self._calculatePageRankLinkWeight(link)) for (link, freq) in links]

        #Normalize the pageRank frequencies
        interimLinks = self._normalize(interimLinks)

        #Combine normalized bag of words and pageRank frequencies
        interimLinks = self._combineFreqParameters(links, interimLinks)

        #Sort and return the top rankLimit items
        sortedLinks = sorted(interimLinks, key= lambda item:item[1])
        #TODO remove this print statement for production
        print "\n".join(["%s=%s" % (link, freq) for (link, freq) in sortedLinks])
        length = len(sortedLinks)
        return [link for (link, freq) in sortedLinks[ length-rankLimit - 1: length]]

    def _calculatePageRankLinkWeight(self, link):
        """
        This calculates the pageRank weights by counting the backLinks that go into a single Link. In other words
        this function returns the number of pages that have links to this page.


        Right now we only count a limited number of backlinks to prune out irrelevant big items (like God or Apple).


        The number of backLinks to read is determined in the
        WikiAdapter class under __CONTINUE_LIMIT.


        Input parameters:
        link: a single link to count the backlinks for


        Returns:
        A number corresponding to the number of pages that link to this page.
        """

        return len(self.wiki.getBackLinks(link))

    def _normalize(self, links):
        """
        A simple normalization function that uses the mean of all the frequencies in the links to normalize frequencies.


        if there are not a lot of links the normalization is not done and the links are returned as inputted.


        Input parameters:
        links: a list of (link,freq) tuples.


        Returns:
        A list of (link,freq) tuples in which freq is the normalized frequency.
        """

        if links:
            mean = float(sum([freq for (link, freq) in links]) / len(links))
            links = [(link, float(freq / mean)) for (link, freq) in links]
        return links

    def _combineFreqParameters(self, list1, list2):
        """
        This function combines the normalized frequencies of two lists passed into it.


        The combining is done by just summing over both frequencies.

        List1 and List2 should be the results of different algorithms


        input Parameters:
         List1: a list containing tuples of (link,freq).
         List2: a list containing tuples of (link,freq).


        Returns:
        A list containing tuples of (link,freq) in which freq is the combined frequency.
        """

        #zipped = [((linkA,freq1),(linkA,freq2)),((linkB,freq3),(linkB,freq4))...]
        zipped = zip(list1, list2)

        return [(combinedTuple[0][0], combinedTuple[0][1] + combinedTuple[1][1]) for combinedTuple in zipped]


if __name__ == "__main__":

    articleName = "Steve Jobs"
    controller = WikiController()

    print "\n".join(controller.getImportantLinks(articleName))
    #articles = {"Steve Jobs","God"}
   #articles = {"Python (programming language)","Java (programming language)"}
#    print controller.findAssociations(articles,AssociationCriteria.ReadArticleFirst)




