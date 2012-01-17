"""
Created on Dec 22, 2011
Wrappers for using Wikipedia API.

Classes:
linkType: Enumerates all types of links to be retrieved from wikipedia
WikiArticleNotFound: An Exception that occurs when the article specified is not valid
WikiConnectionError: An error that happens when connection cannot be established to the wikipedia API
Wiki: The main class for working with wikipedia API
@author: Ali Fathalian
"""
__author__ = 'Ali Fathalian'

import urllib2
from urllib import quote
from xml.dom import minidom


class linkType:
    """Enumerates all types of links that the wiki class can retrieve from wikipedia."""

    Link = "Link"
    BackLink = "BackLink"


class WikiArticleNotFound(Exception):
    """An exception that occurs when the article specified is not available in the wikipedia."""

    pass


class WikiConnectionError(Exception):
    """An exception that occurs when a connection cannot be established with wikipedia API."""

    pass


class Wiki(object):
    """The main class for working with wikipedia API


    Class for retrieving info from wikipedia. This is the main class that works with API.


    The public API for this class are the following functions:
    getArticle: Retrieves an article in wikiText format
    getLinks: Retrieves a list of all the most important links that appear in an article
    getBackLinks: Retrieves a list of all the pages that link to an article
    """

    #API Paths for different API functions
    _articleURL = "http://%s.wikipedia.org/w/index.php?title=%s&action=raw&redirects=true"
    _linkURL = "http://%s.wikipedia.org/w/api.php?action=query&prop=links&" \
                "format=xml&plnamespace=0&pllimit=100&iwurl&titles=%s"
    _backLinksURL = "http://%s.wikipedia.org/w/api.php?action=query&list=backlinks&" \
                     "format=xml&bltitle=%s&blnamespace=0&bllimit=100&redirects"

    #This determines how many items in hundred should be counted for backlinks. For example value 9 means count 1000 items
    #at most.
    _CONTINUE_LIMIT = 9

    def __init__(self, lang = 'en'):
        """
        Pass the language of the wikipages you want to use. The default is english.

        Input parameters:
        lang: The language subset of Wikipedia to run the API on

        """

        self.lang = lang

    def _buildArticleURL(self, articleName):
        """
        Build the article URL of the article you want to retrieve. The raw wikitext is retrieved.


        The article title is normalized into a URL friendly text. This is a call to Wikipedia API.

        InputParameters:
        articleName: The name of the article to retrieve

         Returns:
         A URL to be used to send requests for retrieving wikiText of WikiPedia article specified by articleName
        """

        return self._articleURL % (self.lang, quote(articleName))

    def _buildLinkURL(self, articleName, isContinue=False, plContinue=""):
        """
        Build the URL of the article Links to retrieve. This is a call to Wikipedia API.


        Because the wikipedia API only sends 100 links on the page. If there are more than 100 links on the requested
        pages, the request should be sent in multiple parts. Each part should specify where to continue from the last
        request.


        If isContinue is specified and is true then the built URL contains the parameters to resume a request
        The parameter to resume the request is passed in the optional plContinue parameter.

        Input Parameters:
        articleName: The name of the article to retrieve links for
        isContinue: Determines whether the request should continue a previous query or start from new
        plContinue: If the request is a continuation of a previous request, plContinue determines the place to pick up
        the request from

        Returns:
        A URL to be used to send requests for getting a list of all the links on a wikipedia article specified by
        articleName
        """

        baseLink = self._linkURL % (self.lang, quote(articleName))
        if isContinue:
            baseLink += "&plcontinue=" + quote(plContinue)
        return baseLink

    def _buildBackLinkURL(self, articleName, isContinue=False, blContinue =""):
        """
        Build the URL for the article backlinks. Backlinks are the links that point to a given article.
        This is a call to Wikipedia API.


        Because the wikipedia API only sends 100 links on the page. If there are more than 100 links on the requested
        pages, the request should be sent in multiple parts. Each part should specify where to continue from the last
        request.


        If isContinue is specified and is true then the built URL contains the parameters to resume a request
        The parameter to resume the request is passed in the optional plContinue parameter


        Input Parameters:
        articleName: The name of the article to retrieve backLinks for
        isContinue: Determines whether the request should continue a previous query or start from new
        blContinue: If the request is a continuation of a previous request, blContinue determines the place to pick up
        the request from


        Returns:
        A URL to be used to send requests for getting a list of all the backLinks to a wikipedia article specified by
        articleName
        """

        baseLink = self._backLinksURL % (self.lang,quote(articleName))
        if isContinue:
            baseLink += "&blcontinue=" + quote(blContinue)
        return baseLink

    def _sendRequest(self, requestURL):
        """
        An standard HTTP request to wikipedia api with the requestURL passed in.

        Input Parameters:
        requestURL: The request URL


        Returns:
        The response to the request.


        Throws Exceptions if the article is not found or connection to wikipedia via the api is not possible
        """

        request = urllib2.Request(requestURL)
        request.add_header('User-Agent', 'Mozilla/5.0')
        try:
            result = urllib2.urlopen(request)
        except urllib2.HTTPError, e:
            raise WikiConnectionError(e.code)
        except urllib2.URLError, e:
            raise WikiArticleNotFound(e.reason)

        content = result.read()
        result.close()
        return content

    def _parseLinkContent(self, content, _linkType):
        """
        The API request sent to wikipedia for a list or property will return items in an XML response.
        The xml response has elements of the actual list or property but they are only 100 item in each xml response.
        Depending on the query the elements may be tagged with tags like pl,bl, or etc.


        If there are more items appearing there is a conitnue query item with the related continue parameter.
        This parameter that may be named plcontinue, blcontinue, or etc depending on the type of query can be
        passed to another query to retrieve the rest of the items.


        This function looks into the XML response passed in by content and gets all the item titles and also returns
        hasContinue or continueValue if there are more than 100 items in the response.


        Input parameters:
        content : The XML response for a query for lists or properties
        _linkType : determines what type of item was queried for. Right now the items that this function works with are
        backLinks and forewardLinks


        Returns:
        The function returns a tuple (hasContinue, continueValue, links) :
        hasContinue : determines whether there are more parts to this query or not
        continueValue: If hasContinue is True this determines the place that the query should be resumed from. This value
        should be specified in the next api request in the coressponding xContinue field.
        links: a list of less than 100 links and whether the query should be
        """

        itemName = _linkType == linkType.Link and 'pl' or 'bl'
        continueName = "%scontinue" % itemName
        xmlFile = minidom.parseString(content)
        links = []
        for ref in xmlFile.getElementsByTagName(itemName):
             links.append(ref.attributes['title'].value)
        hasContinue = False
        continueValue = ""

        queryContinue = xmlFile.getElementsByTagName('query-continue')
        if queryContinue:
            hasContinue = True
            continueValue = queryContinue[0].firstChild.attributes[continueName].value
        return hasContinue, continueValue, links

    def _getItemsRecursive(self, articleTitle, linkType, limit, items, hasContinue=False, continueValue=""):
        """
        handles internal recursive calls for retrieving the items of query whether they are a list or a property .


        The main purpose of this is to follow the continue element value to send further requests to retrieve more links
        on a page until all the links are retrieved and the query-continue is not responded.


        Right now this works with backlinks and forwardlinks.


        Input parameters:
        articleTitle: the unique name of the article
        linkType: type of the link to retrieve (backlink or forewardlink)
        hasContinue: whether this query should be continued
        continueValue: if the query should be continued from where it should pick up


        Returns:
        The List of links from the request.
        """

        if not items: items = []
        try:
            buildLink = getattr(self, "_build%sURL" % linkType)
            requestURL = buildLink(articleTitle, hasContinue, continueValue)
        #TODO put the right Exception type
        except Exception:
            #in case of unparsable string
            return []

        content = self._sendRequest(requestURL)
        (hasContinue, continueValue, links) = self._parseLinkContent(content, linkType)
        items.extend(links)
        if hasContinue and limit >= 0:
            self._getItemsRecursive(articleTitle, linkType, limit-1, items, hasContinue, continueValue)
        return items

    def getArticle(self, articleTitle):
        """Retrieves an article in wikiText format


        Retrieves and returns the Raw WikiText of the article specified with the articleTitle.


        The article Title should be a unique article Title as specified in the wikipedia. Redirects are handled
        by wikipedia and the disambiguation page is not shown.


        Input Parameters:
        articleTitle: The title of the article to be retrieved


        Returns:
        The wikiText of the article specified by articleTitle
        """

        requestURL = self._buildArticleURL(articleTitle)
        content = self._sendRequest(requestURL)
        return content

    def getLinks(self, articleTitle):
        """Retrieves a list of all the most important links that appear in an article


        Returns a list of all the available links on a single article page.


        The article Title should be a unique article Title as specified in the wikipedia. Redirects are handled
        by wikipedia and the disambiguation page is not shown.

        Input Parameters:
        articleTitle: The title of the article for which the links are to be retrieved


        Returns:
        A list of links appearing in the article specified by articleTitle.
        """

        return self._getItemsRecursive(articleTitle, linkType.Link, self._CONTINUE_LIMIT, [])

    def getBackLinks(self, articleTitle):
        """Retrieves a list of all the pages that link to an article


        Returns a list of all the available backlinks on a single article page.


        backlinks are links that point to a given page. In other words, this function returns all the articles that have
        links to this article.


        The article Title should be a unique article Title as specified in the wikipedia. Redirects are handled
        by wikipedia and the disambiguation page is not shown.

        Input Parameters:
        articleTitle: The title of the article for which the backLinks are to be retrieved


        Returns:
        A list of backLinks pointing to the article specified by articleTitle.
        """

        items =  self._getItemsRecursive(articleTitle, linkType.BackLink ,self._CONTINUE_LIMIT, [])
        return items


if __name__ == "__main__":

    wiki = Wiki()
    print "Printing a test Article : \n"
    testArticle = "Seattle"
    print wiki.getArticle(testArticle)
    print "Printing all the links in the article : \n"
    links = wiki.getLinks(testArticle)
    print str(len(links)) + " Links Found \n"
    print "\n".join(links)
    print "printing backLinks"
    links = wiki.getBackLinks(testArticle)
    print str(len(links)) + " Links Found \n"
    print "\n".join(links)







