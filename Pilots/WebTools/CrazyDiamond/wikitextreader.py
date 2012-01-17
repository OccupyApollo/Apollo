"""
Created on Dec 22, 2011
This class should be used for parsing a wikipedia text that has been retrieved in wikiText format.


The classes in this module are:
WikiTextReader: Reads wikiText an can return various textual attributes of the wikiText.


@author: Ali Fathalian
"""
__author__ = 'Ali Fathalian'

import re


class WikiTextReader(object):
    """Reads wikiText an can return various textual attributes of the wikiText.


    This is a content reader class. It reads text in wiki format and parses it to retrieve variety of metrics.
    Right now the only implemented method is simple link retrieval. Simple link retrieval retrieves links based on
    their frequency and their aliases frequency.


    The main textReader algorithm is a bag of words algorithm. In the first Scan it collects all the bags of words which
    are the items that appear more _FIRST_SCAN_LIMIT. In the second step, the algorithm goes over the bag of words and
    finds aliases and then looks for the words and their aliases and counts the words that happen more than
    __SECOND_SCAN_LIMIT in the body of text. If the links are less than __MIN_LINKS_LIMIT the __SECOND_SCAN_LIMIT will
    be relaxed until we at least have __MIN_LINKS_LIMIT or no other links. If the links are more than __MAX_LINKS_LIMIT
    the __SECON_SCAN_LIMIT will become tighter until only __MAX_LINKS_LIMIT or less items are gathered.


    The public API for this class has the functions:
    readLinks: Reads and selects the most important links in the article based on the bag of words algorithm
    """

    # The pattern to find links based on
    _linkPattern = re.compile('\[\[(.+?)\]\]')
    _FIRST_SCAN_LIMIT = 1
    _SECOND_SCAN_LIMIT = 2
    _MIN_LINKS_LIMIT = 7
    _MAX_LINKS_LIMIT = 30

    def _searchContent(self, linkTitle, freq, aliases, content):
        """
        Searches a body of wikiText to retrieve the total number of times that a phrase or its aliases appear in the body
        of the text. A single link (linkTitle) may appear under different aliases or names in the text that may or may
        not be similar to the original link. This function finds all the occurrences and adds them up to return the final
        and total frequency of a link or its aliases happening in the text.


        Input Parameters:
            linkTitle: The title of the link or a phrase that we are searching the document for
            freq: the initial frequency that the link appears in the wikiText. This is the number of times that
             A link has appeared between [[ ]]
            aliases: A map of all the aliases that are known. This map is in the form of { "name" = "Alias" } which shows
             an Alias for a given name. The aliases map should be constructed when the text is searched for links
             content:The actual wikiText content to search in. This is the result of the retrieval of RAW data from
             Wikipedia.


        Returns:
            The total number a word (linkTitle) or its link alias has appeared in the content.
        """

        freq += content.count(linkTitle)
        if aliases.has_key(linkTitle) :
            alias = aliases[linkTitle]
            #Single word aliases like 'n' should not be considered in this. There are literally hundreds of them in the
            #and the result will not be accurate
            if len(alias) > 1:
                freq += content.count(alias)
        return freq

    def _selectImportantLinks_Freq(self, interimLinks, aliases, content):
        """
        This class selects the most important links based on the frequency of them or their aliases appearing in the
        text. The output of this function should be bounded and be no more than __MAX_LINKS_LIMIT. This is accomplished by
        tightening the bound of acceptable frequencies for links.


        Input Parameters:
            interimLinks A lists of tuples in the form of (link,freq) that shows the frequency of a single link in
            the text. These are the initial links that have been gathered but should be filtered further.
            aliases: A map of all the aliases that are known. This map is in the form of { "name" = "Alias" } which shows
             an Alias for a given name. The aliases map should be constructed when the text is searched for links
            content: The actual wikiText content to search in. This is the result of the retrieval of RAW data from
             Wikipedia.


        Returns:
             A list of filtered tuples in the form of (link,freq) that are selected from interimLinks based
            on their frequencies.
        """

        #Search the content for the link or its aliases and get the total number of frequencies for each link.
        finalLinks = []
        for (topLink, freq) in interimLinks:
            newFreq = self._searchContent(topLink, freq, aliases, content)
            finalLinks.append((topLink, newFreq))

        #Here we have a large number of links and frequencies. We go through them to just select no more than
        # __MAX_LINKS_LIMIT. Start from __SECOND_SCAN_LIMIT and in each step make the threshold of acceptable frequencies
        #tighter until the number of links is smaller than __MAX_LINKS_LIMIT .
        step = 0
        while  len(finalLinks) >= self._MAX_LINKS_LIMIT:
            finalLinks = [(link, freq) for (link, freq) in finalLinks if freq > self._SECOND_SCAN_LIMIT + step]
            step += 1
        return finalLinks

    def readLinks(self, content):
        """Reads and selects the most important links in the article based on the bag of words algorithm


        This function retrieves the most important links in the body of wikiText. Right now these important links are
        selected based on the frequency of their or their aliases happening in the text. The function starts from a big
        list of links and tightens the acceptable words frequencies until an acceptable number of links are returned.

        Input parameters:
            content: A wikiText describing the article to look for important links in


        Returns:
            importantLinks: A list of the most important links that are no more than __MAX_LINKS_LIMIT and are
            selected based on their frequencies.
        """

        #Use regular expression to find all the text that appear inside [[ ]]. These are our links
        dirtyLinks = self._linkPattern.findall(content)

        #Go through the results one by one. If the link contains ':' it means it is related to a superpage and should
        #not be considered. Links that have '|' in them have aliases. For these links separate the link name and it alias
        #and store the alias with the name key in a map. Finally, if the phrase that you are looking at has been encountered
        #before in the body of the text increase its initial frequency counter; Otherwise, add a new entry to the cleanLinks
        #map. This map contains {nameOfTheLink = frequencyOfTheLink} and will be updated as new links are encountered.
        cleanLinks = {}
        aliases = {}
        for link in dirtyLinks:
            if link.find(':') == -1:
                stringItems = link.split('|')
                key = stringItems[0]
                if cleanLinks.has_key(key):
                    cleanLinks[key] += 1
                else:
                    cleanLinks[key] = 1
                    if len(stringItems) > 1:
                        aliases[key] = stringItems[1]
        #Links that appear only once in the text should be filtered out.
        interimLinks = [ (topLink, freq) for (topLink, freq) in cleanLinks.items() if freq > self._FIRST_SCAN_LIMIT]

        #For the remaining links adjustments must be made so that we get the most important links and the result should
        #be returned.
        return self._selectImportantLinks_Freq(interimLinks, aliases, content)


if __name__ == "__main__":
    from wikiadapter import Wiki
    wiki = Wiki()
    content = wiki.getArticle("Shine On You Crazy Diamond")
    reader = WikiTextReader()
    print len(reader.readLinks(content))
    print ["%s=%s" % (link, freq) for (link, freq) in reader.readLinks(content)]
