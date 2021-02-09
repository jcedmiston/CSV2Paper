from copy import deepcopy
from tkinter.constants import N
from lxml.etree import Element
from lxml import etree
from mailmerge import MailMerge, NAMESPACES

class MailMergeTracking(MailMerge):
    def __init__(self, file, remove_empty_tables=False):
        super().__init__(file, remove_empty_tables=remove_empty_tables)
    
    def merge_templates(self, replacements, separator, queue):
        """
        Duplicate template. Creates a copy of the template, does a merge, and separates them by a new paragraph, a new break or a new section break.
        separator must be :
        - page_break : Page Break. 
        - column_break : Column Break. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - textWrapping_break : Line Break.
        - continuous_section : Continuous section break. Begins the section on the next paragraph.
        - evenPage_section : evenPage section break. section begins on the next even-numbered page, leaving the next odd page blank if necessary.
        - nextColumn_section : nextColumn section break. section begins on the following column on the page. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - nextPage_section : nextPage section break. section begins on the following page.
        - oddPage_section : oddPage section break. section begins on the next odd-numbered page, leaving the next even page blank if necessary.
        """

        #TYPE PARAM CONTROL AND SPLIT
        valid_separators = {'page_break', 'column_break', 'textWrapping_break', 'continuous_section', 'evenPage_section', 'nextColumn_section', 'nextPage_section', 'oddPage_section'}
        if not separator in valid_separators:
            raise ValueError("Invalid separator argument")
        type, sepClass = separator.split("_")
  

        #GET ROOT - WORK WITH DOCUMENT
        for part in self.parts.values():
            root = part.getroot()
            tag = root.tag
            if tag == '{%(w)s}ftr' % NAMESPACES or tag == '{%(w)s}hdr' % NAMESPACES:
                continue
		
            if sepClass == 'section':

                #FINDING FIRST SECTION OF THE DOCUMENT
                firstSection = root.find("w:body/w:p/w:pPr/w:sectPr", namespaces=NAMESPACES)
                if firstSection == None:
                    firstSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)
			
                #MODIFY TYPE ATTRIBUTE OF FIRST SECTION FOR MERGING
                nextPageSec = deepcopy(firstSection)
                for child in nextPageSec:
                #Delete old type if exist
                    if child.tag == '{%(w)s}type' % NAMESPACES:
                        nextPageSec.remove(child)
                #Create new type (def parameter)
                newType = etree.SubElement(nextPageSec, '{%(w)s}type'  % NAMESPACES)
                newType.set('{%(w)s}val'  % NAMESPACES, type)

                #REPLACING FIRST SECTION
                secRoot = firstSection.getparent()
                secRoot.replace(firstSection, nextPageSec)

            #FINDING LAST SECTION OF THE DOCUMENT
            lastSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)

            #SAVING LAST SECTION
            mainSection = deepcopy(lastSection)
            lsecRoot = lastSection.getparent()
            lsecRoot.remove(lastSection)

            #COPY CHILDREN ELEMENTS OF BODY IN A LIST
            childrenList = root.findall('w:body/*', namespaces=NAMESPACES)

            #DELETE ALL CHILDREN OF BODY
            for child in root:
                if child.tag == '{%(w)s}body' % NAMESPACES:
                    child.clear()

            #REFILL BODY AND MERGE DOCS - ADD LAST SECTION ENCAPSULATED OR NOT
            completed_count = 0
            lr = len(replacements)
            lc = len(childrenList)
            parts = []
            for i, repl in enumerate(replacements):
                for (j, n) in enumerate(childrenList):
                    element = deepcopy(n)
                    for child in root:
                        if child.tag == '{%(w)s}body' % NAMESPACES:
                            child.append(element)
                            parts.append(element)
                            if (j + 1) == lc:
                                if (i + 1) == lr:
                                    child.append(mainSection)
                                    parts.append(mainSection)
                                else:
                                    if sepClass == 'section':
                                        intSection = deepcopy(mainSection)
                                        p   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        pPr = etree.SubElement(p, '{%(w)s}pPr'  % NAMESPACES)
                                        pPr.append(intSection)
                                        parts.append(p)
                                    elif sepClass == 'break':
                                        pb   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        r = etree.SubElement(pb, '{%(w)s}r'  % NAMESPACES)
                                        nbreak = Element('{%(w)s}br' % NAMESPACES)
                                        nbreak.attrib['{%(w)s}type' % NAMESPACES] = type
                                        r.append(nbreak)

                    self.merge(parts, **repl)
                
                completed_count += 1
                queue.put((completed_count, "Merging into template...", "determinate"))