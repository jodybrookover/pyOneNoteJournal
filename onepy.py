import json
import win32com.client
from xml.etree import ElementTree

onapp = win32com.client.gencache.EnsureDispatch('OneNote.Application')
NS = "{http://schemas.microsoft.com/office/onenote/2010/onenote}"


def getHierarchyJson():
    """Returns the Notebook Hierarchy as JSON"""
    return(json.dumps(getHierarchy(), indent=4))


def getHierarchy():
    """Returns the Notebook Hierarchy as a Dictionary Array"""
    oneTree = ElementTree.fromstring(onapp.GetHierarchy("",win32com.client.constants.hsPages))

    notebooks = []

    for notebook in oneTree:
        nbk = parseAttributes(notebook)
        if (notebook.getchildren()):
           s, sg = getSections(notebook)
           if (s != []):
               nbk['sections'] = s

           # Removes RecycleBin from SectionGroups and adds as a first class object
           for i in range(len(sg)):
               if ('isRecycleBin' in sg[i]):
                  nbk['recycleBin'] = sg[i]
                  sg.pop(i)

           if (sg != []):
               nbk['sectionGroups'] = sg

        notebooks.append(nbk)

    return notebooks


def getSections(notebook):
    """Takes in a Notebook or SectionGroup  and returns a Dict Array of its Sections & Section Groups"""
    sections = []
    sectionGroups = []
    for section in notebook:
        if (section.tag == NS + "SectionGroup"):
            newSectionGroup = parseAttributes(section)
            if (section.getchildren()):
               s, sg = getSections(section)
               if (sg != []):
                  newSectionGroup['sectionGroups'] = sg
               if (s != []):
                  newSectionGroup['sections'] = s
            sectionGroups.append(newSectionGroup)

        if (section.tag == NS + "Section"):
            newSection = parseAttributes(section)
            if (section.getchildren()):
               newSection['pages'] = getPages(section)
            sections.append(newSection)

    return sections, sectionGroups


def getPages(section):
    """Takes in a Section and returns a Dict Array of its Pages"""
    pages =[]
    for page in section:
        newPage = parseAttributes(page)
        if (page.getchildren()):
            newPage['meta'] = getMeta(page)
        pages.append(newPage)
    return pages


def getMeta (page):
    """Takes in a Page and returns a Dict Array of its Meta properties"""
    metas = []
    for meta in page:
        metas.append(parseAttributes(meta))
    return metas


def parseAttributes(obj):
    """Takes in an object and returns a dictionary of its values"""
    tempDict = {}
    for key,value in obj.items():
        tempDict[key] = value
    return tempDict


def createNewPage(section_id):
    """Creates a new page within the sectin provided"""
    onapp.CreateNewPage(section_id)
    return 1
