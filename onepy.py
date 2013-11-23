import json
import win32com.client
from xml.etree import ElementTree
#from xml.etree.ElementTree import tostring

onapp = win32com.client.gencache.EnsureDispatch('OneNote.Application')
#TODO Make this variable based on your version of one note (/2010/, etc)
NS = "{http://schemas.microsoft.com/office/onenote/2007/onenote}"


#TODO Remove if not needed in the end
class Section():
    def __init__(self, id, type):
        self.id = id
        if type in ('section', 'sectiongroup'):
            self.type = type
        else:
            raise ValueError("Type must be a value of 'section' or 'sectiongroup'")


def getHierarchyJson():
    """Returns the Notebook Hierarchy as JSON"""
    return(json.dumps(getNotebooks(), indent=4))


def getNotebooks():
    """Returns the Notebook Hierarchy as a Dictionary Array"""
    oneTree = ElementTree.fromstring(onapp.GetHierarchy("", win32com.client.constants.hsPages))

    notebooks = []

    for notebook in oneTree:
        nbk = parseAttributes(notebook)
        if notebook.getchildren():
           s, sg = _getSections(notebook)
           if s != []:
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


def getSectionsOfNotebook(notebook_id):
    """Get a list of all the sections and section groups in the notebook, including section group children"""
    oneTree = ElementTree.fromstring(onapp.GetHierarchy(notebook_id, win32com.client.constants.hsPages))
    section_names = []

    for section in oneTree:
        nbk = notebook.attrib
        section_name = nbk['name']
        section_names.append(section_name)  # Includes section groups
    return section_names


def getSectionByName(notebook_id, section_name, section_type):
    """Get a section by its name and type ("sectiongroup" or "section"
    Returns a tuple (id, element) where element is an ElementTree Element"""
    oneTree = ElementTree.fromstring(onapp.GetHierarchy(notebook_id, win32com.client.constants.hsPages))
    lowercase_section_name = section_name.lower()  # For comparison inside a loop
    target_section = None
    for section in oneTree:
        if section.attrib.get('name').lower() == lowercase_section_name:  # TODO Add check for type (group or not)
            target_section = section
            break   # Only get first find, if you have dupes, that's your problem
    if target_section is None:
        raise LookupError('Could not find Section with name: %s', section_name)
    return target_section.attrib.get('ID'), target_section


def createNewPage(section_id, title):
    onapp.CreateNewPage(section_id)



def _getSections(notebook):
    """Takes in a Notebook or SectionGroup  and returns a Dict Array of its Sections & Section Groups"""
    sections = []
    sectionGroups = []
    for section in notebook:
        if section.tag == NS + "SectionGroup":
            newSectionGroup = parseAttributes(section)
            if section.getchildren():
               s, sg = _getSections(section)
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
    for key, value in obj.items():
        tempDict[key] = value
    return tempDict


def createNewPage(section_id, page_title):
    """Creates a new page within the section provided"""
    new_page_id = onapp.CreateNewPage(section_id, win32com.client.constants.npsBlankPageWithTitle)
    oneTree = ElementTree.fromstring(onapp.GetPageContent(new_page_id, win32com.client.constants.piAll))
    xml = onapp.GetPageContent(new_page_id, win32com.client.constants.piAll)
    # HACK ALERT!!
    xml = xml.replace('<![CDATA[]]>', '<![CDATA[%s]]>' % page_title) # HACK!! I hate namespaces
    xml = xml.replace('ns0:', 'one:')
    onapp.UpdatePageContent(xml)
    return 1
