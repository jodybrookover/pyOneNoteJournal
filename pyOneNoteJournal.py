#import win32com.client
import onepy
import datetime
from sys import exit
from pprint import pprint


def find_journal_notebook():
    """Look for a suitable notebook to use.
    We guess it's named 'work' or 'journal'"""
    notebook_name_guesses = ('work', 'journal')  # keep these lowercase since we convert to lowercase to match
    notebooks = onepy.getNotebooks()

    chosen_notebook_id = None
    for notebook in notebooks:
        if notebook['name'].lower() in notebook_name_guesses or notebook['nickname'] in notebook_name_guesses:
            chosen_notebook_id = notebook['ID']
            break
    if chosen_notebook_id is None:
        return None
    return chosen_notebook_id


def get_today():
    """Get date fields necessary"""
    today = datetime.date.today()
    return today.year, today.strftime('%B'), today.day, today.isoformat()


def create_year_section(journal_group_element, year):
    #TODO use UpdateHierarchy method (ew)
    pass


def section_exists(journal_group_element, section_name):
    """Ensure the section exists within the given element"""
    is_exists = False
    for section in journal_group_element:
        if section.attrib.get('name') == str(section_name):
            is_exists = True
    return is_exists


def main():
    #Final structure should look like this
    #Work(or Journal) notebook -- 'Journal' section group -- YYYY section group -- MonthWord Section -- Date Page

    target_notebook_id = find_journal_notebook()
    target_journal_section_id, target_journal_section = onepy.getSectionByName(target_notebook_id, 'Journal', 'sectiongroup')
    # Store some date related values we're gonna need
    year, month_word, day, iso_format_date = get_today()

    assert section_exists(target_journal_section, year), "No section group found for current year: %s" % year
    year_section_id, year_section = onepy.getSectionByName(target_journal_section_id, str(year), 'sectiongroup')
    assert section_exists(year_section, month_word), "No section group found for current month: %s" % month_word
    month_section_id, month_section = onepy.getSectionByName(year_section_id, month_word, 'section')

    month_pages = onepy.getPages(month_section)
    for page in month_pages:
        if page['name'] == iso_format_date:
            print 'Page named "%s" already exists' % iso_format_date
            exit()

    onepy.createNewPage(month_section_id, iso_format_date)

if __name__ == '__main__':
    main()


