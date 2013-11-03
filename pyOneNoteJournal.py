#import win32com.client
import onepy
import datetime


def find_journal_notebook():
    notebook_name_guesses = ('work', 'Journal')  # Hopefully these are case insensitive
    notebooks = onepy.getHierarchy()

    chosen_notebook = None
    for notebook in notebooks:
        if notebook['name'] in notebook_name_guesses or notebook['nickname'] in notebook_name_guesses:
            chosen_notebook = notebook['ID']
            continue
    if chosen_notebook is None:
        return None
    return chosen_notebook


def get_today():
    """Get date fields necessary"""
    today = datetime.date.today()
    return today.year, today.strftime('%B'), today.day, today.isoformat()

def main():
#    onapp = win32com.client.Dispatch('OneNote.Application')
#    win32com.client.gencache.EnsureDispatch('OneNote.Application')

#    notebooks = onapp.GetHierarchy("", win32com.client.constants.hsNotebooks)
#    notebooks_soup = BeautifulSoup(notebooks, 'lxml')
#    print(notebooks_soup.prettify())
    notebooks = onepy.getHierarchy()
    print notebooks

    chosen_notebook_id = find_journal_notebook()
    onepy.getSections(chosen_notebook_id)
    y,m,d,iso = get_today()
    onepy.createNewPage()
    pass


if __name__ == '__main__':
    main()


