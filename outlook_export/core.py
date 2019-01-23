"""Core functions for exporting data from outlook.
"""


import logging
import win32com.client


def show_paths(app=None, the_folders=None):
    all_folders = get_all(app, the_folders)
    return [item.FullFolderPath for item in all_folders]


def get_all(app=None, the_folders=None):
    result = []
    if app is None:
        app = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
    if the_folders is None:
        the_folders = app.Folders
    for thing in the_folders:
        sub_count = len(thing.Folders)
        if sub_count:
            logging.info('Recurse into %s', thing.name)
            result.extend(get_all(app, thing.Folders))
        else:
            result.append(thing)
    return result


def export_all_msgs(path, app):
    all_things = get_all(app=app)
    for thing in all_things:
        if thing.FullFolderPath == path:
            FIXME
            return
    raise ValueError('Could not find path %s in:\n%s\n' % (
        path, '\n'.join([item.FullFolderPath for item in all_things])))

    

    
