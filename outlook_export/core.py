"""Core functions for exporting data from outlook.
"""

import csv
import os
import re
import logging

import json

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


def make_field_map():
    field_map = {n: n for n in [
        'SenderName', 'SenderEmailAddress', 'ReceivedTime',
        'SentOn', 'To', 'Subject', 'Body']}
    return field_map

def export_msg_to_dict(msg):
    field_map = make_field_map()
    data = {}
    for msg_key, tgt_key in field_map.items():
        value = getattr(msg, msg_key)
        xform = getattr(value, 'isoformat', None)
        if xform is not None:
            value = xform()
        assert tgt_key not in data
        data[tgt_key] = value

    return data


def export_msgs_to_json(thing, outdir, max_msgs, names):
    if not len(thing.Items):
        return
    folder_path = thing.FullFolderPath
    for item in thing.Items:
        if max_msgs is not None and len(names) >= max_msgs:
            return
        data = export_msg_to_dict(item)
        basename = '%s__%s' % (data['SentOn'], data['Subject'])
        basename = re.sub('[^-_a-zA-Z_0-9.+@]', '_', basename)
        full_name = os.path.join(outdir, basename)
        logging.info('Export message to: %s', full_name)
        while os.path.exists(full_name):
            full_name += '_'
        names.append((folder_path, full_name))
        with open(full_name, 'w') as out_fd:
            json.dump(data, out_fd)


def export_msgs_to_csv(thing, outdir, max_msgs, names):
    if not len(thing.Items):
        return
    folder_path = thing.FullFolderPath
    out_file = os.path.join(outdir, 'output.csv')
    header = list(sorted(make_field_map()))
    mode = 'a' if os.path.exists(out_file) else 'w'
    with open(out_file, mode, newline='') as out_fd:
        writer = csv.writer(out_fd)
        if mode == 'w':
            writer.writerow(header)
        for item in thing.Items:
            if max_msgs is not None and len(names) >= max_msgs:
                return
            data = export_msg_to_dict(item)
            basename = '%s__%s' % (data['SentOn'], data['Subject'])
            basename = re.sub('[^-_a-zA-Z_0-9.+@]', '_', basename)
            full_name = os.path.join(outdir, basename)
            logging.info('Export message to: %s', full_name)
            while os.path.exists(full_name):
                full_name += '_'
            names.append((folder_path, full_name))
            writer.writerow([data[n] for n in header])


def export_all_msgs(path, app, outdir, max_folders=1, max_msgs=None,
                    fmt='csv'):
    if os.path.exists(outdir):
        raise ValueError('Output directory %s already exists!' % outdir)
    os.makedirs(outdir)
    all_things = get_all(app=app)
    names = []
    folders_exported = 0

    for thing in all_things:
        if thing.FullFolderPath == path:
            if fmt == 'json':
                export_msgs_to_json(thing, outdir, max_msgs, names)
            elif fmt == 'csv':
                export_msgs_to_csv(thing, outdir, max_msgs, names)
            folders_exported += 1
            if folders_exported >= max_folders:
                return names

    return names
