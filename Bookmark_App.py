


import io
import os
import os.path
import sys

from os import environ
from os import curdir, sep

import portalocker
import tempfile
import types

import socket

import uuid
import csv
import json
import pickle
import copy
import operator
import math
import time
import datetime
from datetime import date
from datetime import timedelta
import urllib
from http.server import BaseHTTPRequestHandler, HTTPServer
import shutil
from shutil import copyfile

import xlrd
from xlrd import open_workbook
import xlwt
from wand.image import Image as image
import xlsxwriter 
import openpyxl
import exifread
import piexif
import PIL
from PIL import Image
import glob
import subprocess
import pynotify

import threading
from threading import Lock
import traceback
import logging
import sqlite3
import jinja2

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pdfkit
from bs4 import BeautifulSoup
import codecs

import img2pdf 
import random
import collections
from collections import OrderedDict, defaultdict, Counter
from functools import wraps   

import flask
from flask import Flask, render_template, g, redirect, session, Response, request, url_for, flash, jsonify, abort

from flask_login import LoginManager, UserMixin, AnonymousUserMixin, login_required, login_user, logout_user, current_user
from flask import send_from_directory, send_file
from flask.views import MethodView
from flask_sqlalchemy import SQLAlchemy 
from flask_bcrypt import Bcrypt 
from flask_wtf import Form, RecaptchaField
from wtforms import Form, BooleanField, StringField, RadioField, SelectField, PasswordField, validators, TextAreaField, SubmitField 
from wtforms.validators import Email, DataRequired
from werkzeug.utils import secure_filename
from cachelib.memcached import MemcachedCache       #py3
from cachelib import SimpleCache






















@application.route('/bookmark', methods=['GET', 'POST'])
def bookmark():

    user_id='demo1234567890'
    email='demoemail@demo.com'

    filenames = load_pickles(user_id, 'FileName'+user_id+'.pickle', file_lock)
    fileauthors = load_pickles(user_id, 'FileAuthor'+user_id+'.pickle', file_lock)

    filedata = load_pickles(user_id, 'FileData'+user_id+'.pickle', file_lock)
    filetitles = load_pickles(user_id, 'FileTitle'+user_id+'.pickle', file_lock)
    filenames = load_pickles(user_id, 'FileName'+user_id+'.pickle', file_lock)
    filesizes = load_pickles(user_id, 'FileSize'+user_id+'.pickle', file_lock)
    filetypes = load_pickles(user_id, 'FileType'+user_id+'.pickle', file_lock)
    filetypeList = load_pickles(user_id, 'FileTypeList'+user_id+'.pickle', file_lock)
    fileauthors = load_pickles(user_id, 'FileAuthor'+user_id+'.pickle', file_lock)
    fileauthorList = load_pickles(user_id, 'FileAuthorList'+user_id+'.pickle', file_lock)
    sharelink = load_pickles(user_id, 'ShareLink'+user_id+'.pickle', file_lock)

    datelist = load_pickles(user_id, 'DateList'+user_id+'.pickle', file_lock)
    filelist = load_pickles(user_id, 'DateFile'+user_id+'.pickle', file_lock)

    notelist = load_pickles(user_id, 'NoteList'+user_id+'.pickle', file_lock)
    bookmarklist = load_pickles(user_id, 'BookmarkList'+user_id+'.pickle', file_lock)
    statuslist = load_pickles(user_id, 'StatusList'+user_id+'.pickle', file_lock)
    progresslist = load_pickles(user_id, 'ProgressList'+user_id+'.pickle', file_lock)
    prioritylist = load_pickles(user_id, 'PriorityList'+user_id+'.pickle', file_lock)
    commentlist = load_pickles(user_id, 'CommentList'+user_id+'.pickle', file_lock)


    templist={}
    for commentid, comment in bookmarklist.items():
        if isinstance(comment, (tuple, list)):
            templist[commentid] = {}
        else:
            templist[commentid] = comment

    bookmarklist=templist
    
    tempmarklist=OrderedDict()

    for dates in datelist:
        for fileid in datefile.get(dates,{}):
            if fileid in bookmarklist:
                tempmarklist[fileid]=bookmarklist[fileid]

    bookmarklist=OrderedDict(sorted(tempmarklist.items(), key=lambda t: t[0], reverse=True))           
    
                
    return render_template('homebookmark.html', bookmarklist=bookmarklist)

@application.route('/saveBookmarks', methods=['GET', 'POST'])  # not found
@login_required
def saveBookmarks():
    
    user_id='demo1234567890'
    email='demoemail@demo.com'

    filenames = load_pickles(user_id, 'FileName'+user_id+'.pickle', file_lock)
    fileauthors = load_pickles(user_id, 'FileAuthor'+user_id+'.pickle', file_lock)

    filedata = load_pickles(user_id, 'FileData'+user_id+'.pickle', file_lock)
    filetitles = load_pickles(user_id, 'FileTitle'+user_id+'.pickle', file_lock)
    filenames = load_pickles(user_id, 'FileName'+user_id+'.pickle', file_lock)
    filesizes = load_pickles(user_id, 'FileSize'+user_id+'.pickle', file_lock)
    filetypes = load_pickles(user_id, 'FileType'+user_id+'.pickle', file_lock)
    filetypeList = load_pickles(user_id, 'FileTypeList'+user_id+'.pickle', file_lock)
    fileauthors = load_pickles(user_id, 'FileAuthor'+user_id+'.pickle', file_lock)
    fileauthorList = load_pickles(user_id, 'FileAuthorList'+user_id+'.pickle', file_lock)
    sharelink = load_pickles(user_id, 'ShareLink'+user_id+'.pickle', file_lock)

    datelist = load_pickles(user_id, 'DateList'+user_id+'.pickle', file_lock)
    filelist = load_pickles(user_id, 'DateFile'+user_id+'.pickle', file_lock)

    notelist = load_pickles(user_id, 'NoteList'+user_id+'.pickle', file_lock)
    bookmarklist = load_pickles(user_id, 'BookmarkList'+user_id+'.pickle', file_lock)
    statuslist = load_pickles(user_id, 'StatusList'+user_id+'.pickle', file_lock)
    progresslist = load_pickles(user_id, 'ProgressList'+user_id+'.pickle', file_lock)
    prioritylist = load_pickles(user_id, 'PriorityList'+user_id+'.pickle', file_lock)
    commentlist = load_pickles(user_id, 'CommentList'+user_id+'.pickle', file_lock)


    timestamp  = datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')[:-3]
    
    datestamp = int(timestamp[:8])     #dates = int(actime[:8])
    
    if request.method == 'POST':
        fileNames = request.form.get('addtitles','')
        filename = fileNames.split('.')[0]
        fileNames = convert_to_unicode(fileNames)
        URLS = request.form.get('addurls','').decode('utf-8-sig').encode('utf-8')

        #thumbnail_url = request.form.get('thumbnail_url','').decode('utf-8-sig').encode('utf-8')
 
        status = request.form.get('addstatuss','UNREAD')
        status = convert_to_unicode(key5)#(status)
        progress = request.form.get('addprogress',0)
        progress = int(progress)
        relevants = request.form.get('addrelevants','')
        relevants = convert_to_unicode(relevants)
        prioritys = request.form.get('addprioritys','')
        prioritys = convert_to_unicode(prioritys)
        newcomment = request.form.get('addcomment','')
        newcomment = convert_to_unicode(newcomment)

        loc1 = request.form.get('addloc1','')
        loc1 = str(loc1)
        key1 = request.form.get('addkey1','')
        key1 = str(key1)

        loc2 = request.form.get('addloc2','')
        loc2 = str(loc2)
        key2 = request.form.get('addkey2','')
        key2 = str(key2)

        loc3 = request.form.get('addloc3','')
        loc3 = str(loc3)
        key3 = request.form.get('addkey3','')
        key3 = str(key3)

        loc4 = request.form.get('addloc4','')
        loc4 = str(loc4)
        key4 = request.form.get('addkey4','')
        key4 = str(key4)

        loc5 = request.form.get('addloc5','')
        loc5 = str(loc5)
        key5 = request.form.get('addkey5','')
        key5 = str(key5)

        loc6 = request.form.get('addloc6','')
        loc6 = str(loc6)
        key6 = request.form.get('addkey6','')
        key6 = str(key6)

        sharelink = str(uuid.uuid4())

        title=filename
        filetype='html'

        fileurl=URLS
        if fileurl == "underfined":
            return redirect(url_for('savebookmarks'))
        if fileurl != "underfined":

            fileid = str(uuid.uuid4())
            #tripletime = actime+actime+actime

            if status in statuslist:
                statuslist[status].append(fileid)
            else:
                statuslist[status]=[fileid]

            progresslist[fileid]=progress

            filenames[fileid]=filename
            filetypes[fileid]=filetype
            if filetype in filetypelist:
                filetypelist[filetype].append(fileid)
            else:
                filetypelist[filetype]=[fileid]
            fileauthors[Tripletime]=email#username#user_id
            if email in fileauthorlist:
                fileauthorlist[email].append(fileid)
            else:
                fileauthorlist[email]=[fileid]

            sharelinks[sharelink]=fileid

            if dates not in datelist:
                datelist.append(dates)

            if dates in datefile:
                if fileid not in datefile[dates]:
                    datefile[dates][fileid]=[fileid]
                #else:
                    #   datefile[dates][Tripletime].append(Tripletime)
            else:
                datefile[dates]={fileid:[fileid]}

            author=[email]

            driver = webdriver.PhantomJS(executable_path=r'/usr/local/bin/phantomjs')#/home/serve624/node_modules/phantomjs/lib/phantom/bin/phantomjs')   # this is now in PATH
            driver.get(fileurl)
            webcontent = driver.page_source
            
            driver.maximize_window()
            driver.save_screenshot(os.path.join(pickle_dir, fileid+'.thumbnail.png'))

            html_path=os.path.join(pickle_dir, fileid+'.html')
            markfile = open(html_path, 'wb')
            markfile.write(webcontent.encode('utf-8'))
            markfile.close
            driver.quit()

            response = requests.get(fileurl)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, 'html.parser')

            body_content = soup.body.get_text(strip=True)

            if not isinstance(body_content, unicode):  
                body_content = body_content.decode('utf-8', 'ignore')

            text_path = os.path.join(pickle_dir, fileid+'.txt')
            with open(text_path, 'wb') as textfiles:
                textfiles.write(body_content.encode('utf-8'))
            
            config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')

            pdf_path = os.path.join(pickle_dir, group_id, tripletime+'.pdf')

            pdfkit.from_string(body_content, pdf_path, configuration=config)


            #need to save html copy

            
            thumbnail = fileid+'.thumbnail.jpg'
            fullfile = fileid+'.'+filetype

            filepath=os.path.join(pickle_dir, fileid+'.'+filetype)

            file_stats = os.stat(filepath)

            file_stats = os.stat(os.path.join(pickle_dir, fileid+'.'+filetype))

            filesize = file_stats.st_size                
            filesizes[fileid] = float(filesize)

            bookmark = {'bookmarkid':fileid, 'filename':filename, 'fileurl':fileurl, 'comment':newcomment, 'status':status, 'priority': prioritys, 'author': author, 'original':fileid}
            bookmarklist[fileid] = bookmark

            if newcomment:
                commentlist[fileid]=newcomment

            secure_save(user_id, 'FileTitle'+user_id, filetitles, file_lock)
            secure_save(user_id, 'FileName'+user_id, filenames, file_lock)
            secure_save(user_id, 'FileSize'+user_id, filesizes, file_lock)
            secure_save(user_id, 'FileType'+user_id, filetypes, file_lock)
            secure_save(user_id, 'FileTypeList'+user_id, filetypelist, file_lock)
            secure_save(user_id, 'FileAuthor'+user_id, fileauthors, file_lock)
            secure_save(user_id, 'FileAuthorList'+user_id, fileauthorlist, file_lock)
            secure_save(user_id, 'ShareLink'+user_id, sharelink, file_lock)

            secure_save(user_id, 'DateList'+user_id, datelist, file_lock)
            secure_save(user_id, 'DateFile'+user_id, filelist, file_lock)

            secure_save(user_id, 'NoteList'+user_id, notelist, file_lock)
            secure_save(user_id, 'BookmarkList'+user_id, bookmarklist, file_lock)
            secure_save(user_id, 'StatusList'+user_id, statuslist, file_lock)
            secure_save(user_id, 'ProgressList'+user_id, progresslist, file_lock)
            secure_save(user_id, 'PriorityList'+user_id, prioritylist, file_lock)
            secure_save(user_id, 'CommentList'+user_id, commentlist, file_lock)

            return render_template('bookmarks_save.html')

    return render_template('bookmarks_save.html')

pickle_dir=os.path.join(os.path.dirname(__file__), str('static'))
pickle_direct=os.path.join(os.path.dirname(__file__), str('static'))
file_lock = Lock()


def normalize_for_json(data):

    if isinstance(data, dict):
        new = {}
        for k, v in data.items():

            # keys → force string
            if isinstance(k, unicode):
                k = k.encode('utf-8')
            else:
                k = str(k)

            new[k] = normalize_for_json(v)

        return new

    elif isinstance(data, list):
        return [normalize_for_json(i) for i in data]

    else:
        return data

def atomic_replace(src, dst):

    try:
        os.rename(src, dst)
    except OSError:
        if os.path.exists(dst):
            os.remove(dst)
        os.rename(src, dst)

def write_atomic_text(path, text_data):

    dir_name = os.path.dirname(path) or '.'
    fd, temp_path = tempfile.mkstemp(dir=dir_name)

    try:
        with os.fdopen(fd, 'w') as tmp:
            tmp.write(text_data)
            tmp.flush()
            os.fsync(tmp.fileno())

        atomic_replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def write_atomic_binary(path, binary_data):

    dir_name = os.path.dirname(path) or '.'
    fd, temp_path = tempfile.mkstemp(dir=dir_name)

    try:
        with os.fdopen(fd, 'wb') as tmp:
            tmp.write(binary_data)
            tmp.flush()
            os.fsync(tmp.fileno())

        atomic_replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def secure_save(user_id, filename, data, filelock):

    filepath = get_teammate_file_path(user_id, filename)

    json_path = filepath + ".json"
    pickle_path = filepath + ".pickle"
    lock_path = filepath + ".lock"

    try:
        if data is None:
            print("Refusing to save None data:", filename)
            return False

        with open(lock_path, 'w') as lock_file:
            portalocker.lock(lock_file, portalocker.LOCK_EX)

            if os.path.exists(json_path):
                backup_path = json_path + ".bak"
                shutil.copy2(json_path, backup_path)

            safe_data = normalize_for_json(data)

            json_text = json.dumps(
                safe_data,
                indent=4,
                sort_keys=True,
                ensure_ascii=False
            )

            pickle_bytes = pickle.dumps(
                data,
                protocol=2
            )

            write_atomic_text(json_path, json_text)
            write_atomic_binary(pickle_path, pickle_bytes)

        return True

    except Exception as e:
        print("Critical Failure:", str(e))
        return False
        
def load_pickles(user_id, name, file_lock):
    file_path = get_teammate_file_path(user_id, name)
    return load_pickle_file(file_path, {})

def load_pickle_file(file_path, default_data=None, lock=None):

    if default_data is None:
        default_data = {}

    def _load():
        try:
            with open(file_path, 'rb') as f:
                try:
                    return pickle.load(f)
                except UnicodeDecodeError:
                    f.seek(0)
                    return pickle.load(f, encoding='latin1')

        except  IOError:
            dir_path = os.path.dirname(file_path)
            if dir_path and not os.path.exists(dir_path):
                os.makedirs(dir_path)
                
            with open(file_path, 'wb') as f:
                pickle.dump(default_data, f, protocol=2)

            return copy.deepcopy(default_data)

        except (pickle.PickleError, EOFError, OSError) as e:
            logger.exception("Failed loading pickle %s", file_path)
            return copy.deepcopy(default_data)

    if lock:
        with lock:
            return _load()

    return _load()

def save_pickle_file(file_path, data, file_lock=None):
    def _save_operation():
        try:
            dir_path = os.path.dirname(file_path)
            if dir_path and not os.path.exists(dir_path):
                os.makedirs(dir_path)
            
            protocol = 2 
            
            with open(file_path, 'wb') as f:
                pickle.dump(data, f, protocol=protocol)
                
                f.flush()
                os.fsync(f.fileno())
                
        except Exception as e:
            err_msg = str(e) if sys.version_info[0] < 3 else e
            print("Error saving {}: {}".format(file_path, err_msg))
            traceback.print_exc()
            raise 

    if file_lock:
        with file_lock:
            _save_operation()
    else:
        _save_operation()

def get_teammate_file_path(viewer_user_id, file_name)
    
    viewer_user_id=str(viewer_user_id)

    file_path=os.path.join(pickle_dir, str(viewer_user_id))

    if not os.path.exists(file_path):
        os.makedirs(file_path)

    return os.path.join(file_path, str(file_name))





if __name__ == "__main__": 
    application.TEMPLATES_AUTO_RELOAD = True 
    application.run(threaded=True)
    application.run(host='0.0.0.0', port=5000, debug=False)

















