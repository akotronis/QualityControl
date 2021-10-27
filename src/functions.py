import base64
from math import log
import numpy as np
import os
import pandas as pd
import PySimpleGUI as sg
import subprocess
import sys
import tempfile
import webbrowser





def mycprint(w):
    '''A closure that carries the window object to use from printing to the app console'''
    def cprint(*args, **kwargs):
        message = args[0]
        ML_KEY = '-ML-'
        if kwargs.get('u'):
            w[ML_KEY].update('')
            if not message.strip():
                return cprint
        cp = sg.cprint
        # If l!= False, print a line ABOVE the output message
        if kwargs.get('l'):
            cp(100*'=')
        if kwargs.get('cons'):
            print(message)
        kwargs = {k:v for k,v in kwargs.items() if k not in ['w','u','cons','l']}
        cp(*args, **kwargs)
    return cprint


def timer(start, end):
    '''Input start and end in seconds.
       Returns a string of the form ..h:..m:..s
    '''
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    if not (hours or minutes):
        txt, vals = "{:05.2f}sec", (seconds,)
    elif not hours:
        txt, vals = "{}min:{:05.2f}sec", (int(minutes),seconds)
    else:
        txt, vals = "{}hour:{}min:{:05.2f}sec", (int(hours),int(minutes),seconds)
    return txt.format(*vals)


def join_columns(df, col_list, sep='-'):
    new_col = df[col_list].apply(lambda r:sep.join([str(x) for x in r]), axis='columns')
    return new_col


def diffs_old(r):
    output = []
    for i,v in enumerate(r):
        if i < len(r) - 1:
            if r[i] == 0 or r[i+1] == 0:
                item = np.nan
            else:
                item = r[i] * log(r[i] / r[i+1]) + r[i+1] - r[i]
            output.append(item)
    return output


def diffs(df):
    diff_columns = ['Diff_{}'.format(i+1) for i in range(len(df.columns)-1)]
    records = [diffs_old(row.values) for _, row in df.iterrows()]
    return pd.DataFrame.from_records(records, index=df.index, columns=diff_columns)


def newton(f, Df, x0, epsilon=1e-6, max_iter=100):
    # https://www.math.ubc.ca/~pwalls/math-python/roots-optimization/newton/
    '''Approximate solution of f(x)=0 by Newton's method.
    Parameters
    ----------
    f : function
        Function for which we are searching for a solution f(x)=0.
    Df : function
        Derivative of f(x).
    x0 : number
        Initial guess for a solution f(x)=0.
    epsilon : number
        Stopping criteria is abs(f(x)) < epsilon.
    max_iter : integer
        Maximum number of iterations of Newton's method.
    Returns
    -------
    xn : number
        Implement Newton's method: compute the linear approximation
        of f(x) at xn and find x intercept by the formula
            x = xn - f(xn)/Df(xn)
        Continue until abs(f(xn)) < epsilon and return xn.
        If Df(xn) == 0, return None. If the number of iterations
        exceeds max_iter, then return None.
    Examples
    --------
    >>> f = lambda x: x**2 - x - 1
    >>> Df = lambda x: 2*x - 1
    >>> newton(f,Df,1,1e-8,10)
    Found solution after 5 iterations.
    1.618033988749989
    '''
    xn = x0
    try:
        for n in range(0,max_iter):
            fxn = f(xn)
            if abs(fxn) < epsilon:
                # print('Found solution after',n,'iterations.')
                return xn
            Dfxn = Df(xn)
            if Dfxn == 0:
                # print('Zero derivative. No solution found.')
                return None
            xn = xn - fxn/Dfxn
        # print('Exceeded maximum iterations. No solution found.')
        return None
    except:
        return None


def progress_bar(key, iterable, *args, title='', **kwargs):
    """
    Takes your iterable and adds a progress meter onto it
    :param key: Progress Meter key
    :param iterable: your iterable
    :param args: To be shown in one line progress meter
    :param title: Title shown in meter window
    :param kwargs: Other arguments to pass to one_line_progress_meter
    :return:
    """
    sg.set_options(element_padding=((5, 5),(5, 5)))
    sg.one_line_progress_meter(title, 0, len(iterable), key, *args, **kwargs)
    for i, val in enumerate(iterable):
        yield val
        if not sg.one_line_progress_meter(title, i+1, len(iterable), key, *args, **kwargs):
            break


def toggle_menu(menu_def, window, enable=True):
    for item in menu_def:
        if item[0].startswith('!') and enable:
            item[0] = item[0][1:]
        if not item[0].startswith('!') and not enable:
            item[0] = '!' + item[0]
    window['-MN-'].Update(menu_def)
    return menu_def


def popup_yes_no(title='', message=''):
    '''A confirmation popup
    '''
    message += '\n\nConfirm?\n'
    sg.set_options(auto_size_buttons=False)
    layout = [
        [sg.Text(message, auto_size_text=True, justification='center')],
        [sg.Button('Yes'), sg.Button('No')]
    ]
    window = sg.Window(title, layout, finalize=True, modal=True, element_justification='c', size=(300,150))
    event, values = window.read()
    window.close()
    return event == 'Yes'


def subprocess_call(*args, **kwargs):
    #also works for Popen. It creates a new *hidden* window, so it will work in frozen apps (.exe).
    IS_WIN32 = 'win32' in str(sys.platform).lower()
    if IS_WIN32:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags = subprocess.CREATE_NEW_CONSOLE | subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        kwargs['startupinfo'] = startupinfo
    retcode = subprocess.call(*args, **kwargs)
    return retcode


def show_docs(_type='html'):
    if _type == 'html':
        from constants import HTML_DOCS
        with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html') as f:
            _path = f'file://{f.name}'
            f.write(HTML_DOCS)
        webbrowser.open(_path)
    elif _type == 'pdf':
        from constants import PDF_BASE64_DOCS
        with tempfile.NamedTemporaryFile('wb', delete=False, suffix='.pdf') as f:
            _path = f'{f.name}'
            f.write(base64.b64decode(PDF_BASE64_DOCS))
        webbrowser.open_new(_path)
    print(_path)