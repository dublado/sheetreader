#!/usr/bin/env python
# encoding: utf-8

'''
Functions for reading csv and Excel files.
'''

from xlrd import open_workbook,biffh,xldate_as_tuple
import csv,_csv
import re
from StringIO import StringIO

def parsevalue(v):
    r'''
    Returns a number, date or string, according to value format.
    >>> parsevalue('1.2')
    1.2
    >>> parsevalue('\xe3')
    u'\xe3'
    >>> parsevalue('\xc3\xa3')
    u'\xe3'
    >>> parsevalue('3,2')
    3.2
    >>> parsevalue('R$ 3,2')
    3.2
    >>> parsevalue('01-12-17')
    u'2017-12-01'
    '''
    if isinstance(v,str):
        v=re.sub(r'^(?:R\$ )?(\d+),(\d+)$',r'\1.\2',v)
        v=re.sub(r'^([0123]?\d)[/\-\.]([01]?\d)[/\-\.](\d?\d)$',r'20\3-\2-\1',v)
    try:
        return float(v)
    except:pass
    try:
        return v.decode('utf-8')
    except:pass
    return v.decode('latin-1')

def higherfreq(t,chars):
    '''
    Returns the most repeat character in a string.
    >>> higherfreq('banana','ban')
    'a'
    '''
    top=0
    char=''
    for c in chars:
        if t.count(c)>top:
            top=t.count(c)
            char=c
    return char

def parse(f):
    r'''
    >>> parse('Teste,1,2,3\nA,12/12/17,"R$ 4,50","Teste\nTeste"')
    [[u'Teste', 1.0, 2.0, 3.0], [u'A', u'2017-12-12', 4.5, u'Teste\nTeste']]
    >>> parse('Teste;1;2;3,0\nA;12/12/17;"R$ 4,50";"Teste\nTeste"')
    [[u'Teste', 1.0, 2.0, 3.0], [u'A', u'2017-12-12', 4.5, u'Teste\nTeste']]
    '''
    if isinstance(f,str):
        try:
            f=open(f).read()
        except: pass
    else:
        f=f.read()
    try:
        wb=open_workbook(file_contents=f)
        sheet=wb.sheets()[0]
        lines=sheet.nrows
        cols=sheet.ncols
        results=[]
        for row in range(lines):
            values=[]
            for col in range(cols):
                v=sheet.cell(row,col).value
                if sheet.cell(row,col).ctype==3:
                    v='%04i-%02i-%02i %02i:%02i:%02i' % xldate_as_tuple(v,wb.datemode)
                    v=v.replace(' 00:00:00','')
                values.append(v)
            results.append(values)
    except biffh.XLRDError:
        delimiter=higherfreq(f,',;\t')
        quotechar=higherfreq(f,'\'"')
        results=list(csv.reader(StringIO(f), delimiter=delimiter, quotechar=quotechar))
    results=[[parsevalue(v) for v in l] for l in results]

    return results

def sheet2dict(lines):
    '''
    Translates a spreadsheet data list into a dict.
    >>> sheet2dict([['name','email'],['John','john@doe.com'],['Foo','foo@bar.com']])
    [{'name': 'John', 'email': 'john@doe.com'}, {'name': 'Foo', 'email': 'foo@bar.com'}]
    '''
    headers=lines.pop(0)
    return [dict(map(None,headers,i)) for i in lines]

if __name__ == '__main__':
    import sys,json
    if '--test' in sys.argv:
        import doctest
        doctest.testmod()
    for i in sys.argv[1:]:
        if not i.startswith('-'):
            print json.dumps(parse(i),indent=2)

