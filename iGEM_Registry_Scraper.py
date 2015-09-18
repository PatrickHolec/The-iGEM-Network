import datetime
import urllib
import lxml.html
import itertools

from openpyxl import Workbook
from bs4 import BeautifulSoup

def GeneralExcel(data,fname):
    wb = Workbook()
    ws = wb.active
    ws['A1'] = datetime.datetime.now()
    for d in data:
        ws.append(d)
    wb.save(fname)


urls = [['http://igem.org/Team_Parts?year=2005',2005],
        ['http://igem.org/Team_Parts?year=2006',2006],
        ['http://igem.org/Team_Parts?year=2007',2007],
        ['http://igem.org/Team_Parts?year=2008',2008],
        ['http://igem.org/Team_Parts?year=2009',2009],
        ['http://igem.org/Team_Parts?year=2010',2010],
        ['http://igem.org/Team_Parts?year=2011',2011],
        ['http://igem.org/Team_Parts?year=2012',2012],
        ['http://igem.org/Team_Parts?year=2013',2013],
        ['http://igem.org/Team_Parts?year=2014',2014]]

final_parts = []

for url,year in urls:
    connection = urllib.urlopen(url)
    dom =  lxml.html.fromstring(connection.read())
    groups = []

    for link in dom.xpath('//a/@href'): # select the url in href for all a tags(links)
        if 'group' in link:
            groups.append(link)
    
    for i,group in enumerate(groups):
        if i < 254 and year == 2014:
            pass
        else:
            print 'Beginning anaylsis of group',str(i+1),'of year',year,'...'
            connection = urllib.urlopen(group)
            dom =  lxml.html.fromstring(connection.read())
            parts = []
            for link in dom.xpath('//a/@href'): # select the url in href for all a tags(links)
                if 'BBa_' in link:
                    parts.append([link,0,'-','Team '+str(i+1),year,'Edges:'])

            parts.sort()
            list(parts for parts,_ in itertools.groupby(parts))
            
            for j,part in enumerate(parts):
                f = urllib.urlopen(part[0])
                myfile = f.read()
                dom =  lxml.html.fromstring(myfile)
                for link in dom.xpath('//a/@href'): # select the url in href for all a tags(links)
                    if 'uses.cgi' in link:
                        ff = urllib.urlopen(link)
                        dom_f =  lxml.html.fromstring(ff.read())
                        for l in dom_f.xpath('//a/@href'):
                            if 'BBa_' in l and not part[0][part[0].index('BBa_'):] in l:
                                parts[j].append(l[l.index('BBa_'):])
                if ' Uses</a></div>' in myfile:
                    ind = myfile.index(' Uses</a></div>')
                    temp = myfile[ind-10:ind]
                    count = int(temp[temp.index("\'>")+2:])
                    parts[j][1] = count
                    print 'Uses for',part[0][part[0].index('BBa_'):],':',count
                elif 'Not Used</div>' in myfile:
                    print 'Uses for',part[0][part[0].index('BBa_'):],':',0
                else:
                    print 'Unrecognized.'
                print 'Total edges:',len(parts[j])-6
                #for i in parts[j][5]: print i

                final_parts.append(parts[j])
            print 'Publishing...'
            try:
                GeneralExcel(final_parts,'iGEM Scrapping Results v2.xlsx')
                print 'Saved.'
            except:
                print 'File in use, not saving.'
        




