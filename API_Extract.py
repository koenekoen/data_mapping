import csv
import os
from andy.remote import Andy

print "Let's have a look!"
#help(Andy)

andy = Andy('/api/v1/',id="koen.peters",password="4GtrdxQV1")

target = '/pipeline\activities'
##params = {"serializer": "std",
##          "name__icontains": "ita"}
##filtered = andy.filter('/geo/country', **params)
table = andy.filter(target)

script_dir = os.path.dirname(os.path.abspath(__file__))
dest_dir = os.path.join(script_dir, 'API Extracts')
try:
    os.makedirs(dest_dir) # create ..\API Extracts\  subfolder
except OSError:
    pass
dest_dir = os.path.join(dest_dir, target[1:].replace("/","_") + ".csv")

out = open(dest_dir,"wb")
c = csv.writer(out, dialect='excel')
header = []
for i in table:
    for col in i.keys():
        if col.startswith("u'"):
            col = col[2:]
        if col.endswith("'"):
            col = col[:-1]
        header.append(col)
    c.writerow(header)
    break

for dictionary in table:
    row = []
    for i in dictionary.items():
        v = unicode(i[1]).encode('utf8')
##        if v.startswith("http:/"):
##            v = andy.get(v[26:])
        row.append(v)
    c.writerow(row)

out.close()
print "Job's done"