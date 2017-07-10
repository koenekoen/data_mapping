import json
import urllib
import csv
import time

charTestList = ['a','e','i','o','u','y']

def stringContainsCharacters(string,charList):
    containsAny = False
    for c in charList:
        if c in string or c.upper() in string:
            containsAny = True
            continue
    return containsAny

def getCityData(name):
    url = 'https://maps.googleapis.com/maps/api/geocode/json?address=' + name.replace(" ", "+")
    j = json.loads(urllib.urlopen(url).read())
    info = []

    if j['status']=='OK': # Found location
        try:
            temp = j['results'][0]['address_components']
            country = str(temp[len(temp)-1]['long_name'])
            if not any(c.isalpha() for c in country):
                country = str(temp[len(temp)-2]['long_name']) # it's either in the last or second-to-last
                if country == 'United States':
                    country = '111' # Something's fishy in that case, probably US is the wrong location.
            if any(c.isalpha() for c in country):
                lat = j['results'][0]['geometry']['location']['lat']
                lng = j['results'][0]['geometry']['location']['lng']
                city = name
                SCIPSLoc = city.upper() + ' (' + country.upper() + ')';

                info = [line[0],SCIPSLoc,city,country,lat,lng,0]
                #break
        except:
            pass
    elif j['status']=='OVER_QUERY_LIMIT':
        print "Reached Daily Query Limit..."
        info = "Done for today"

    time.sleep(0.1)
    return info


with open('NDPs to be mapped.csv', 'r') as f:
    with open('Mapped NDPs.csv', 'wb') as g:
        r = csv.reader(f)
        r.next()
        w = csv.writer(g)
        w.writerow(['NDP (input)', 'SCIPS Location', 'City', 'Country', 'Latitude', 'Longtitude', 'Mapped correctly'])
        nLines = 0
        tStart = time.time()
        done = 0
        for line in r:
            if done == 1:
                print "Done for today :("
                break
            nLines += 1
            name = line[0]
            nonletters = [',','/','-','_','.','(',')',"'"]
            for s in nonletters:
                name = name.replace(' '+s+' ',' ')
                name = name.replace(s+' ',' ')
                name = name.replace(' '+s,' ')
                name = name.replace(s,' ')
            potentialNames = [] # Captures list of names to try to find the lcoation for.

            print line[0]

            exceptions = ["NGO","WAREHOUSE","WFP","MILL","EDP","ICRC","SUPPLIER","OFFICE","CAMP"] # Sub-names that often occur but don't make sense to map
            for item in name.split(' '):
                if item.upper() not in exceptions and len(item)>2:
                    potentialNames.append(item)

            tempRange = range(len(potentialNames)-2)
            for k in range(len(potentialNames)-1):
                potentialNames.append(potentialNames[k] + ' ' + potentialNames[k+1]) # In case some city names are two words, like Abu Dhabi
            for k in tempRange:
                potentialNames.append(potentialNames[k] + ' ' + potentialNames[k+1] + ' ' + potentialNames[k+2]) # In case some city names are two words, like KAFR EL SHEIKH

            foundEntry = False
            for name in potentialNames:
                newInfo = getCityData(name)
                if newInfo!=[]:
                    if newInfo == "Done for today":
                        done = 1
                        break
                    foundEntry = True
                    w.writerow(newInfo)
            if not foundEntry:
                w.writerow([line[0]])

            print "Time elapsed: " + str(int(time.time()-tStart)) + ' seconds.'
            print "Entry: " + str(nLines)
            print "Average time per entry: " + str(round((time.time()-tStart)/nLines,2)) + ' seconds.'
            print ""

