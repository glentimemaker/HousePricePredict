from bs4 import BeautifulSoup
import urllib2
import re
import urllib
import os
import csv


# Get all the house information urls in searching results
def make_url(url, county, city, zip):
    #quoted_query = urllib.quote(url)
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'}
        req = urllib2.Request(url, headers=headers)
    except IOError:
        print("Error: no page_2")
    else:
        thepage = urllib2.urlopen(req).read()
        soupdata = BeautifulSoup(thepage, "html.parser")

    # re.compile("/homedetails/")
    # , limit=2

        house_urls = soupdata.find_all('article')
        i = 1
    # house_href_initial = " "

        list = []
        for house_href in house_urls:
            house_price_i = house_href.find('span', class_="zsg-photo-card-price")
            if house_price_i:
                house_price = house_price_i.string
                house_zipid = house_href.get('data-zpid')
                house_lat_i = house_href.find('meta', itemprop="latitude")
                house_lat = house_lat_i.get('content')
                house_lon_i = house_href.find('meta', itemprop="longitude")
                house_lon = house_lon_i.get('content')
                house_href2 = house_href.find('a')['href']

            # if house_href2!=house_href_initial:
            #   house_href_initial = house_href2
                url = "http://www.zillow.com" + house_href2
                slist = [house_zipid, county, city, zip, house_lat, house_lon, url, house_price]
                list.append(slist)

                make_soup(url, house_zipid, county, city, zip)
                i += 1
                print house_href2

        with open('DataBaseLess\\' + county + '-ca-testless.csv', 'a') as file:
            writer = csv.writer(file)
            writer.writerows(list)


    #next_page_b = soupdata.find('li', class_="zsg-pagination-next")
    #if next_page_b:
    #    next_page = "http://www.zillow.com" + next_page_b.find('a')['href']
    #    print(next_page)
    #    make_url(next_page, county, city, zip)


# Download images of each house
def make_soup(url, i, county, city, zip):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'}
    req = urllib2.Request(url, headers=headers)
    thepage = urllib2.urlopen(req).read()
    soupdata = BeautifulSoup(thepage, "html.parser")

    imglist = soupdata.find_all(href=re.compile("p_c"))

    foldername = str(i)
    picpath = 'E:\\ImageDownloadTestLess\\' + county + '\\' + zip + '\\%s' % (foldername)
    if not os.path.exists(picpath):
        os.makedirs(picpath)

    x = 0
    for imgurl in imglist:
        s = list(imgurl.get('href'))
        s[34] = 'f'
        a = "".join(s)
        # https: // photos.zillowstatic.com / p_c / IS6qji5piwpwhu0000000000.jpg
        target = picpath + '\\%s.jpg' % x
        try:
            urllib.urlretrieve(a, target)
        except:
            continue
        x += 1
        # print(imglist)


if __name__ == '__main__':
    # --------------- Change this string to get the desire result -----------#
    ifile = open('zipcode.csv', 'rb')
    reader = csv.reader(ifile)
    rownum = 0
    for row in reader:
        if rownum == 0:
            zipcode_header = row #Save header row
        else:
            colnum = 0
            for col in row:
                zipcode_header[colnum] = col
                colnum += 1
            zip = zipcode_header[0]
            city = zipcode_header[1]
            county = zipcode_header[2]
            #for zip in range(92101, 92105):  # Doesn't include the last one of the range
            city_state_zip = city + "-ca-" + str(zip)
            make_url("http://www.zillow.com/" + urllib2.quote(city_state_zip), str(county), str(city), str(zip))
            make_url( "http://www.zillow.com/" + urllib2.quote(city_state_zip) + "/2_p", str(county), str(city), str(zip))
            # + "/2_p/"
        rownum += 1
