from instagramy import InstagramUser
import statistics
import pandas as pd
from urllib import request
from bs4 import BeautifulSoup
import requests
from PIL import Image
from instagramy.plugins.download import *
import re
import os
import unicodedata

path = 'filed/instagram/'

def googleSearch(name):
    googleTrendsUrl = 'https://google.com'
    response = requests.get(googleTrendsUrl)
    if response.status_code == 200:
        g_cookies = response.cookies.get_dict()
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5)\
        AppleWebKit/537.36 (KHTML, like Gecko) Cafari/537.36'}
    google = f"https://www.google.com/search?q={name}"
    html = requests.get(google, headers=headers, cookies=g_cookies)
    scoup = BeautifulSoup(html.text, 'html.parser')
    for tag in scoup.find_all('h3'):
        return tag.text
    return name

def wikiSearch(name):
    name = name.replace(' ', '_')
    print(name)
    if name.isascii() == False:
        name = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore')
        name = name.decode('utf-8')
    website = f"https://en.wikipedia.org/wiki/{name}"
    html = request.urlopen(website).read()
    scoup = BeautifulSoup(html, 'html.parser')
    text = scoup.get_text()
    age  = re.findall(r'\(age.+[0-9]{2}\)', text)
    return int(re.findall(r'[0-9]+', age[0])[0])

def getAgeLocation():
    pass

def download_post_images(name, url, num):
    with open(f'{path}{name}_post{num}.png', 'wb') as f:
        f.write(requests.get(url).content)
        f.close()


def getUserDetails(name):
    details = {}
    posts = []
    add_posts_details = {}
    user = InstagramUser(name)
    details['fullname'] = user.fullname
    name = user.username
    details['username'] = user.username
    details['age'] = wikiSearch(googleSearch(name))
    details['Topics'] = user.biography.split('.')[0]
    details['followers'] = user.number_of_followers
    details['following'] = user.number_of_followings
    details['Avg_Likes'] = statistics.mean([i[0] for i in user.posts])
    details['Avg_Comment'] = statistics.mean([i[1] for i in user.posts])
    details['Engagement_Rate'] = f"{round(((details['Avg_Likes'] + details['Avg_Comment'])/details['followers'])*100, 2)}%"
    details['Total_posts'] = user.number_of_posts
    details['Email'] = user.user_data['business_email']
    details['Website'] = user.user_data['external_url']
    details['Biography'] = user.biography
    details['Gender'] = None
    details['Location'] = None
    details['Language'] = None
    details['YouTube'] = getSocialmediaLinks(details['Website'], 'youtube')
    details['Twitter'] = getSocialmediaLinks(details['Website'], 'twitter')
    details['TikTok'] = getSocialmediaLinks(details['Website'], 'tiktok')
    details['Facebook'] = getSocialmediaLinks(details['Website'], 'facebook')
    details['profile_pic_url'] = user.profile_picture_url
    details['Picture'] = None
    details['Verified'] = ("True" if user.is_verified else "False")
    details['Account Type'] = ("Business" if user.user_data['is_business_account'] else "Creator" if user.user_data['is_professional_account'] else "Personal")
    download_profile_pic(username=user.username, filepath=f'{path}{name}.png')
    for num, post in enumerate(user.posts):
        post_details = []
        add_posts_details['Posts'] = None
        post_details.append(add_posts_details['Posts'])
        add_posts_details['Description'] = post[2]
        post_details.append(add_posts_details['Description'])
        add_posts_details['Hastags'] = None
        post_details.append(add_posts_details['Hastags'])
        add_posts_details['Tags'] = None
        post_details.append(add_posts_details['Tags'])
        add_posts_details['Likes'] = post[0]
        post_details.append(add_posts_details['Likes'])
        add_posts_details['Comments'] = post[1]
        post_details.append(add_posts_details['Comments'])
        add_posts_details['Image/Video URL'] = post[8]
        post_details.append(add_posts_details['Image/Video URL'])
        add_posts_details['Post URL'] = post[7]
        post_details.append(add_posts_details['Post URL'])
        download_post_images(name, post[9], num)
        posts.append(post_details)
    details['no_of_posts'] = posts
    return details
        
def getSocialmediaLinks(website = None, args=None):
    link = ''
    if website != None or website != '':
        html = request.urlopen(website)
        scoup = BeautifulSoup(html, 'html.parser')
        for tag in scoup('a'):
            link = tag.get('href', None)
            if(link != None):
                if args == 'youtube' and args in link and args != None:
                    return link
                elif args == 'twitter' and args in link and args != None:
                    return link
                elif args == 'tiktok' and args in link and args != None:
                    return link
                elif args == 'facebook' and args in link and args != None:
                    return link
    return link

def saveExcel(excelFileName, add_details):
    fieldnames = ['fullname', 'username', 'age', 'Topics', 'followers', 'following', 'Avg_Likes', 'Avg_Comment', 'Engagement_Rate', 'Total_posts', 'Email', 'Website', 'Biography', 'Gender', 'Location', 'Language', 'YouTube', 'Twitter', 'TikTok', 'Facebook', 'profile_pic_url', 'Picture','Verified', 'Account Type', 'no_of_posts']
    second_fieldnames = ['Posts', 'Description', 'Hastags', 'Tags', 'likes', 'comments', 'Image/Video URL', 'Post URL']
    df = pd.DataFrame(add_details, columns=fieldnames)
    print("************ Saving data to excel file ************")
    with pd.ExcelWriter(f'{path}{excelFileName}', engine='xlsxwriter') as writer:
        for i in range(len(df)):
            name = df.loc[i, 'username']
            df1 = df.loc[i, 'fullname':'Account Type']
            pd.DataFrame(df1).T.to_excel(writer, sheet_name=name)
            df2 = pd.DataFrame(df.loc[i, 'no_of_posts'], columns=second_fieldnames)
            df2.to_excel(writer, sheet_name=name, columns=second_fieldnames, startrow=5)

            workbook  = writer.book
            worksheet = writer.sheets[name]
            img = Image.open(f'{path}{name}.png')
            img.thumbnail((100, 100))
            img.save(f'{path}{name}.png')
            image_width = img.width
            image_height = img.height
            cell_width = image_width*0.75
            cell_height = image_width*0.75
            x_scale = cell_width/image_width
            y_scale = cell_height/image_height
            worksheet.set_column(22,22,10)
            worksheet.set_row(1, 55)
            worksheet.insert_image('W2', f'{path}{name}.png', {'x_scale': x_scale, 'y_scale': y_scale})

            for j in range(len(df2)):
                try:
                    image = Image.open(f'{path}{name}_post{j}.png')
                except Exception as e:
                    print(e)
                    continue
                image.thumbnail((100, 100))
                image.save(f'{path}{name}_post{j}.png')
                image_width = image.width
                image_height = image.height
                cell_width = image_width*0.75
                cell_height = image_width*0.75
                x_scale = cell_width/image_width
                y_scale = cell_height/image_height
                worksheet.set_column(1,1,10)
                worksheet.set_row(6+j, 55)
                worksheet.insert_image(f'B{7+j}', f'{path}{name}_post{j}.png', {'x_scale': x_scale, 'y_scale': y_scale})




# Connecting the profile
if __name__ == '__main__':
    names = []
    with open(f'{path}influencers.txt', 'r') as name:
        names = name.readlines()

    influencers = list(map(lambda name : name.strip(), names))
    add_details = []
    for user in influencers:
        add_details.append(getUserDetails(user))

    saveExcel('Influencers.xlsx', add_details)

    filename = os.listdir(path)
    for name in filename:
        if '.png' in name:
            os.remove(os.path.join(path, name))