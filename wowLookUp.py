#!python3
#wowLookUp.py - search your WoW buddies.
#This script looks for a wow character and outputs race, class and item level


import bs4, requests, random

realm = input('On what realm is your buddy? ')
name = input('What\'s his name? ')

adjectives = ['fearless', 'brave', 'silly', 'devious', 'shy', 'powerful', 'skillful']

wowUrl = (f'https://www.wowprogress.com/character/eu/{realm}/{name}')
res = requests.get(wowUrl)
try:
    res.raise_for_status
except:
    print('Coudn\'t connect to server')
soup = bs4.BeautifulSoup(res.text, 'html.parser')
try:
    BasicInfoElem = soup.select('.primary > div:nth-child(4) > div:nth-child(1) > i:nth-child(4)')
    BasicInfo = BasicInfoElem[0].getText().split() #First is race, second is class
    if BasicInfo[0] == 'blood':
        BasicInfo[0] = 'Blood Elf'
        BasicInfo.remove('elf')
    if 'demon' in BasicInfo:
        index = BasicInfo.index('demon')
        BasicInfo.remove('demon')
        BasicInfo.insert(index, 'Demon hunter')

    charRace = BasicInfo[0][0].upper() + BasicInfo[0][1:] #So the race starts with a capital letter.
    print(name + ' is a ' + random.choice(adjectives) + ' ' + charRace + ' '+ BasicInfo[1] )

    GearLevelElem = soup.select('.gearscore')
    print('Has an impressive ' + GearLevelElem[0].getText().lower())
except:
    print('No character found :(')