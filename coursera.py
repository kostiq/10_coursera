import requests
import random
import re
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_list(count=20):
    courses_xml = requests.get(
        'https://www.coursera.org/sitemap~www~courses.xml').content
    tree = etree.fromstring(courses_xml)
    courses_list = [link.find('{*}loc').text for link in tree]
    return random.sample(courses_list, count)


def get_title(html, soup):
    return (soup.find('h1').text)


def get_start_date(html, soup):
    if soup.find('div', {'class': re.compile('startdate')}):
        return soup.find('div', {'class': re.compile('startdate')}).text
    else:
        return 'No info about startdate'


def get_duration(html, soup):
    duration = ''
    class_list = soup.find_all('td', {'class': 'td-data'})
    if re.search('\d+-\d+', class_list[0].text):
        duration = class_list[0].text
    elif re.search('\d+-\d+', class_list[1].text):
        duration = class_list[1].text
    else:
        duration = 'No info about duration'
    return duration


def get_avr_star(html, soup):
    if soup.find('div', {'class': re.compile('ratings-text')}):
        return soup.find('div', {'class': re.compile('ratings-text')}).text
    else:
        return 'No info about avr rating'


def get_language(html, soup):
    return soup.find('div', {'class': 'language-info'}).text


def get_course_info(course_slug):
    course_info = []
    html = requests.get(course_slug).content
    soup = BeautifulSoup(html, 'html.parser')
    course_info = [course_slug,
                   get_title(html, soup),
                   get_start_date(html, soup),
                   get_duration(html, soup),
                   get_language(html, soup),
                   get_avr_star(html, soup)]
    return course_info


def output_courses_info_to_xlsx(filepath, courses_info):
    wb = Workbook()
    sheet1 = wb.create_sheet(title='Coursera')

    for x, course_info in enumerate(courses_info):
        for y, info in enumerate(course_info):
            sheet1.cell(row=x+1, column=y+1, value=info)

    wb.save(filename=filepath)


if __name__ == '__main__':
    overall_info = []
    for link in get_courses_list():
        overall_info.append(get_course_info(link))
    output_courses_info_to_xlsx('info.xlsx', overall_info)
