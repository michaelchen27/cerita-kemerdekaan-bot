from bs4 import BeautifulSoup
import requests
import selenium.webdriver as webdriver
import openpyxl as xl
import os
import time

edited_excel_path = "D:\\IG_Project\\poyo.xlsx"

if os.path.exists(edited_excel_path):
    os.remove(edited_excel_path)

wb = xl.load_workbook('sorted_form.xlsx')
sheet = wb['Form Responses 1']

counter = 0
for row in range(2, sheet.max_row + 1):
    url = sheet.cell(row, 12).value

    driver = webdriver.Chrome(executable_path="D:\\Downloads\\chromedriver.exe")
    driver.get(url)

    soup = BeautifulSoup(driver.page_source, features="lxml")

    i = 0
    tagExist = False

# Cek Tag @djikp
    for tag in soup.find_all('a', class_='notranslate'):
        if tag.text == "@djikp":
            tagExist = True

    if tagExist:
        sheet.cell(row, 17).value = "ADA"
    else:
        sheet.cell(row, 17).value = "TIDAK ADA"

    wb.save("poyo.xlsx")
# Cek Tag @djikp

# Cek hashtag
    for x in soup.find_all('a', {'class': 'xil3i'}):
        if x.text.lower() == "#ceritakemerdekaan" or x.text.lower() == "#lombapodcastkominfo":
            i += 1

    if i >= 2:
        sheet.cell(row, 20).value = "ADA"
        i = 0

    else:
        sheet.cell(row, 20).value = "TIDAK ADA"
# Cek hashtag

# Cek Durasi Video
    video = soup.find('video')
    if video:
        video_url = video['src']
        driver.get(video_url)


    else:
        print("Video doesn't exist!")
    # sheet.cell(row, 21).value = video.duration

# Cek Durasi Video

    wb.save("poyo.xlsx")
    counter += 1
    print(counter)
