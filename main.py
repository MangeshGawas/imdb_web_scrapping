from bs4 import BeautifulSoup
import requests,openpyxl


excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = "Top Rated TV show"
print(excel.sheetnames)

sheet.append(['Show Rank',"Show Name", "Year of Release","Rating"])

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
session = requests.Session()



try:
    
    source = session.get('https://m.imdb.com/chart/toptv/', headers=headers)
    source.raise_for_status()
    soup = BeautifulSoup(source.text,'html.parser')

    show_unordered_list = soup.find('ul',class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-a1e81754-0 eBRbsI compact-list-view ipc-metadata-list--base")
    show_list = show_unordered_list.find_all('li')
    print(len(show_list))

    for show in show_list:
        name_element = show.find('div', class_="ipc-metadata-list-summary-item__c").find('div', class_="ipc-metadata-list-summary-item__tc").find('div', class_="sc-b189961a-0 hBZnfJ cli-children").find('div', class_="ipc-title ipc-title--base ipc-title--title ipc-title-link-no-icon ipc-title--on-textPrimary sc-b189961a-9 iALATN cli-title").find('a', class_="ipc-title-link-wrapper").h3
        if name_element:
            name_and_rank = name_element.get_text(strip=True)
            name = name_element.get_text(strip=True).split(".", 1)[-1].strip()
            rank = name_and_rank.split(".")[0].strip()
            year = show.find('div', class_="ipc-metadata-list-summary-item__c").find('div', class_="ipc-metadata-list-summary-item__tc").find('div', class_="sc-b189961a-0 hBZnfJ cli-children").find('div',class_="sc-b189961a-7 feoqjK cli-title-metadata").find('span',class_="sc-b189961a-8 kLaxqf cli-title-metadata-item").text
            rating = show.find('div', class_="ipc-metadata-list-summary-item__c").find('div', class_="ipc-metadata-list-summary-item__tc").find('div', class_="sc-b189961a-0 hBZnfJ cli-children").find('span',class_="sc-b189961a-1 kcfvgk").find('div').text
            print(rank, name,year,rating)
            sheet.append([rank, name,year,rating])
            # break
    # print(show_unordered_list)

except Exception as e:
    print(e)


excel.save("IMDB TV show Rating.xlsx")