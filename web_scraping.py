from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
# this is where we load out data.
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB rating'])
print(excel.sheetnames)

try:
    source = requests.get('https://www.imdb.com/chart/top/')
    # this is a good practice, whenever a link doesn't exist, it will throw an error.
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')
    # find is going to fetch the first match. Then, find_all 'tr' tags, then returning a list.
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    # now iterate to each movie within 'tr' tag.
    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text # you are getting the text inside <a href>
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0] # spliting the ranking number and movie_name, turn it into list.
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        rating = movie.find('td',class_="ratingColumn imdbRating").strong.text
        print(rank, name, year, rating)
        #load in the excel
        sheet.append([rank, name, year, rating])
    
except Exception as e:
    print(e)

excel.save('Movie_Ratings.xlsx')

