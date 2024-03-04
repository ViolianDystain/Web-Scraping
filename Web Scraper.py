import requests
from bs4 import BeautifulSoup
import re
import pandas as pd
from openpyxl import Workbook, load_workbook

blacklist = ["ntv", "emis", "bizi", "senoven", "dnb", "bloomberg",
             "instagram", "wikipedia", "youtube", "pdf", "linkedin", "update.company"]


def read_excel(filepath, column):

    sheet_name = 'Sheet1'
    column_name = 'A'

    df = pd.read_excel(filepath, sheet_name=sheet_name)

    column_data = df.iloc[:, column]

    return column_data


def google_search(query):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    response = requests.get(
        f"https://www.google.com/search?q={query}", headers=headers)

    try:

        if response.status_code == 200:
            return response.text
        else:
            return "Error in the Google Search"
    except requests.Timeout:
        return "Request timed out"

    except requests.ConnectionError:
        return "Connection error"

    except Exception as e:
        return f"An error occurred: {e}"


def get_request(query):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    response = requests.get(
        query, headers=headers, timeout=5)

    if response.status_code == 200:
        return response.text
    else:
        return "Error in url request"

# egMi0 kCrYT


def check_if_member_in_string(array, string):
    for item in array:
        if item in string:
            return True  # Return True as soon as any array member is found in the string
    return False  # Return False if no array members are found in the string


def parse_results(html):
    soup = BeautifulSoup(html, 'html.parser')
    # Find the first search result link
    results = []
    regex = re.compile(r'(.*)(htt.+)(&ved.+)')

    for result_div in soup.find_all('div', class_='egMi0 kCrYT', limit=7):
        if not check_if_member_in_string(blacklist, str(result_div)):
            link = result_div.find('a')['href']
            match = regex.search(link)
            url = match.group(2)
            results.append(url)

    replaced = [regex.sub('', item) for item in results]
    return results


def parse_result_urls(html, url):
    soup = BeautifulSoup(html, 'html.parser')
    # Find the first search result link
    results = []
    regex = re.compile(r'(\d| |\(|\)|\+){5,}')

    texts = soup.find_all(string=regex)
    if texts:
        for text in texts:
            regex.match()
            print(text + " " + url)
            print("---------------------------------------------------------------------------------------------------------------")

    return 0


def search_keywords(html, keywords):
    soup = BeautifulSoup(html, 'html.parser')
    for key in keywords:
        if isinstance(key, str):
            case_insensitive_key = re.compile(key, re.IGNORECASE)
            if soup.find_all(string=case_insensitive_key):
                return True
                break
    return False


def insert_url(filepath, row, url):
    workbook = load_workbook(filepath)

    sheet = workbook['Sheet1']
    coordinate = "C" + str(row)
    sheet[coordinate] = url

    workbook.save(filepath)


def main():
    row = 1

    companies = read_excel(
        r"C:\Users\utku.ondin\Downloads\Taranacak Liste 1 (1).xlsx", 0)

    keywords = read_excel(
        r"C:\Users\utku.ondin\Downloads\Taranacak Liste 1 (1).xlsx", 5)

    for company in companies:
        row += 1
        print("Company " + str(row) + ": " + company)

        resulturls = parse_results(google_search(company))
        for resulturl in resulturls:
            print("url: " + resulturl)
            try:
                if search_keywords(get_request(resulturl), keywords):
                    insert_url(
                        r"C:\Users\utku.ondin\Downloads\Taranacak Liste 1 (1).xlsx", row, resulturl)
                    break
            except Exception as e:
                pass


if __name__ == "__main__":
    main()
