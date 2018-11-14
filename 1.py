import requests
from bs4 import BeautifulSoup

n=1
url="http://www.letpub.com.cn/index.php?journalid="+str(n)+"&page=journalapp&view=detail"
code="utf-8"

r=requests.get(url)
r.raise_for_status()
r.encoding=code
#print(r.text)
soup = BeautifulSoup(r)
soup.find_all( 'tr' )
