from os.path import join
from os.path import dirname
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

chromedriver = "/chromedriver"

# driver = webdriver.Chrome()
# driver = webdriver.Edge()
# driver = webdriver.Firefox()
option = webdriver.ChromeOptions()
option.binary_location = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
service = Service(chromedriver)
driver = webdriver.Chrome(service=service, options=option)
driver.get("https://www.wikifolio.com/de/de/w/wf00124816#trades")

# This will get the initial html - before javascript
html1 = driver.page_source

# This will get the html after on-load javascript
html2 = driver.execute_script("return document.documentElement.innerHTML;")

with open(join(dirname(__file__), "html1.html"), "w", encoding="utf-8") as file:
    file.write(html1)

with open(join(dirname(__file__), "html2.html"), "w", encoding="utf-8") as file:
    file.write(html2)
