from selenium import webdriver
# seleciona o navegador
driver = webdriver.Chrome()
driver.get("https://www.selenium.dev/selenium/web/web-form.html")

title = driver.title

driver.implicitly_wait(10)