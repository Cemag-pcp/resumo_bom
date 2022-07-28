opts = FirefoxOptions()
opts.add_argument("--headless")
browser = webdriver.Chrome()

browser.get('http://example.com')
st.write(browser.page_source)
