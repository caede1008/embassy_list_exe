# process.py
class Process:
    def __init__(self, driver):
        self.driver = driver

    def goPage(self):
        self.driver.get("https://www.mofa.go.jp/mofaj/link/emblist/europe.html")