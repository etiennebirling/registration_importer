import os
import email
from bs4 import BeautifulSoup
import openpyxl

pilgrims = []
files = []

months = {
    "janv." : "January",
    "föv." : "February",
    "mars" : "March",
    "avr." : "April",
    "mai" : "April",
    "juin" : "June",
    "juil." : "July",
    "août" : "August",
    "sept." : "September",
    "oct." : "October",
    "nov." : "November",
    "dec." : "December",
}

class Pilgrim:

  def __init__(self, _registerNumber, _firstname, _lastname, _familyStatus, _birthday, _chapterName, _chapterInfo, _email, _price):
      self.registerNumber = _registerNumber
      self.firstname = _firstname
      self.lastname = _lastname
      self.familyStatus = _familyStatus
      self.birthday = _birthday
      self.chapterName = _chapterName
      self.chapterInfo = _chapterInfo
      self.email = _email
      self.price = _price

def getBirthday(soup, tags):
    if tags[0].get_text() =="1":
        return getBirthdayOfMaster(soup)
    else:
        birthday = removeAsciiHexCode(tags[4].get_text())
        return birthday.replace(birthday.split(" ")[0], months[birthday.split(" ")[0]])


def getBirthdayOfMaster(soup):
    tds = soup.find_all("td")

    for td in tds:
        text = td.get_text()
        first = text.find("/")
        second = text.rfind("/")
        if first == 2 and second == 5:
            return text

def getLastName(text):
    result = ""
    if text.find("\n") > -1:
        result = text[0: text.find("\n")]
    else:
        result = text
    return result

def getEmail(text):
    result = ""
    if text.find("<br>") > -1:
        result = text[0:text.find("<br>")]
    if text.find("+") > -1:
        result = text[0:text.find("+")]
    return result

def getChapterInfo(text):
    result = ""
    if text.rfind("/") > -1:
        result = text[0:text.rfind("/")]
    return result

def getChapterName(text):
    result = ""
    if text.rfind("/") > -1:
        result = text[text.rfind("/") + 2:len(text)]
    return result

def removeAsciiHexCode(text):
    result = text

    if result.find("=") > -1:
        if result.find("=\nr") > -1:
            result = result.replace("=\nr", "")
        if result.find("=\n") > -1:
            result = result.replace("=\n", " ")
        if result.find("=C3") > -1:
            result = result.replace("=C3", "ö")
        if result.find("=B6") > -1:
            result = result.replace("=B6", "")
        if result.find("=A9") > -1:
            result = result.replace("=A9", "")
        if result.find("=BB") > -1:
            result = result.replace("=BB", "")
        if result.find("=AB") > -1:
            result = result.replace("=AB", "")
        if result.find("=A4") > -1:
            result = result.replace("=A4", "")
        if result.find("=BC") > -1:
            result = result.replace("=BC", "")
        if result.find("=9") > -1:
            result = result.replace("=9", "")

    if result.find("=") > -1:
        hex = result[result.find("=")+1:result.find("=")+3]
        try:
            value = bytearray.fromhex(hex).decode()
        except():
            print("An exception occurred")

        result = result.replace("=" + hex, value)

    if result.find("\n") > -1:
        result = result.replace("\n", "")
    if result.find("\t") > -1:
        result = result.replace("\t", "")

    return result

def getAllFiles():
    dir_path = "input/"
    for file in os.listdir(dir_path):
        if file.find(".eml") > 0:
            files.append(file)
    print(files)

def parseFiles():
    for file in files:
        with open("input/" + file, "r",  encoding="utf-8") as email_file:
            email_message = email.message_from_file(email_file)

        payload = email_message.get_payload()
        soup = BeautifulSoup(payload, "html.parser")

        #getBirthay from master
        birthdayMaster = getBirthdayOfMaster(soup)

        # Extract the text within the div tag
        trs = soup.find_all("tr")

        # Convert the extracted text to a JSON object
        for tr in trs:
            _tags = tr.find_all("td")
            if _tags:
                if _tags[0].get_text() ==  "1" or _tags[0].get_text() ==  "2" or _tags[0].get_text() ==  "3" or _tags[0].get_text() ==  "4" or _tags[0].get_text() ==  "5" or _tags[0].get_text() ==  "6" or _tags[0].get_text() ==  "7":
                    print("1", _tags[1].get_text())
                    print("2", _tags[2].get_text())
                    print("3", _tags[3].get_text())
                    print("4", _tags[4].get_text())
                    print("5", _tags[5].get_text())
                    print("6", _tags[6].get_text())
                    print("8", _tags[8].get_text())

                    p = Pilgrim(
                        removeAsciiHexCode(file[file.find("=")+1:len(file)-5]),
                        removeAsciiHexCode(_tags[1].get_text()[_tags[1].get_text().find(" ") + 1:len(_tags[1].get_text())]),
                        getLastName(removeAsciiHexCode(_tags[2].get_text())),
                        removeAsciiHexCode(_tags[3].get_text()),
                        getBirthday(soup, _tags),
                        getChapterInfo(removeAsciiHexCode(_tags[5].get_text())),
                        getChapterName(removeAsciiHexCode(_tags[5].get_text())),
                        getEmail(removeAsciiHexCode(_tags[6].get_text())),
                        removeAsciiHexCode(_tags[8].get_text())
                    )
                    pilgrims.append(p)

        print(pilgrims)

def extractToExcel():
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()
    # Select the default sheet (usually named 'Sheet')
    sheet = workbook.active
    # Add data to the Excel sheet
    data = [
        ["Register Number", "First name", "Last name", "Family status", "Birthday", "Chapter info", "Chapter name", "Price"],
    ]
    for pilgrim in pilgrims:
        data.append([
            pilgrim.registerNumber,
            pilgrim.firstname,
            pilgrim.lastname,
            pilgrim.familyStatus,
            pilgrim.birthday,
            pilgrim.chapterInfo,
            pilgrim.chapterName,
            pilgrim.price])
    for row in data:
        sheet.append(row)
    # Save the workbook to a file
    workbook.save("output/registrations.xlsx")
    # Print a success message
    print("Excel file created successfully!")

if __name__ == '__main__':
    getAllFiles()
    parseFiles()
    extractToExcel()


