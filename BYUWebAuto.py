import os
import zipfile
from datetime import datetime
from titlecase import titlecase
from tkinter import *


def GetFileName(
        slug):  # this takes a slug and returns the name of the cooresponding file, capitalization doesn't matter
    setups = os.listdir('//Users/scot/Box/BYU Radio/Top of Mind/2020 Setup Sheets/')
    setupsLower = [x.lower() for x in setups]  # makes all titles in listdir lowercase for cap-free searching
    fileName = str([s for s in setupsLower if slug.lower() in s])
    if len(fileName) == 2:
        setupSheet = "NO FILE FOUND WITH \"" + slug + "\" IN NAME"
    elif fileName.count(".docx") > 1:
        setupSheet = "MORE THAN ONE POSSIBLE FILE: " + fileName
    else:
        cleanName = setupsLower.index(fileName[2:len(fileName) - 2])
        setupSheet = setups[cleanName]  # returns properly capitalized filename based on index
    return setupSheet


def CleanWord(fileSlug):
    fileName = GetFileName(fileSlug)
    doc = zipfile.ZipFile('//Users/scot/Box/BYU Radio/Top of Mind/2020 Setup Sheets/' + fileName)
    content = doc.read('word/document.xml').decode(
        'utf-8')  # .decode() turns type bytes into string for easy manipulation
    cleaned = re.sub('<(.|\n)*?>', '', content)
    return cleaned


def GetHeadline(cleaned):
    # get headline
    headline = cleaned[cleaned.find("HEADLINE: ") + len("HEADLINE: "):cleaned.find("SEGMENT")]
    return headline


def GetGuest(cleaned):
    # get name and title of guest--should be reviewed for proper formatting
    nameAndTitle = "Name and Title of Guest(s): "
    if cleaned.find(
            "Pronunciation:") != -1:  # if "Pronunciation has been deleted, the string will stop at "Pre-Record:"
        guest = cleaned[cleaned.find(nameAndTitle) + len(nameAndTitle):cleaned.find("Pronunciation:")]
    elif cleaned.find("Pre-Record") != -1:
        guest = cleaned[cleaned.find(nameAndTitle) + len(nameAndTitle):cleaned.find("Pre-Record:")]
    else:
        guest = cleaned[cleaned.find(nameAndTitle) + len(nameAndTitle):cleaned.find("FOR AIR:")]
    return guest


def GetIntroCopy(cleaned, fileSlug):  # EDIT THIS SO IT DOESN'T HAVE TO TAKE fileSlug
    # gets intro, goes until it sees "QUESTIONS:" or "OUTRO COPY", whichever comes first
    introCopy = "INTRO COPY (LIVE-READ, WRITTEN-TO-SOUND"
    # guestName = guest[0:guest.find(" ")] #gives first name of guest to be used as stop point for intro
    indexIntroCopyLeft = cleaned.find(introCopy) + len(introCopy) + 3
    if cleaned.find("QUESTIONS:") < cleaned.find("OUTRO COPY") and cleaned.find("QUESTIONS:") != -1:
        intro = cleaned[indexIntroCopyLeft:indexIntroCopyLeft + cleaned[indexIntroCopyLeft:len(cleaned)].find(
            "QUESTIONS:")]  # guestName)]
    elif cleaned.find("Welcome.") < cleaned.find("OUTRO COPY") and cleaned.find("Welcome.") != -1:
        intro = cleaned[indexIntroCopyLeft:indexIntroCopyLeft + cleaned[indexIntroCopyLeft:len(cleaned)].find(
            "Welcome.")]  # guestName)]
    else:
        intro = cleaned[
                indexIntroCopyLeft:indexIntroCopyLeft + cleaned[indexIntroCopyLeft:len(cleaned)].find("OUTRO COPY")]

    # if it's for Prime Cuts, this gives the original air date
    # may need some debugging for different scenarios
    if GetFileName(fileSlug).find("PRIME CUTS") != -1:
        forAir = cleaned[cleaned.find("FOR AIR") + len("FOR AIR: "):len(cleaned)]
        origAired = "(Originally aired" + forAir[0:forAir.find("PRIME")] + ")"
    else:
        origAired = ""

    return intro + origAired


def ScrapeWord(fileSlug):  # this goes into a word document and returns titlecase headline, guest info, and intro copy
    cleaned = CleanWord(fileSlug)

    headline = GetHeadline(cleaned)
    guest = GetGuest(cleaned)
    intro = GetIntroCopy(cleaned, fileSlug)

    # concatenate info, titlecase headline and guest
    info = titlecase(headline) + "\n" + "Guest: " + titlecase(guest) + "\n" + intro + "\n\n"
    return info


def GiveEpisodeInfo(fileName): # returns titlecase headline, guest info, and intro copy from word
    cleaned = CleanWord(fileName)

    # get name and title of guest--should be reviewed for proper formatting
    nameAndTitle = "Name and Title of Guest(s): "
    if cleaned.find(
            "Pronunciation:") != -1:  # if "Pronunciation has been deleted, the string will stop at "Pre-Record:"
        guest = cleaned[cleaned.find(nameAndTitle) + len(nameAndTitle):cleaned.find("Pronunciation:")]
    else:
        guest = cleaned[cleaned.find(nameAndTitle) + len(nameAndTitle):cleaned.find("Pre-Record:")]

    info = guest[0:guest.find(",")] + " of " + guest[guest.find(","):len(guest)] + " on " + fileName.lower() + "."
    return info


# Work below here, may need some functions above

def ListSlugs(): #TODO change this to adjust number of slugs, or don't get
    slugs = [slug_entry_1.get(), slug_entry_2.get(), slug_entry_3.get(),
             slug_entry_4.get(), slug_entry_5.get(), slug_entry_6.get()]
    return slugs


def PrintSlugs():
    slugs = ListSlugs()
    j = 0
    for i in slugs:
        j += 1
        if len(i) == 0:
            Label(root, text="NO SLUG ENTERED").grid(row=j, column=3)
        else:
            Label(root, text=GetFileName(i)).grid(row=j, column=3)
    create_txt_done.grid_forget()


def InsertTxt(webText):  # this opens/creates a new .txt file, writes webText to it, closes .txt file
    now = datetime.now()
    today = now.strftime("%m-%d-%Y")
    file = "//Users/scot/Desktop/Setups/WebsiteAuto/" + today + ".txt"
    f = open(file, "w+")
    f.write(webText)
    f.close()


def EpInfo():
    slugs = ListSlugs()
    episodeInfo = ''
    for i in slugs:
        episodeInfo += GiveEpisodeInfo(GetFileName(i))
    e = Entry(root)
    e.grid(row=7, column=3)
    e.insert(0, episodeInfo)
    create_txt_done.grid_forget()


def CreateTxt():
    slugs = ListSlugs()
    text = ""
    for x in slugs:
        text += ScrapeWord(x)
    InsertTxt(text)
    create_txt_done.grid(row=8, column=1)

root = Tk()
root.title("BYU Top of Mind Web Automated")

title = Label(root, text="BYU Web Automated", compound=CENTER)
show_name = Label(root, text="Top of Mind with Julie Rose", compound=CENTER)

slug_label_1 = Label(root, text="Enter Slug 1:")
slug_label_2 = Label(root, text="Enter Slug 2:")
slug_label_3 = Label(root, text="Enter Slug 3:")
slug_label_4 = Label(root, text="Enter Slug 4:")
slug_label_5 = Label(root, text="Enter Slug 5:")
slug_label_6 = Label(root, text="Enter Slug 6:")

slug_entry_1 = Entry(root)
slug_entry_2 = Entry(root)
slug_entry_3 = Entry(root)
slug_entry_4 = Entry(root)
slug_entry_5 = Entry(root)
slug_entry_6 = Entry(root)
get_filenames_button = Button(root, text="Check filenames", command=PrintSlugs)
create_txt = Button(root, text="Create .txt file", command=CreateTxt)
create_txt_done = Label(root, text="Done!")
get_ep_info = Button(root, text="Get Episode Info", command=EpInfo)
#TODO add a clear button to destroy the filename labels...or just clear them anytime you call Check filename

title.grid(row=0, column=0)
show_name.grid(row=0, column=1)
slug_label_1.grid(row=1, column=0)
slug_label_2.grid(row=2, column=0)
slug_label_3.grid(row=3, column=0)
slug_label_4.grid(row=4, column=0)
slug_label_5.grid(row=5, column=0)
slug_label_6.grid(row=6, column=0)
slug_entry_1.grid(row=1, column=1)
slug_entry_2.grid(row=2, column=1)
slug_entry_3.grid(row=3, column=1)
slug_entry_4.grid(row=4, column=1)
slug_entry_5.grid(row=5, column=1)
slug_entry_6.grid(row=6, column=1)

get_filenames_button.grid(row=7, column=0)
create_txt.grid(row=7, column=1)
get_ep_info.grid(row=7, column=2)

root.mainloop()
