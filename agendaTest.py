from mailmerge import MailMerge
import datetime
import time
import os
import json

dir_path = os.path.dirname(os.path.realpath(__file__))


# main loop
def main():

    # load config file
    try:
        config = LoadFile('config.json')
    except Exception as e:
        print("config file error?" + e)
        input("Seek help!")
        exit
    config = config['DEFAULT']


    # initialise values
    vals = { 'theDate' : '0', 'theTime' : '0000', 'theDuration' : '000', 'kdp' : [] }
    usrInput = ''

    # enter main loop
    while(usrInput != 'Q'):
        print('Input meeting <date>\t\t<Load> meeting details\t\t\t<Create> doc')
        print('Input meeting <time>\t\t<Add> agenda item\t\t\t<Quit>')
        print('Input meeting <duration>\t<Edit> agenda item')
        print('<Display> meeting details\t<Sort> agenda items\n')



        usrInput = input('--> ').upper()

        if('DATE'.startswith(usrInput)):
            vals['theDate'] = dateInput()
            input('success!')

        elif('TIME'.startswith(usrInput)):
            vals['theTime'] = timeInput()
            input('success!')

        elif('DURATION'.startswith(usrInput)):  # returns the time added to vals['theTime'] in 24 hour format
            duration = inputDuration('Enter the duration of the meeting in minutes: ',vals['theTime'])
            vals['theDuration'] = addTime(vals['theTime'],duration)
            input('success!')

        elif('DISPLAY'.startswith(usrInput)): # print all fields
            print('Meeting date: ' + vals['theDate'])
            print('Meeting start time: ' + vals['theTime'])
            print('Meeting end time: ' + vals['theDuration'])
            if(len(vals['kdp']) > 0):
                DisplayItems(config,vals)

            input('...')

        elif('LOAD'.startswith(usrInput)):
            if(UserConfirmed('Overwrite current data with saved data (Y/N): ')): # confirm we want to overwrite
                try:
                    vals = LoadFile('save.json') # if a failure occurs, don't overwrite the current vals
                except Exception as e: # need more explicit exceptions?
                    print("failed: " + str(e))
                    print(vals)

        elif('ADD'.startswith(usrInput)):
            item = AddItem(config, vals['theTime'], len(vals['kdp']))
            vals['kdp'].append(item)


        elif('EDIT'.startswith(usrInput)):
            EditItem(config,vals)

        elif('SORT'.startswith(usrInput)):
            SortItems(config,vals)

        elif('CREATE'.startswith(usrInput)):
            MakeDoc(vals, config)
            input('success!')

        elif('QUIT'.startswith(usrInput)):
            usrInput = 'Q'

        else:
            print("Invalid input")


        os.system('cls')

    # save vals to json file
    saveFile = vals

    with open('save.json','w') as outfile:
        json.dump(saveFile,outfile,indent=4)

def EditItem(config,vals):
    DisplayItems(config,vals)
    configItem = config['key_decisions']
    listLen = len(vals['kdp']) + 1
    options = list(range(1,listLen))
    x = 0
    while(x == 0):

        idx = xDigitInput('Select the ID you with to edit ' + str(options) + ': ', 1, len(str(listLen)))
        idx = int(idx)
        if idx <= listLen:
            vals['kdp'][idx-1] = AddItem(config,vals['theTime'], idx-2)
            x += 1
        else:
            print("invalid option")


def SortItems(config,vals):
    DisplayItems(config,vals)
    configItem = config['key_decisions']
    x = 0
    while(x == 0):
        options = list(range(1,len(vals['kdp'])+1))
        prefOptions = []

        while (len(options) > 0):
            print(options)
            usrInput = xDigitInput("Enter the next ID: ",1,len(str(len(vals['kdp']))))
            if int(usrInput) in options:
                prefOptions.append(int(usrInput))
                options.remove(int(usrInput))
            else:
                print("invalid option")
        print(list(range(1,len(vals['kdp'])+1)))
        print(prefOptions)
        if(UserConfirmed('Sort into new order (Y/N)? ')):
            newList = []
            for n in prefOptions:
                newList.append(vals['kdp'][n-1])
            for i in range(len(newList)):
                newList[i][configItem['id']] = configItem['code'] + '.' + str(i+1)
            vals['kdp'] = newList
            x += 1

def DisplayItems(config, vals):
    configItem = config['key_decisions']

    print('|-- ID | TOPIC | PRESENTER | ACTION | TIME --|')
    print('L------I-------I-----------I--------I--------J\n')
    for n in vals['kdp']:
        print('r-----')
        print('| ' +
            n[configItem['id']] + ' + ' +
            n[configItem['topic']] + ' + ' +
            n[configItem['presenter']] + ' + ' +
            n[configItem['action']] + ' + ' +
            n['time'])
        print('L_____')

def AddItem(config, theTime, listLen):
    item = {}
    configItem = config['key_decisions']

    item[configItem['id']] = configItem['code'] + '.' + str(listLen + 1)
    item[configItem['topic']] = input('enter the name of the item: ')
    item[configItem['presenter']] = input('enter the initials of the presenter: ').upper()
    item['time'] = xDigitInput('Enter the duration of the item in minutes: ',1,3)
    item[configItem['time']] = '0'
    item[configItem['action']] = SelectAction()

    return item


def SelectAction():
    print("Select Action: \nA)pproval\nE)ndorsement\nD)iscussion")

    tooMuchMan = {'A' : 'Approval',
                'E' : 'Endorsement',
                'D' : 'Discussion'}
    x = ''
    while (True):
        x = input('--> ').upper()
        if x in tooMuchMan.keys():
            return tooMuchMan[x]
        else:
            print("Invalid option")



# function to input the date
def dateInput():
    valid = 0

    while (valid == 0):
        dayInput = xDigitInput("enter the date as <dd>: ", 2, 2)
        monInput = xDigitInput("enter the month as <mm>: ", 2, 2)
        yearInput = xDigitInput("enter the year as <yyyy>: ", 4, 4)

        date = dayInput + '/' + monInput + '/' + yearInput
        try:
            datetime.datetime.strptime(date,'%d/%m/%Y')
            valid += 1
        except ValueError:
            print("Invalid date. Try again, buddy.")

    return date

# function to input the time in 24 hour format
def timeInput():
    valid = 0

    while (valid == 0):
        usrInput = xDigitInput("enter the time in 24HR format <####>: ", 4, 4)
        if(int(usrInput) > 2400 or (int(usrInput[2:]) > 59)):
            print("Invalid entry")
        else:
            valid +=1

    if(int(usrInput) == 2400):
        usrInput = '0000'

    return usrInput

# function to convert 24 hour time to 12 hour time
def to12hour(theTime):
    t = time.strptime(theTime,'%H%M')
    time12hour = time.strftime("%I:%M %p",t)

    return time12hour

def inputDuration(query, theTime):
    valid = 0

    while (valid == 0):
        usrInput = int(xDigitInput(query,1,3))
        if (usrInput > 720):
            print("cannot be longer than 12 hours, for reasons.")
        else:
            valid += 1
    return usrInput

# function to add time (in minutes) to 24 hour time
def addTime(theTime, duration):
    hour = int(theTime[:2])
    mins = int(theTime[2:])


    hour += int(duration) // 60
    mins += int(duration) % 60

    if (mins > 59):
        hour += 1
        mins -= 60
    if (mins < 10):
        mins = '0' + str(mins)

    if (hour > 23):
        hour -= 24
    if (len(str(hour)) < 2):
        hour = '0' + str(hour)


    return str(hour) + str(mins)

# positive whole numbers within a range of factors (length)
def xDigitInput(inputString,minLen,maxLen):
    valid = 0

    while (valid == 0):
        usrInput = input(inputString)
        if(not usrInput.isdigit() or
            (len(usrInput) < minLen) or
            (len(usrInput) > maxLen)):
            print("Invalid entry")
        else:
            valid += 1

    return usrInput

# function to create a new word doc based on a template, filling in the mergeFields within the source file
def MakeDoc(vals, config):

    try:
        agendaFile = dir_path + "\\" + config['docFile']

        print(agendaFile)
        template = {}
        template[config['meeting_date']] = vals['theDate']
        template[config['meeting_time']] = to12hour(vals['theTime'])
        template[config['meeting_duration']] = to12hour(vals['theDuration'])
        kdpItems = vals['kdp']

        configItem = config['key_decisions']
        totalTime = int(config['lead_time'])
        print(totalTime)

        for n in kdpItems:
            temp = n['time']
            n[configItem['time']] = to12hour((addTime(vals['theTime'],int(totalTime)))) + ' - ' + str(n['time']) + ' mins'

            totalTime += int(temp)

        doc = MailMerge(agendaFile)

        doc.merge(**template)
        doc.merge_rows(config['key_decisions']['id'],kdpItems)

        docName = '1.0BWCC ' + vals['theDate'].replace('/','-') + '.docx'
        doc.write(docName)
    except Exception as e:
        print('Error in making doc-- ' + str(e))

def LoadFile(fileName):

    with open(fileName,'r') as f:
        vals = json.load(f)

    return vals

def UserConfirmed(query):
    x = ''
    while not (x == 'Y' or x == 'N'):
        x = input(query).upper()

    return (x == 'Y')



try:
    main()
except Exception as e:
    input("over:" + str(e))
