from mailmerge import MailMerge
import datetime
import time
import os
import json

# main loop
def main():

    # load config file
    try:
        config = LoadFile('config.json')
    except Exception as e:
        print("config file error?" + e)
        input("Seek help!")
        return
    config = config['DEFAULT']

    # initialise values
    dir_path = os.path.dirname(os.path.realpath(__file__))
    vals = { 'theDate' : '0', 'theTime' : '0', 'theDuration' : '0', 'kdp' : [] }
    usrInput = ''

    # enter main loop
    while(usrInput != 'e'):
        print('1) Input meeting date')
        print('2) Input meeting time')
        print('3) Input meeting duration')
        print('4) Display meeting details')
        print('5) Load meeting details')
        print('6) Add agenda item')
        print('c) Create doc')
        print('e) Exit')

        usrInput = input('--> ')

        if(usrInput == '1'):
            vals['theDate'] = dateInput()
            input('success!')

        elif(usrInput == '2'):
            vals['theTime'] = timeInput()
            input('success!')

        elif(usrInput == '3'):  # returns the time added to vals['theTime'] in 24 hour format
            vals['theDuration'] = addTime('Enter the duration of the meeting in minutes: ',vals['theTime'])
            input('success!')

        elif(usrInput == '4'): # print all fields
            print('Meeting date: ' + vals['theDate'])
            print('Meeting start time: ' + vals['theTime'])
            print('Meeting end time: ' + vals['theDuration'])
            if(len(vals['kdp']) > 0):
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
                        n[configItem['time']])
                    print('L_____')

            input('...')

        elif(usrInput == '5'):
            if(UserConfirmed()): # confirm we want to overwrite
                try:
                    vals = LoadFile('save.json') # if a failure occurs, don't overwrite the current vals
                except Exception as e: # need more explicit exceptions?
                    print("failed: " + str(e))
                    print(vals)

        elif(usrInput == '6'):
            item = AddItem(config, vals['theTime'] ,len(vals['kdp']))
            vals['kdp'].append(item)
            print(vals)

        elif(usrInput == 'c'):
            MakeDoc(vals)
            input('success!')

        elif(usrInput == 'e'):
            pass

        else:
            print("Invalid input")


        os.system('cls')


    # save vals to json file
    saveFile = vals

    with open('save.json','w') as outfile:
        json.dump(saveFile,outfile)

def AddItem(config, theTime,listLen):
    item = {}
    configItem = config['key_decisions']

    item[configItem['id']] = configItem['code'] + '.' + str(listLen + 1)
    item[configItem['topic']] = input('enter the name of the item: ')
    item[configItem['presenter']] = input('enter the initials of the presenter: ')
    item[configItem['time']] = xDigitInput('Enter the duration of the item in minutes: ',1,3)
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

# function to add time (in minutes) to 24 hour time
def addTime(query, theTime):
    valid = 0
    hour = int(theTime[:2])
    mins = int(theTime[2:])

    while (valid == 0):
        usrInput = int(xDigitInput(query,1,3))
        if (usrInput > 720):
            print("cannot be longer than 12 hours, for reasons.")
        else:
            valid += 1
    hour += usrInput // 60
    mins += usrInput % 60

    if (mins > 59):
        hour += 1
        mins -= 60
    if (mins < 10):
        mins = '0' + str(mins)

    if (hour > 23):
        hour -= 24
    if (len(str(hour)) < 4):
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
def MakeDoc(vals):

    try:
        agendaFile = dir_path + "\\" + config['docFile']

        print(agendaFile)
        template = {}
        template[config['meeting_date']] = vals['theDate']
        template[config['meeting_time']] = to12hour(vals['theTime'])
        template[config['meeting_duration']] = to12hour(vals['theDuration'])

        print(template)

        doc = MailMerge(agendaFile)

        doc.merge(**template)
        docName = '1.0BWCC ' + vals['theDate'].replace('/','-') + '.docx'
        doc.write(docName)
    except Exception as e:
        print(e)

def LoadFile(fileName):

    with open(fileName,'r') as f:
        vals = json.load(f)

    return vals

def UserConfirmed():
    x = ''
    while not (x == 'Y' or x == 'N'):
        x = input('Overwrite current data with saved data (Y/N): ').upper()

    return (x == 'Y')



try:
    main()
except Exception as e:
    input("over:" + str(e))
