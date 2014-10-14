# -*- coding: utf-8 -*-
from psychopy import visual, core, data, event, logging, gui, sound
from psychopy.constants import *
import os, random, math, copy, xlwt, numpy, csv
from xlrd import open_workbook
from xlutils.copy import copy as xlcopy
import cPickle as pickle
from random import shuffle, choice
from datetime import datetime
from Math_Task import LDRH_math_obj_4_choice_w_gradeinfo as Math_Script
from Tones_Task import LDRH_tones_obj_w_gradeinfo as Tones_Script
from Dots_Task import LDRH_panamath_boxes_obj_w_gradeinfo as Dots_Script
from Reading_Task import LDRH_reading_obj_4buttons_inst_gradeinfo as Reading_Script
from Phonology_Task import LDRH_phonology_obj_w_gradeinfo as Phonology_Script
from Star_Task import LDRH_spatial_obj_w_gradeinfo as Star_Script
from Feedback import feedback

try:
    import taskversion as VERSION
except ImportError as e:
    VERSION = "no_version"

#enable pickling of data
pickle_enabled = False

#if you want to skip the first homing phase of the task
just_choice = False

#touchscreen? if False, uses conventional mouse
touchscreen = True

#which tasks to run
task_names=['Spatial','Phonology','Math','Music','Reading','Dots']

#store info about the experiment session
expName='REaCh Task'; expInfo={'participant':'','grade':'(k,1,2,3,4,or 5)'}
dlg=gui.DlgFromDict(dictionary=expInfo,title=expName)
if dlg.OK==False: core.quit() #user pressed cancel
expInfo['date']=data.getDateStr(); expInfo['expName']=expName
# Setup files for saving
if not os.path.isdir('data'):
    os.makedirs('data')  # if this fails (e.g. permissions) we will get error
grade = str(expInfo['grade'])
filename = 'data' + os.path.sep + '%s_%s' %(expInfo['participant'], expInfo['date'])
ppt = expInfo['participant']
logFile = logging.LogFile(filename+'.log', level=logging.EXP)
logging.console.setLevel(logging.WARNING)  # this outputs to the screen, not a file

# Check for pickle
if os.path.isfile('data.p') and pickle_enabled:
    f = open('data.p', 'r')
    pdata = pickle.load(f)
    if ppt in pdata.keys():
        dates = pdata[ppt].keys()
        dlg = gui.Dlg(title="Found previous data")
        if len(dates)==1:
            dlg.addText('An incomplete dataset was found for this participant from {}, would you like to resume that session?'.format(dates[0]))
        elif len(dates)>1:
            query = '{} incomplete datasets were found for this participant from the following dates:'.format(len(dates))
            for date in dates: query = query+'\n\t{}'.format(date)
            query = query+'\nWould you like to resume one of those sessions?'
            dlg.addText(query)
        dlg.addField('Resume:', choices=['Yes','No'])
        dlg.show()  # show dialog and wait for OK or Cancel
        if dlg.OK:  # then the user pressed OK
            print dlg.data
            use_pickle=dlg.data[0]=='Yes'
            if len(dates)==1:
                this_pdata = pdata.pop(ppt)[dates[0]]
                old_date = dates[0]
            elif len(dates)>1:
                datedlg = gui.Dlg(title="Choose session")
                datedlg.addText('Please choose the incomplete session for {} that you would like to resume:'.format(ppt))
                datedlg.addField('Session:', choices=dates)
                datedlg.show()
                if datedlg.OK:
                    this_pdata=pdata[ppt].pop(datedlg.data[0])
                    old_date = datedlg.data[0]
                else: core.quit()
        else: core.quit()
        f.close()
        f = open('data.p', 'w')
        pickle.dump(pdata, f)
        f.close()
        pdata = this_pdata
        old_filename = os.path.join('data', '{}_{}.xls'.format(ppt, old_date))
        if not os.path.isfile(old_filename):
            print 'could not find old data-- looked for {}'.format(old_filename)
            core.quit()
    else: pdata=None
else: pdata=None

# Initialize things regardless of pickle
win = visual.Window(size=(1100, 700), allowGUI=True, monitor=u'testMonitor', color=[-1,-1,-1], colorSpace=u'rgb', units=u'pix', fullscr=True) #Window
trialClock=core.Clock()
image_choice_path = 'Images/Choice/'
audio_path = 'Audio/'
retry_instructions = visual.TextStim(win=win, text='Touch anywhere to try again.', height=28)
choice_instructions = visual.TextStim(win=win, height=28, wrapWidth=800, text=
    "Now we are going to play all of the games together. In this next part you can choose which game you want to play by touching one of the game buttons on the screen. Each time you play, you will earn points that will fill up the colored bar at the top of the screen. Each game button will have colored rings. The more rings there are, the more points you’ll earn for playing that game. For example, a game button with four rings will give you  a lot of points. But another game that has less rings or no rings will give you less points. You can still play the game button that has no rings. You will win when the colored bar on top is fully colored! \n\n\n\n\nTouch anywhere on the screen to play.")
math_icon = visual.ImageStim(win=win, image = image_choice_path + 'math.png', units = 'pix', ori = 0, pos = [0,0], size = [120, 120], opacity = 1, mask =None, interpolate = True)
dots_icon = visual.ImageStim(win=win, image = image_choice_path + 'panamath.png', units = 'pix', ori = 0, pos = [0,0], size = [120, 120], opacity = 1, mask =None, interpolate = True)
reading_icon = visual.ImageStim(win=win, image = image_choice_path + 'reading.png', units = 'pix', ori = 0, pos = [0,0], size = [126, 120], opacity = 1, mask =None, interpolate = True)
phonology_icon = visual.ImageStim(win=win, image = image_choice_path + 'phonology2.png', units = 'pix', ori = 0, pos = [0,0], size = [120, 120], opacity = 1, mask =None, interpolate = True)
spatial_icon = visual.ImageStim(win=win, image = image_choice_path + 'stars.png', units = 'pix', ori = 0, pos = [0,0], size = [120, 120], opacity = 1, mask =None, interpolate = True)
music_icon = visual.ImageStim(win=win, image = image_choice_path + 'music.png', units = 'pix', ori = 0, pos = [0,0], size = [120, 120], opacity = 1, mask =None, interpolate = True)
selection_circle = visual.ImageStim(win=win, image = image_choice_path + 'selection_circle.png', units = 'pix', ori = 0, pos = [0,0], size = [120, 120], opacity = 1, mask =None, interpolate = True)
progress_frame = visual.Rect(win=win, units='pix',pos=[0,300],size=[1206,56],lineColor='white',fillColor=None,lineWidth=3)
progress_fill = visual.Rect(win=win, units='pix',pos=[0,300],size=[0,30], fillColor='lime', lineColor='lime')
progress_animation = visual.Rect(win=win, units='pix',pos=[0,300],size=[0,30], fillColor='lime')
math_operations = ['addition','subtraction','multiplication','division']
math_benchmarks = {'subtraction': {'addition': {'thresh': 3, 'count': 0}}, 'multiplication': {'addition': {'thresh': 5, 'count': 0}, 'subtraction': {'thresh': 3, 'count': 0}}, 'division': {'multiplication': {'thresh': 3, 'count': 0}}}

congratulations_text = visual.TextStim(win=win, text="You did it! You win!", height=38, pos = [0,200])
fireworks = visual.MovieStim(win=win, filename=audio_path + 'fireworks.mp4', loop=True, pos = [0,-100])
applause = sound.Sound(audio_path + 'applause.wav')
applause.setVolume(0.6)

score = visual.TextStim(win, units = 'pix', ori=0, font=u'Arial', pos=[0, -10], color=u'white', text='000')
cash_register = sound.Sound(value= audio_path + 'cash_register.wav')
cash_register.setVolume(0.2)
mouse=event.Mouse(win=win)
mouse.setVisible(0)
mouse.getPos()
point_intervals=12

#create rings and store in dictionary
colors_for_rings=['red','orange','light_orange','yellow']
all_rings = {'Math': {}, 'Dots': {}, 'Reading': {}, 'Phonology': {}, 'Spatial': {}, 'Music': {}}
for ring in range(len(colors_for_rings)):
    all_rings['Math'][ring] = visual.ImageStim(win=win, image = image_choice_path + '%s_ring.png' %colors_for_rings[ring], units = 'pix', ori = 0, pos = [0,0], size = [134+(ring*17), 134+(ring*17)], opacity = 1, mask =None, interpolate = True)
    all_rings['Dots'][ring] = visual.ImageStim(win=win, image = image_choice_path + '%s_ring.png' %colors_for_rings[ring], units = 'pix', ori = 0, pos = [0,0], size = [134+(ring*17), 134+(ring*17)], opacity = 1, mask =None, interpolate = True)#copy.copy(all_rings['Math'][ring])
    all_rings['Reading'][ring] = visual.ImageStim(win=win, image = image_choice_path + '%s_ring.png' %colors_for_rings[ring], units = 'pix', ori = 0, pos = [0,0], size = [134+(ring*17), 134+(ring*17)], opacity = 1, mask =None, interpolate = True)#copy.copy(all_rings['Math'][ring])
    all_rings['Phonology'][ring] = visual.ImageStim(win=win, image = image_choice_path + '%s_ring.png' %colors_for_rings[ring], units = 'pix', ori = 0, pos = [0,0], size = [134+(ring*17), 134+(ring*17)], opacity = 1, mask =None, interpolate = True)#copy.copy(all_rings['Math'][ring])
    all_rings['Spatial'][ring] = visual.ImageStim(win=win, image = image_choice_path + '%s_ring.png' %colors_for_rings[ring], units = 'pix', ori = 0, pos = [0,0], size = [134+(ring*17), 134+(ring*17)], opacity = 1, mask =None, interpolate = True)#copy.copy(all_rings['Math'][ring])
    all_rings['Music'][ring] = visual.ImageStim(win=win, image = image_choice_path + '%s_ring.png' %colors_for_rings[ring], units = 'pix', ori = 0, pos = [0,0], size = [134+(ring*17), 134+(ring*17)], opacity = 1, mask =None, interpolate = True)#copy.copy(all_rings['Math'][ring])

def can_evaluate(value):
    try:
        eval(value)
        return True
    except:
        return False

def importConditions(path):
    out = []
    with open(path, 'rU') as csvfile:
        reader = csv.reader(csvfile)
        headers = reader.next()
        for row in reader:
             out.append(dict(zip(headers,[eval(cell) if can_evaluate(cell) else cell for cell in row])))
    return out

#conditions files to read in
all_conditions = {
    'Math': {
        'addition': importConditions('Math_Task/math_stims_addition.csv'),
        'subtraction': importConditions('Math_Task/math_stims_subtraction.csv'),
        'multiplication': importConditions('Math_Task/math_stims_multiplication.csv'),
        'division': importConditions('Math_Task/math_stims_division.csv')
        },
    'Dots': importConditions('Dots_Task/dots_conds2.csv'),
    'Reading': importConditions('Reading_Task/stimulus_gradelist_newlysorted4_more_g4.csv'),
    'Phonology': importConditions('Phonology_Task/phonology_stims_new.csv'),
    'Spatial': None,
    'Music': importConditions('Tones_Task/tones_stims_new.csv')}

#initialize games
all_games = {'Math': Math_Script.Math_Game(win, all_conditions['Math']),
    'Music': Tones_Script.Tones_Game(win, all_conditions['Music']),
    'Dots': Dots_Script.Dots_Game(win, all_conditions['Dots']),
    'Reading': Reading_Script.Reading_Game(win, all_conditions['Reading']),
    'Phonology': Phonology_Script.Phonology_Game(win, all_conditions['Phonology']),
    'Spatial': Star_Script.Star_Game(win)}

#dictionary of icons
all_icons = {'Math': math_icon, 'Music': music_icon, 'Reading': reading_icon, 'Dots': dots_icon, 'Phonology': phonology_icon, 'Spatial': spatial_icon}
#[Music, Phonology, Dots, Reading, Spatial]
low_thresh = {'Music':10,'Phonology':3,'Dots':39,'Reading':5,'Spatial':150}

#load pickled data if applicable
if pdata:
    for k in pdata.keys():
        exec('{} = pdata[k]'.format(k))
    wb_read = open_workbook(old_filename)
    wb = xlcopy(wb_read)
    all_sheets = {}
    for i, sheet in enumerate(wb_read.sheet_names()):
        ws_read = wb_read.sheet_by_index(i)
        all_sheets[sheet] = dict(sheet = wb.get_sheet(i), headers=[ws_read.cell_value(0,col) for col in range(ws_read.ncols)], row=ws_read.nrows)

else:
    print 'loading things outside of pickle'
    points = 0
    thesePoints=0
    first_pass=True
    trial_number = 1

    #create output structure
    wb = xlwt.Workbook()
    all_sheets = {'Main': dict(sheet = wb.add_sheet('Main'), headers=['Trial Number', 'Game', 'Difficulty', 'Score','Type','Icon_Pos', 'Task Version'], row=1),
        'Math': dict(sheet = wb.add_sheet('Math'), headers = ['Trial Number', 'Operation', 'Difficulty','Stimulus','Target','Foil1','Foil2','Foil3','Score','Resp Time', 'Task Version'], row=1),
        'Dots': dict(sheet = wb.add_sheet('Dots'), headers = ['Trial Number', 'Difficulty','Correct','Incorrect','Ratio','Score','Resp Time', 'Task Version'], row=1),
        'Reading': dict(sheet = wb.add_sheet('Reading'), headers = ['Trial Number', 'Difficulty','Grade','Criteria','Target_2b','Foil_2b','Target_4b','Foil_4b1','Foil_4b2','Foil_4b3','Foil_4b4','Response','Score','Resp Time', 'Task Version'], row=1),
        'Phonology': dict(sheet = wb.add_sheet('Phonology'), headers = ['Trial Number', 'Difficulty','Stim1','Stim2','Response','Correct Response','Score','Resp Time','POA_steps','VOT_steps','VOT_or_POA','Difference Position','Distance','Number of Phonemes','Phoneme Difference', 'Task Version'], row=1),
        'Spatial': dict(sheet = wb.add_sheet('Spatial'), headers = ['Difficulty', 'Score', 'First_Click_Time', 'Second_Click_Time', 'Resp Time', 'Star_Pos', 'Resp_Pos', 'Resp_Distance', 'Task Version'], row=1),
        'Music': dict(sheet = wb.add_sheet('Music'), headers = ['Trial Number', 'Difficulty', 'soundA','soundB','Details','Contour','Notes Different','Root','Response','Correct Response','Score','Resp Time', 'Task Version'], row=1),
        'Task_Times': dict(sheet = wb.add_sheet('Task_Times'), headers = ['Task','Instructions','Practice','Staircase','Total', 'Task Version'],row=1)}

    #initialize headers for each sheet
    for key in all_sheets.keys():
        for col, header in enumerate(all_sheets[key]['headers']): all_sheets[key]['sheet'].write(0,col,header)

    #create handlers
    all_handlers = {
        'Math': {
                'addition': data.StairHandler(startVal= len(all_conditions['Math']['addition'])-1, stepSizes=[2,1,1,1],
                    minVal=0, maxVal=len(all_conditions['Math']['addition'])-1, nUp=1, nDown=3, nTrials=10, stepType = 'lin'),
                'subtraction': data.StairHandler(startVal= len(all_conditions['Math']['subtraction'])-1, stepSizes=[2,1,1,1],
                    minVal=0, maxVal=len(all_conditions['Math']['subtraction'])-1, nUp=1, nDown=3, nTrials=10, stepType = 'lin'),
                'multiplication': data.StairHandler(startVal= len(all_conditions['Math']['multiplication'])-1, stepSizes=[2,1,1,1],
                    minVal=0, maxVal=len(all_conditions['Math']['multiplication'])-1, nUp=1, nDown=3, nTrials=10, stepType = 'lin'),
                'division': data.StairHandler(startVal= len(all_conditions['Math']['division'])-1, stepSizes=[2,1,1,1],
                    minVal=0, maxVal=len(all_conditions['Math']['division'])-1, nUp=1, nDown=3, nTrials=10, stepType = 'lin')
                },
        'Music': data.StairHandler(startVal = 14, stepType = 'lin', stepSizes=[2,1,1,1,1,1], #reduce step size every two reversals
            minVal=0, maxVal=len(all_conditions['Music'])-1, nUp=1, nDown=3,  #will home in on the 80% threshold
            nTrials = 10),
        'Dots': data.StairHandler(startVal = 35, stepType = 'lin', stepSizes=[5,3,2,2,1,1], #reduce step size every two reversals
            minVal=0, maxVal=len(all_conditions['Dots'])-1, nUp=1, nDown=3,  #will home in on the 80% threshold
            nTrials = 10),
        'Reading': data.StairHandler(startVal = 5, stepType = 'lin', stepSizes=[1,1,1,1], #reduce step size every two reversals
            minVal=0, maxVal=len(all_conditions['Reading'])-1, nUp=1, nDown=3,  #will home in on the 80% threshold
            nTrials = 10, nReversals = 0),
        'Phonology': data.StairHandler(startVal = 4, stepType = 'lin', stepSizes=[1,1,1,1], #reduce step size every two reversals
            minVal=0, maxVal=len(all_conditions['Phonology'])-1, nUp=1, nDown=3,  #will home in on the 80% threshold
            nTrials = 10, nReversals = 0),
        'Spatial': data.StairHandler(startVal = 150,
            stepType = 'db', stepSizes=[3,3,2,2,1,1],#[8,4,4,2,2,1,1], #reduce step size every two reversals
            minVal=0, maxVal=350, nUp=1, nDown=3,  #will home in on the 80% threshold
            nTrials = 10)
        }

    #dictionry of ring tracking
    num_rings = {'Math': 4, 'Music': 4, 'Reading': 4, 'Dots': 4, 'Phonology': 4, 'Spatial': 4}

    #dictionary of threshold tracking
    all_thresholds = {}

    #create list to randomize order of presentation
    # shuffle(task_names)
    staircased = []

#method to get clicks
if touchscreen:
    def click():
        if mouse.mouseMoved(): return True
        else: return False
elif not touchscreen:
    def click():
        if mouse.getPressed()==[1,0,0]: return True
        else: return False

#save data to xls then quit
def save_and_quit(complete=False):
    wb.save('data/'+ '%s_%s' %(ppt, expInfo['date']+'.xls'))
    if complete:
        wb.save('data/complete_data/'+ '%s_%s' %(ppt, expInfo['date']+'.xls'))
    core.quit()

#pickle data then save_and_quit
def pickle_and_quit():
    if not pickle_enabled: save_and_quit()
    # load previous pickle if exists
    if os.path.isfile('data.p'):
        f = open('data.p', 'r')
        pdata = pickle.load(f)
        f.close()
        f = open('data.p', 'w')
    else:
        pdata = {}
        f = open('data.p', 'w')

    # create dictionary to pickle
    this_pdata = dict(points=points, thesePoints=thesePoints, first_pass=first_pass, trial_number=trial_number,
        all_handlers=all_handlers, num_rings=num_rings,
        all_thresholds=all_thresholds, task_names=task_names)

    # nest dictionary into outer dicts with ppt and date as keys
    if ppt in pdata.keys():
        pdata[ppt][expInfo['date']] = this_pdata
    else:
        pdata[ppt] = {expInfo['date']: this_pdata}

    # dump, save, and quit
    pickle.dump(pdata, f)
    f.close()
    print 'pickled data to {}'.format(f)
    save_and_quit()

def run_staircase(task, operation=None):
    global trial_number
    if operation: handler = all_handlers[task][operation]
    else: handler = all_handlers[task]
    pos_streak=0
    neg_streak=0

    try:
        thisIncrement = handler.next()
        print 'thisIncrement:', thisIncrement
        #run game-- output is a dictionary of values
        if operation: output = all_games[task].run_game(win, grade, operation, thisIncrement)
        else: output = all_games[task].run_game(win, grade, thisIncrement)
        if output=='QUIT': pickle_and_quit()

        #first write trial number to output, then write the output variables
        all_sheets[task]['sheet'].write(all_sheets[task]['row'], 0, trial_number)
        for col,header in enumerate(all_sheets[task]['headers'][1:]):
            if header=="Task Version":
                all_sheets[task]['sheet'].write(all_sheets[task]['row'],col+1, VERSION)
            else:
                all_sheets[task]['sheet'].write(all_sheets[task]['row'],col+1, output[header])
        #increment row on output structure
        all_sheets[task]['row'] += 1

        #write output for main sheet
        main_output = {'Trial Number':trial_number, 'Game': task, 'Difficulty': output['Difficulty'],'Score':output['Score'],'Type':'threshold','Icon_Pos':''}
        for col,header in enumerate(all_sheets['Main']['headers']):
            all_sheets['Main']['sheet'].write(trial_number, col, main_output[header])

        #update handler only if not a "same" trial from tones or phonology
        #if 'Correct Response' in output.keys() and output['Correct Response'].lower() != 'same': all_handlers[task].addData(output['Score'])
        #else: print 'same trial-- did not update stairhandler'
        handler.addData(output['Score'])

        # This code is to boost the staircase level of 'lower' operations when a student is successful on
        # more difficult operations. As a first pass, this simply records a success in *all* operations up to the one
        # that the student achieves the success in.
        if operation and output['Score']:
            operations = ['addition', 'subtraction', 'multiplication', 'division']
            for operation_name in operations[0:operations.index(operation)]:
                all_handlers[task][operation_name].addData(output['Score'])

        #increment trial number
        trial_number+=1

        #keep track of how many correct/incorrect in a row
        if output['Score'] and all_conditions[task] and output['Difficulty']==len(all_conditions[task]): output['pos_streak']=1
        else: output['pos_streak']=0
        if not output['Score'] and all_conditions[task] and output['Difficulty']==1: output['neg_streak']=1
        else: output['neg_streak']=0
        print "output['pos_streak']:", output['pos_streak']

        output['thisIncrement'] = thisIncrement

    except StopIteration:
        output={}
        output['Score'] = 'StopIteration'
        output['thisIncrement'] = handler.intensities[-1]

    return output


#STAIRCASING SECTION

# set up timers
instructions_times = {}
practice_times = {}
staircasing_times = {}
total_times = {}

if not just_choice:
    #run through instructions, practice, and staircase for each task
    for task in [task for task in task_names if task not in all_thresholds.keys()]: #only loops through tasks with no threshold yet
        #display icon for the task
        all_icons[task].draw()
        win.flip()
        start_time=trialClock.getTime()
        while trialClock.getTime()<start_time+3:
            if event.getKeys(keyList=['q', 'escape']): pickle_and_quit()

        #run instructions for task
        instructions_start = trialClock.getTime()
        if all_games[task].run_instructions(win)=='QUIT': pickle_and_quit()
        instructions_times[task] = trialClock.getTime() - instructions_start

        #run practice for task
        practice_start = trialClock.getTime()
        if hasattr(all_games[task], 'run_practice'):
            if all_games[task].run_practice(win, grade)=='QUIT': pickle_and_quit()
        practice_times[task] = trialClock.getTime() - practice_start

        #run staircase; math needs special circumstances
        staircasing_start = trialClock.getTime()
        if task=='Math':
            if 'Math' not in all_thresholds.keys(): all_thresholds['Math']={}
            streaks = {'addition': {'pos':0,'neg':0}, 'subtraction': {'pos':0,'neg':0}, 'multiplication': {'pos':0,'neg':0}, 'division': {'pos':0, 'neg':0}}
            add_count_for_mult = 0
            active_operations = ['addition']

            while active_operations:
                for operation in math_operations:
                    if operation in active_operations:
                        #one trial of staircase is run here
                        output = run_staircase(task, operation=operation)

                        for new_operation, reqs in math_benchmarks.items():
                            if operation in reqs.keys() and output['Score'] and (len(all_conditions[task][operation]) - output['thisIncrement']) >= reqs[operation]['thresh']: reqs[operation]['count']+=1
                        for new_operation, reqs in math_benchmarks.items():
                            if False not in [benchmark['count'] >=3 for req_operation, benchmark in reqs.items()]:
                                active_operations.append(new_operation)
                                math_benchmarks.pop(new_operation)
                                print '{} is now active'.format(new_operation)
                        #separate logic for OR case with multiplication
                        if operation=='addition' and output['Score'] and (len(all_conditions[task]['multiplication']) - output['thisIncrement']) >= 6: add_count_for_mult+=1
                        if 'multiplication' not in active_operations and 'multiplication' in math_benchmarks.keys() and add_count_for_mult >= 3:
                            active_operations.append('multiplication')
                            math_benchmarks.pop('multiplication')
                            print '{} is now active'.format(new_operation)

                        #handle StopIterations
                        if output['Score']=='StopIteration':
                            #record threshold and remove operation
                            all_thresholds[task][operation] = output['thisIncrement']
                            active_operations.remove(operation)
                            continue

                        #keep track of streaks
                        streaks[operation]['pos'] = streaks[operation]['pos']*output['pos_streak']+output['pos_streak'] #adds 1 or sets to 0
                        streaks[operation]['neg'] = streaks[operation]['neg']*output['neg_streak']+output['neg_streak']
                        #handle streak breaking
                        if streaks[operation]['pos'] >= 8:
                            all_thresholds[task][operation] = output['thisIncrement']
                            active_operations.remove(operation)
                        #remove operation from being active, don't record a threshold
                        if streaks[operation]['neg'] >=2:
                            active_operations.remove(operation)

                #add new operation if applicable
                for new_operation, reqs in math_benchmarks.items():
                    if new_operation in all_thresholds[task].keys(): continue

                    if False not in [req_operation in all_thresholds[task].keys() and (len(all_conditions[task][req_operation]) - all_thresholds[task][req_operation]) >= benchmark for req_operation, benchmark in reqs.items()]:
                        active_operations.append(new_operation)
                        print 'added', new_operation

        else:
            pos_streak=0
            neg_streak=0
            while True:
                #one trial of staircase is run here
                output = run_staircase(task)

                #handle StopIterations
                if output['Score']=='StopIteration':
                    #record threshold and remove operation
                    all_thresholds[task] = output['thisIncrement']
                    break

                #keep track of streaks
                pos_streak = pos_streak*output['pos_streak']+output['pos_streak'] #adds 1 or sets to 0
                neg_streak = neg_streak*output['neg_streak']+output['neg_streak']
                print 'pos_streak:', pos_streak
                #handle streak breaking
                if pos_streak >= 8:
                    all_thresholds[task] = output['thisIncrement']
                    break
                if neg_streak >=2: break
        staircasing_times[task] = trialClock.getTime() - staircasing_start
        total_times[task] = instructions_times[task]+practice_times[task]+staircasing_times[task]

    print 'instruction times', instructions_times
    print 'practice_times', practice_times
    print 'staircasing times', staircasing_times
    print 'total times', total_times
    #record task times
    for task in task_names:
        all_sheets['Task_Times']['sheet'].write(all_sheets['Task_Times']['row'],0, task)
        all_sheets['Task_Times']['sheet'].write(all_sheets['Task_Times']['row'],1, instructions_times[task])
        all_sheets['Task_Times']['sheet'].write(all_sheets['Task_Times']['row'],2, practice_times[task])
        all_sheets['Task_Times']['sheet'].write(all_sheets['Task_Times']['row'],3, staircasing_times[task])
        all_sheets['Task_Times']['sheet'].write(all_sheets['Task_Times']['row'],4, total_times[task])
        all_sheets['Task_Times']['row']+=1

elif just_choice:
    for task in task_names:
        if task=='Math':
            all_thresholds[task] = {}
            for operation in math_operations:
                all_thresholds[task][operation] = all_handlers[task][operation].startVal
        else:
            all_thresholds[task] = all_handlers[task].startVal


#CHOICE SECTION

#method to draw all icons, rings, and progress bar
def draw_main_screen():
    progress_frame.draw()
    progress_fill.draw()
    for num,task in enumerate(task_names):
        all_icons[task].draw()
        for ring in range(0, num_rings[task]): all_rings[task][ring].draw()

#present instructions for choice task
choice_start = trialClock.getTime()
mouse.getPos()
choice_instructions.draw()
win.flip()
while True:
    if click(): break
    if event.getKeys(['escape','q']): pickle_and_quit()
win.flip()

while True:
    thesePoints=0
    progress_fill.setSize([points,50])
    #starting point =-250; position should be = -250 + 1/2 pos
    progress_fill.setPos([-300+ (points/4), 300])

    xy = [('left',[-200,0]),('top-left',[-100,173]),('top-right',[100,173]),('right',[200,0]),('bottom-right',[100,-173]),('bottom-left',[-100,-173])]
    shuffle(xy)
    for num,task in enumerate(task_names):
        all_icons[task].setPos(xy[num][1])
        for ring in range(0, num_rings[task]): #will give us the number of rings we want to display
            all_rings[task][ring].setPos(xy[num][1])
    draw_main_screen()
    win.flip()

    #draw animation of progress bar if applicable
    progress_colors=['red','orangered','orange','gold','yellow']
    if not first_pass:
        start_time = trialClock.getTime()
        t = trialClock.getTime()
        while t<start_time+0.35:
            t = trialClock.getTime()
            if event.getKeys(["escape"]): pickle_and_quit()
        progress_frame.setLineColor('aqua')
        progress_animation.setFillColor(progress_colors[num_rings[this_task]+1])
        progress_animation.setSize([points,50])
        progress_animation.setPos([-300+ (points/4), 300])
        progress_animation.setOpacity(1)
        steps=points_to_add/point_intervals
        for step in range(1,steps+1):
            start_time=trialClock.getTime()
            cash_register.play()
            while t<start_time+0.35:
                t = trialClock.getTime()
                progress_animation.setSize([points+(points_to_add*step/steps),50])
                progress_animation.setPos([-300+ ((points+(points_to_add*step/steps))/4), 300])
                draw_main_screen()
                progress_frame.draw()
                progress_animation.draw()
                progress_fill.draw()
                win.flip()
                if event.getKeys(["escape"]): pickle_and_quit()


        #add points, create normal progress bar
        points+=points_to_add
        progress_frame.setLineColor('white')
        progress_fill.setSize([points,50])
        progress_fill.setPos([-300+ (points/4), 300])
        if progress_fill.size[0]>=1200: #break if progress bar is filled
            progress_frame.draw()
            progress_fill.draw()
            win.flip()
            core.wait(0.75)
            break
        draw_main_screen()
        win.flip()

    this_task = None
    mouse.getPos()
    while this_task==None:
        if click():
            if math_icon.contains(mouse): this_task='Math'
            elif dots_icon.contains(mouse): this_task='Dots'
            elif reading_icon.contains(mouse): this_task='Reading'
            elif phonology_icon.contains(mouse): this_task='Phonology'
            elif spatial_icon.contains(mouse): this_task='Spatial'
            elif music_icon.contains(mouse): this_task='Music'
        if event.getKeys(['escape','q']): pickle_and_quit()

    print 'task chosen:', this_task

    #show selection screen and wait 0.5 seconds
    #draw_main_screen()
    all_icons[this_task].draw()
    progress_frame.draw()
    progress_fill.draw()
    selection_circle.setPos(all_icons[this_task].pos)
    selection_circle.draw()
    win.flip()
    start_time = trialClock.getTime()
    while True:
        if event.getKeys(['escape','q']): pickle_and_quit()
        if trialClock.getTime() - start_time > 0.5: break

    if this_task!='Math':
        all_thresholds[this_task] = all_thresholds[this_task] if this_task in all_thresholds else low_thresh[this_task]

    #run game until get a correct answer
    score = None
    while score!=1:
        if this_task=='Math':
            print all_thresholds['Math'].keys()
            operation = choice(all_thresholds['Math'].keys())
            output = all_games[this_task].run_game(win, grade, operation, all_thresholds[this_task][operation])
        else:
            output = all_games[this_task].run_game(win, grade, all_thresholds[this_task]) #None, all_sheets[this_task]['sheet'])

        score = output['Score']
        thesePoints += score*(num_rings[this_task]+1)*point_intervals

        #first write trial number to output
        all_sheets[this_task]['sheet'].write(all_sheets[this_task]['row'], 0, trial_number)

        #next write the output variables
        for col,header in enumerate(all_sheets[this_task]['headers'][1:]):
            all_sheets[this_task]['sheet'].write(all_sheets[this_task]['row'],col+1,output[header])

        #increment row for output records
        all_sheets[this_task]['row'] += 1

        #write output for main sheet
        main_output = {'Trial Number':trial_number, 'Game': this_task, 'Difficulty': output['Difficulty'],'Score':output['Score'],'Type':'choice','Icon_Pos':[tup[0] for tup in xy if tup[1][0]==all_icons[this_task].pos[0] and tup[1][1]==all_icons[this_task].pos[1]][0]}
        for col,header in enumerate(all_sheets['Main']['headers']):
            all_sheets['Main']['sheet'].write(trial_number, col, main_output[header])

        #increment trial number
        trial_number+=1

        #display retry screen if incorrect trial
        if score==0:
            mouse.getPos()
            retry_instructions.draw()
            win.flip()
            while True:
                if click(): break
                if event.getKeys(keyList=['q', 'escape']): pickle_and_quit()

    if thesePoints!=0: points_to_add=thesePoints

    #decrease rings by 1
    if num_rings[this_task]>0: num_rings[this_task]-=1
    first_pass=False

#record choice time
all_sheets['Task_Times']['sheet'].write(all_sheets['Task_Times']['row'],0, 'Choice Section')
all_sheets['Task_Times']['sheet'].write(all_sheets['Task_Times']['row'],4, trialClock.getTime()-choice_start)
print 'choice time:', trialClock.getTime()-choice_start

#fireworks yay!!
start_time = trialClock.getTime()
applause.play()
while start_time + 20 > trialClock.getTime():
    fireworks.draw()
    congratulations_text.draw()
    win.flip()
    if event.getKeys(keyList=['q', 'escape']): save_and_quit(complete=True)
start_time = trialClock.getTime()
applause.fadeOut(5000)
while start_time + 5 > trialClock.getTime():
    fireworks.draw()
    congratulations_text.draw()
    win.flip()
    if event.getKeys(keyList=['q', 'escape']): save_and_quit(complete=True)

save_and_quit(complete=True)
